"""
Backup Cloud Functions — save/load user mapping backups to Firestore.
Gated by Firebase Auth + active subscription (same as API endpoints).

Structure: backups/{uid}/versions/{timestamp_id}
  - Each version has _savedAt, _name, _pinned fields
  - Up to 2 versions can be pinned (protected from auto-delete)
  - On save, oldest unpinned versions are pruned to keep total <= MAX_VERSIONS
"""

import json
import logging
from datetime import datetime, timezone
from typing import Optional

import firebase_admin
from firebase_admin import auth as firebase_auth, firestore
from firebase_functions import https_fn

try:
    firebase_admin.get_app()
except ValueError:
    firebase_admin.initialize_app()

_db = None
MAX_VERSIONS = 5  # total cap (up to 2 pinned + 3 unpinned)


def db():
    global _db
    if _db is None:
        _db = firestore.client()
    return _db

ACTIVE_STATUSES = {"active", "trialing"}


def _cors_headers(origin: str = "*") -> dict:
    return {
        "Access-Control-Allow-Origin": origin,
        "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
        "Access-Control-Allow-Headers": "Authorization, Content-Type",
        "Access-Control-Max-Age": "3600",
    }


def _err(message: str, status: int = 400, origin: str = "*") -> https_fn.Response:
    return https_fn.Response(
        json.dumps({"ok": False, "detail": message}),
        status=status,
        headers={**_cors_headers(origin), "Content-Type": "application/json"},
    )


def _ok(data: dict, origin: str = "*") -> https_fn.Response:
    return https_fn.Response(
        json.dumps({"ok": True, **data}),
        status=200,
        headers={**_cors_headers(origin), "Content-Type": "application/json"},
    )


def _get_uid(req: https_fn.Request) -> Optional[str]:
    auth_header = req.headers.get("Authorization", "")
    if not auth_header.startswith("Bearer "):
        return None
    id_token = auth_header[len("Bearer "):]
    try:
        return firebase_auth.verify_id_token(id_token)["uid"]
    except Exception:
        return None


def _has_active_subscription(uid: str) -> bool:
    doc = db().collection("subscriptions").document(uid).get()
    if not doc.exists:
        return False
    return (doc.to_dict() or {}).get("status") in ACTIVE_STATUSES


def _guard(req: https_fn.Request):
    """Returns (uid, None) on success or (None, error_response) on failure."""
    origin = req.headers.get("Origin", "*")
    uid = _get_uid(req)
    if not uid:
        return None, _err("Sign in to your Monarch Bridge account first.", 401, origin)
    if not _has_active_subscription(uid):
        return None, _err("Active subscription required.", 402, origin)
    return uid, None


def _versions_ref(uid: str):
    return db().collection("backups").document(uid).collection("versions")


def _get_version_list(uid: str) -> list[dict]:
    """Return all versions sorted newest first."""
    docs = _versions_ref(uid).order_by("_savedAt", direction=firestore.Query.DESCENDING).stream()
    return [
        {
            "id": d.id,
            "savedAt": (d.to_dict() or {}).get("_savedAt", ""),
            "name": (d.to_dict() or {}).get("_name", "Untitled"),
            "pinned": bool((d.to_dict() or {}).get("_pinned", False)),
        }
        for d in docs
    ]


def _prune_versions(uid: str):
    """Delete oldest unpinned versions to stay within MAX_VERSIONS total."""
    all_docs = list(
        _versions_ref(uid).order_by("_savedAt", direction=firestore.Query.DESCENDING).stream()
    )
    if len(all_docs) <= MAX_VERSIONS:
        return
    # Walk from oldest to newest, delete unpinned until we're at the limit
    for doc in reversed(all_docs):
        if len(all_docs) <= MAX_VERSIONS:
            break
        data = doc.to_dict() or {}
        if not data.get("_pinned"):
            doc.reference.delete()
            all_docs.remove(doc)


@https_fn.on_request()
def save_backup(req: https_fn.Request) -> https_fn.Response:
    origin = req.headers.get("Origin", "*")
    if req.method == "OPTIONS":
        return https_fn.Response("", status=204, headers=_cors_headers(origin))

    uid, err = _guard(req)
    if err:
        return err

    try:
        body = req.get_json(silent=True) or {}
        backup = body.get("backup")
        if not backup or not isinstance(backup, dict):
            return _err("Missing or invalid backup payload.", 400, origin)

        now = datetime.now(timezone.utc)
        version_id = now.strftime("%Y%m%dT%H%M%SZ")
        backup["_savedAt"] = now.isoformat()
        backup["_name"] = body.get("name", "Untitled")
        backup["_pinned"] = False

        _versions_ref(uid).document(version_id).set(backup)
        _prune_versions(uid)

        return _ok({"versionId": version_id}, origin)
    except Exception:
        logging.exception("Backup save error")
        return _err("Something went wrong saving your backup. Please try again.", 500, origin)


@https_fn.on_request()
def load_backup(req: https_fn.Request) -> https_fn.Response:
    origin = req.headers.get("Origin", "*")
    if req.method == "OPTIONS":
        return https_fn.Response("", status=204, headers=_cors_headers(origin))

    uid, err = _guard(req)
    if err:
        return err

    try:
        body = req.get_json(silent=True) or {}
        version_id = body.get("versionId")
        versions = _versions_ref(uid)

        backup = None
        if version_id:
            doc = versions.document(version_id).get()
            backup = doc.to_dict() if doc.exists else None
        elif not body.get("listOnly"):
            docs = list(versions.order_by("_savedAt", direction=firestore.Query.DESCENDING).limit(1).stream())
            backup = docs[0].to_dict() if docs else None

        version_list = _get_version_list(uid)
        return _ok({"backup": backup, "versions": version_list}, origin)
    except Exception:
        logging.exception("Backup load error")
        return _err("Something went wrong loading your backup. Please try again.", 500, origin)


@https_fn.on_request()
def update_backup(req: https_fn.Request) -> https_fn.Response:
    origin = req.headers.get("Origin", "*")
    if req.method == "OPTIONS":
        return https_fn.Response("", status=204, headers=_cors_headers(origin))

    uid, err = _guard(req)
    if err:
        return err

    try:
        body = req.get_json(silent=True) or {}
        version_id = body.get("versionId")
        if not version_id:
            return _err("versionId required.", 400, origin)

        doc_ref = _versions_ref(uid).document(version_id)
        doc = doc_ref.get()
        if not doc.exists:
            return _err("Backup version not found.", 404, origin)

        updates = {}
        if "name" in body:
            updates["_name"] = body["name"]
        if "pinned" in body:
            # Enforce max 2 pinned
            if body["pinned"]:
                all_versions = _get_version_list(uid)
                pinned_count = sum(1 for v in all_versions if v["pinned"] and v["id"] != version_id)
                if pinned_count >= 2:
                    return _err("Maximum 2 pinned backups allowed.", 400, origin)
            updates["_pinned"] = bool(body["pinned"])

        if updates:
            doc_ref.update(updates)

        return _ok({}, origin)
    except Exception:
        logging.exception("Backup update error")
        return _err("Something went wrong updating your backup. Please try again.", 500, origin)


@https_fn.on_request()
def delete_backup(req: https_fn.Request) -> https_fn.Response:
    origin = req.headers.get("Origin", "*")
    if req.method == "OPTIONS":
        return https_fn.Response("", status=204, headers=_cors_headers(origin))

    uid, err = _guard(req)
    if err:
        return err

    try:
        body = req.get_json(silent=True) or {}
        version_id = body.get("versionId")
        if not version_id:
            return _err("versionId required.", 400, origin)

        doc_ref = _versions_ref(uid).document(version_id)
        doc = doc_ref.get()
        if not doc.exists:
            return _err("Backup version not found.", 404, origin)

        data = doc.to_dict() or {}
        if data.get("_pinned"):
            return _err("Unpin the backup before deleting.", 400, origin)

        doc_ref.delete()
        return _ok({}, origin)
    except Exception:
        logging.exception("Backup delete error")
        return _err("Something went wrong deleting your backup. Please try again.", 500, origin)
