"""
Monarch API proxy Cloud Functions — gated by Firebase Auth + active subscription.

Every endpoint:
  1. Requires Authorization: Bearer <Firebase ID token>
  2. Checks Firestore subscriptions/{uid}.status == 'active'
  3. Requires X-Monarch-Token header (the long-lived session token stored in the extension)
  4. Proxies the request to the Monarch GraphQL API via monarch_client

This mirrors the local monarch.py endpoints so the extension can swap base URLs
between self-hosted (localhost:47821) and Firebase (functions URL) with no other changes.
"""

import asyncio
import json
from datetime import date, timedelta
from typing import Optional

import firebase_admin
from firebase_admin import auth as firebase_auth, firestore
from firebase_functions import https_fn

import monarch_client

try:
    firebase_admin.get_app()
except ValueError:
    firebase_admin.initialize_app()

_db = None

def db():
    global _db
    if _db is None:
        _db = firestore.client()
    return _db

ACTIVE_STATUSES = {"active", "trialing"}


# ---------------------------------------------------------------------------
# Auth / subscription guard
# ---------------------------------------------------------------------------

def _get_uid(req: https_fn.Request) -> str | None:
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


def _guard(req: https_fn.Request) -> tuple[Optional[str], Optional[https_fn.Response]]:
    """
    Returns (monarch_token, None) if the request is authorized.
    Returns (None, error_response) if not.
    """
    origin = req.headers.get("Origin", "*")

    uid = _get_uid(req)
    if not uid:
        return None, _err("Sign in to your Monarch Bridge account first (use the Account button in the header).", 401, origin)

    if not _has_active_subscription(uid):
        return None, _err("Active subscription required. Use the Account button in the header to manage billing.", 402, origin)

    monarch_token = req.headers.get("X-Monarch-Token", "").strip()
    if not monarch_token:
        return None, _err("X-Monarch-Token header required", 401, origin)

    return monarch_token, None


# ---------------------------------------------------------------------------
# Response helpers
# ---------------------------------------------------------------------------

def _cors_headers(origin: str = "*") -> dict:
    return {
        "Access-Control-Allow-Origin": origin,
        "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type, Authorization, X-Monarch-Token",
    }


def _ok(data, origin: str = "*") -> https_fn.Response:
    return https_fn.Response(
        json.dumps(data),
        status=200,
        headers={**_cors_headers(origin), "Content-Type": "application/json"},
    )


def _err(message: str, status: int = 400, origin: str = "*") -> https_fn.Response:
    return https_fn.Response(
        json.dumps({"ok": False, "detail": message}),
        status=status,
        headers={**_cors_headers(origin), "Content-Type": "application/json"},
    )


def _handle_monarch_error(e: Exception, origin: str) -> https_fn.Response:
    if isinstance(e, PermissionError) and "TOKEN_EXPIRED" in str(e):
        return _err("TOKEN_EXPIRED: Your Monarch session has expired. Re-authenticate.", 401, origin)
    import logging
    logging.exception("Monarch API error")
    return _err("Something went wrong fetching data from Monarch. Please try again.", 500, origin)


def _options_response(origin: str) -> https_fn.Response:
    return https_fn.Response("", status=204, headers=_cors_headers(origin))


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------

@https_fn.on_request(region="us-central1")
def get_accounts(req: https_fn.Request) -> https_fn.Response:
    origin = req.headers.get("Origin", "*")
    if req.method == "OPTIONS":
        return _options_response(origin)
    token, err = _guard(req)
    if err:
        return err
    try:
        return _ok(asyncio.run(monarch_client.get_accounts(token)), origin)
    except Exception as e:
        return _handle_monarch_error(e, origin)


@https_fn.on_request(region="us-central1")
def get_budgets(req: https_fn.Request) -> https_fn.Response:
    origin = req.headers.get("Origin", "*")
    if req.method == "OPTIONS":
        return _options_response(origin)
    token, err = _guard(req)
    if err:
        return err
    today = date.today()
    start_date = req.args.get("start_date", (today - timedelta(days=365)).strftime("%Y-%m-%d"))
    end_date   = req.args.get("end_date",   today.strftime("%Y-%m-%d"))
    try:
        return _ok(asyncio.run(monarch_client.get_budgets(token, start_date, end_date)), origin)
    except Exception as e:
        return _handle_monarch_error(e, origin)


@https_fn.on_request(region="us-central1")
def get_transactions(req: https_fn.Request) -> https_fn.Response:
    origin = req.headers.get("Origin", "*")
    if req.method == "OPTIONS":
        return _options_response(origin)
    token, err = _guard(req)
    if err:
        return err
    filters = {
        "limit":  int(req.args.get("limit", 100)),
        "offset": int(req.args.get("offset", 0)),
        "startDate": req.args.get("start_date"),
        "endDate":   req.args.get("end_date"),
        "search":    req.args.get("search", ""),
    }
    try:
        return _ok(asyncio.run(monarch_client.get_transactions(token, **filters)), origin)
    except Exception as e:
        return _handle_monarch_error(e, origin)


@https_fn.on_request(region="us-central1")
def get_cashflow(req: https_fn.Request) -> https_fn.Response:
    origin = req.headers.get("Origin", "*")
    if req.method == "OPTIONS":
        return _options_response(origin)
    token, err = _guard(req)
    if err:
        return err
    try:
        return _ok(
            asyncio.run(monarch_client.get_cashflow(
                token,
                start_date=req.args.get("start_date"),
                end_date=req.args.get("end_date"),
                limit=int(req.args.get("limit", 100)),
            )),
            origin,
        )
    except Exception as e:
        return _handle_monarch_error(e, origin)


@https_fn.on_request(region="us-central1")
def get_cashflow_summary(req: https_fn.Request) -> https_fn.Response:
    """Alias — returns same data as get_cashflow (summary field)."""
    return get_cashflow(req)


@https_fn.on_request(region="us-central1")
def get_categories(req: https_fn.Request) -> https_fn.Response:
    origin = req.headers.get("Origin", "*")
    if req.method == "OPTIONS":
        return _options_response(origin)
    token, err = _guard(req)
    if err:
        return err
    try:
        return _ok(asyncio.run(monarch_client.get_transaction_categories(token)), origin)
    except Exception as e:
        return _handle_monarch_error(e, origin)


@https_fn.on_request(region="us-central1")
def get_tags(req: https_fn.Request) -> https_fn.Response:
    origin = req.headers.get("Origin", "*")
    if req.method == "OPTIONS":
        return _options_response(origin)
    token, err = _guard(req)
    if err:
        return err
    try:
        return _ok(asyncio.run(monarch_client.get_transaction_tags(token)), origin)
    except Exception as e:
        return _handle_monarch_error(e, origin)


@https_fn.on_request(region="us-central1")
def get_recurring_transactions(req: https_fn.Request) -> https_fn.Response:
    origin = req.headers.get("Origin", "*")
    if req.method == "OPTIONS":
        return _options_response(origin)
    token, err = _guard(req)
    if err:
        return err
    try:
        return _ok(
            asyncio.run(monarch_client.get_recurring_transactions(
                token,
                start_date=req.args.get("start_date"),
                end_date=req.args.get("end_date"),
            )),
            origin,
        )
    except Exception as e:
        return _handle_monarch_error(e, origin)


# ---------------------------------------------------------------------------
# /expense-budgets — combined categories + budgets (mirrors monarch.py)
# ---------------------------------------------------------------------------

@https_fn.on_request(region="us-central1")
def get_expense_budgets(req: https_fn.Request) -> https_fn.Response:
    origin = req.headers.get("Origin", "*")
    if req.method == "OPTIONS":
        return _options_response(origin)
    token, err = _guard(req)
    if err:
        return err

    today = date.today()
    start_date = req.args.get("start_date", (today - timedelta(days=365)).strftime("%Y-%m-%d"))
    end_date   = req.args.get("end_date",   today.strftime("%Y-%m-%d"))

    async def _fetch():
        return await asyncio.gather(
            monarch_client.get_transaction_categories(token),
            monarch_client.get_budgets(token, start_date, end_date),
        )

    try:
        categories_data, budgets_data = asyncio.run(_fetch())
    except Exception as e:
        return _handle_monarch_error(e, origin)

    expense_categories = {
        c["id"]: {"name": c["name"], "groupName": (c.get("group") or {}).get("name", "")}
        for c in categories_data.get("categories", [])
        if (c.get("group") or {}).get("type", "").lower() == "expense"
    }

    result = []
    for item in budgets_data.get("budgetData", {}).get("monthlyAmountsByCategory", []):
        category_id = (item.get("category") or {}).get("id")
        if category_id in expense_categories:
            result.append({
                "categoryId":    category_id,
                "categoryName":  expense_categories[category_id]["name"],
                "groupName":     expense_categories[category_id]["groupName"],
                "monthlyAmounts": item.get("monthlyAmounts", []),
            })

    return _ok(result, origin)
