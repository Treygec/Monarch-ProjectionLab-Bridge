"""
Stripe billing Cloud Functions.

Endpoints:
  POST /createCheckoutSession  — creates a Stripe Checkout session for a new subscriber
  POST /stripeWebhook          — receives Stripe events, updates Firestore subscriptions
  POST /createPortalSession    — creates a Stripe Customer Portal session for managing billing

Environment secrets (set via Firebase secrets, NOT plain env vars):
  STRIPE_SECRET_KEY    — Stripe secret API key (sk_live_... or sk_test_...)
  STRIPE_WEBHOOK_SECRET — Stripe webhook signing secret (whsec_...)
  STRIPE_PRICE_ID      — Stripe Price ID for the monthly subscription (price_...)
"""

import json
import os
from datetime import datetime, timezone

import stripe
import firebase_admin
from firebase_admin import auth as firebase_auth, firestore
from firebase_functions import https_fn
from firebase_functions.params import SecretParam

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

# Secrets — defined as params for Firebase secret manager
STRIPE_SECRET_KEY     = SecretParam("STRIPE_SECRET_KEY")
STRIPE_WEBHOOK_SECRET = SecretParam("STRIPE_WEBHOOK_SECRET")
STRIPE_PRICE_ID       = SecretParam("STRIPE_PRICE_ID")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _cors_headers(origin: str = "*") -> dict:
    return {
        "Access-Control-Allow-Origin": origin,
        "Access-Control-Allow-Methods": "POST, OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type, Authorization",
    }


def _ok(data: dict, origin: str = "*") -> https_fn.Response:
    return https_fn.Response(
        json.dumps(data),
        status=200,
        headers={**_cors_headers(origin), "Content-Type": "application/json"},
    )


def _err(message: str, status: int = 400, origin: str = "*") -> https_fn.Response:
    return https_fn.Response(
        json.dumps({"ok": False, "error": message}),
        status=status,
        headers={**_cors_headers(origin), "Content-Type": "application/json"},
    )


def _get_uid_from_request(req: https_fn.Request) -> str | None:
    """Extract and verify Firebase Auth UID from the Authorization: Bearer <id_token> header."""
    auth_header = req.headers.get("Authorization", "")
    if not auth_header.startswith("Bearer "):
        return None
    id_token = auth_header[len("Bearer "):]
    try:
        decoded = firebase_auth.verify_id_token(id_token)
        return decoded["uid"]
    except Exception:
        return None


# ---------------------------------------------------------------------------
# POST /createCheckoutSession
# ---------------------------------------------------------------------------

@https_fn.on_request(
    region="us-central1",
    secrets=[STRIPE_SECRET_KEY, STRIPE_PRICE_ID],
)
def create_checkout_session(req: https_fn.Request) -> https_fn.Response:
    """
    Creates a Stripe Checkout session.
    Requires: Authorization: Bearer <Firebase ID token>
    Body: { success_url: str, cancel_url: str }
    """
    origin = req.headers.get("Origin", "*")

    if req.method == "OPTIONS":
        return https_fn.Response("", status=204, headers=_cors_headers(origin))

    uid = _get_uid_from_request(req)
    if not uid:
        return _err("Sign in to your Monarch Bridge account first.", 401, origin)

    success_url = "https://app.projectionlab.com"
    cancel_url  = "https://app.projectionlab.com"

    stripe.api_key = STRIPE_SECRET_KEY.value

    # Check if this user already has a Stripe customer ID
    sub_doc = db().collection("subscriptions").document(uid).get()
    customer_id = (sub_doc.to_dict() or {}).get("stripe_customer_id") if sub_doc.exists else None

    session_params = {
        "mode": "subscription",
        "line_items": [{"price": STRIPE_PRICE_ID.value, "quantity": 1}],
        "success_url": success_url + "?checkout=success",
        "cancel_url":  cancel_url  + "?checkout=cancelled",
        "metadata": {"firebase_uid": uid},
        "subscription_data": {"metadata": {"firebase_uid": uid}},
        "allow_promotion_codes": True,
    }

    if customer_id:
        session_params["customer"] = customer_id
    else:
        # Pre-fill email from Firebase Auth
        try:
            user = firebase_auth.get_user(uid)
            if user.email:
                session_params["customer_email"] = user.email
        except Exception:
            pass

    try:
        session = stripe.checkout.Session.create(**session_params)
    except stripe.StripeError:
        import logging
        logging.exception("Stripe checkout error")
        return _err("Unable to create checkout session. Please try again.", 500, origin)

    return _ok({"ok": True, "url": session.url}, origin)


# ---------------------------------------------------------------------------
# POST /stripeWebhook
# ---------------------------------------------------------------------------

@https_fn.on_request(
    region="us-central1",
    secrets=[STRIPE_SECRET_KEY, STRIPE_WEBHOOK_SECRET],
)
def stripe_webhook(req: https_fn.Request) -> https_fn.Response:
    """
    Stripe webhook receiver.
    Verifies signature, processes subscription events, writes to Firestore.
    Configure in Stripe dashboard: https://dashboard.stripe.com/webhooks
    Events to subscribe:
      - checkout.session.completed
      - customer.subscription.updated
      - customer.subscription.deleted
      - invoice.payment_failed
    """
    stripe.api_key = STRIPE_SECRET_KEY.value
    payload = req.get_data(as_text=False)
    sig_header = req.headers.get("Stripe-Signature", "")

    try:
        event = stripe.Webhook.construct_event(
            payload, sig_header, STRIPE_WEBHOOK_SECRET.value
        )
    except stripe.errors.SignatureVerificationError:
        return https_fn.Response("Invalid signature", status=400)
    except Exception:
        import logging
        logging.exception("Stripe webhook error")
        return https_fn.Response("Webhook processing error", status=400)

    event_type = event["type"]
    data_obj   = event["data"]["object"]

    if event_type == "checkout.session.completed":
        uid             = (data_obj.get("metadata") or {}).get("firebase_uid")
        customer_id     = data_obj.get("customer")
        subscription_id = data_obj.get("subscription")
        if uid:
            db().collection("subscriptions").document(uid).set({
                "status":           "active",
                "stripe_customer_id":    customer_id,
                "stripe_subscription_id": subscription_id,
                "updated_at":       datetime.now(timezone.utc),
            }, merge=True)

    elif event_type in ("customer.subscription.updated", "customer.subscription.deleted"):
        subscription_id = data_obj.get("id")
        status          = data_obj.get("status")
        customer_id     = data_obj.get("customer")
        uid             = (data_obj.get("metadata") or {}).get("firebase_uid")

        if not uid:
            # Look up uid by stripe_customer_id
            docs = db().collection("subscriptions").where("stripe_customer_id", "==", customer_id).limit(1).stream()
            for doc in docs:
                uid = doc.id
                break

        if uid:
            db().collection("subscriptions").document(uid).set({
                "status":                status,
                "stripe_subscription_id": subscription_id,
                "updated_at":            datetime.now(timezone.utc),
            }, merge=True)

    elif event_type == "invoice.payment_failed":
        customer_id = data_obj.get("customer")
        docs = db().collection("subscriptions").where("stripe_customer_id", "==", customer_id).limit(1).stream()
        for doc in docs:
            db().collection("subscriptions").document(doc.id).set({
                "status":     "past_due",
                "updated_at": datetime.now(timezone.utc),
            }, merge=True)
            break

    return https_fn.Response(json.dumps({"ok": True}), status=200,
                             headers={"Content-Type": "application/json"})


# ---------------------------------------------------------------------------
# POST /createPortalSession
# ---------------------------------------------------------------------------

@https_fn.on_request(
    region="us-central1",
    secrets=[STRIPE_SECRET_KEY],
)
def create_portal_session(req: https_fn.Request) -> https_fn.Response:
    """
    Creates a Stripe Customer Portal session so the user can manage their subscription.
    Requires: Authorization: Bearer <Firebase ID token>
    Body: { return_url: str }
    """
    origin = req.headers.get("Origin", "*")

    if req.method == "OPTIONS":
        return https_fn.Response("", status=204, headers=_cors_headers(origin))

    uid = _get_uid_from_request(req)
    if not uid:
        return _err("Sign in to your Monarch Bridge account first.", 401, origin)

    sub_doc = db().collection("subscriptions").document(uid).get()
    if not sub_doc.exists:
        return _err("No subscription found", 404, origin)

    customer_id = (sub_doc.to_dict() or {}).get("stripe_customer_id")
    if not customer_id:
        return _err("No Stripe customer found", 404, origin)

    return_url = "https://app.projectionlab.com"

    stripe.api_key = STRIPE_SECRET_KEY.value

    try:
        session = stripe.billing_portal.Session.create(
            customer=customer_id,
            return_url=return_url,
        )
    except stripe.StripeError:
        import logging
        logging.exception("Stripe portal error")
        return _err("Unable to open billing portal. Please try again.", 500, origin)

    return _ok({"ok": True, "url": session.url}, origin)
