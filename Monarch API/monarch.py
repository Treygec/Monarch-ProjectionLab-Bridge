import asyncio
from contextvars import ContextVar
from datetime import date, timedelta
from typing import List, Optional

from fastapi import FastAPI, HTTPException
from monarchmoney import MonarchMoney
from monarchmoney.monarchmoney import RequireMFAException
from pydantic import BaseModel
from starlette.middleware.base import BaseHTTPMiddleware

app = FastAPI()
mm = MonarchMoney()

# Pending MFA state — held only between /auth/initiate and /auth/complete
_pending_email:    Optional[str] = None
_pending_password: Optional[str] = None

# Per-request context vars (populated by middleware)
_req_email:    ContextVar[Optional[str]] = ContextVar('req_email',    default=None)
_req_password: ContextVar[Optional[str]] = ContextVar('req_password', default=None)
_req_mfa:      ContextVar[Optional[str]] = ContextVar('req_mfa',      default=None)
_req_token:    ContextVar[Optional[str]] = ContextVar('req_token',    default=None)


class CredentialMiddleware(BaseHTTPMiddleware):
    """Populate per-request context vars from incoming headers."""
    async def dispatch(self, request, call_next):
        _req_email.set(request.headers.get('x-monarch-email'))
        _req_password.set(request.headers.get('x-monarch-password'))
        _req_mfa.set(request.headers.get('x-monarch-mfa'))
        _req_token.set(request.headers.get('x-monarch-token'))
        return await call_next(request)


app.add_middleware(CredentialMiddleware)


def _apply_token():
    """Set the session token on mm from the current request context."""
    token = _req_token.get()
    if not token:
        raise HTTPException(
            status_code=401,
            detail="X-Monarch-Token required. Authenticate via POST /auth/initiate first.",
        )
    mm.set_token(token)
    mm._headers["Authorization"] = f"Token {token}"


async def mm_call(func, *args, **kwargs):
    """Apply session token then call a MonarchMoney method."""
    _apply_token()
    try:
        return await func(*args, **kwargs)
    except Exception as e:
        if any(k in str(e).lower() for k in ("unauthorized", "401", "not authenticated", "not logged in")):
            raise HTTPException(
                status_code=401,
                detail="TOKEN_EXPIRED: Your Monarch session has expired. Re-authenticate.",
            )
        raise


class MFARequest(BaseModel):
    code: str


@app.get("/auth/status")
async def auth_status():
    return {"ok": True}


@app.post("/auth/initiate")
async def initiate_login():
    """
    Begin login with email + password (from X-Monarch-Email / X-Monarch-Password headers).
    Returns { requires_mfa: false, token } on success, or { requires_mfa: true } if MFA is needed.
    Optionally pass X-Monarch-MFA with a TOTP secret to auto-complete MFA without user input.
    """
    global _pending_email, _pending_password
    email    = _req_email.get()
    password = _req_password.get()
    mfa      = _req_mfa.get()
    if not email or not password:
        raise HTTPException(
            status_code=401,
            detail="Provide X-Monarch-Email and X-Monarch-Password headers.",
        )
    _pending_email = _pending_password = None
    try:
        await mm.login(
            email=email,
            password=password,
            save_session=False,
            use_saved_session=False,
            mfa_secret_key=mfa or None,
        )
        return {"ok": True, "requires_mfa": False, "token": mm.token}
    except RequireMFAException:
        _pending_email   = email
        _pending_password = password
        return {"ok": True, "requires_mfa": True}


@app.post("/auth/complete")
async def complete_mfa(body: MFARequest):
    """
    Complete MFA using the 6-digit code from the user's authenticator app.
    Must be called after /auth/initiate returned requires_mfa: true.
    Returns { ok: true, token } on success.
    """
    global _pending_email, _pending_password
    if not _pending_email:
        raise HTTPException(
            status_code=400,
            detail="No pending MFA session. Call /auth/initiate first.",
        )
    code = body.code.strip()
    if not code:
        raise HTTPException(status_code=400, detail="MFA code is required.")
    await mm.multi_factor_authenticate(
        _pending_email, _pending_password, code, trusted_device=True
    )
    token = mm.token
    _pending_email = _pending_password = None
    return {"ok": True, "token": token}


@app.get("/accounts")
async def get_accounts():
    return await mm_call(mm.get_accounts)


@app.get("/accounts/type-options")
async def get_account_type_options():
    return await mm_call(mm.get_account_type_options)


@app.get("/accounts/recent-balances")
async def get_recent_account_balances(start_date: Optional[str] = None):
    return await mm_call(mm.get_recent_account_balances, start_date=start_date)


@app.get("/accounts/snapshots-by-type")
async def get_account_snapshots_by_type(start_date: str, timeframe: str):
    return await mm_call(mm.get_account_snapshots_by_type, start_date=start_date, timeframe=timeframe)


@app.get("/accounts/aggregate-snapshots")
async def get_aggregate_snapshots(
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    account_type: Optional[str] = None,
):
    return await mm_call(
        mm.get_aggregate_snapshots,
        start_date=start_date, end_date=end_date, account_type=account_type,
    )


@app.get("/accounts/{account_id}/holdings")
async def get_account_holdings(account_id: int):
    return await mm_call(mm.get_account_holdings, account_id=account_id)


@app.get("/accounts/{account_id}/history")
async def get_account_history(account_id: int):
    return await mm_call(mm.get_account_history, account_id=account_id)


@app.get("/institutions")
async def get_institutions():
    return await mm_call(mm.get_institutions)


@app.get("/subscription")
async def get_subscription_details():
    return await mm_call(mm.get_subscription_details)


@app.get("/budgets")
async def get_budgets(
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    use_legacy_goals: Optional[bool] = False,
    use_v2_goals: Optional[bool] = True,
):
    return await mm_call(
        mm.get_budgets,
        start_date=start_date, end_date=end_date,
        use_legacy_goals=use_legacy_goals, use_v2_goals=use_v2_goals,
    )


@app.get("/transactions")
async def get_transactions(
    limit: int = 100,
    offset: Optional[int] = 0,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    search: str = "",
    category_ids: Optional[List[str]] = None,
    account_ids: Optional[List[str]] = None,
    tag_ids: Optional[List[str]] = None,
    has_attachments: Optional[bool] = None,
    has_notes: Optional[bool] = None,
    hidden_from_reports: Optional[bool] = None,
    is_split: Optional[bool] = None,
    is_recurring: Optional[bool] = None,
    imported_from_mint: Optional[bool] = None,
    synced_from_institution: Optional[bool] = None,
):
    return await mm_call(
        mm.get_transactions,
        limit=limit, offset=offset,
        start_date=start_date, end_date=end_date,
        search=search,
        category_ids=category_ids or [],
        account_ids=account_ids or [],
        tag_ids=tag_ids or [],
        has_attachments=has_attachments,
        has_notes=has_notes,
        hidden_from_reports=hidden_from_reports,
        is_split=is_split,
        is_recurring=is_recurring,
        imported_from_mint=imported_from_mint,
        synced_from_institution=synced_from_institution,
    )


@app.get("/transactions/summary")
async def get_transactions_summary():
    return await mm_call(mm.get_transactions_summary)


@app.get("/transactions/{transaction_id}")
async def get_transaction_details(transaction_id: str, redirect_posted: bool = True):
    return await mm_call(
        mm.get_transaction_details,
        transaction_id=transaction_id, redirect_posted=redirect_posted,
    )


@app.get("/transactions/{transaction_id}/splits")
async def get_transaction_splits(transaction_id: str):
    return await mm_call(mm.get_transaction_splits, transaction_id=transaction_id)


@app.get("/categories")
async def get_categories():
    return await mm_call(mm.get_transaction_categories)


@app.get("/expense-budgets")
async def get_expense_budgets(
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
):
    today = date.today()
    if not end_date:
        end_date = today.strftime("%Y-%m-%d")
    if not start_date:
        start_date = (today - timedelta(days=365)).strftime("%Y-%m-%d")

    _apply_token()

    try:
        categories_data, budgets_data = await asyncio.gather(
            mm.get_transaction_categories(),
            mm.get_budgets(start_date=start_date, end_date=end_date),
        )
    except Exception as e:
        if any(k in str(e).lower() for k in ("unauthorized", "401", "not authenticated", "not logged in")):
            raise HTTPException(status_code=401, detail="TOKEN_EXPIRED: Your Monarch session has expired. Re-authenticate.")
        raise

    expense_categories = {
        c["id"]: {"name": c["name"], "groupName": c.get("group", {}).get("name", "")}
        for c in categories_data.get("categories", [])
        if c.get("group", {}).get("type", "").lower() == "expense"
    }

    result = []
    for item in budgets_data.get("budgetData", {}).get("monthlyAmountsByCategory", []):
        category_id = item.get("category", {}).get("id")
        if category_id in expense_categories:
            result.append({
                "categoryId": category_id,
                "categoryName": expense_categories[category_id]["name"],
                "groupName": expense_categories[category_id]["groupName"],
                "monthlyAmounts": item.get("monthlyAmounts", []),
            })

    return result


@app.get("/category-groups")
async def get_transaction_category_groups():
    return await mm_call(mm.get_transaction_category_groups)


@app.get("/tags")
async def get_transaction_tags():
    return await mm_call(mm.get_transaction_tags)


@app.get("/cashflow")
async def get_cashflow(
    limit: int = 100,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
):
    return await mm_call(mm.get_cashflow, limit=limit, start_date=start_date, end_date=end_date)


@app.get("/cashflow/summary")
async def get_cashflow_summary(
    limit: int = 100,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
):
    return await mm_call(mm.get_cashflow_summary, limit=limit, start_date=start_date, end_date=end_date)


@app.get("/recurring-transactions")
async def get_recurring_transactions(
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
):
    return await mm_call(mm.get_recurring_transactions, start_date=start_date, end_date=end_date)


@app.get("/accounts/refresh-status")
async def is_accounts_refresh_complete(account_ids: Optional[List[str]] = None):
    return await mm_call(mm.is_accounts_refresh_complete, account_ids=account_ids)
