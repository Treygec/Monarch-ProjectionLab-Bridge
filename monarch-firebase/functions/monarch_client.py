"""
Direct Monarch Money client — no monarchmoneycommunity library dependency.
Uses httpx for async HTTP. Credentials (token) are passed per-call; nothing is stored here.
Mirrors the auth approach from the monarchmoney-enhanced Python library.
"""

import certifi
import httpx
import uuid

MONARCH_API_URL = "https://api.monarch.com/graphql"
MONARCH_AUTH_URL = "https://api.monarch.com/auth/login/"  # httpx follow_redirects handles 301 from /auth/login

_DEFAULT_HEADERS = {
    "Content-Type": "application/json",
    "Accept": "application/json",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Client-Platform": "web",
    "Origin": "https://app.monarch.com",
}

# Use certifi's CA bundle for SSL verification (required in Cloud Run environments)
_SSL_VERIFY = certifi.where()

# ---------------------------------------------------------------------------
# Low-level GraphQL helper
# ---------------------------------------------------------------------------

async def _gql(query: str, variables: dict, token: str) -> dict:
    """Execute a single GraphQL query against the Monarch API."""
    headers = {**_DEFAULT_HEADERS, "Authorization": f"Token {token}"}
    async with httpx.AsyncClient(timeout=30, verify=_SSL_VERIFY) as client:
        resp = await client.post(
            MONARCH_API_URL,
            json={"query": query, "variables": variables},
            headers=headers,
        )
    if resp.status_code == 401:
        raise PermissionError("TOKEN_EXPIRED")
    resp.raise_for_status()
    body = resp.json()
    if "errors" in body:
        msgs = [e.get("message", "") for e in body["errors"]]
        if any("not authenticated" in m.lower() or "not logged in" in m.lower() for m in msgs):
            raise PermissionError("TOKEN_EXPIRED")
        raise RuntimeError(f"GraphQL errors: {msgs}")
    return body.get("data", {})


# ---------------------------------------------------------------------------
# Auth — login / MFA
# Tries REST /auth/login/ first, falls back to GraphQL mutation (like
# monarchmoney-enhanced library). REST may 404 from some server environments.
# ---------------------------------------------------------------------------

LOGIN_MUTATION = """
mutation LoginMutation($email: String!, $password: String!, $totpToken: String) {
  login(email: $email, password: $password, totpToken: $totpToken) {
    token
    user { id email }
    errors { field messages }
  }
}
"""


async def _rest_login(email: str, password: str, totp: str = None) -> dict:
    """REST login attempt. Returns parsed JSON or raises."""
    headers = {**_DEFAULT_HEADERS, "device-uuid": str(uuid.uuid4())}
    payload = {
        "username": email,
        "password": password,
        "supports_mfa": True,
        "trusted_device": False,
    }
    if totp:
        payload["totp"] = totp
    async with httpx.AsyncClient(timeout=30, verify=_SSL_VERIFY, follow_redirects=True) as client:
        resp = await client.post(MONARCH_AUTH_URL, json=payload, headers=headers)
    return resp


async def _gql_login(email: str, password: str, totp: str = None) -> dict:
    """GraphQL login attempt (fallback)."""
    headers = {**_DEFAULT_HEADERS, "device-uuid": str(uuid.uuid4())}
    payload = {
        "query": LOGIN_MUTATION,
        "variables": {"email": email, "password": password, "totpToken": totp},
    }
    async with httpx.AsyncClient(timeout=30, verify=_SSL_VERIFY) as client:
        resp = await client.post(MONARCH_API_URL, json=payload, headers=headers)
    return resp


async def login(email: str, password: str) -> dict:
    """
    Attempt Monarch login via GraphQL mutation.
    REST /auth/login/ is blocked from GCP by Cloudflare, so we go straight to GraphQL.
    Returns {'token': str} on success.
    Returns {'mfa_required': True, 'email': str, 'password': str} if MFA is needed.
    Raises ValueError on bad credentials.
    """
    resp = await _gql_login(email, password)
    result = _parse_gql_login_response(resp)
    if result.get("mfa_required"):
        result["email"] = email
        result["password"] = password
    return result


async def complete_mfa(email: str, password: str, code: str) -> str:
    """
    Complete MFA via GraphQL mutation with totpToken.
    Returns the long-lived session token.
    """
    resp = await _gql_login(email, password, totp=code)
    result = _parse_gql_login_response(resp)
    token = result.get("token")
    if not token:
        raise ValueError("MFA failed: no token returned")
    return token


def _parse_gql_login_response(resp) -> dict:
    """Parse GraphQL login mutation response."""
    if resp.status_code != 200:
        # Include response body for debugging
        try:
            detail = resp.text[:300]
        except Exception:
            detail = ""
        raise ValueError(f"Login failed (HTTP {resp.status_code}): {detail}")
    body = resp.json()

    if "errors" in body:
        msgs = [e.get("message", "") for e in body["errors"]]
        raise ValueError(f"Login failed: {'; '.join(msgs)}")

    data = body.get("data", {}).get("login", {})
    if not data:
        raise ValueError("Login failed: unexpected response")

    errors = data.get("errors") or []
    for e in errors:
        messages = e.get("messages", [])
        field = e.get("field", "")
        if "multi_factor" in field.lower() or any("mfa" in m.lower() or "multi" in m.lower() for m in messages):
            return {"mfa_required": True}
        raise ValueError("; ".join(messages) if messages else "Login failed")

    token = data.get("token")
    if not token:
        raise ValueError("Login failed: no token returned")
    return {"token": token}


# ---------------------------------------------------------------------------
# Accounts
# ---------------------------------------------------------------------------

ACCOUNTS_QUERY = """
query GetAccounts {
  accountTypeSummaries {
    type { name display }
    accounts {
      id name displayName type { name display }
      subtype { name display }
      currentBalance displayBalance signedBalance isHidden isAsset
      institution { name url logo }
      updatedAt
    }
  }
}
"""


async def get_accounts(token: str) -> dict:
    """Fetch accounts and flatten from grouped accountTypeSummaries into a flat list."""
    data = await _gql(ACCOUNTS_QUERY, {}, token)
    flat = []
    for group in data.get("accountTypeSummaries", []):
        for acct in group.get("accounts", []):
            flat.append(acct)
    return {"accounts": flat}


# ---------------------------------------------------------------------------
# Budgets / expense categories
# ---------------------------------------------------------------------------

BUDGETS_QUERY = """
query GetJointPlanningData(
  $startDate: Date!
  $endDate: Date!
) {
  budgetData(
    startMonth: $startDate
    endMonth: $endDate
  ) {
    monthlyAmountsByCategory {
      category { id name }
      monthlyAmounts { month plannedCashFlowAmount actualAmount }
    }
    totalsByMonth {
      month
      totalIncome { plannedAmount actualAmount }
      totalExpenses { plannedAmount actualAmount }
    }
  }
}
"""

CATEGORIES_QUERY = """
query GetCategories {
  categories {
    id name order isSystemCategory isDisabled
    group { id name type }
  }
}
"""


async def get_transaction_categories(token: str) -> dict:
    return await _gql(CATEGORIES_QUERY, {}, token)


async def get_budgets(token: str, start_date: str, end_date: str) -> dict:
    return await _gql(
        BUDGETS_QUERY,
        {
            "startDate": start_date,
            "endDate": end_date,
        },
        token,
    )


# ---------------------------------------------------------------------------
# Transactions
# ---------------------------------------------------------------------------

TRANSACTIONS_QUERY = """
query GetTransactionsList(
  $offset: Int
  $limit: Int
  $startDate: String
  $endDate: String
  $search: String
  $categoryIds: [ID!]
  $accountIds: [ID!]
  $tagIds: [ID!]
  $hasAttachments: Boolean
  $hasNotes: Boolean
  $hiddenFromReports: Boolean
  $isSplit: Boolean
  $isRecurring: Boolean
  $importedFromMint: Boolean
  $syncedFromInstitution: Boolean
) {
  allTransactions(
    filters: {
      offset: $offset
      limit: $limit
      startDate: $startDate
      endDate: $endDate
      search: $search
      categoryIds: $categoryIds
      accountIds: $accountIds
      tagIds: $tagIds
      hasAttachments: $hasAttachments
      hasNotes: $hasNotes
      hiddenFromReports: $hiddenFromReports
      isSplit: $isSplit
      isRecurring: $isRecurring
      importedFromMint: $importedFromMint
      syncedFromInstitution: $syncedFromInstitution
    }
  ) {
    totalCount
    results {
      id date amount signedAmount merchant { name }
      category { id name }
      account { id displayName }
      isRecurring isPending hasSplitTransactions isSplitTransaction
      notes tags { id name }
    }
  }
}
"""


async def get_transactions(token: str, **filters) -> dict:
    return await _gql(TRANSACTIONS_QUERY, filters, token)


# ---------------------------------------------------------------------------
# Cash flow
# ---------------------------------------------------------------------------

CASHFLOW_QUERY = """
query GetCashFlow($startDate: String, $endDate: String, $limit: Int) {
  cashFlow(filters: {startDate: $startDate, endDate: $endDate, limit: $limit}) {
    summary {
      sumIncome sumExpense savings savingsRate
      month { month year }
    }
    byCategory {
      category { id name group { id name type } }
      totalAmount transactionsCount
    }
  }
}
"""


async def get_cashflow(token: str, start_date: str = None, end_date: str = None, limit: int = 100) -> dict:
    return await _gql(CASHFLOW_QUERY, {"startDate": start_date, "endDate": end_date, "limit": limit}, token)


# ---------------------------------------------------------------------------
# Tags
# ---------------------------------------------------------------------------

TAGS_QUERY = """
query GetTransactionTags {
  householdTransactionTags { id name order }
}
"""


async def get_transaction_tags(token: str) -> dict:
    return await _gql(TAGS_QUERY, {}, token)


# ---------------------------------------------------------------------------
# Recurring transactions
# ---------------------------------------------------------------------------

RECURRING_QUERY = """
query GetRecurringTransactionItems($startDate: String, $endDate: String) {
  recurringTransactionItems(startDate: $startDate, endDate: $endDate) {
    stream {
      id name frequency amount category { id name }
      merchant { id name }
    }
    date amount isPast
  }
}
"""


async def get_recurring_transactions(token: str, start_date: str = None, end_date: str = None) -> dict:
    return await _gql(RECURRING_QUERY, {"startDate": start_date, "endDate": end_date}, token)
