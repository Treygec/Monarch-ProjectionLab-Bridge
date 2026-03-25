"""
Monarch Bridge — Firebase Cloud Functions entry point.
Imports and re-exports all function groups.
"""

# Stripe functions (checkout, webhook, customer portal)
from billing import create_checkout_session, stripe_webhook, create_portal_session  # noqa: F401

# Monarch API proxy functions (gated by Firebase Auth + active subscription)
from api import (  # noqa: F401
    get_accounts,
    get_budgets,
    get_expense_budgets,
    get_transactions,
    get_cashflow,
    get_categories,
    get_tags,
    get_recurring_transactions,
    get_cashflow_summary,
)

# User backup functions (gated by Firebase Auth + active subscription)
from backup import save_backup, load_backup, update_backup, delete_backup  # noqa: F401
