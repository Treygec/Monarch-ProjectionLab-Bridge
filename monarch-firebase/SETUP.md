# Firebase Backend — Setup Guide

This file walks through every browser/console step needed before deploying.
All code is already written — you just need to wire in project IDs and API keys.

---

## Step 1 — Firebase CLI login

```bash
firebase login
```

Opens a browser window. Sign in with the Google account you want to own the projects.

---

## Step 2 — Create Firebase projects

Create **two** projects (one for dev/testing, one for production):

1. Go to [https://console.firebase.google.com](https://console.firebase.google.com)
2. Click **Add project**
3. Name: `monarch-bridge-dev` → Continue → Disable Google Analytics (optional) → Create
4. Repeat for `monarch-bridge-prod`

Note the **Project ID** for each (shown in project settings). Update `.firebaserc`:

```json
{
  "projects": {
    "dev":  "YOUR-DEV-PROJECT-ID",
    "prod": "YOUR-PROD-PROJECT-ID"
  }
}
```

---

## Step 3 — Enable Firebase services (do this in BOTH projects)

In each project console:

| Service | Where |
|---|---|
| **Authentication → Email/Password** | Build → Authentication → Sign-in method → Email/Password → Enable |
| **Firestore Database** | Build → Firestore Database → Create database → Production mode → us-central1 |
| **Cloud Functions** | Build → Functions → Upgrade to Blaze (pay-as-you-go) plan required |
| ~~**Hosting**~~ | Not needed — extension calls Cloud Functions URLs directly |

---

## Step 4 — Create a Stripe account

1. Go to [https://dashboard.stripe.com](https://dashboard.stripe.com) → Create account
2. Add a product: **Products → Add product**
   - Name: `Monarch Bridge Pro`
   - Pricing: Recurring, Monthly, set your price (e.g. $4.99/mo)
   - Save → note the **Price ID** (`price_...`)
3. Get your API keys: **Developers → API keys**
   - Note the **Secret key** (`sk_live_...` or `sk_test_...` for testing)
4. Create a webhook: **Developers → Webhooks → Add endpoint**
   - URL: `https://us-central1-YOUR-PROD-PROJECT-ID.cloudfunctions.net/stripeWebhook`
   - Events to listen for:
     - `checkout.session.completed`
     - `customer.subscription.updated`
     - `customer.subscription.deleted`
     - `invoice.payment_failed`
   - Save → note the **Signing secret** (`whsec_...`)

---

## Step 5 — Set Firebase secrets

Run these in the `monarch-firebase/` directory for the dev project:

```bash
firebase use dev
firebase functions:secrets:set STRIPE_SECRET_KEY
# paste: sk_test_...

firebase functions:secrets:set STRIPE_WEBHOOK_SECRET
# paste: whsec_...

firebase functions:secrets:set STRIPE_PRICE_ID
# paste: price_...
```

Repeat for prod (`firebase use prod`) with live Stripe keys.

---

## Step 6 — Deploy

```bash
cd monarch-firebase/

# Deploy to dev first
firebase use dev
firebase deploy --only functions,firestore

# Once tested, deploy to prod
firebase use prod
firebase deploy --only functions,firestore
```

---

## Step 7 — Wire the extension

After deploying, get your Functions base URL:

```
https://us-central1-YOUR-PROJECT-ID.cloudfunctions.net
```

This goes into the extension's Setup tab → Monarch API URL field when using the
hosted (paid) mode. The extension will detect it's not a localhost URL and use
Firebase Auth headers automatically (Phase 3).

---

## Firestore TTL index

The `mfa_sessions` collection uses a TTL on the `expireAt` field (auto-deletes after
10 minutes). This is configured in `firestore.indexes.json` and deployed automatically.
It may take up to 24 hours to become active on a new project — this is normal.
