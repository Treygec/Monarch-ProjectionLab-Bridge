# Privacy Policy — Monarch → ProjectionLab Bridge

**Last updated:** March 24, 2026

## Overview

Monarch → ProjectionLab Bridge ("the Extension") syncs financial data from Monarch Money into ProjectionLab. This privacy policy explains what data is collected, how it is used, and how it is stored.

## Data We Collect

### Authentication Information
- **Monarch Money credentials** (email, password): Used to authenticate with Monarch Money's API. In self-hosted mode, credentials are stored locally in your browser and sent only to your local proxy. In hosted mode, credentials are sent directly from your browser to Monarch's API (`api.monarch.com`) — they are never sent to or stored on our servers.
- **Monarch Bridge account credentials** (email, password): Used to authenticate with the hosted service via Firebase Authentication. Passwords are handled entirely by Google Firebase and are never accessible to us.
- **MFA/TOTP secrets**: If configured, stored locally in your browser's extension storage. TOTP codes are generated entirely within your browser using the Web Crypto API — your TOTP secret is never sent to our servers.
- **Session tokens**: Monarch session tokens are stored locally in your browser to maintain authenticated sessions. In hosted mode, your Monarch password is cleared from local storage after a session token is successfully obtained.

### Financial Information
- **Account balances**: Retrieved from Monarch Money (account names, types, balances) and pushed to ProjectionLab via its Plugin API.
- **Budget data**: Budget category names, amounts, and groupings retrieved from Monarch Money.
- **ProjectionLab data**: Account and expense data exported from ProjectionLab via its Plugin API to facilitate mapping.

### User Configuration
- **Mappings and settings**: Your account mappings, merge groups, budget mappings, and extension settings are stored locally in your browser's extension storage. Hosted users may optionally back up this configuration to our cloud service.

## How Data Is Used

All data is used exclusively for the extension's single purpose: syncing financial data from Monarch Money to ProjectionLab.

- Financial data is retrieved from Monarch and written to ProjectionLab — it is not used for any other purpose.
- No data is sold, rented, or shared with third parties.
- No data is used for advertising, analytics, tracking, or profiling.
- No data is used to determine creditworthiness or for lending purposes.

## Data Storage

### Local Storage (All Users)
All user data — credentials, mappings, settings, and cached account data — is stored in `chrome.storage.local`, which is sandboxed to the extension on your device.

### Cloud Storage (Hosted Mode Only)
- **Firebase Authentication**: Email and hashed password managed by Google Firebase.
- **Firestore**: Subscription status and optional configuration backups stored in Google Cloud Firestore, associated with your Firebase user ID.
- **Stripe**: Payment and billing information is handled entirely by Stripe. We store only your Stripe customer ID and subscription status.

### Self-Hosted Mode
In self-hosted mode, your Monarch credentials are sent from the extension to your local proxy via HTTP request headers. The proxy then authenticates with Monarch Money's API on your behalf. No data is sent to our servers — all communication stays between your machine and Monarch.

## Data Retention

- Local extension data persists until you uninstall the extension or clear extension data.
- Cloud backups are retained until you delete them (maximum 5 versions per user).
- If you cancel your hosted subscription, your cloud data may be deleted after 90 days.

## Third-Party Services

The extension interacts with the following third-party services:

| Service | Purpose | Data Sent |
|---|---|---|
| Monarch Money (`api.monarch.com`) | Retrieve financial data | Credentials (from browser), session token |
| ProjectionLab (`app.projectionlab.com`) | Write synced data | Account balances, budget amounts via Plugin API |
| Google Firebase | Authentication and cloud storage (hosted mode) | Email, configuration backups |
| Stripe | Payment processing (hosted mode) | Handled entirely by Stripe — we never see card details |

## Your Rights

- **Export**: You can export all your mappings and settings to a JSON file at any time from the Settings tab.
- **Delete**: You can clear all local data by uninstalling the extension. Hosted users can delete cloud backups from the Settings tab and cancel their subscription from the Account menu.
- **Access**: All stored data is visible and accessible through the extension's UI.

## Open Source

This extension is open source. You can review the complete source code to verify these privacy practices:
https://github.com/Treygec/Monarch-ProjectionLab-Bridge

## Contact

For privacy questions or concerns, please open an issue on the GitHub repository or contact treycorple@gmail.com.

## Changes

We may update this privacy policy from time to time. Changes will be reflected in the "Last updated" date above and committed to the source repository.
