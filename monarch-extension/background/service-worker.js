// ============================================================
// Monarch → ProjectionLab Bridge  —  Service Worker v4
// ============================================================

const DEFAULT_MONARCH_API       = 'http://localhost:47821';
const FIREBASE_API_KEY          = 'AIzaSyBlQFWbPQ1bhREni8ZhfKbRsd3lKOPPO0s';
const FIREBASE_FUNCTIONS_BASE   = 'https://us-central1-monarch-bridge-prod.cloudfunctions.net';

async function isHostedMode() {
  const stored = await chrome.storage.local.get('connectionMode');
  return (stored.connectionMode || 'self-hosted') === 'hosted';
}

async function getFirebaseToken() {
  const stored = await chrome.storage.local.get(['firebaseIdToken', 'firebaseRefreshToken', 'firebaseTokenExpiry']);
  const now = Date.now();
  // Valid if not expiring within the next 5 minutes
  if (stored.firebaseIdToken && stored.firebaseTokenExpiry > now + 5 * 60 * 1000) {
    return stored.firebaseIdToken;
  }
  if (!stored.firebaseRefreshToken) return null;
  try {
    const res = await fetch(
      `https://securetoken.googleapis.com/v1/token?key=${FIREBASE_API_KEY}`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: `grant_type=refresh_token&refresh_token=${encodeURIComponent(stored.firebaseRefreshToken)}`,
      }
    );
    if (!res.ok) return null;
    const data = await res.json();
    const newToken   = data.id_token;
    const newRefresh = data.refresh_token;
    const expiry     = Date.now() + (parseInt(data.expires_in, 10) || 3600) * 1000;
    await chrome.storage.local.set({ firebaseIdToken: newToken, firebaseRefreshToken: newRefresh, firebaseTokenExpiry: expiry });
    return newToken;
  } catch { return null; }
}

async function monarchHeaders() {
  const stored = await chrome.storage.local.get(['monarchEmail', 'monarchPassword', 'monarchSessionToken', 'settings']);
  const headers = { 'Accept': 'application/json' };
  if (stored.monarchSessionToken) headers['X-Monarch-Token']    = stored.monarchSessionToken;
  if (stored.monarchEmail)        headers['X-Monarch-Email']    = stored.monarchEmail;
  if (stored.monarchPassword)     headers['X-Monarch-Password'] = stored.monarchPassword;
  const s = stored.settings || {};
  if (s.monarchMfaEnabled && s.monarchMfaSecret) headers['X-Monarch-MFA'] = s.monarchMfaSecret;
  if (await isHostedMode()) {
    const idToken = await getFirebaseToken();
    if (idToken) headers['Authorization'] = `Bearer ${idToken}`;
  }
  return headers;
}

// ---------------------------------------------------------------------------
// TOTP generator (RFC 6238) using Web Crypto API
// ---------------------------------------------------------------------------
function base32Decode(str) {
  const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ234567';
  str = str.replace(/[\s=-]+/g, '').toUpperCase();
  let bits = '';
  for (const c of str) {
    const val = alphabet.indexOf(c);
    if (val === -1) continue;
    bits += val.toString(2).padStart(5, '0');
  }
  const bytes = new Uint8Array(Math.floor(bits.length / 8));
  for (let i = 0; i < bytes.length; i++) {
    bytes[i] = parseInt(bits.slice(i * 8, i * 8 + 8), 2);
  }
  return bytes;
}

async function generateTOTP(secret, period = 30, digits = 6) {
  const key = base32Decode(secret);
  const counter = Math.floor(Date.now() / 1000 / period);
  const counterBytes = new Uint8Array(8);
  let tmp = counter;
  for (let i = 7; i >= 0; i--) {
    counterBytes[i] = tmp & 0xff;
    tmp = Math.floor(tmp / 256);
  }
  const cryptoKey = await crypto.subtle.importKey(
    'raw', key, { name: 'HMAC', hash: 'SHA-1' }, false, ['sign']
  );
  const sig = new Uint8Array(await crypto.subtle.sign('HMAC', cryptoKey, counterBytes));
  const offset = sig[sig.length - 1] & 0x0f;
  const code = (
    ((sig[offset] & 0x7f) << 24) |
    ((sig[offset + 1] & 0xff) << 16) |
    ((sig[offset + 2] & 0xff) << 8) |
    (sig[offset + 3] & 0xff)
  ) % (10 ** digits);
  return code.toString().padStart(digits, '0');
}

async function getMfaSecretIfEnabled() {
  const stored = await chrome.storage.local.get('settings');
  const s = stored.settings || {};
  if (s.monarchMfaEnabled && s.monarchMfaSecret) return s.monarchMfaSecret;
  return null;
}

async function monarchBase() {
  if (await isHostedMode()) return FIREBASE_FUNCTIONS_BASE;
  const stored = await chrome.storage.local.get('monarchApiUrl');
  return (stored.monarchApiUrl || DEFAULT_MONARCH_API).replace(/\/$/, '');
}

// Map local proxy paths to Firebase Cloud Function names
const FIREBASE_PATH_MAP = {
  '/accounts':        '/get_accounts',
  '/expense-budgets': '/get_expense_budgets',
  '/budgets':         '/get_budgets',
  '/transactions':    '/get_transactions',
  '/cashflow':        '/get_cashflow',
  '/categories':      '/get_categories',
  '/tags':            '/get_tags',
  '/recurring':       '/get_recurring_transactions',
};

async function monarchUrl(localPath) {
  const base = await monarchBase();
  if (await isHostedMode()) {
    return base + (FIREBASE_PATH_MAP[localPath] || localPath);
  }
  return base + localPath;
}

async function monarchErrorMessage(res) {
  try {
    const body = await res.json();
    const detail = body?.detail || '';
    // Strip internal prefix if present, use the human-readable part
    if (detail.includes('MFA_REQUIRED:')) return detail.replace('MFA_REQUIRED: ', '');
    if (detail) return detail;
  } catch {}
  return `Monarch API returned ${res.status}`;
}

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  handle(msg).then(sendResponse).catch(e => sendResponse({ error: e.message || String(e) }));
  return true;
});

async function handle(msg) {
  switch (msg.type) {

    case 'FETCH_MONARCH_ACCOUNTS': {
      const url = await monarchUrl('/accounts');
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), 10000);
      let res;
      try {
        res = await fetch(url, { headers: await monarchHeaders(), signal: controller.signal });
      } catch (e) {
        if (e.name === 'AbortError') throw new Error('Monarch API timed out after 10s — is the local server running?');
        throw e;
      } finally { clearTimeout(timeout); }
      if (!res.ok) throw new Error(await monarchErrorMessage(res));
      const body = await res.json();
      const raw = Array.isArray(body) ? body : (body?.accounts ?? null);
      if (!Array.isArray(raw)) throw new Error('Unexpected response from Monarch API');
      const accounts = raw.map(a => ({
        ...a,
        name:    a.displayName    ?? a.name    ?? '',
        balance: a.displayBalance ?? a.currentBalance ?? a.balance ?? 0,
      }));
      return { accounts };
    }

    case 'MONARCH_INITIATE_LOGIN': {
      const hosted = await isHostedMode();
      if (hosted) {
        // Login directly to Monarch from the browser (REST endpoint).
        // Firebase/GCP can't reach this endpoint (Cloudflare blocks server IPs),
        // but the extension runs in the user's browser so it works fine.
        const stored = await chrome.storage.local.get(['monarchEmail', 'monarchPassword']);
        const controller = new AbortController();
        const timeout = setTimeout(() => controller.abort(), 20000);
        let res;
        try {
          res = await fetch('https://api.monarch.com/auth/login/', {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              'Client-Platform': 'web',
            },
            credentials: 'omit',
            body: JSON.stringify({
              username: stored.monarchEmail || '',
              password: stored.monarchPassword || '',
              supports_mfa: true,
              trusted_device: false,
            }),
            signal: controller.signal,
          });
        } catch (e) {
          if (e.name === 'AbortError') throw new Error('Login timed out');
          throw e;
        } finally { clearTimeout(timeout); }

        // 403 = MFA required — auto-complete if TOTP secret is configured
        if (res.status === 403) {
          const mfaSecret = await getMfaSecretIfEnabled();
          if (mfaSecret) {
            const totp = await generateTOTP(mfaSecret);
            const mfaRes = await fetch('https://api.monarch.com/auth/login/', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json', 'Client-Platform': 'web' },
              credentials: 'omit',
              body: JSON.stringify({
                username: stored.monarchEmail || '',
                password: stored.monarchPassword || '',
                supports_mfa: true,
                trusted_device: false,
                totp,
              }),
            });
            if (!mfaRes.ok) {
              const mfaBody = await mfaRes.json().catch(() => ({}));
              throw new Error(mfaBody?.detail || `Auto-MFA failed: ${mfaRes.status}`);
            }
            const mfaData = await mfaRes.json();
            if (mfaData.token) {
              await chrome.storage.local.set({ monarchSessionToken: mfaData.token });
            }
            return { ok: true, requires_mfa: false, token: mfaData.token };
          }
          return { ok: true, requires_mfa: true };
        }
        if (!res.ok) {
          const body = await res.json().catch(() => ({}));
          throw new Error(body?.detail || `Login failed: ${res.status}`);
        }
        const data = await res.json();
        if (data.token) {
          await chrome.storage.local.set({ monarchSessionToken: data.token });
        }
        return { ok: true, requires_mfa: false, token: data.token };
      }
      // Self-hosted: use local proxy
      const url2 = await monarchUrl('/auth/initiate');
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), 20000);
      let res;
      try {
        const hdrs = await monarchHeaders();
        hdrs['Content-Type'] = 'application/json';
        res = await fetch(url2, { method: 'POST', headers: hdrs, body: '{}', signal: controller.signal });
      } catch (e) {
        if (e.name === 'AbortError') throw new Error('Login timed out — is the proxy running?');
        throw e;
      } finally { clearTimeout(timeout); }
      if (!res.ok) throw new Error(await monarchErrorMessage(res));
      return await res.json();
    }

    case 'MONARCH_COMPLETE_MFA': {
      const hosted2 = await isHostedMode();
      if (hosted2) {
        // MFA: re-POST to same Monarch REST endpoint with totp code
        const stored = await chrome.storage.local.get(['monarchEmail', 'monarchPassword']);
        const controller = new AbortController();
        const timeout = setTimeout(() => controller.abort(), 20000);
        let res;
        try {
          res = await fetch('https://api.monarch.com/auth/login/', {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              'Client-Platform': 'web',
            },
            credentials: 'omit',
            body: JSON.stringify({
              username: stored.monarchEmail || '',
              password: stored.monarchPassword || '',
              supports_mfa: true,
              trusted_device: false,
              totp: msg.code,
            }),
            signal: controller.signal,
          });
        } catch (e) {
          if (e.name === 'AbortError') throw new Error('MFA verification timed out');
          throw e;
        } finally { clearTimeout(timeout); }
        if (!res.ok) {
          const body = await res.json().catch(() => ({}));
          throw new Error(body?.detail || `MFA failed: ${res.status}`);
        }
        const data = await res.json();
        if (data.token) {
          await chrome.storage.local.set({ monarchSessionToken: data.token });
        }
        return { ok: true, token: data.token };
      }
      // Self-hosted: use local proxy
      const url3 = await monarchUrl('/auth/complete');
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), 20000);
      let res;
      try {
        const hdrs = await monarchHeaders();
        hdrs['Content-Type'] = 'application/json';
        res = await fetch(url3, {
          method: 'POST', headers: hdrs,
          body: JSON.stringify({ code: msg.code }),
          signal: controller.signal,
        });
      } catch (e) {
        if (e.name === 'AbortError') throw new Error('MFA verification timed out');
        throw e;
      } finally { clearTimeout(timeout); }
      if (!res.ok) throw new Error(await monarchErrorMessage(res));
      return await res.json();
    }

    case 'MONARCH_REFRESH_SESSION': {
      const urlR = await monarchUrl('/auth/refresh');
      const hdrs = await monarchHeaders();
      hdrs['Content-Type'] = 'application/json';
      const controllerR = new AbortController();
      const timeoutR = setTimeout(() => controllerR.abort(), 15000);
      let resR;
      try {
        resR = await fetch(urlR, { method: 'POST', headers: hdrs, signal: controllerR.signal });
      } catch (e) {
        if (e.name === 'AbortError') throw new Error('Monarch session refresh timed out — is the proxy running?');
        throw e;
      } finally { clearTimeout(timeoutR); }
      if (!resR.ok) throw new Error(`Session refresh failed: ${resR.status}`);
      return { ok: true };
    }

    case 'PL_EXPORT_DATA': {
      return execInPL(msg.key, async (key) => {
        return await window.projectionlabPluginAPI.exportData({ key });
      });
    }

    case 'PL_UPDATE_ACCOUNT': {
      return execInPL(msg.key, async (key, plAccountId, data) => {
        await window.projectionlabPluginAPI.updateAccount(plAccountId, data, { key });
        return { ok: true };
      }, [msg.plAccountId, msg.data]);
    }

    // Create new PL accounts by exporting, appending, then restoring current finances
    case 'PL_CREATE_ACCOUNTS': {
      return execInPL(msg.key, async (key, newAccounts) => {
        function uuid() { return crypto.randomUUID(); }

        function noteFor(a) {
          const names = a.monarchNames;
          if (!names || !names.length) return {};
          return { notes: 'Monarch: ' + names.join(' + '), hasNotes: true };
        }

        function makeSavings(a) {
          return {
            id: uuid(), name: a.name, title: 'Savings', type: 'savings',
            balance: a.balance, owner: 'me',
            icon: 'mdi-piggy-bank', color: 'teal',
            dividendType: 'none', dividendRate: 0,
            investmentGrowthType: 'fixed', investmentGrowthRate: 0,
            ...noteFor(a),
          };
        }

        function makeInvestment(a) {
          const typeIconMap = {
            'taxable':    { icon: 'mdi-chart-line',    color: 'green',             title: 'Taxable Brokerage' },
            '401k':       { icon: 'mdi-briefcase',     color: 'blue-darken-2',     title: '401(k)' },
            'roth-401k':  { icon: 'mdi-briefcase',     color: 'blue-darken-2',     title: 'Roth 401(k)' },
            'roth-ira':   { icon: 'mdi-shield-account',color: 'cyan-darken-1',     title: 'Roth IRA' },
            'ira':        { icon: 'mdi-shield-account',color: 'cyan-darken-1',     title: 'Traditional IRA' },
            '403b':       { icon: 'mdi-briefcase',     color: 'blue-darken-2',     title: '403(b)' },
            'hsa':        { icon: 'mdi-hospital-box',  color: 'green-darken-1',    title: 'HSA' },
            'crypto':     { icon: 'mdi-currency-btc',  color: 'purple-lighten-2',  title: 'Cryptocurrency' },
            '529':        { icon: 'mdi-school',        color: 'teal-darken-1',     title: '529 Plan' },
          };
          const meta = typeIconMap[a.plType] || typeIconMap['taxable'];
          return {
            id: uuid(), name: a.name, title: meta.title, type: a.plType || 'taxable',
            balance: a.balance, costBasis: 0, owner: 'me',
            icon: meta.icon, color: meta.color, subtitle: '',
            dividendType: 'none', dividendRate: 0,
            dividendTaxType: 'qualified', dividendReinvestment: true,
            dividendsArePassiveIncome: false, isPassiveIncome: false,
            investmentGrowthType: 'plan', investmentGrowthRate: 0,
            yearlyFee: 0, yearlyFeeType: '%', liquid: false,
            ...noteFor(a),
          };
        }

        function makeDebt(a) {
          const DEBT_META = {
            'debt':             { title: 'Unsecured Debt',  icon: 'mdi-lock',           color: 'orange-lighten-1' },
            'student-loans':    { title: 'Student Loans',   icon: 'mdi-school',          color: 'blue-lighten-1' },
            'medical-debt':     { title: 'Medical Debt',    icon: 'mdi-hospital-box',    color: 'red-lighten-1' },
            'credit-card-debt': { title: 'Credit Card Debt',icon: 'mdi-credit-card',     color: 'deep-orange-lighten-1' },
          };
          const meta = DEBT_META[a.plType] || DEBT_META['debt'];
          return {
            id: uuid(), name: a.name,
            title: meta.title, type: a.plType || 'debt',
            amount: a.balance, amountType: 'today$', owner: 'me',
            icon: meta.icon, color: meta.color, order: 0,
            additionalFields: { disabled: false, whitelist: ['fundWithAccounts','taxDeductible','itemized'] },
            start: { type: 'keyword', value: 'beforeCurrentYear' },
            end:   { type: 'keyword', value: 'never' },
            frequency: 'monthly', frequencyChoices: true,
            yearlyChange: { amount: 0, amountType: 'today$', type: 'none', limitEnabled: false, limit: 0, limitType: 'today$' },
            monthlyPayment: 0, monthlyPaymentType: 'today$',
            interestRate: 0, interestType: 'compound', compounding: 'daily',
            effectiveDate: { type: 'keyword', value: 'beforeCurrentYear' },
            hasForgiveness: false,
            forgiveAt: { type: 'keyword', value: 'now', modifier: 'include' },
            planPath: 'expenses',
            ...noteFor(a),
          };
        }

        function makeAsset(a) {
          // Per-type metadata: [icon, color, title, yearlyChangeType, yearlyChangeAmt, brokersFee]
          const ASSET_META = {
            'real-estate':         { icon:'mdi-home',               color:'indigo-lighten-1',    title:'House',             chgType:'appreciate', chgAmt:3,   fee:5.5 },
            'car':                 { icon:'mdi-car',                 color:'deep-purple',         title:'Car',               chgType:'depreciate', chgAmt:8,   fee:0   },
            'rental-property':     { icon:'mdi-home-city',           color:'indigo-darken-1',     title:'Rental Property',   chgType:'appreciate', chgAmt:3,   fee:5.5 },
            'land':                { icon:'mdi-terrain',             color:'brown-lighten-1',     title:'Land',              chgType:'appreciate', chgAmt:2,   fee:5.5 },
            'building':            { icon:'mdi-office-building',     color:'blue-grey',           title:'Building',          chgType:'appreciate', chgAmt:2,   fee:5.5 },
            'commercial-property': { icon:'mdi-office-building',     color:'blue-grey-darken-1',  title:'Commercial Property',chgType:'appreciate',chgAmt:2,   fee:5.5 },
            'motorcycle':          { icon:'mdi-motorbike',           color:'cyan-darken-1',       title:'Motorcycle',        chgType:'depreciate', chgAmt:10,  fee:0   },
            'boat':                { icon:'mdi-sail-boat',           color:'light-blue-darken-1', title:'Boat',              chgType:'depreciate', chgAmt:8,   fee:0   },
            'jewelry':             { icon:'mdi-diamond-stone',       color:'pink-lighten-1',      title:'Jewelry',           chgType:'appreciate', chgAmt:1,   fee:0   },
            'precious-metals':     { icon:'mdi-gold',                color:'amber-darken-1',      title:'Precious Metals',   chgType:'appreciate', chgAmt:2,   fee:0   },
            'furniture':           { icon:'mdi-sofa',                color:'brown',               title:'Furniture',         chgType:'depreciate', chgAmt:10,  fee:0   },
            'instrument':          { icon:'mdi-music',               color:'deep-purple-lighten-1',title:'Instrument',       chgType:'depreciate', chgAmt:5,   fee:0   },
            'machinery':           { icon:'mdi-engine',              color:'grey-darken-1',       title:'Machinery',         chgType:'depreciate', chgAmt:10,  fee:0   },
            'custom':              { icon:'mdi-shape',               color:'blue-grey-lighten-1', title:'Custom Asset',      chgType:'appreciate', chgAmt:0,   fee:0   },
          };
          const meta = ASSET_META[a.plType] || ASSET_META['real-estate'];

          // writeField determines which field gets the Monarch balance.
          // All other fields are populated from manualValues (user-entered overrides),
          // defaulting to 0. This avoids a race condition where post-creation updateAccount
          // calls would fire before restoreCurrentFinances fully commits the new account.
          const wf  = a.writeField || 'amount';  // 'amount'|'balance'|'initialValue'
          const val = a.balance ?? 0;
          const mv  = a.manualValues || {};
          const assetObj = {
            id: uuid(), name: a.name,
            title: meta.title, type: a.plType || 'real-estate',
            amount:       wf === 'amount'       ? val : (mv.amount       || 0),
            balance:      wf === 'balance'      ? val : (mv.balance      || 0),
            initialValue: wf === 'initialValue' ? val : (mv.initialValue || 0),
            owner: 'me',
            icon: meta.icon, color: meta.color,
            planPath: 'assets',
            paymentMethod: a.paymentStatus || (wf === 'balance' ? 'financed' : 'financed'),
            start: { type: 'keyword', value: 'beforeCurrentYear' },
            end:   { type: 'keyword', value: 'never', modifier: 'include' },
            repeat: false,
            amountType: 'today$', initialValueType: 'today$', balanceType: 'today$',
            downPayment: 0, downPaymentType: 'today$',
            monthlyPayment: 0, monthlyPaymentType: 'today$',
            monthlyHOA: 0,
            interestRate: 0, interestType: 'simple', compounding: 'monthly',
            taxRate: 0, taxRateType: '%',
            insuranceRate: 0, insuranceRateType: '%',
            maintenanceRate: 0, maintenanceRateType: '%',
            improvementRate: 0, improvementRateType: '%',
            incomeRate: 0, incomeRateType: '%',
            managementRate: 0, managementRateType: '%',
            yearlyChange: { amount: meta.chgAmt, type: meta.chgType,
                            amountType: 'today$', limitEnabled: false, limit: 0, limitType: 'today$' },
            brokersFee: meta.fee,
            generateIncome: false, isPassiveIncome: false,
            selfEmployment: false, cancelRent: false,
            estimateRentalDeductions: false, percentRented: 0,
            initialBuildingValue: 0, initialBuildingValueType: 'today$',
            excludeLoanFromLNW: false,
            ...noteFor(a),
          };
          return assetObj;
        }

        const exported = await window.projectionlabPluginAPI.exportData({ key });
        const today = exported.today;

        const createdIds = [];
        for (const a of newAccounts) {
          let obj;
          if      (a.plCategory === 'savings')    { obj = makeSavings(a);    today.savingsAccounts    = [...(today.savingsAccounts    || []), obj]; }
          else if (a.plCategory === 'investment')  { obj = makeInvestment(a); today.investmentAccounts = [...(today.investmentAccounts || []), obj]; }
          else if (a.plCategory === 'debt')        { obj = makeDebt(a);       today.debts              = [...(today.debts              || []), obj]; }
          else if (a.plCategory === 'asset')       { obj = makeAsset(a);      today.assets             = [...(today.assets             || []), obj]; }
          if (obj) createdIds.push({ monarchId: a.monarchId, plId: obj.id, plCategory: a.plCategory, writeField: a.writeField || null });
        }

        await window.projectionlabPluginAPI.restoreCurrentFinances(today, { key });
        return { ok: true, created: createdIds };

      }, [msg.accounts]);
    }

    // Write notes to existing PL accounts via export→inject→restoreCurrentFinances.
    // This is the only way to add notes/hasNotes to accounts that don't already have those fields.
    // PL_SYNC_ALL: one export → apply all account updates + notes + plan expenses → two restores max.
    // accountUpdates: [{ plAccountId, data }]
    // notesMap: { [plRealId]: 'Name1 + Name2' }
    // expenseUpdates: same as PL_WRITE_PLAN_EXPENSES
    case 'PL_SYNC_ALL': {
      return execInPL(msg.key, async (key, accountUpdates, notesMap, expenseUpdates) => {
        const exported = await window.projectionlabPluginAPI.exportData({ key });
        const today    = exported.today;

        // ── Account field updates ────────────────────────────────
        const acctArrays = [
          today.savingsAccounts    || [],
          today.investmentAccounts || [],
          today.debts              || [],
          today.assets             || [],
        ];
        let acctUpdated = 0;
        for (const { plAccountId, data } of accountUpdates) {
          for (const arr of acctArrays) {
            const acct = arr.find(a => String(a.id) === String(plAccountId));
            if (acct) { Object.assign(acct, data); acctUpdated++; break; }
          }
        }

        // ── Notes injection ──────────────────────────────────────
        const MONARCH_START = '--- Monarch ---';
        const MONARCH_END   = '--- End Monarch ---';
        let notesUpdated = 0;
        if (notesMap && Object.keys(notesMap).length) {
          for (const arr of acctArrays) {
            for (const acct of arr) {
              const names = notesMap[String(acct.id)];
              if (!names) continue;
              const monarchBlock = MONARCH_START + '\n' + names + '\n' + MONARCH_END;
              const existing = (acct.notes || '').replace(/\n?--- Monarch ---[\s\S]*?--- End Monarch ---\n?/g, '').trimEnd();
              acct.notes    = existing ? existing + '\n\n' + monarchBlock : monarchBlock;
              acct.hasNotes = true;
              notesUpdated++;
            }
          }
        }

        // ── Plan expense updates ─────────────────────────────────
        const MONARCH_S = '--- Monarch ---';
        const MONARCH_E = '--- End Monarch ---';

        function applyExpenseNotes(event, monarchNames) {
          if (!monarchNames || !monarchNames.length) return;
          const block    = MONARCH_S + '\n' + monarchNames.join(' + ') + '\n' + MONARCH_E;
          const existing = (event.notes || '').replace(/\n?--- Monarch ---[\s\S]*?--- End Monarch ---\n?/g, '').trimEnd();
          event.notes    = existing ? existing + '\n\n' + block : block;
          event.hasNotes = true;
        }

        const baseYearlyChange = (type) => ({
          amount: 0, amountType: 'today$', limitEnabled: false, limit: 0, limitType: 'today$',
          type: (type === 'debt' || type === 'charity' || type === 'wedding' || type === 'other') ? 'none' : 'match-inflation',
        });

        let expUpdated = 0;
        if (expenseUpdates && expenseUpdates.length) {
          // Pass 0: create brand-new expenses
          for (const { planId, expenseId, newExpenseId, newExpenseName, newFrequency, newExpType, newExpTypeParams, amount, writeField, monarchNames } of expenseUpdates) {
            if (expenseId !== '__CREATE_NEW__') continue;
            const targetPlan = (exported.plans || []).find(p => p.id === planId);
            if (!targetPlan) continue;
            if (!targetPlan.expenses) targetPlan.expenses = { events: [] };
            if (!targetPlan.expenses.events) targetPlan.expenses.events = [];
            if (targetPlan.expenses.events.some(e => e.id === newExpenseId)) {
              targetPlan.expenses.events.find(e => e.id === newExpenseId)[writeField || 'amount'] = amount;
              expUpdated++; continue;
            }
            const expType = newExpType || 'living-expenses';
            const freq    = newFrequency || 'monthly';
            const params  = newExpTypeParams || {};
            let event;
            switch (expType) {
              case 'health-care': event = { id: newExpenseId, name: newExpenseName, title: 'Health Care', type: 'health-care', frequency: freq, frequencyChoices: true, planPath: 'expenses', amount, amountType: 'today$', spendingType: 'essential', taxDeductible: params.taxDeductible ?? false, itemized: params.itemized ?? false, deductFromIncomeId: '', yearlyChange: baseYearlyChange(expType), icon: 'mdi-heart-plus', color: 'pink-lighten-1', owner: 'me', key: Math.random(), start: { type: 'keyword', value: 'beforeCurrentYear' }, end: { modifier: 'include', type: 'keyword', value: 'endOfPlan' } }; break;
              case 'debt': event = { id: newExpenseId, name: newExpenseName, title: 'Debt', type: 'debt', frequency: 'monthly', frequencyChoices: true, planPath: 'expenses', amount: params.debtBalance || 0, monthlyPayment: amount, monthlyPaymentType: 'today$', amountType: 'today$', interestRate: params.interestRate ?? 0, interestType: 'compound', compounding: params.compounding || 'daily', hasForgiveness: false, effectiveDate: { type: 'keyword', value: 'beforeCurrentYear' }, forgiveAt: { value: 'now', type: 'keyword', modifier: 'include' }, yearlyChange: baseYearlyChange(expType), icon: 'mdi-lock', color: 'orange-lighten-1', owner: 'me', key: Math.random(), start: { type: 'keyword', value: 'beforeCurrentYear' }, end: { type: 'keyword', value: 'never' } }; break;
              case 'charity': event = { id: newExpenseId, name: newExpenseName, title: 'Charity', type: 'charity', frequency: freq, frequencyChoices: true, planPath: 'expenses', amount, amountType: 'today$', spendingType: 'discretionary', taxDeductible: params.taxDeductible ?? ['federal','state'], itemized: params.itemized ?? true, repeat: false, yearlyChange: baseYearlyChange(expType), icon: 'mdi-charity', color: 'blue-lighten-1', owner: 'me', key: Math.random(), start: { type: 'keyword', value: 'beforeCurrentYear' }, end: { modifier: 'include', type: 'keyword', value: 'endOfPlan' } }; break;
              case 'wedding': event = { id: newExpenseId, name: newExpenseName, title: 'Wedding', type: 'wedding', frequency: 'once', frequencyChoices: true, planPath: 'expenses', amount, amountType: 'today$', spendingType: 'discretionary', switchToMarried: false, repeat: false, yearlyChange: baseYearlyChange(expType), icon: 'mdi-diamond-stone', color: 'pink-lighten-1', owner: 'me', key: Math.random(), start: { type: 'date', value: new Date().toISOString().slice(0,10) }, end: { modifier: 'include', type: 'keyword', value: 'endOfPlan' } }; break;
              case 'other': event = { id: newExpenseId, name: newExpenseName, title: 'Custom Expense', type: 'other', frequency: 'once', frequencyChoices: true, planPath: 'expenses', amount, amountType: 'today$', spendingType: 'essential', taxDeductible: false, itemized: false, deductFromIncomeId: '', repeat: false, yearlyChange: baseYearlyChange(expType), icon: 'mdi-currency-usd-circle', color: 'brown-lighten-1', owner: 'me', key: Math.random(), start: { type: 'date', value: new Date().toISOString().slice(0,10) }, end: { modifier: 'include', type: 'keyword', value: 'endOfPlan' } }; break;
              default: event = { id: newExpenseId, name: newExpenseName, title: expType === 'dependent-support' ? 'Dependent Support' : 'Living Expenses', type: expType === 'dependent-support' ? 'dependent-support' : 'living-expenses', frequency: freq, frequencyChoices: true, planPath: 'expenses', amount, amountType: 'today$', spendingType: 'essential', yearlyChange: baseYearlyChange(expType), icon: 'mdi-home-city', color: 'blue-lighten-1', owner: 'me', key: Math.random(), start: { type: 'keyword', value: 'beforeCurrentYear' }, end: { modifier: 'include', type: 'keyword', value: 'endOfPlan' } };
            }
            targetPlan.expenses.events.push(event);
            applyExpenseNotes(event, monarchNames);
            expUpdated++;
          }

          // Pass 1: update existing expenses
          for (const { planId, expenseId, amount, writeField, monarchNames } of expenseUpdates) {
            if (expenseId === '__CREATE_NEW__') continue;
            const plan  = (exported.plans || []).find(p => p.id === planId);
            const event = plan?.expenses?.events?.find(e => e.id === expenseId);
            if (!event) continue;
            event[writeField || 'amount'] = amount;
            applyExpenseNotes(event, monarchNames);
            expUpdated++;
          }

          // Pass 2: createIfMissing — clone into plans that don't have the expense yet
          for (const { planId, expenseId, createIfMissing, amount, writeField, monarchNames } of expenseUpdates) {
            if (!createIfMissing || expenseId === '__CREATE_NEW__') continue;
            const targetPlan = (exported.plans || []).find(p => p.id === planId);
            if (!targetPlan) continue;
            if (!targetPlan.expenses) targetPlan.expenses = { events: [] };
            if (!targetPlan.expenses.events) targetPlan.expenses.events = [];
            const fieldToWrite = writeField || 'amount';
            const existingEvent = targetPlan.expenses.events.find(e => e.id === expenseId);
            if (existingEvent) { existingEvent[fieldToWrite] = amount; applyExpenseNotes(existingEvent, monarchNames); expUpdated++; continue; }
            let sourceEvent = null;
            for (const plan of (exported.plans || [])) {
              sourceEvent = plan.expenses?.events?.find(e => e.id === expenseId);
              if (sourceEvent) break;
            }
            if (!sourceEvent) continue;
            const cloned = JSON.parse(JSON.stringify(sourceEvent));
            cloned[fieldToWrite] = amount;
            applyExpenseNotes(cloned, monarchNames);
            targetPlan.expenses.events.push(cloned);
            expUpdated++;
          }
        }

        // ── Two restores max ─────────────────────────────────────
        if (acctUpdated + notesUpdated > 0) {
          await window.projectionlabPluginAPI.restoreCurrentFinances(today, { key });
        }
        if (expUpdated > 0) {
          await window.projectionlabPluginAPI.restorePlans(exported.plans, { key });
        }

        return { ok: true, acctUpdated, notesUpdated, expUpdated, exported };
      }, [msg.accountUpdates, msg.notesMap, msg.expenseUpdates]);
    }

    case 'PL_WRITE_NOTES': {
      return execInPL(msg.key, async (key, notesMap) => {
        // notesMap: { [plRealId]: 'Name1 + Name2' }  (just the account names, no prefix)
        const MONARCH_START = '--- Monarch ---';
        const MONARCH_END   = '--- End Monarch ---';

        const exported = await window.projectionlabPluginAPI.exportData({ key });
        const today    = exported.today;

        const arrays = [
          today.savingsAccounts    || [],
          today.investmentAccounts || [],
          today.debts              || [],
          today.assets             || [],
        ];

        let updated = 0;
        for (const arr of arrays) {
          for (const acct of arr) {
            const names = notesMap[String(acct.id)];
            if (!names) continue;

            const monarchBlock = MONARCH_START + '\n' + names + '\n' + MONARCH_END;

            // Strip any existing Monarch block from the current notes
            const existing = (acct.notes || '').replace(
              /\n?--- Monarch ---[\s\S]*?--- End Monarch ---\n?/g, ''
            ).trimEnd();

            // Append fresh Monarch block, separated from existing content
            acct.notes    = existing ? existing + '\n\n' + monarchBlock : monarchBlock;
            acct.hasNotes = true;
            updated++;
          }
        }

        if (updated > 0) {
          await window.projectionlabPluginAPI.restoreCurrentFinances(today, { key });
        }
        return { ok: true, updated };
      }, [msg.notesMap]);
    }

    case 'FETCH_MONARCH_BUDGETS': {
      const url2 = await monarchUrl('/expense-budgets');
      const controller2 = new AbortController();
      const timeout2 = setTimeout(() => controller2.abort(), 10000);
      let res2;
      try {
        res2 = await fetch(url2, { headers: await monarchHeaders(), signal: controller2.signal });
      } catch (e) {
        if (e.name === 'AbortError') throw new Error('Monarch API timed out after 10s');
        throw e;
      } finally { clearTimeout(timeout2); }
      if (!res2.ok) throw new Error(await monarchErrorMessage(res2));
      const budgets = await res2.json();
      if (!Array.isArray(budgets)) throw new Error('Unexpected budget response from Monarch API');
      return { budgets };
    }

    // Update expense amounts in PL plans.
    // expenseUpdates: [{ planId, expenseId, amount }]
    case 'PL_WRITE_PLAN_EXPENSES': {
      return execInPL(msg.key, async (key, expenseUpdates) => {
        const exported = await window.projectionlabPluginAPI.exportData({ key });

        let updated = 0;

        // Pass 0: create brand-new expenses (expenseId === '__CREATE_NEW__')
        for (const { planId, expenseId, newExpenseId, newExpenseName, newFrequency, newExpType, newExpTypeParams, amount, writeField } of expenseUpdates) {
          if (expenseId !== '__CREATE_NEW__') continue;
          const targetPlan = (exported.plans || []).find(p => p.id === planId);
          if (!targetPlan) continue;
          if (!targetPlan.expenses) targetPlan.expenses = { events: [] };
          if (!targetPlan.expenses.events) targetPlan.expenses.events = [];
          // Don't duplicate if already created in a prior sync
          if (targetPlan.expenses.events.some(e => e.id === newExpenseId)) {
            const existing = targetPlan.expenses.events.find(e => e.id === newExpenseId);
            existing[writeField || 'amount'] = amount;
            updated++; continue;
          }

          const expType  = newExpType || 'living-expenses';
          const freq     = newFrequency || 'monthly';
          const params   = newExpTypeParams || {};

          const baseYearlyChange = (type) => ({
            amount: 0, amountType: 'today$', limitEnabled: false, limit: 0, limitType: 'today$',
            type: (type === 'debt' || type === 'charity' || type === 'wedding' || type === 'other') ? 'none' : 'match-inflation',
          });

          let event;
          switch (expType) {
            case 'health-care':
              event = {
                id: newExpenseId, name: newExpenseName, title: 'Health Care', type: 'health-care',
                frequency: freq, frequencyChoices: true, planPath: 'expenses',
                amount, amountType: 'today$', spendingType: params.spendingType || 'essential',
                taxDeductible: params.taxDeductible ?? false,
                itemized: params.itemized ?? false,
                deductFromIncomeId: params.deductFromIncomeId || '',
                yearlyChange: baseYearlyChange(expType),
                icon: 'mdi-heart-plus', color: 'pink-lighten-1', owner: 'me',
                key: Math.random(),
                start: { type: 'keyword', value: 'beforeCurrentYear' },
                end:   { modifier: 'include', type: 'keyword', value: 'endOfPlan' },
              };
              break;

            case 'debt':
              event = {
                id: newExpenseId, name: newExpenseName, title: 'Debt', type: 'debt',
                frequency: 'monthly', frequencyChoices: true, planPath: 'expenses',
                amount: params.debtBalance || 0,       // balance field
                monthlyPayment: amount,                // Monarch budget amount goes here
                monthlyPaymentType: 'today$',
                amountType: 'today$',
                interestRate: params.interestRate ?? 0,
                interestType: 'compound',
                compounding: params.compounding || 'daily',
                hasForgiveness: false,
                effectiveDate: { type: 'keyword', value: 'beforeCurrentYear' },
                forgiveAt: { value: 'now', type: 'keyword', modifier: 'include' },
                yearlyChange: baseYearlyChange(expType),
                icon: 'mdi-lock', color: 'orange-lighten-1', owner: 'me',
                key: Math.random(),
                start: { type: 'keyword', value: 'beforeCurrentYear' },
                end:   { type: 'keyword', value: 'never' },
              };
              break;

            case 'charity':
              event = {
                id: newExpenseId, name: newExpenseName, title: 'Charity', type: 'charity',
                frequency: freq, frequencyChoices: true, planPath: 'expenses',
                amount, amountType: 'today$', spendingType: 'discretionary',
                taxDeductible: params.taxDeductible ?? ['federal', 'state'],
                itemized: params.itemized ?? true,
                repeat: false,
                yearlyChange: baseYearlyChange(expType),
                icon: 'mdi-charity', color: 'blue-lighten-1', owner: 'me',
                key: Math.random(),
                start: { type: 'keyword', value: 'beforeCurrentYear' },
                end:   { modifier: 'include', type: 'keyword', value: 'endOfPlan' },
              };
              break;

            case 'wedding':
              event = {
                id: newExpenseId, name: newExpenseName, title: 'Wedding', type: 'wedding',
                frequency: 'once', frequencyChoices: true, planPath: 'expenses',
                amount, amountType: 'today$',
                spendingType: 'discretionary',
                switchToMarried: false,
                repeat: false,
                yearlyChange: baseYearlyChange(expType),
                icon: 'mdi-diamond-stone', color: 'pink-lighten-1', owner: 'me',
                key: Math.random(),
                start: { type: 'date', value: new Date().toISOString().slice(0,10) },
                end:   { modifier: 'include', type: 'keyword', value: 'endOfPlan' },
              };
              break;

            case 'other':
              event = {
                id: newExpenseId, name: newExpenseName, title: 'Custom Expense', type: 'other',
                frequency: 'once', frequencyChoices: true, planPath: 'expenses',
                amount, amountType: 'today$',
                spendingType: 'essential',
                taxDeductible: false,
                itemized: false,
                deductFromIncomeId: '',
                repeat: false,
                yearlyChange: baseYearlyChange(expType),
                icon: 'mdi-currency-usd-circle', color: 'brown-lighten-1', owner: 'me',
                key: Math.random(),
                start: { type: 'date', value: new Date().toISOString().slice(0,10) },
                end:   { modifier: 'include', type: 'keyword', value: 'endOfPlan' },
              };
              break;

            default: // living-expenses + dependent-support
              event = {
                id: newExpenseId, name: newExpenseName,
                title: expType === 'dependent-support' ? 'Dependent Support' : 'Living Expenses',
                type: expType === 'dependent-support' ? 'dependent-support' : 'living-expenses',
                frequency: freq, frequencyChoices: true, planPath: 'expenses',
                amount, amountType: 'today$',
                spendingType: 'essential',
                yearlyChange: baseYearlyChange(expType),
                icon: 'mdi-home-city', color: 'blue-lighten-1', owner: 'me',
                key: Math.random(),
                start: { type: 'keyword', value: 'beforeCurrentYear' },
                end:   { modifier: 'include', type: 'keyword', value: 'endOfPlan' },
              };
          }

          targetPlan.expenses.events.push(event);
          updated++;
        }

        // Pass 1: update amounts for expenses that already exist in their target plan
        for (const { planId, expenseId, amount, writeField } of expenseUpdates) {
          if (expenseId === '__CREATE_NEW__') continue;
          const plan = (exported.plans || []).find(p => p.id === planId);
          if (!plan) continue;
          const events = plan.expenses?.events || [];
          const event = events.find(e => e.id === expenseId);
          if (!event) continue;
          // Use writeField to target the correct property (amount for most, monthlyPayment for debt)
          const fieldToWrite = writeField || 'amount';
          event[fieldToWrite] = amount;
          updated++;
        }

        // Pass 2: for createIfMissing entries, update if the event already exists in the
        // exported plan (S.plExpenses can be stale), clone only if truly not there.
        for (const { planId, expenseId, createIfMissing, amount, writeField } of expenseUpdates) {
          if (!createIfMissing || expenseId === '__CREATE_NEW__') continue;
          const targetPlan = (exported.plans || []).find(p => p.id === planId);
          if (!targetPlan) continue;
          if (!targetPlan.expenses) targetPlan.expenses = { events: [] };
          if (!targetPlan.expenses.events) targetPlan.expenses.events = [];
          const fieldToWrite = writeField || 'amount';

          const existingEvent = targetPlan.expenses.events.find(e => e.id === expenseId);
          if (existingEvent) {
            existingEvent[fieldToWrite] = amount;
            updated++;
            continue;
          }

          let sourceEvent = null;
          for (const plan of (exported.plans || [])) {
            sourceEvent = (plan.expenses?.events || []).find(e => e.id === expenseId);
            if (sourceEvent) break;
          }
          if (!sourceEvent) continue;
          const cloned = JSON.parse(JSON.stringify(sourceEvent));
          cloned[fieldToWrite] = amount;
          targetPlan.expenses.events.push(cloned);
          updated++;
        }

        if (updated === 0) return { ok: true, updated };

        await window.projectionlabPluginAPI.restorePlans(exported.plans, { key });
        return { ok: true, updated };
      }, [msg.expenseUpdates]);
    }

    case 'FIREBASE_SIGN_IN':
    case 'FIREBASE_SIGN_UP': {
      const endpoint = msg.type === 'FIREBASE_SIGN_UP'
        ? `https://identitytoolkit.googleapis.com/v1/accounts:signUp?key=${FIREBASE_API_KEY}`
        : `https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=${FIREBASE_API_KEY}`;
      const res = await fetch(endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email: msg.email, password: msg.password, returnSecureToken: true }),
      });
      const data = await res.json();
      if (!res.ok) {
        const errMsg = data?.error?.message || 'Authentication failed';
        throw new Error(errMsg.replace(/_/g, ' ').toLowerCase().replace(/^\w/, c => c.toUpperCase()));
      }
      const expiry = Date.now() + (parseInt(data.expiresIn, 10) || 3600) * 1000;
      await chrome.storage.local.set({
        firebaseIdToken:      data.idToken,
        firebaseRefreshToken: data.refreshToken,
        firebaseTokenExpiry:  expiry,
        firebaseEmail:        data.email,
      });
      return { ok: true, email: data.email };
    }

    case 'FIREBASE_SIGN_OUT': {
      await chrome.storage.local.remove(['firebaseIdToken', 'firebaseRefreshToken', 'firebaseTokenExpiry', 'firebaseEmail']);
      return { ok: true };
    }

    case 'FIREBASE_OPEN_BILLING': {
      const idToken = await getFirebaseToken();
      if (!idToken) throw new Error('Not signed in');
      const authHeader = { 'Content-Type': 'application/json', 'Authorization': `Bearer ${idToken}` };

      async function openOrReuse(url) {
        // Reuse an existing tab with the same origin, or open a new one
        const origin = new URL(url).origin;
        const existing = await chrome.tabs.query({ url: origin + '/*' });
        if (existing.length) {
          await chrome.tabs.update(existing[0].id, { url, active: true });
        } else {
          await chrome.tabs.create({ url });
        }
      }

      // Try portal session first (for existing subscribers)
      const portalRes = await fetch(FIREBASE_FUNCTIONS_BASE + '/create_portal_session', {
        method: 'POST', headers: authHeader, body: '{}',
      });
      if (portalRes.ok) {
        const portalData = await portalRes.json();
        if (portalData.url) { await openOrReuse(portalData.url); return { ok: true }; }
      }
      // 404 = no subscription yet → fall through to checkout
      if (portalRes.status >= 500) {
        const portalBody = await portalRes.json().catch(() => ({}));
        console.warn('Portal session error:', portalBody?.error || portalRes.status);
      }

      // No subscription yet — create a checkout session
      const checkoutRes = await fetch(FIREBASE_FUNCTIONS_BASE + '/create_checkout_session', {
        method: 'POST', headers: authHeader, body: '{}',
      });
      const checkoutBody = await checkoutRes.json().catch(() => ({}));
      if (!checkoutRes.ok) {
        throw new Error(checkoutBody?.error || `Billing error ${checkoutRes.status}`);
      }
      if (checkoutBody.url) { await openOrReuse(checkoutBody.url); return { ok: true }; }
      throw new Error('No checkout URL returned');
    }

    case 'FIREBASE_SAVE_BACKUP': {
      const idToken = await getFirebaseToken();
      if (!idToken) throw new Error('Not signed in');
      const res = await fetch(FIREBASE_FUNCTIONS_BASE + '/save_backup', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${idToken}` },
        body: JSON.stringify({ backup: msg.backup, name: msg.name || 'Untitled' }),
      });
      const body = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(body?.detail || `Backup save failed (${res.status})`);
      return { ok: true };
    }

    case 'FIREBASE_LOAD_BACKUP': {
      const idToken = await getFirebaseToken();
      if (!idToken) throw new Error('Not signed in');
      const payload = {};
      if (msg.versionId) payload.versionId = msg.versionId;
      const res = await fetch(FIREBASE_FUNCTIONS_BASE + '/load_backup', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${idToken}` },
        body: JSON.stringify(payload),
      });
      const body = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(body?.detail || `Backup load failed (${res.status})`);
      return { ok: true, backup: body.backup, versions: body.versions };
    }

    case 'FIREBASE_UPDATE_BACKUP': {
      const idToken = await getFirebaseToken();
      if (!idToken) throw new Error('Not signed in');
      const payload = { versionId: msg.versionId };
      if (msg.name !== undefined) payload.name = msg.name;
      if (msg.pinned !== undefined) payload.pinned = msg.pinned;
      const res = await fetch(FIREBASE_FUNCTIONS_BASE + '/update_backup', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${idToken}` },
        body: JSON.stringify(payload),
      });
      const body = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(body?.detail || `Update failed (${res.status})`);
      return { ok: true };
    }

    case 'FIREBASE_DELETE_BACKUP': {
      const idToken = await getFirebaseToken();
      if (!idToken) throw new Error('Not signed in');
      const res = await fetch(FIREBASE_FUNCTIONS_BASE + '/delete_backup', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${idToken}` },
        body: JSON.stringify({ versionId: msg.versionId }),
      });
      const body = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(body?.detail || `Delete failed (${res.status})`);
      return { ok: true };
    }

    case 'STORAGE_SET': {
      await chrome.storage.local.set(msg.data);
      return { ok: true };
    }
    case 'STORAGE_GET': {
      return chrome.storage.local.get(msg.keys);
    }
    case 'STORAGE_CLEAR': {
      await chrome.storage.local.clear();
      return { ok: true };
    }

    default:
      throw new Error(`Unknown message type: ${msg.type}`);
  }
}

async function execInPL(key, fn, extraArgs = []) {
  const tabs = await chrome.tabs.query({ url: 'https://app.projectionlab.com/*' });
  if (!tabs.length) throw new Error('ProjectionLab is not open. Please open app.projectionlab.com first.');

  let targetTab = tabs[0];
  if (tabs.length > 1) {
    // Sort by lastAccessed descending — use the most recently focused PL tab
    tabs.sort((a, b) => (b.lastAccessed || 0) - (a.lastAccessed || 0));
    targetTab = tabs[0];
    // Notify the popup so it can surface a warning toast
    chrome.runtime.sendMessage({ type: 'WARN_MULTI_TAB', tabTitle: targetTab.title || 'ProjectionLab' }).catch(() => {});
  }

  const [{ result }] = await chrome.scripting.executeScript({
    target: { tabId: targetTab.id },
    world:  'MAIN',
    func: (serialisedFn, key, extraArgs) => {
      return new Promise((resolve) => {
        try {
          const fn = new Function(`return (${serialisedFn})`)();
          fn(key, ...extraArgs).then(resolve).catch(e => resolve({ error: e.message }));
        } catch(e) {
          resolve({ error: e.message });
        }
      });
    },
    args: [fn.toString(), key, extraArgs],
  });

  if (result?.error) {
    const errMsg = result.error;
    // Detect CSP blocking new Function() / eval — signals PL has added a Content Security Policy
    if (errMsg && (errMsg.includes('Content Security Policy') || errMsg.includes('unsafe-eval') || errMsg.toLowerCase().includes('evalerror'))) {
      throw new Error('CSP_BLOCKED: ' + errMsg);
    }
    throw new Error(errMsg);
  }
  return result;
}
