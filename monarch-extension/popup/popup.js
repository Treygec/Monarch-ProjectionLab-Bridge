// ============================================================
// Monarch → ProjectionLab Bridge  —  Popup v4
// ============================================================

const FIREBASE_API_KEY        = 'AIzaSyBlQFWbPQ1bhREni8ZhfKbRsd3lKOPPO0s';
const FIREBASE_FUNCTIONS_BASE = 'https://us-central1-monarch-bridge-prod.cloudfunctions.net';

// ── State ────────────────────────────────────────────────────
const S = {
  monarchAccounts:  [],  // [{name, balance, id, isHidden, type, subtype, ...}]
  plAccounts:       [],  // [{id, name, category, writeField, balance, ...}]
  mappings:         {},  // monarchId → {plId, plRealId, plField}
  pendingCreations: {},  // monarchId → {name, balance, plCategory, plType, monarchId}
  mergeGroups:      [],  // [{id, plId, plRealId, plField, monarchIds}]
  monarchBudgets:   [],  // [{categoryId, categoryName, groupName, monthlyAmount}]
  plExpenses:       [],  // [{planId, planName, expenseId, expenseName, amount, frequency}]
  budgetMappings:   {},  // categoryId → [{planId, expenseId}]
  budgetMergeGroups:          [], // [{id, name, categoryIds, mappings:[{planId,expenseId,createIfMissing}], sourceOverride}]
  budgetSourceOverrides:      {}, // categoryId → {mode, lookback, aggregation}
  budgetGroupSourceOverrides: {}, // groupName  → {mode, lookback, aggregation}
  plApiKey:         '',
  monarchEmail:        '',
  monarchPassword:     '',
  monarchSessionToken: '',
  monarchApiUrl:       'http://localhost:47821',
  connectionMode:      'self-hosted',  // 'self-hosted' | 'hosted'
  firebaseEmail:       '',
  firebaseIdToken:     '',
  firebaseRefreshToken:'',
  firebaseTokenExpiry: 0,
  settings:         { showHidden: false, writeNotes: true, monarchMfaEnabled: false, monarchMfaSecret: '', budgetSource: { mode: 'planned', lookback: 3, aggregation: 'average' } },
};

// ── PL investment type options ────────────────────────────────
const INV_TYPES = [
  { value: 'taxable',   label: 'Taxable Brokerage' },
  { value: '401k',      label: '401(k)' },
  { value: 'roth-401k', label: 'Roth 401(k)' },
  { value: 'roth-ira',  label: 'Roth IRA' },
  { value: 'ira',       label: 'Traditional IRA' },
  { value: '403b',      label: '403(b)' },
  { value: 'hsa',       label: 'HSA' },
  { value: 'crypto',    label: 'Cryptocurrency' },
  { value: '529',       label: '529 Plan' },
];

const ASSET_TYPES = [
  { value: 'real-estate',         label: 'House' },
  { value: 'car',                 label: 'Car' },
  { value: 'rental-property',     label: 'Rental Property' },
  { value: 'land',                label: 'Land' },
  { value: 'building',            label: 'Building' },
  { value: 'commercial-property', label: 'Commercial Property' },
  { value: 'motorcycle',          label: 'Motorcycle' },
  { value: 'boat',                label: 'Boat' },
  { value: 'jewelry',             label: 'Jewelry' },
  { value: 'precious-metals',     label: 'Precious Metals' },
  { value: 'furniture',           label: 'Furniture' },
  { value: 'instrument',          label: 'Instrument' },
  { value: 'machinery',           label: 'Machinery' },
  { value: 'custom',              label: 'Custom Asset' },
];

const DEBT_TYPES = [
  { value: 'debt',             label: 'Debt (Generic)' },
  { value: 'student-loans',    label: 'Student Loans' },
  { value: 'medical-debt',     label: 'Medical Debt' },
  { value: 'credit-card-debt', label: 'Credit Card Debt' },
];

// Asset write-field options — what the Monarch balance represents inside the PL asset record
const ASSET_WRITE_FIELDS = [
  { value: 'amount',       label: 'Current Value' },
  { value: 'balance',      label: 'Loan Balance' },
  { value: 'initialValue', label: 'Purchase Price' },
];

const PL_CATEGORIES = [
  { value: 'savings',    label: 'Savings / Cash' },
  { value: 'investment', label: 'Investment' },
  { value: 'asset',      label: 'Real Asset' },
  { value: 'debt',       label: 'Unsecured Debt' },
];

const CATEGORY_COLORS = {
  savings: 'var(--blue)', investment: 'var(--green)',
  debt: 'var(--red)',     asset: 'var(--amber)',
};
const CATEGORY_LABELS = {
  savings: 'Savings', investment: 'Investment', debt: 'Debt', asset: 'Real Asset',
};

// ── Category inference from Monarch account data ─────────────
function inferCategory(a) {
  const typeName    = a.type?.name    || '';
  const subtypeName = a.subtype?.name || '';
  const nameLower   = (a.name || '').toLowerCase();

  // ── Real Estate ──────────────────────────────────────────────
  // real_estate/primary_home, real_estate/rental, real_estate/* → real-estate asset
  if (typeName === 'real_estate') {
    return { category: 'asset', type: 'real-estate', writeField: 'amount' };
  }

  // ── Valuables ────────────────────────────────────────────────
  // valuables/art, /jewelry, /collectibles, /furniture, /other
  if (typeName === 'valuables') {
    const valuablesMap = {
      'art':          'custom',
      'jewelry':      'jewelry',
      'collectibles': 'custom',
      'furniture':    'furniture',
      'other':        'custom',
    };
    return { category: 'asset', type: valuablesMap[subtypeName] || 'custom', writeField: 'amount' };
  }

  // ── Investments (Brokerage) ───────────────────────────────────
  // All brokerage subtypes seen in real data:
  //   brokerage (taxable), stock_plan (RSU/ESPP), cryptocurrency,
  //   st_401k, st_403b, health_savings_account, roth (Roth IRA),
  //   non_taxable_brokerage_account (needs name disambiguation),
  //   ira (Traditional IRA)
  if (typeName === 'brokerage') {
    const brokerageMap = {
      'brokerage':          'taxable',
      'stock_plan':         'taxable',    // RSUs, ESPPs — taxable brokerage in PL
      'cryptocurrency':     'crypto',
      'st_401k':            '401k',
      'st_403b':            '403b',
      'health_savings_account': 'hsa',
      'roth':               'roth-ira',
      'ira':                'ira',
    };
    if (brokerageMap[subtypeName]) {
      return { category: 'investment', type: brokerageMap[subtypeName] };
    }
    // non_taxable_brokerage_account: disambiguate by account name
    if (subtypeName === 'non_taxable_brokerage_account') {
      if (nameLower.includes('roth'))  return { category: 'investment', type: 'roth-401k' };
      if (nameLower.includes('401k')) return { category: 'investment', type: '401k' };
      return { category: 'investment', type: '401k' }; // safest default
    }
    // Any other brokerage subtype defaults to taxable
    return { category: 'investment', type: 'taxable' };
  }

  // ── Cash / Depository ────────────────────────────────────────
  // depository/checking, /savings, /cash_management → all Savings in PL
  if (typeName === 'depository') {
    return { category: 'savings', type: 'savings' };
  }

  // ── Credit Cards ─────────────────────────────────────────────
  // credit/credit_card, credit/student_loan
  if (typeName === 'credit') {
    if (subtypeName === 'student_loan') return { category: 'debt', type: 'student-loans' };
    return { category: 'debt', type: 'credit-card-debt' };
  }

  // ── Loans ────────────────────────────────────────────────────
  // loan/mortgage → generic debt (user will typically model as asset+loan in PL)
  // loan/student, loan/medical, loan/* → appropriate debt types
  if (typeName === 'loan') {
    if (subtypeName === 'mortgage')  return { category: 'asset', type: 'real-estate', writeField: 'balance' };
    if (subtypeName === 'student')   return { category: 'debt', type: 'student-loans' };
    if (subtypeName === 'medical')   return { category: 'debt', type: 'medical-debt' };
    return { category: 'debt', type: 'debt' };
  }

  // ── Fallback: use Monarch's isAsset flag as the tiebreaker ───
  // Better than blindly defaulting to savings for unknown types
  if (a.isAsset === true)  return { category: 'asset',   type: 'custom',  writeField: 'amount' };
  if (a.isAsset === false) return { category: 'debt',    type: 'debt' };
  return { category: 'savings', type: 'savings' };
}

// ── Boot ─────────────────────────────────────────────────────
// ── Permission helpers ───────────────────────────────────────
function isLocalhostUrl(url) {
  try {
    const u = new URL(url);
    return u.hostname === 'localhost' || u.hostname === '127.0.0.1';
  } catch { return false; }
}

function originFromUrl(url) {
  try { return new URL(url).origin + '/*'; } catch { return null; }
}

async function checkUrlPermission(url) {
  if (!url || isLocalhostUrl(url)) return { needed: false, granted: true };
  const origin = originFromUrl(url);
  if (!origin) return { needed: true, granted: false };
  const granted = await chrome.permissions.contains({ origins: [origin] });
  return { needed: true, granted };
}

async function requestUrlPermission(url) {
  const origin = originFromUrl(url);
  if (!origin) return false;
  return chrome.permissions.request({ origins: [origin] });
}

// Update permission status indicator in the Setup tab
async function updatePermissionUI(url) {
  const permStatus = byId('monarchPermStatus');
  const grantBtn   = byId('grantPermissionBtn');
  const fetchBtn   = byId('fetchMonarchAccounts');
  if (!permStatus || !grantBtn) return;

  const { needed, granted } = await checkUrlPermission(url);
  if (!needed) {
    permStatus.textContent = '';
    permStatus.className = 'status-line';
    grantBtn.style.display = 'none';
    fetchBtn.disabled = false;
  } else if (granted) {
    permStatus.textContent = '✓ Permission granted for this URL';
    permStatus.className = 'status-line ok';
    grantBtn.style.display = 'none';
    fetchBtn.disabled = false;
  } else {
    permStatus.textContent = '⚠ Permission required before fetching from this URL';
    permStatus.className = 'status-line err';
    grantBtn.style.display = '';
    fetchBtn.disabled = true;
  }
}

// Listen for warnings from the service worker
chrome.runtime.onMessage.addListener((msg) => {
  if (msg.type === 'WARN_MULTI_TAB') {
    toast(`⚠ Multiple PL tabs open — using most recent: "${msg.tabTitle}"`, 'err');
  }
});

document.addEventListener('DOMContentLoaded', async () => {
  await loadFromStorage();
  setupTabs();
  setupSetup();
  setupMapping();
  setupMappingListDelegation();
  setupMerge();
  setupMergeSourceListDelegation();
  setupMergeGroupListDelegation();
  setupBudgetMergeBuilderDelegation();
  setupBudgetMapping();
  setupBudgetMappingListDelegation();
  setupSync();
  setupSettings();
  renderAll();
  checkPLStatus();
});

// ── Storage ──────────────────────────────────────────────────
async function loadFromStorage() {
  const d = await msg('STORAGE_GET', {
    keys: ['plApiKey','monarchEmail','monarchPassword','monarchSessionToken','monarchApiUrl',
           'connectionMode','firebaseEmail','firebaseIdToken','firebaseRefreshToken','firebaseTokenExpiry',
           'monarchAccounts','plAccounts',
           'mappings','pendingCreations','mergeGroups','settings',
           'monarchBudgets','plExpenses','budgetMappings','budgetMergeGroups',
           'budgetSourceOverrides','budgetGroupSourceOverrides'],
  });
  if (d.plApiKey)          S.plApiKey          = d.plApiKey;
  if (d.monarchEmail)           S.monarchEmail           = d.monarchEmail;
  if (d.monarchPassword)        S.monarchPassword        = d.monarchPassword;
  if (d.monarchSessionToken)    S.monarchSessionToken    = d.monarchSessionToken;
  if (d.monarchApiUrl)          S.monarchApiUrl          = d.monarchApiUrl;
  if (d.connectionMode)         S.connectionMode         = d.connectionMode;
  if (d.firebaseEmail)          S.firebaseEmail          = d.firebaseEmail;
  if (d.firebaseIdToken)        S.firebaseIdToken        = d.firebaseIdToken;
  if (d.firebaseRefreshToken)   S.firebaseRefreshToken   = d.firebaseRefreshToken;
  if (d.firebaseTokenExpiry)    S.firebaseTokenExpiry    = d.firebaseTokenExpiry;
  if (d.monarchAccounts)   S.monarchAccounts   = d.monarchAccounts;
  if (d.plAccounts)        S.plAccounts        = d.plAccounts;
  if (d.mappings)          S.mappings          = d.mappings;
  if (d.pendingCreations)  S.pendingCreations  = d.pendingCreations;
  if (d.mergeGroups)       S.mergeGroups       = d.mergeGroups;
  if (d.settings)          S.settings          = { ...S.settings, ...d.settings };
  if (d.monarchBudgets)    S.monarchBudgets    = d.monarchBudgets;
  if (d.plExpenses)        S.plExpenses        = d.plExpenses;
  if (d.budgetMappings)    S.budgetMappings    = d.budgetMappings;
  if (d.budgetMergeGroups)     S.budgetMergeGroups     = d.budgetMergeGroups;
  if (d.budgetSourceOverrides)      S.budgetSourceOverrides      = sanitizeBudgetSourceOverrides(d.budgetSourceOverrides);
  if (d.budgetGroupSourceOverrides) S.budgetGroupSourceOverrides = sanitizeBudgetSourceOverrides(d.budgetGroupSourceOverrides);

  byId('plApiKey').value         = S.plApiKey;
  byId('monarchEmail').value     = S.monarchEmail;
  byId('monarchPassword').value  = S.monarchPassword;
  byId('monarchApiUrl').value    = S.monarchApiUrl;
  // Check permission status for the stored URL on load
  if (S.monarchApiUrl) updatePermissionUI(S.monarchApiUrl);
  applyConnectionModeUI();
}

async function save() {
  await msg('STORAGE_SET', { data: {
    plApiKey: S.plApiKey, monarchEmail: S.monarchEmail, monarchPassword: S.monarchPassword, monarchSessionToken: S.monarchSessionToken, monarchApiUrl: S.monarchApiUrl,
    connectionMode: S.connectionMode,
    monarchAccounts: S.monarchAccounts, plAccounts: S.plAccounts,
    mappings: S.mappings, pendingCreations: S.pendingCreations,
    mergeGroups: S.mergeGroups, settings: S.settings,
    monarchBudgets: S.monarchBudgets, plExpenses: S.plExpenses,
    budgetMappings: S.budgetMappings,
    budgetMergeGroups: S.budgetMergeGroups,
    budgetSourceOverrides: S.budgetSourceOverrides,
    budgetGroupSourceOverrides: S.budgetGroupSourceOverrides,
  }});
}

// ── Connection mode UI ───────────────────────────────────────
function applyConnectionModeUI() {
  const hosted = S.connectionMode === 'hosted';
  byId('monarchUrlSection').classList.toggle('hidden', hosted);

  // Show donate chip for self-hosted, account chip for hosted
  byId('donateChip').classList.toggle('hidden', hosted);
  byId('accountChipWrap').classList.toggle('hidden', !hosted);

  // Show/hide the inline banner on Setup tab
  const banner = byId('firebaseAuthBanner');
  banner.classList.toggle('hidden', !hosted);

  const signedIn = hosted && !!S.firebaseEmail;

  // Update header chip state
  const chip = byId('accountChip');
  const chipLabel = byId('accountChipLabel');
  if (signedIn) {
    chip.classList.add('signed-in');
    chipLabel.textContent = S.firebaseEmail.split('@')[0];
  } else {
    chip.classList.remove('signed-in');
    chipLabel.textContent = 'Account';
  }

  // Update dropdown panels
  byId('accountDropdownSignedOut').classList.toggle('hidden', signedIn);
  byId('accountDropdownSignedIn').classList.toggle('hidden', !signedIn);
  if (signedIn) {
    byId('acctDdCurrentEmail').textContent = S.firebaseEmail;
    banner.classList.add('signed-in');
    byId('firebaseAuthBannerText').innerHTML =
      `Signed in as <strong>${esc(S.firebaseEmail)}</strong> — manage your account from the header.`;
  } else {
    banner.classList.remove('signed-in');
    byId('firebaseAuthBannerText').innerHTML =
      'Hosted mode — sign in to your Monarch Bridge account using the <strong>Account</strong> button in the header.';
  }

  // Update Monarch credentials hint based on mode
  byId('monarchCredentialsHint').textContent = hosted
    ? 'Enter your Monarch Money credentials. These are stored locally in your browser and sent directly to Monarch — never to our servers.'
    : 'Enter your Monarch Money credentials. These are stored locally in your browser and sent only to your proxy — never to our servers.';

  // Keep the hosted mode toggle in sync (may be called before setupSettings runs)
  const cb = byId('settingHostedMode');
  if (cb) cb.checked = hosted;
}

// ── Tab system ───────────────────────────────────────────────
function setupTabs() {
  document.querySelectorAll('.tab').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      byId(`panel-${btn.dataset.tab}`).classList.add('active');
      if (btn.dataset.tab === 'sync') renderSync();
    });
  });

  // Sub-tab wiring inside Mapping panel
  document.querySelectorAll('.sub-tab').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.sub-tab').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.sub-panel').forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      byId(`sub-panel-${btn.dataset.sub}`).classList.add('active');
      if (btn.dataset.sub === 'budgets') renderBudgetMapping();
    });
  });
}

// ── PL status indicator ──────────────────────────────────────
async function checkPLStatus() {
  const dot = byId('plStatus');
  try {
    const tabs = await chrome.tabs.query({ url: 'https://app.projectionlab.com/*' });
    dot.className = 'status-dot ' + (tabs.length ? 'ready' : 'error');
  } catch { dot.className = 'status-dot error'; }
}

// ── SETUP tab ────────────────────────────────────────────────

function isCspError(msg) {
  if (!msg) return false;
  const m = msg.toLowerCase();
  return msg.startsWith('CSP_BLOCKED:')
    || m.includes('content security policy')
    || m.includes('unsafe-eval')
    || m.includes('evalerror');
}

function showCspWarning() {
  // Replace the PL section with a prominent warning that explains the situation
  const existingWarning = byId('cspWarning');
  if (existingWarning) return; // already shown

  const warning = document.createElement('div');
  warning.id = 'cspWarning';
  warning.className = 'csp-warning';
  warning.innerHTML = `
    <div class="csp-warning-icon">⚠</div>
    <div class="csp-warning-body">
      <div class="csp-warning-title">ProjectionLab API Access Blocked</div>
      <div class="csp-warning-msg">
        ProjectionLab appears to have added a Content Security Policy that prevents
        this extension from communicating with it. This is a compatibility issue that
        requires an extension update.
      </div>
      <div class="csp-warning-msg" style="margin-top:6px">
        Please <a href="https://github.com" target="_blank" style="color:var(--amber)">open a GitHub issue</a>
        or check for an updated version of this extension.
      </div>
    </div>
  `;

  // Insert after the PL API key section
  const plSection = byId('fetchPlAccounts')?.closest('.section');
  if (plSection) plSection.after(warning);
}

function setupSetup() {
  byId('toggleApiKey').addEventListener('click', () => {
    const f = byId('plApiKey');
    f.type = f.type === 'password' ? 'text' : 'password';
  });

  byId('fetchPlAccounts').addEventListener('click', async () => {
    const key = byId('plApiKey').value.trim();
    if (!key) { toast('Enter your PL API key first', 'err'); return; }
    // Save key before fetching so it's persisted even if fetch fails
    S.plApiKey = key;
    await save();
    setLoading('fetchPlAccounts', true);
    status('plFetchStatus', 'Connecting to ProjectionLab…', 'loading');
    try {
      const data = await msg('PL_EXPORT_DATA', { key: S.plApiKey });
      if (data?.error) throw new Error(data.error);
      S.plAccounts = extractPLAccounts(data);
      S.plExpenses = extractPLExpenses(data);
      await save();
      const counts = S.plAccounts.reduce((acc, a) => { acc[a.category]=(acc[a.category]||0)+1; return acc; }, {});
      const summary = Object.entries(counts).map(([k,v]) => `${v} ${k}`).join(', ');
      const planCount = new Set(S.plExpenses.map(e => e.planId)).size;
      status('plFetchStatus', `✓ ${planCount} plan${planCount !== 1 ? 's' : ''}, ${S.plAccounts.length} accounts: ${summary}`, 'ok');
      byId('plStatus').className = 'status-dot ready';
      detectStaleMappings();
      renderMapping();
    } catch (e) {
      if (isCspError(e.message)) {
        byId('plStatus').className = 'status-dot error';
        showCspWarning();
      }
      status('plFetchStatus', `✗ ${e.message}`, 'err');
    }
    finally { setLoading('fetchPlAccounts', false); }
  });

  byId('clearStaleMappings').addEventListener('click', async () => {
    const removed = clearStaleMappings();
    await save();
    byId('staleWarningRow').classList.add('hidden');
    renderMapping(); renderMerge(); renderBudgetMapping();
    toast(`Cleared ${removed} stale reference${removed !== 1 ? 's' : ''}`, 'ok');
  });

  // Monarch credentials
  byId('monarchEmail').addEventListener('change', async () => {
    S.monarchEmail = byId('monarchEmail').value.trim();
    await save();
  });
  byId('monarchPassword').addEventListener('change', async () => {
    S.monarchPassword = byId('monarchPassword').value;
    await save();
  });
  byId('toggleMonarchPassword').addEventListener('click', () => {
    const f = byId('monarchPassword');
    f.type = f.type === 'password' ? 'text' : 'password';
  });

  // Monarch API URL permission check
  byId('monarchApiUrl').addEventListener('input', () => {
    const url = byId('monarchApiUrl').value.trim();
    if (url) updatePermissionUI(url);
  });

  byId('grantPermissionBtn').addEventListener('click', async () => {
    const url = byId('monarchApiUrl').value.trim();
    setLoading('grantPermissionBtn', true);
    try {
      const granted = await requestUrlPermission(url);
      toast(granted ? 'Permission granted ✓' : 'Permission denied — cannot fetch from this URL', granted ? 'ok' : 'err');
      await updatePermissionUI(url);
    } catch (e) {
      toast(e.message, 'err');
    } finally { setLoading('grantPermissionBtn', false); }
  });

  // Shared: load accounts + budgets after session is established
  async function loadMonarchData() {
    const data = await msg('FETCH_MONARCH_ACCOUNTS');
    if (data?.error) throw new Error(data.error);
    S.monarchAccounts = data.accounts;
    try {
      const budgetData = await msg('FETCH_MONARCH_BUDGETS');
      if (!budgetData?.error && Array.isArray(budgetData?.budgets)) {
        S.monarchBudgets = budgetData.budgets.map(b => {
          const rawMonths = b.monthlyAmounts || [];
          const latest = rawMonths[rawMonths.length - 1] || {};
          return { categoryId: b.categoryId, categoryName: b.categoryName, groupName: b.groupName || '', monthlyAmounts: rawMonths, monthlyAmount: latest.plannedCashFlowAmount ?? 0 };
        }).filter(b => b.monthlyAmount !== 0 || true);
        const validCategoryIds = new Set(S.monarchBudgets.map(b => b.categoryId));
        let budgetPruned = 0;
        for (const cid of Object.keys(S.budgetMappings)) {
          if (!validCategoryIds.has(cid)) { delete S.budgetMappings[cid]; budgetPruned++; }
        }
        S.budgetMergeGroups = S.budgetMergeGroups.filter(g => {
          const before = g.categoryIds.length;
          g.categoryIds = g.categoryIds.filter(cid => validCategoryIds.has(cid));
          budgetPruned += before - g.categoryIds.length;
          return g.categoryIds.length >= 1;
        });
        if (budgetPruned > 0) console.info(`[Monarch Bridge] Pruned ${budgetPruned} stale budget category reference(s)`);
      }
    } catch (_) { /* non-fatal */ }
    const currentIds = new Set(S.monarchAccounts.map(a => a.id));
    let pruned = 0;
    for (const id of Object.keys(S.pendingCreations)) { if (!currentIds.has(id)) { delete S.pendingCreations[id]; pruned++; } }
    for (const id of Object.keys(S.mappings)) { if (!currentIds.has(id)) { delete S.mappings[id]; pruned++; } }
    for (const g of S.mergeGroups) {
      const before = g.monarchIds.length;
      g.monarchIds = g.monarchIds.filter(id => currentIds.has(id));
      if (g.fieldMappings) g.fieldMappings = g.fieldMappings.filter(fm => currentIds.has(fm.monarchId));
      pruned += before - g.monarchIds.length;
    }
    await save();
    const pruneNote = pruned ? ` (${pruned} stale reference${pruned>1?'s':''} removed)` : '';
    const catNote = S.monarchBudgets.length ? `, ${S.monarchBudgets.length} budget categor${S.monarchBudgets.length===1?'y':'ies'}` : '';
    status('monarchFetchStatus', `✓ ${S.monarchAccounts.length} accounts${catNote} loaded${pruneNote}`, 'ok');
    byId('monarchStatus').className = 'status-dot ready';
    renderMapping();
    if (pruned) renderMerge();
  }

  byId('fetchMonarchAccounts').addEventListener('click', async () => {
    S.monarchApiUrl = (byId('monarchApiUrl').value.trim() || 'http://localhost:47821').replace(/\/accounts$/, '').replace(/\/$/, '');
    const { needed, granted } = await checkUrlPermission(S.monarchApiUrl);
    if (needed && !granted) {
      status('monarchFetchStatus', '✗ Permission not granted for this URL — click Grant Permission first', 'err');
      return;
    }
    if (!S.monarchEmail || !S.monarchPassword) {
      status('monarchFetchStatus', '✗ Enter your Monarch email and password above', 'err');
      return;
    }
    byId('mfaCodeRow').classList.add('hidden');
    setLoading('fetchMonarchAccounts', true);
    status('monarchFetchStatus', 'Connecting to Monarch…', 'loading');
    try {
      // In hosted mode, skip login if we already have a session token
      const hosted = S.connectionMode === 'hosted';
      if (hosted && S.monarchSessionToken) {
        status('monarchFetchStatus', 'Loading data…', 'loading');
      } else {
        const auth = await msg('MONARCH_INITIATE_LOGIN');
        if (auth?.error) throw new Error(auth.error);
        if (auth.requires_mfa) {
          status('monarchFetchStatus', 'Check your authenticator app and enter the 6-digit code below', 'pending');
          byId('mfaCodeRow').classList.remove('hidden');
          byId('monarchMfaCode').value = '';
          byId('monarchMfaCode').focus();
          return;
        }
        S.monarchSessionToken = auth.token;
        // In hosted mode, clear password from storage now that we have a session token
        if (hosted) { S.monarchPassword = ''; await msg('STORAGE_SET', { data: { monarchPassword: '' } }); }
        await save();
        status('monarchFetchStatus', 'Authenticated — loading data…', 'loading');
      }
      await loadMonarchData();
    } catch (e) {
      const expired = e.message?.includes('TOKEN_EXPIRED');
      if (expired) { S.monarchSessionToken = ''; await save(); }
      status('monarchFetchStatus', `✗ ${expired ? 'Session expired — re-authenticate' : e.message}`, 'err');
      byId('monarchStatus').className = 'status-dot error';
    } finally { setLoading('fetchMonarchAccounts', false); }
  });

  byId('monarchMfaCode').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); byId('verifyMfaCode').click(); }
  });

  byId('verifyMfaCode').addEventListener('click', async () => {
    const code = byId('monarchMfaCode').value.replace(/\s/g, '');
    if (!code || code.length < 6) { status('mfaVerifyStatus', '✗ Enter the 6-digit code', 'err'); return; }
    setLoading('verifyMfaCode', true);
    status('mfaVerifyStatus', 'Verifying…', 'loading');
    try {
      const result = await msg('MONARCH_COMPLETE_MFA', { code });
      if (result?.error) throw new Error(result.error);
      S.monarchSessionToken = result.token;
      // In hosted mode, clear password from storage now that we have a session token
      if (S.connectionMode === 'hosted') { S.monarchPassword = ''; await msg('STORAGE_SET', { data: { monarchPassword: '' } }); }
      await save();
      byId('mfaCodeRow').classList.add('hidden');
      status('mfaVerifyStatus', '', '');
      status('monarchFetchStatus', 'MFA verified — loading data…', 'loading');
      await loadMonarchData();
    } catch (e) {
      status('mfaVerifyStatus', `✗ ${e.message}`, 'err');
    } finally { setLoading('verifyMfaCode', false); }
  });

  // ── Donate button (self-hosted mode) ────────────────────────
  byId('donateChip').addEventListener('click', () => {
    window.open('https://buy.stripe.com/7sY28s6SXcM176D0JSfbq00', '_blank');
  });

  // ── Account dropdown (header) ──────────────────────────────
  const _chipBtn = byId('accountChip');
  const _dropdown = byId('accountDropdown');
  _chipBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    _dropdown.classList.toggle('hidden');
  });
  _dropdown.addEventListener('click', (e) => {
    e.stopPropagation(); // don't close when clicking inside dropdown
  });
  document.addEventListener('click', () => {
    _dropdown.classList.add('hidden');
  });

  byId('acctDdTogglePw').addEventListener('click', () => {
    const f = byId('acctDdPassword');
    f.type = f.type === 'password' ? 'text' : 'password';
  });

  async function doFirebaseAuth(type) {
    const email    = byId('acctDdEmail').value.trim();
    const password = byId('acctDdPassword').value;
    if (!email || !password) { status('acctDdStatus', '✗ Enter email and password', 'err'); return; }
    const btnId = type === 'FIREBASE_SIGN_UP' ? 'acctDdSignUp' : 'acctDdSignIn';
    setLoading(btnId, true);
    status('acctDdStatus', type === 'FIREBASE_SIGN_UP' ? 'Creating account…' : 'Signing in…', 'loading');
    try {
      const result = await msg(type, { email, password });
      if (result?.error) throw new Error(result.error);
      S.firebaseEmail = result.email;
      await save();
      status('acctDdStatus', '', '');
      applyConnectionModeUI();
      document.dispatchEvent(new Event('firebase-auth-changed'));
      toast(`Signed in as ${result.email}`, 'ok');
    } catch (e) {
      status('acctDdStatus', `✗ ${e.message}`, 'err');
    } finally { setLoading(btnId, false); }
  }

  byId('acctDdSignIn').addEventListener('click', () => doFirebaseAuth('FIREBASE_SIGN_IN'));
  byId('acctDdSignUp').addEventListener('click', () => doFirebaseAuth('FIREBASE_SIGN_UP'));

  // Enter key to submit on email/password fields
  byId('acctDdEmail').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); byId('acctDdSignIn').click(); }
  });
  byId('acctDdPassword').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); byId('acctDdSignIn').click(); }
  });

  byId('acctDdSignOut').addEventListener('click', async () => {
    await msg('FIREBASE_SIGN_OUT');
    S.firebaseEmail = '';
    S.firebaseIdToken = '';
    S.firebaseRefreshToken = '';
    S.firebaseTokenExpiry = 0;
    await save();
    applyConnectionModeUI();
    document.dispatchEvent(new Event('firebase-auth-changed'));
    byId('acctDdEmail').value = '';
    byId('acctDdPassword').value = '';
    byId('acctDdCurrentEmail').textContent = '';
    byId('accountDropdown').classList.add('hidden');
    toast('Signed out', 'ok');
  });

  byId('acctDdBilling').addEventListener('click', async () => {
    setLoading('acctDdBilling', true);
    try {
      await msg('FIREBASE_OPEN_BILLING');
    } catch (e) {
      toast(e.message, 'err');
    } finally { setLoading('acctDdBilling', false); }
  });

  byId('debugExportBtn').addEventListener('click', async () => {
    if (!S.plApiKey) { toast('Enter your PL API key first', 'err'); return; }
    if (!S.plAccounts.length) { toast('Fetch from PL first', 'err'); return; }
    setLoading('debugExportBtn', true);
    try {
      const data = await msg('PL_EXPORT_DATA', { key: S.plApiKey });
      if (data?.error) throw new Error(data.error);
      const blob = new Blob([JSON.stringify(data, null, 2)], {type:'application/json'});
      window.open(URL.createObjectURL(blob), '_blank');
    } catch (e) { toast(e.message, 'err'); }
    finally { setLoading('debugExportBtn', false); }
  });
}

// ── PL Account Extractor ─────────────────────────────────────
// Returns a count of stale references and shows/hides the warning row.
function detectStaleMappings() {
  const validPlRealIds = new Set(S.plAccounts.map(a => a.id.replace(/__balance$|__initialValue$/, '')));
  const validExpenseKeys = new Set(S.plExpenses.map(e => e.planId + '::' + e.expenseId));

  let stale = 0;

  // Account mappings
  for (const [, m] of Object.entries(S.mappings)) {
    if (m?.plRealId && !validPlRealIds.has(m.plRealId)) stale++;
  }
  // Merge groups (non-pending)
  for (const g of S.mergeGroups) {
    if (g.plRealId && !g.pendingCreate && !validPlRealIds.has(g.plRealId)) stale++;
  }
  // Budget mappings
  for (const targets of Object.values(S.budgetMappings)) {
    for (const t of targets) {
      if (t.expenseId !== '__CREATE_NEW__' && !validExpenseKeys.has(t.planId + '::' + t.expenseId)) stale++;
    }
  }
  // Budget merge group mappings
  for (const g of S.budgetMergeGroups) {
    for (const t of g.mappings || []) {
      if (t.expenseId !== '__CREATE_NEW__' && !validExpenseKeys.has(t.planId + '::' + t.expenseId)) stale++;
    }
  }

  const row = byId('staleWarningRow');
  if (stale > 0) {
    byId('staleWarningText').textContent = `⚠ ${stale} stale mapping${stale !== 1 ? 's' : ''} detected`;
    row.classList.remove('hidden');
  } else {
    row.classList.add('hidden');
  }
  return stale;
}

// Removes all mappings pointing at PL IDs that no longer exist. Returns count removed.
function clearStaleMappings() {
  const validPlRealIds = new Set(S.plAccounts.map(a => a.id.replace(/__balance$|__initialValue$/, '')));
  const validExpenseKeys = new Set(S.plExpenses.map(e => e.planId + '::' + e.expenseId));
  let removed = 0;

  // Account mappings
  for (const monarchId of Object.keys(S.mappings)) {
    const m = S.mappings[monarchId];
    if (m?.plRealId && !validPlRealIds.has(m.plRealId)) {
      delete S.mappings[monarchId]; removed++;
    }
  }
  // Merge groups
  for (const g of S.mergeGroups) {
    if (g.plRealId && !g.pendingCreate && !validPlRealIds.has(g.plRealId)) {
      g.plId = null; g.plRealId = null; g.plField = null; removed++;
    }
  }
  // Budget mappings
  for (const [cid, targets] of Object.entries(S.budgetMappings)) {
    const before = targets.length;
    S.budgetMappings[cid] = targets.filter(t =>
      t.expenseId === '__CREATE_NEW__' || validExpenseKeys.has(t.planId + '::' + t.expenseId)
    );
    removed += before - S.budgetMappings[cid].length;
    if (!S.budgetMappings[cid].length) delete S.budgetMappings[cid];
  }
  // Budget merge group mappings
  for (const g of S.budgetMergeGroups) {
    const before = (g.mappings || []).length;
    g.mappings = (g.mappings || []).filter(t =>
      t.expenseId === '__CREATE_NEW__' || validExpenseKeys.has(t.planId + '::' + t.expenseId)
    );
    removed += before - g.mappings.length;
  }

  return removed;
}

function extractPLAccounts(exportData) {
  const today = exportData?.today ?? exportData?.currentFinances ?? exportData;
  if (!today) return [];
  const results = [];
  for (const a of (today.savingsAccounts    || [])) {
    if (!a.id || !a.name) continue;
    results.push({ id: String(a.id), name: a.name.trim(), category: 'savings',    writeField: 'balance', balance: a.balance ?? 0 });
  }
  for (const a of (today.investmentAccounts || [])) {
    if (!a.id || !a.name) continue;
    results.push({ id: String(a.id), name: a.name.trim(), category: 'investment', writeField: 'balance', balance: a.balance ?? 0 });
  }
  for (const a of (today.debts              || [])) {
    if (!a.id || !a.name) continue;
    results.push({ id: String(a.id), name: a.name.trim(), category: 'debt',       writeField: 'amount',  balance: a.amount  ?? 0 });
  }
  for (const a of (today.assets             || [])) {
    if (!a.id || !a.name) continue;
    const base = { id: String(a.id), name: a.name.trim(), category: 'asset', balance: a.amount ?? 0, amount: a.amount ?? 0, initialValue: a.initialValue ?? 0 };
    results.push({ ...base, writeField: 'amount',  subLabel: 'Current Value', displayBal: base.amount });
    results.push({ ...base, writeField: 'balance', subLabel: 'Loan Balance',  displayBal: base.balance, id: base.id+'__balance' });
  }
  return results.sort((a, b) => {
    const o = { savings:0, investment:1, debt:2, asset:3 };
    const co = (o[a.category]??9) - (o[b.category]??9);
    return co !== 0 ? co : a.name.localeCompare(b.name);
  });
}

// ── PL Plan Expense Extractor ────────────────────────────────
function extractPLExpenses(exportData) {
  const plans = exportData?.plans || [];
  const results = [];
  for (const plan of plans) {
    const events = plan.expenses?.events || [];
    for (const e of events) {
      if (!e.id || !e.name) continue;
      // Skip debt-linked events (they have debtId) — those are synced via account mapping
      if (e.debtId) continue;
      results.push({
        planId:         plan.id,
        planName:       plan.name || plan.id,
        expenseId:      e.id,
        expenseName:    e.name,
        amount:         e.amount ?? 0,
        frequency:      e.frequency || 'monthly',
        expType:        e.type || 'living-expenses',
        monthlyPayment: e.monthlyPayment ?? null, // debt type only
        key:            plan.id + '::' + e.id,
      });
    }
  }
  // Sort by plan name then expense name
  return results.sort((a, b) =>
    a.planName.localeCompare(b.planName) || a.expenseName.localeCompare(b.expenseName)
  );
}

// Compute the correct PL write field and sync amount for a budget→expense mapping.
// Different PL expense types store their "payment amount" differently:
//   debt:     monthlyPayment field (not amount, which is the balance)
//   once:     amount, but a monthly budget can't meaningfully map to a one-time expense
//   yearly:   amount, converted monthly→yearly
//   monthly:  amount, 1:1
function computeBudgetSync(expDef, monarchMonthlyAmount) {
  if (!expDef) return { writeField: 'amount', syncAmount: monarchMonthlyAmount, freqLabel: '/mo', warn: null };

  const freq    = expDef.frequency || 'monthly';
  const expType = expDef.expType   || 'living-expenses';

  // Debt: budget maps to monthlyPayment, not the balance (amount field)
  if (expType === 'debt') {
    return { writeField: 'monthlyPayment', syncAmount: monarchMonthlyAmount, freqLabel: '/mo', warn: null };
  }

  // One-time expenses: no meaningful periodic conversion — warn, write as-is
  if (freq === 'once') {
    return {
      writeField: 'amount', syncAmount: monarchMonthlyAmount, freqLabel: '',
      warn: 'One-time expense — Monarch monthly amount written as-is. Consider setting this manually in PL.',
    };
  }

  // Find the matching PL freq option and apply its conversion
  const opt = PL_FREQ_OPTIONS.find(o => o.value === freq);
  if (opt) {
    const syncAmount = opt.convert(monarchMonthlyAmount);
    const freqLabel  = `/${(freq === 'yearly' || freq === 'yearly-lump-sum') ? 'yr' : freq} (${opt.hint})`;
    return { writeField: 'amount', syncAmount, freqLabel, warn: null };
  }

  // Fallback: monthly 1:1
  return { writeField: 'amount', syncAmount: monarchMonthlyAmount, freqLabel: '/mo', warn: null };
}

// Resolve the monthly amount for a budget category.
// override: per-item {mode,lookback,aggregation} — if null/undefined, falls back to global settings.
// Planned mode: defaults to most-recent month; supports lookback + aggregation over plannedCashFlowAmount.
// Actual mode:  supports lookback + aggregation over actualAmount (abs value — outflows are negative).
const VALID_BUDGET_MODES = new Set(['', 'planned', 'actual']);
const VALID_AGGREGATIONS = new Set(['average', 'median']);

// Sanitize a single override object, returning null if it's not a valid object.
function sanitizeBudgetSourceOverride(obj) {
  if (!obj || typeof obj !== 'object') return null;
  const mode = VALID_BUDGET_MODES.has(obj.mode) ? obj.mode : '';
  const rawLookback = Number(obj.lookback);
  const lookback = Number.isFinite(rawLookback) ? Math.max(1, Math.min(12, Math.round(rawLookback))) : (mode === 'actual' ? 3 : 1);
  const aggregation = VALID_AGGREGATIONS.has(obj.aggregation) ? obj.aggregation : 'average';
  return { mode, lookback, aggregation };
}

// Sanitize the full budgetSourceOverrides map loaded from storage.
function sanitizeBudgetSourceOverrides(raw) {
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return {};
  const out = {};
  for (const [cid, val] of Object.entries(raw)) {
    const sanitized = sanitizeBudgetSourceOverride(val);
    if (sanitized) out[cid] = sanitized;
  }
  return out;
}

function resolveMonarchAmount(b, override = null) {
  const src = override != null ? override : (S.settings.budgetSource || {});
  const mode = VALID_BUDGET_MODES.has(src.mode) ? (src.mode || 'planned') : 'planned';
  const months = b.monthlyAmounts || [];

  if (!months.length) return b.monthlyAmount ?? 0;

  if (mode === 'planned') {
    // Default: most-recent month only (lookback = 1). User can override.
    const rawLookback = Number(src.lookback);
    const lookback = Math.max(1, Math.min(Number.isFinite(rawLookback) ? rawLookback : 1, months.length));
    const recent = months.slice(-lookback);
    const values = recent.map(m => m.plannedCashFlowAmount ?? 0);
    if (!values.length) return b.monthlyAmount ?? 0;
    const agg = VALID_AGGREGATIONS.has(src.aggregation) ? src.aggregation : 'average';
    if (agg === 'median') {
      const sorted = [...values].sort((a, bv) => a - bv);
      const mid = Math.floor(sorted.length / 2);
      return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
    }
    return values.reduce((s, v) => s + v, 0) / values.length;
  }

  // Actual mode: actualAmount is typically negative (cash outflow), take abs value
  const rawLookback = Number(src.lookback);
  const lookback = Math.max(1, Math.min(Number.isFinite(rawLookback) ? rawLookback : 3, months.length));
  const recent = months.slice(-lookback);
  const values = recent
    .map(m => (m.actualAmount != null ? Math.abs(m.actualAmount) : null))
    .filter(v => v !== null);

  if (!values.length) return b.monthlyAmount ?? 0;

  const agg = VALID_AGGREGATIONS.has(src.aggregation) ? src.aggregation : 'average';
  if (agg === 'median') {
    const sorted = [...values].sort((a, bv) => a - bv);
    const mid = Math.floor(sorted.length / 2);
    return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
  }
  return values.reduce((s, v) => s + v, 0) / values.length;
}

// Short human-readable label describing the resolved source.
// override: per-item override or null to use global settings.
function budgetSourceLabel(override = null) {
  const src = override != null ? override : (S.settings.budgetSource || {});
  const mode = src.mode || 'planned';
  const defaultLookback = mode === 'actual' ? 3 : 1;
  const lookback = src.lookback || defaultLookback;
  const agg = src.aggregation || 'average';
  const aggShort = { average: 'avg', median: 'med' }[agg] || agg;
  if (mode === 'planned') {
    return lookback === 1 ? 'planned (latest)' : `${aggShort} ${lookback}mo planned`;
  }
  return `${aggShort} ${lookback}mo actual`;
}

// Resolve the effective source override for one category within a merge group.
// Respects group.sourceModeAll: true/undefined = use group-level override; false = per-category.
function getGroupCategoryOverride(g, categoryId) {
  if (g.sourceModeAll === false) return S.budgetSourceOverrides[categoryId] || null;
  return g.sourceOverride || null;
}

// Resolve the effective source override for an individual budget category row.
// Priority: group-level override → per-category override → null (global default).
function getEffectiveBudgetOverride(categoryId) {
  const b = S.monarchBudgets.find(x => x.categoryId === categoryId);
  const groupName = b?.groupName || 'Other';
  return S.budgetGroupSourceOverrides[groupName] || S.budgetSourceOverrides[categoryId] || null;
}

function resolvePlMapping(plId) {
  if (plId.endsWith('__balance'))      return { realId: plId.replace('__balance',''),       field: 'balance' };
  if (plId.endsWith('__initialValue')) return { realId: plId.replace('__initialValue',''), field: 'initialValue' };
  const acct = S.plAccounts.find(a => a.id === plId);
  // Debts use 'amount'; everything else uses 'balance'. Never default to 'balance' for debts.
  const field = acct ? acct.writeField : 'balance';
  return { realId: plId, field };
}

// Re-resolve the correct write field for a saved mapping against current S.plAccounts.
// Corrects stale mappings that were saved with the wrong plField (e.g. 'balance' for a debt).
function correctPlField(plId, storedField) {
  if (plId.endsWith('__balance'))      return 'balance';
  if (plId.endsWith('__initialValue')) return 'initialValue';
  const acct = S.plAccounts.find(a => a.id === plId);
  if (acct) return acct.writeField;  // always trust live PL account data over stored value
  return storedField;                // fall back to stored if PL not loaded
}

// ── Visible accounts (respects showHidden setting) ────────────
function visibleMonarchAccounts() {
  if (S.settings.showHidden) return S.monarchAccounts;
  return S.monarchAccounts.filter(a => a.isHidden !== true);
}

// ── MAPPING tab ──────────────────────────────────────────────
function setupMapping() {
  byId('mappingSearch').addEventListener('input', renderMapping);
  byId('autoMapBtn').addEventListener('click', autoMap);

  async function doSave() {
    await save();
    const status = byId('mappingSaveStatus');
    if (status) { status.textContent = '✓ Saved'; setTimeout(() => { status.textContent = ''; }, 3500); }
    toast('Mappings saved!', 'ok');
  }

  byId('saveMappings').addEventListener('click', doSave);
  byId('floatingSaveBtn').addEventListener('click', doSave);
  byId('createAllBtn').addEventListener('click', openBulkCreatePanel);

  // Show floating save button when the header-row save button scrolls out of view
  const legend = document.querySelector('#sub-panel-accounts .mapping-legend');
  const floatBar = byId('floatingSaveBar');
  document.addEventListener('scroll', () => {
    if (!floatBar || !legend) return;
    floatBar.classList.toggle('hidden', legend.getBoundingClientRect().bottom > 0);
  }, { passive: true });
}

function buildPlOptionHTML(selectedId, showAssetSubtypes = true) {
  let html = '<option value="">— unmapped —</option>';
  html += `<option value="__CREATE_NEW__" class="create-opt">✦ Create new PL account…</option>`;

  if (!S.plAccounts.length) return html;

  const groups = {};
  for (const p of S.plAccounts) {
    if (!groups[p.category]) groups[p.category] = [];
    if (p.category === 'asset') {
      if (showAssetSubtypes) {
        // Mapping page: show Current Value + Loan Balance, skip Purchase Price
        if (p.writeField === 'initialValue') continue;
      } else {
        // Merge page: only the primary entry (no sub-entries at all)
        if (p.writeField !== 'amount') continue;
      }
    }
    groups[p.category].push(p);
  }
  for (const [cat, items] of Object.entries(groups)) {
    html += `<optgroup label="── ${CATEGORY_LABELS[cat] || cat} ──">`;
    for (const p of items) {
      const label = (showAssetSubtypes && p.subLabel) ? `${esc(p.name)} → ${p.subLabel}` : esc(p.name);
      html += `<option value="${esc(p.id)}" ${p.id === selectedId ? 'selected' : ''}>${label}</option>`;
    }
    html += '</optgroup>';
  }
  return html;
}

function updateCreateAllBtn() {
  const mergedIds = new Set(S.mergeGroups.flatMap(g => g.monarchIds));
  const unmappedCount = visibleMonarchAccounts().filter(a =>
    !S.mappings[a.id] && !mergedIds.has(a.id) && !S.pendingCreations[a.id]
  ).length;
  const btn = byId('createAllBtn');
  btn.textContent = `✦ Create all unmapped (${unmappedCount})`;
  btn.disabled = !unmappedCount;
}

function renderMapping() {
  const q     = byId('mappingSearch').value.toLowerCase();
  const list  = byId('mappingList');
  const mergedIds = new Set(S.mergeGroups.flatMap(g => g.monarchIds));

  const filtered = visibleMonarchAccounts().filter(a =>
    !q || a.name.toLowerCase().includes(q)
  );

  if (!filtered.length) {
    list.innerHTML = '<div class="empty-state">No accounts — fetch from Setup tab</div>';
    return;
  }

  updateCreateAllBtn();

  list.innerHTML = filtered.map(a => {
    const isMerged  = mergedIds.has(a.id);
    const creation  = S.pendingCreations[a.id];
    const curMapping = S.mappings[a.id] || {};
    const plId      = curMapping.plId || '';
    const plAcct    = S.plAccounts.find(p => p.id === plId);
    const catColor  = plAcct ? (CATEGORY_COLORS[plAcct.category] || 'var(--txt2)') : 'var(--txt2)';

    // If this account has a pending creation, show the creation badge
    const creationBadge = creation
      ? `<span class="create-badge">${CATEGORY_LABELS[creation.plCategory] || creation.plCategory}</span>`
      : '';

    return `
      <div class="map-row ${isMerged ? 'merged' : ''} ${creation ? 'has-creation' : ''}"
           data-monarch-id="${a.id}">
        <div class="map-source">
          <div class="map-source-name">${esc(a.name)}${creationBadge}</div>
          <div class="map-source-bal">${fmt(a.balance)}${isMerged ? ' · <span style="color:var(--amber)">in merge</span>' : ''}
          </div>
        </div>
        <div class="map-arrow">→</div>
        <div class="map-dest">
          ${(() => {
            if (creation) {
              const wfLabels = {amount:'Current Value',balance:'Loan Balance',initialValue:'Purchase Price'};
              if (creation.failedAt) {
                // Failed creation — show error with Retry and Clear actions
                return `<div class="creation-error">
                  <span class="creation-error-label">✗ Failed: ${esc(creation.name)}</span>
                  <span class="creation-error-msg">${esc(creation.errorMsg || 'Unknown error')}</span>
                  <div class="creation-error-actions">
                    <button class="btn btn-sm btn-ghost creation-retry" data-mid="${a.id}">↺ Retry</button>
                    <button class="btn btn-sm btn-danger creation-clear" data-mid="${a.id}">✕ Clear</button>
                  </div>
                </div>`;
              }
              return `<div class="creation-summary" style="cursor:pointer" data-edit-mid="${a.id}" title="Click to edit">
                <span class="creation-label">New: ${esc(creation.name)} <span style="font-size:12px;color:var(--txt2)">&#9998; edit</span></span>
                <span class="creation-cat" style="color:${CATEGORY_COLORS[creation.plCategory]}">${CATEGORY_LABELS[creation.plCategory]}${creation.plType && creation.plType !== creation.plCategory ? ' &middot; '+esc(creation.plType) : ''}${creation.writeField ? ' &middot; '+(wfLabels[creation.writeField]||'') : ''}</span>
              </div>`;
            }
            if (isMerged) {
              const mergeGroup = S.mergeGroups.find(g => g.monarchIds.includes(a.id));
              const wfLabels = {amount:'Current Value',balance:'Loan Balance',initialValue:'Purchase Price'};
              let mergeLabel = 'in merge';
              if (mergeGroup && mergeGroup.pendingCreate) {
                const fm = mergeGroup.fieldMappings && mergeGroup.fieldMappings.find(f => f.monarchId === a.id);
                mergeLabel = 'New: ' + esc(mergeGroup.pendingCreate.name) + (fm ? ' \u2192 ' + (wfLabels[fm.writeField]||fm.writeField) : '');
              } else if (mergeGroup && mergeGroup.plId) {
                const plAcct = S.plAccounts.find(p => p.id === mergeGroup.plId);
                mergeLabel = plAcct ? esc(plAcct.name) : 'in merge';
              }
              return `<div style="font-size:11px;color:var(--amber);padding:4px 0;font-weight:600">${mergeLabel}</div>`;
            }
            return `<select class="select-input pl-acct-select"
                       style="font-size:11px;padding:4px 24px 4px 8px;border-left:2px solid ${catColor}"
                       data-monarch-id="${a.id}">
                 ${buildPlOptionHTML(plId)}
               </select>`;
          })()}
        </div>
        ${isMerged
          ? '<span class="map-info-icon" title="This account is managed from the Merge tab">ⓘ</span>'
          : `<button class="map-remove" data-monarch-id="${a.id}" title="Clear">✕</button>`
        }
      </div>
      <!-- inline creation form placeholder -->
      <div class="creation-form hidden" id="creation-form-${a.id}"></div>`;
  }).join('');

}

// One-time delegated listeners for mappingList — called from setupMapping()
function setupMappingListDelegation() {
  const list = byId('mappingList');

  // Clamp asset override value inputs to >= 0
  list.addEventListener('change', e => {
    const overrideInp = e.target.closest('.cf-override-val');
    if (overrideInp) {
      const v = parseFloat(overrideInp.value);
      overrideInp.value = (isFinite(v) && v >= 0) ? v : 0;
      return;
    }
  });

  list.addEventListener('change', e => {
    const sel = e.target.closest('.pl-acct-select');
    if (sel) {
      const mid = sel.dataset.monarchId;
      if (sel.value === '__CREATE_NEW__') {
        sel.value = '';
        showCreationForm(mid);
        return;
      }
      const plId = sel.value;
      if (plId) {
        const { realId, field } = resolvePlMapping(plId);
        S.mappings[mid] = { plId, plRealId: realId, plField: field };
      } else {
        delete S.mappings[mid];
      }
      const plAcct = S.plAccounts.find(p => p.id === plId);
      sel.style.borderLeft = `2px solid ${plAcct ? CATEGORY_COLORS[plAcct.category] : 'var(--txt2)'}`;
      updateCreateAllBtn();
    }

    // Payment status / write-field changes inside inline creation forms
    if (e.target.classList.contains('cf-payment-status')) {
      const formEl = e.target.closest('.creation-form');
      if (formEl) { syncAssetRowsForPayment(formEl, e.target.value); syncOverrideRows(formEl); }
    }
    if (e.target.classList.contains('cf-write-field')) {
      const formEl = e.target.closest('.creation-form');
      if (formEl) syncOverrideRows(formEl);
    }
  });

  list.addEventListener('click', async e => {
    // Edit creation summary → reopen form
    const summary = e.target.closest('.creation-summary[data-edit-mid]');
    if (summary) { showCreationForm(summary.dataset.editMid); return; }

    // Retry failed creation
    const retry = e.target.closest('.creation-retry');
    if (retry) {
      const mid = retry.dataset.mid;
      if (S.pendingCreations[mid]) {
        delete S.pendingCreations[mid].failedAt;
        delete S.pendingCreations[mid].errorMsg;
        await save(); renderMapping();
        toast('Creation re-armed — push to sync to retry', 'ok');
      }
      return;
    }

    // Clear failed creation
    const clear = e.target.closest('.creation-clear');
    if (clear) {
      delete S.pendingCreations[clear.dataset.mid];
      await save(); renderMapping();
      toast('Creation cleared', '');
      return;
    }

    // Remove mapping
    const remove = e.target.closest('.map-remove');
    if (remove) {
      delete S.mappings[remove.dataset.monarchId];
      delete S.pendingCreations[remove.dataset.monarchId];
      renderMapping();
      return;
    }

    // Save / cancel inline creation form buttons
    const cfSave = e.target.closest('.cf-save');
    if (cfSave) {
      const formEl = cfSave.closest('.creation-form');
      const monarchId = formEl?.id?.replace('creation-form-', '');
      if (monarchId && formEl) saveCreationForm(monarchId, formEl);
      return;
    }
    const cfCancel = e.target.closest('.cf-cancel');
    if (cfCancel) {
      const formEl = cfCancel.closest('.creation-form');
      if (formEl) { formEl.classList.add('hidden'); renderMapping(); }
      return;
    }
  });
}

// ── Inline creation form ──────────────────────────────────────
function showCreationForm(monarchId, closeOthers = true) {
  if (closeOthers) {
    document.querySelectorAll('.creation-form:not(.hidden)').forEach(f => f.classList.add('hidden'));
  }

  const a        = S.monarchAccounts.find(ac => ac.id === monarchId);
  if (!a) return;
  const inferred = inferCategory(a);
  const formEl   = byId(`creation-form-${monarchId}`);
  if (!formEl) return;

  formEl.innerHTML = buildCreationFormHTML(monarchId, a, inferred);
  formEl.classList.remove('hidden');

  // Wire up category → type sub-select visibility
  const catSel   = formEl.querySelector('.cf-category');
  const typeRow  = formEl.querySelector('.cf-type-row');
  const typeSel  = formEl.querySelector('.cf-type');
  const fieldRow = formEl.querySelector('.cf-field-row');

  // Category and type changes use direct listeners since they interact with
  // each other synchronously (category change must immediately update type options)
  catSel.addEventListener('change', () => {
    updateTypeOptions(catSel, typeRow, typeSel);
  });
  typeSel.addEventListener('change', () => {
    updateFieldRow(catSel.value, typeSel.value, fieldRow);
    syncOverrideRows(formEl);
  });

  // Set initial state
  updateTypeOptions(catSel, typeRow, typeSel, inferred.type);
  syncOverrideRows(formEl);
  // cf-payment-status, cf-write-field, cf-save, cf-cancel are handled by
  // the delegated listener on mappingList (setupMappingListDelegation)
}

function buildCreationFormHTML(monarchId, a, inferred) {
  const catOptions = PL_CATEGORIES.map(c =>
    `<option value="${c.value}" ${c.value === inferred.category ? 'selected' : ''}>${c.label}</option>`
  ).join('');
  const inferredWriteField = inferred.category === 'asset' ? (inferred.writeField || 'amount') : '';

  return `
    <div class="cf-inner">
      <div class="cf-row">
        <label class="cf-label">Name in ProjectionLab</label>
        <input class="text-input cf-name" value="${esc(a.name)}" style="font-size:12px;padding:5px 8px"/>
      </div>
      <div class="cf-row">
        <label class="cf-label">Category</label>
        <select class="select-input cf-category" style="font-size:12px;padding:5px 24px 5px 8px">
          ${catOptions}
        </select>
      </div>
      <div class="cf-row cf-type-row hidden">
        <label class="cf-label cf-type-label">Account Type</label>
        <select class="select-input cf-type" style="font-size:12px;padding:5px 24px 5px 8px"></select>
      </div>
      <!-- Asset-only rows — shown for all Real Assets -->
      <div class="cf-row cf-payment-row hidden">
        <label class="cf-label">Payment Status</label>
        <select class="select-input cf-payment-status" style="font-size:12px;padding:5px 24px 5px 8px">
          <option value="financed" selected>Financed</option>
          <option value="pay-in-full">Fully Owned</option>
        </select>
      </div>
      <div class="cf-row cf-field-row hidden">
        <label class="cf-label">This Monarch balance represents</label>
        <select class="select-input cf-write-field" style="font-size:12px;padding:5px 24px 5px 8px">
          ${ASSET_WRITE_FIELDS.map(f =>
            `<option value="${f.value}" ${f.value === inferredWriteField ? 'selected' : ''}>${f.label}</option>`
          ).join('')}
        </select>
      </div>
      <!-- Manual value overrides for unmapped asset fields (hidden until asset category selected) -->
      <div class="cf-asset-overrides hidden">
        <div class="cf-label" style="margin-top:4px;margin-bottom:4px">
          Unmapped field values <span style="font-size:12px;font-weight:400;color:var(--txt2)">set values not covered by the Monarch balance</span>
        </div>
        <div class="cf-override-row" data-field="amount">
          <span class="cf-override-label">Current Value</span>
          <input type="number" class="text-input cf-override-val" data-field="amount" placeholder="0" min="0" style="font-size:11px;padding:3px 6px"/>
        </div>
        <div class="cf-override-row cf-override-balance" data-field="balance">
          <span class="cf-override-label">Loan Balance</span>
          <input type="number" class="text-input cf-override-val" data-field="balance" placeholder="0" min="0" style="font-size:11px;padding:3px 6px"/>
        </div>
        <div class="cf-override-row" data-field="initialValue">
          <span class="cf-override-label">Purchase Price</span>
          <input type="number" class="text-input cf-override-val" data-field="initialValue" placeholder="0" min="0" style="font-size:11px;padding:3px 6px"/>
        </div>
      </div>
      <div class="cf-actions">
        <button class="btn btn-primary btn-sm cf-save">Confirm Create</button>
        <button class="btn btn-ghost btn-sm cf-cancel">Cancel</button>
      </div>
    </div>`;
}

// updateTypeOptions: fires when the CATEGORY select changes.
// Populates the type sub-select and hides/shows rows appropriately.
// Does NOT touch cf-field-row — that's handled by updateFieldRow.
// preselectedType: optional — when provided, that option is selected after populating.
// This is critical: without it, typeSel.value always defaults to the first option
// (real-estate), so the field row would always show for assets.
function updateTypeOptions(catSel, typeRow, typeSel, preselectedType) {
  const cat      = catSel.value;
  const formEl   = catSel.closest('.cf-inner');
  const fieldRow = formEl?.querySelector('.cf-field-row');

  if (cat === 'investment') {
    typeSel.innerHTML = INV_TYPES.map(t =>
      `<option value="${t.value}" ${preselectedType && t.value === preselectedType ? 'selected' : ''}>${t.label}</option>`
    ).join('');
    typeRow.classList.remove('hidden');
  } else if (cat === 'asset') {
    typeSel.innerHTML = ASSET_TYPES.map(t =>
      `<option value="${t.value}" ${preselectedType && t.value === preselectedType ? 'selected' : ''}>${t.label}</option>`
    ).join('');
    typeRow.classList.remove('hidden');
  } else if (cat === 'debt') {
    typeSel.innerHTML = DEBT_TYPES.map(t =>
      `<option value="${t.value}" ${preselectedType && t.value === preselectedType ? 'selected' : ''}>${t.label}</option>`
    ).join('');
    typeRow.classList.remove('hidden');
  } else {
    // savings — no sub-type
    typeRow.classList.add('hidden');
  }
  // Evaluate field row with the NOW-CORRECT typeSel.value (pre-selection already applied above)
  if (fieldRow) updateFieldRow(cat, typeSel.value, fieldRow);
}

// updateFieldRow — show asset rows for ALL real assets, hide for everything else.
// Also manages payment-status row, overrides section, and loan balance visibility.
function updateFieldRow(cat, plType, fieldRow) {
  const formEl       = fieldRow.closest('.cf-inner');
  const paymentRow   = formEl?.querySelector('.cf-payment-row');
  const overridesEl  = formEl?.querySelector('.cf-asset-overrides');
  const paymentSel   = formEl?.querySelector('.cf-payment-status');

  if (cat === 'asset') {
    fieldRow.classList.remove('hidden');
    if (paymentRow)  paymentRow.classList.remove('hidden');
    if (overridesEl) overridesEl.classList.remove('hidden');
    // Sync the write-field options and override rows to current payment status
    syncAssetRowsForPayment(formEl, paymentSel?.value || 'financed');
  } else {
    fieldRow.classList.add('hidden');
    if (paymentRow)  paymentRow.classList.add('hidden');
    if (overridesEl) overridesEl.classList.add('hidden');
  }
}

// Show/hide Loan Balance option and override row based on payment status.
function syncAssetRowsForPayment(formEl, paymentStatus) {
  const fullyOwned   = paymentStatus === 'pay-in-full';
  const writeFieldSel = formEl?.querySelector('.cf-write-field');
  const loanOverride = formEl?.querySelector('.cf-override-balance');

  // Remove or restore Loan Balance option in the write-field select
  if (writeFieldSel) {
    const curVal = writeFieldSel.value;
    writeFieldSel.innerHTML = ASSET_WRITE_FIELDS
      .filter(f => !(fullyOwned && f.value === 'balance'))
      .map(f => `<option value="${f.value}" ${f.value === curVal ? 'selected' : ''}>${f.label}</option>`)
      .join('');
    if (fullyOwned && curVal === 'balance') writeFieldSel.value = 'amount';
  }

  // Hide/show the Loan Balance manual override row
  if (loanOverride) {
    fullyOwned ? loanOverride.classList.add('hidden') : loanOverride.classList.remove('hidden');
  }
}

// Sync override rows: hide the row whose field is already covered by the write-field select.
function syncOverrideRows(formEl) {
  const writeFieldSel = formEl?.querySelector('.cf-write-field');
  const covered = writeFieldSel?.value;
  formEl?.querySelectorAll('.cf-override-row').forEach(row => {
    if (row.dataset.field === covered) {
      row.classList.add('hidden');
    } else {
      row.classList.remove('hidden');
    }
  });
  // Also keep loan balance hidden if fully owned
  const paymentSel = formEl?.querySelector('.cf-payment-status');
  const loanRow = formEl?.querySelector('.cf-override-balance');
  if (loanRow && paymentSel?.value === 'pay-in-full') {
    loanRow.classList.add('hidden');
  }
}

function saveCreationForm(monarchId, formEl) {
  const a        = S.monarchAccounts.find(ac => ac.id === monarchId);
  const name     = formEl.querySelector('.cf-name').value.trim();
  const cat      = formEl.querySelector('.cf-category').value;
  const typeRow  = formEl.querySelector('.cf-type-row');
  const fieldRow = formEl.querySelector('.cf-field-row');

  const plType = typeRow.classList.contains('hidden')
    ? (cat === 'savings' ? 'savings' : 'debt')
    : formEl.querySelector('.cf-type').value;

  // writeField and paymentStatus apply to ALL real assets
  const writeField = (cat === 'asset' && fieldRow && !fieldRow.classList.contains('hidden'))
    ? (formEl.querySelector('.cf-write-field')?.value || 'amount')
    : null;

  const paymentStatus = cat === 'asset'
    ? (formEl.querySelector('.cf-payment-status')?.value || 'financed')
    : null;

  // Collect manual override values for asset fields not covered by writeField
  const manualValues = {};
  if (cat === 'asset') {
    formEl.querySelectorAll('.cf-override-row:not(.hidden) .cf-override-val').forEach(inp => {
      const val = parseFloat(inp.value) || 0;
      manualValues[inp.dataset.field] = val;
    });
  }

  if (!name) { toast('Enter an account name', 'err'); return; }

  S.pendingCreations[monarchId] = {
    monarchId, name, balance: a?.balance ?? 0,
    plCategory: cat, plType,
    writeField,     // null for non-assets; 'amount'|'balance'|'initialValue' for assets
    paymentStatus,  // null for non-assets; 'financed'|'pay-in-full' for assets
    manualValues,   // {} for non-assets; override values for unmapped asset fields
  };
  delete S.mappings[monarchId];
  renderMapping();
  toast(`"${name}" queued for creation`, 'ok');
}

// ── Auto-map ──────────────────────────────────────────────────
function autoMap() {
  const primaryPL = S.plAccounts.filter(p => !p.subLabel || p.subLabel === 'Current Value');
  let matched = 0;

  function nameWords(name) {
    return new Set(name.toLowerCase().split(/[\s\-_.,()&]+/).filter(w => w.length >= 2));
  }
  function jaccard(a, b) {
    const wa = nameWords(a), wb = nameWords(b);
    if (!wa.size || !wb.size) return 0;
    let inter = 0;
    for (const w of wa) if (wb.has(w)) inter++;
    return inter / (wa.size + wb.size - inter);
  }

  const THRESHOLD = 0.5;

  // Track PL IDs already assigned (existing mappings + newly assigned this run)
  const assignedPlIds = new Set(
    Object.values(S.mappings).map(m => m.plRealId).filter(Boolean)
  );

  for (const ma of visibleMonarchAccounts()) {
    if (S.mappings[ma.id]) continue;

    let bestPL = null, bestScore = 0;
    for (const p of primaryPL) {
      // Skip PL accounts already claimed by another Monarch account
      if (assignedPlIds.has(p.id)) continue;
      const score = jaccard(ma.name, p.name);
      if (score > bestScore) { bestScore = score; bestPL = p; }
    }

    if (bestPL && bestScore >= THRESHOLD) {
      const { realId, field } = resolvePlMapping(bestPL.id);
      S.mappings[ma.id] = { plId: bestPL.id, plRealId: realId, plField: field };
      assignedPlIds.add(bestPL.id); // reserve for this run
      delete S.pendingCreations[ma.id];
      matched++;
    }
  }
  renderMapping();
  toast(`Auto-mapped ${matched} account${matched !== 1 ? 's' : ''}`, matched ? 'ok' : '');
}

// ── Bulk Create Panel ─────────────────────────────────────────
// Bulk create: iterate all visible unmapped accounts, infer category/type,
// and add them to pendingCreations exactly as the individual inline form does.
// This keeps the UX identical — each ends up as the same amber-badged creation row.
function openBulkCreatePanel() {
  const mergedIds = new Set(S.mergeGroups.flatMap(g => g.monarchIds));
  const unmapped  = visibleMonarchAccounts().filter(a =>
    !S.mappings[a.id] && !mergedIds.has(a.id) && !S.pendingCreations[a.id]
  );
  if (!unmapped.length) { toast('No unmapped accounts', ''); return; }

  let queued = 0;
  for (const a of unmapped) {
    const { category, type: plType } = inferCategory(a);
    S.pendingCreations[a.id] = {
      monarchId:  a.id,
      name:       a.name,
      balance:    a.balance ?? 0,
      plCategory: category,
      plType,
    };
    delete S.mappings[a.id];
    queued++;
  }

  renderMapping();
  // Open every inline form simultaneously so the user can review/override each inferred category.
  // closeOthers=false so they all stay open at once.
  for (const a of unmapped) {
    showCreationForm(a.id, false);
  }
  toast(`${queued} accounts queued — review categories below, then confirm each`, 'ok');
}

// ── MERGE tab ────────────────────────────────────────────────
let editingMergeId = null;

function setupMerge() {
  byId('addMergeGroup').addEventListener('click', () => openMergeBuilder());
  byId('addBudgetMergeGroup').addEventListener('click', () => openBudgetMergeBuilder());
  byId('saveBudgetMergeGroup').addEventListener('click', saveBudgetMergeGroup);
  byId('closeBudgetMergeBuilder').addEventListener('click', closeBudgetMergeBuilder);

  // Floating buttons wire to same actions
  byId('floatingAccountMergeBtn').addEventListener('click', () => openMergeBuilder());
  byId('floatingBudgetMergeBtn').addEventListener('click', () => openBudgetMergeBuilder());

  // Show floating buttons when merge-list scrolls — panel-merge has overflow:hidden
  // so IntersectionObserver can't see the legend leave the viewport
  byId('mergeGroupList').addEventListener('scroll', () => {
    const scrolled = byId('mergeGroupList').scrollTop > 10;
    byId('floatingAccountMergeBtn').classList.toggle('hidden', !scrolled);
    byId('floatingBudgetMergeBtn').classList.toggle('hidden', !scrolled);
  });
  byId('saveMergeGroup').addEventListener('click', saveMergeGroup);
  byId('floatingMergeSaveBtn').addEventListener('click', saveMergeGroup);
  byId('floatingBudgetMergeSaveBtn').addEventListener('click', saveBudgetMergeGroup);

  // Clamp manual-value override inputs (purchase price, current value, loan balance) to >= 0
  ['mergeValAmount', 'mergeValBalance', 'mergeValInitialValue'].forEach(id => {
    byId(id)?.addEventListener('change', e => {
      const v = parseFloat(e.target.value);
      e.target.value = (isFinite(v) && v >= 0) ? v : 0;
    });
  });



  // Show floating save when account builder header scrolls out of view
  // Show floating save when account builder body has scrolled
  byId('mergeBuilderBody').addEventListener('scroll', () => {
    const scrolled = byId('mergeBuilderBody').scrollTop > 10;
    byId('mergeFloatingBar').classList.toggle('hidden', !scrolled);
  });
  byId('closeMergeBuilder').addEventListener('click', closeMergeBuilder);
  byId('mergeSourceList').addEventListener('change', updateMergePreview);

  // PL target select: show/hide creation sub-form
  byId('mergePlTarget').addEventListener('change', onMergePlTargetChange);

  // Category/type changes inside the merge create sub-form
  byId('mergeCreateCategory').addEventListener('change', onMergeCreateCategoryChange);
  byId('mergeCreateType').addEventListener('change', onMergeCreateTypeChange);
  byId('mergeCreatePaymentStatus').addEventListener('change', () => {
    refreshMergeSourceList(null); // re-render checklist with/without loan balance option
    updateMergeManualOverrides(); // re-evaluate which override rows show
  });
}

// Is the merge currently in "create new" + "multi-field asset" mode?
// True for ALL real asset types.
function mergeIsMultiField() {
  const val = byId('mergePlTarget').value;
  if (val === '__CREATE_NEW__') return byId('mergeCreateCategory').value === 'asset';
  // Existing PL asset — look up in plAccounts
  const acct = S.plAccounts.find(p => p.id === val);
  return acct?.category === 'asset';
}

// Returns true if the merge asset is "Fully Owned"
function mergeIsFullyOwned() {
  return byId('mergeCreatePaymentStatus')?.value === 'pay-in-full';
}

function onMergePlTargetChange() {
  const val = byId('mergePlTarget').value;
  const createForm = byId('mergeCreateForm');
  if (val === '__CREATE_NEW__') {
    createForm.classList.remove('hidden');
    onMergeCreateCategoryChange();
  } else {
    createForm.classList.add('hidden');
  }
  // If an existing asset is selected, update the hint to reflect field-mapping mode
  const hint = byId('mergeSourceHint');
  if (hint && mergeIsMultiField()) {
    hint.textContent = '— assign each to a field in the existing asset';
  }
  refreshMergeSourceList(null);
  updateMergePreview();
}

function onMergeCreateCategoryChange() {
  const cat     = byId('mergeCreateCategory').value;
  const typeRow = byId('mergeCreateTypeRow');
  const typeSel = byId('mergeCreateType');
  const paymentRow = byId('mergeCreatePaymentRow');
  if (cat === 'investment') {
    typeSel.innerHTML = INV_TYPES.map(t => `<option value="${t.value}">${t.label}</option>`).join('');
    typeRow.classList.remove('hidden');
    if (paymentRow) paymentRow.classList.add('hidden');
  } else if (cat === 'asset') {
    typeSel.innerHTML = ASSET_TYPES.map(t => `<option value="${t.value}">${t.label}</option>`).join('');
    typeRow.classList.remove('hidden');
    if (paymentRow) paymentRow.classList.remove('hidden');
  } else if (cat === 'debt') {
    typeSel.innerHTML = DEBT_TYPES.map(t => `<option value="${t.value}">${t.label}</option>`).join('');
    typeRow.classList.remove('hidden');
    if (paymentRow) paymentRow.classList.add('hidden');
  } else {
    typeRow.classList.add('hidden');
    if (paymentRow) paymentRow.classList.add('hidden');
  }
  refreshMergeSourceList(null);
  updateMergeManualOverrides();
  updateMergePreview();
}

function onMergeCreateTypeChange() {
  refreshMergeSourceList(null);
  updateMergeManualOverrides();
  updateMergePreview();
}

// Refresh the checklist. existingFieldMappings allows pre-selecting field values when editing.
function refreshMergeSourceList(existingFieldMappings) {
  const multiField = mergeIsMultiField();
  const hint = byId('mergeSourceHint');
  hint.textContent = multiField
    ? '— assign each to a field in the new asset'
    : '— balances will be summed';

  const checkedIds = existingFieldMappings
    ? new Set(existingFieldMappings.map(fm => fm.monarchId))
    : new Set([...document.querySelectorAll('#mergeSourceList input[type=checkbox]:checked')].map(c => c.value));

  byId('mergeSourceList').innerHTML = visibleMonarchAccounts().map(a => {
    const isChecked = checkedIds.has(a.id);
    const existingField = existingFieldMappings?.find(fm => fm.monarchId === a.id)?.writeField || 'amount';
    if (multiField) {
      const fullyOwned = mergeIsFullyOwned();
      const fieldOpts  = ASSET_WRITE_FIELDS
        .filter(f => f.value !== 'initialValue')  // Purchase Price is manual override only
        .filter(f => !(fullyOwned && f.value === 'balance'))
        .map(f => `<option value="${f.value}" ${f.value === existingField ? 'selected' : ''}>${f.label}</option>`)
        .join('');
      return `
        <div class="check-item merge-field-row" data-monarch-id="${a.id}">
          <input type="checkbox" value="${a.id}" ${isChecked ? 'checked' : ''} class="merge-check"/>
          <span class="check-item-name">${esc(a.name)}</span>
          <span class="check-item-bal">${fmt(a.balance)}</span>
          <select class="select-input merge-field-sel ${isChecked ? '' : 'hidden'}"
                  data-monarch-id="${a.id}"
                  style="font-size:11px;padding:2px 20px 2px 6px;margin-left:auto;width:auto;max-width:120px">
            ${fieldOpts}
          </select>
        </div>`;
    } else {
      return `
        <label class="check-item">
          <input type="checkbox" value="${a.id}" ${isChecked ? 'checked' : ''}/>
          <span class="check-item-name">${esc(a.name)}</span>
          <span class="check-item-bal">${fmt(a.balance)}</span>
        </label>`;
    }
  }).join('');

  // Delegation handles all mergeSourceList interactions — see setupMergeSourceListDelegation()
  updateMergeManualOverrides();
}

// One-time delegated listeners for mergeSourceList — called from setupMerge()
function setupMergeSourceListDelegation() {
  const sourceList = byId('mergeSourceList');

  sourceList.addEventListener('change', e => {
    // Multi-field checkbox toggle: show/hide field selector
    const cb = e.target.closest('.merge-check');
    if (cb) {
      const row = cb.closest('.merge-field-row');
      const sel = row?.querySelector('.merge-field-sel');
      if (sel) sel.classList.toggle('hidden', !cb.checked);
      updateMergeManualOverrides();
      updateMergePreview();
      return;
    }
    // Field selector change
    if (e.target.classList.contains('merge-field-sel')) {
      updateMergeManualOverrides();
    }
    // Single-field checkbox: just update preview
    if (e.target.type === 'checkbox') {
      updateMergePreview();
    }
  });
}

// Show manual override inputs for any House asset fields not claimed by a Monarch account.
// Called any time the checklist or field assignments change.
function updateMergeManualOverrides() {
  const overridesEl = byId('mergeManualOverrides');
  if (!overridesEl) return;

  const isMultiField = mergeIsMultiField();
  if (!isMultiField) {
    overridesEl.classList.add('hidden');
    return;
  }

  overridesEl.classList.remove('hidden');

  // Collect which writeFields are currently covered by checked accounts
  const coveredFields = new Set();
  byId('mergeSourceList').querySelectorAll('.merge-field-row').forEach(row => {
    const cb  = row.querySelector('.merge-check');
    const sel = row.querySelector('.merge-field-sel');
    if (cb?.checked && sel) coveredFields.add(sel.value);
  });

  const fullyOwned = mergeIsFullyOwned();

  // Show override row only when: field not covered by a mapped account AND (not loan balance if fully owned)
  ['amount', 'balance', 'initialValue'].forEach(field => {
    const rowEl = byId('mergeOverride' + field.charAt(0).toUpperCase() + field.slice(1));
    if (rowEl) {
      const hide = coveredFields.has(field) || (field === 'balance' && fullyOwned);
      hide ? rowEl.classList.add('hidden') : rowEl.classList.remove('hidden');
    }
  });
}

function renderMerge() {
  renderCombinedMergeList();
}

function renderCombinedMergeList() {
  const list = byId('mergeGroupList');
  const hasAny = S.mergeGroups.length || S.budgetMergeGroups.length;
  if (!hasAny) {
    list.innerHTML = '<div class="empty-state" style="flex:none;padding:20px">No merge groups yet — use the buttons below to create one</div>';
    return;
  }
  const WRITE_FIELD_LABELS = { amount:'Current Value', balance:'Loan Balance', initialValue:'Purchase Price' };

  const accountCards = S.mergeGroups.map(g => {
    const isCreate = g.pendingCreate != null;
    let plName;
    if (isCreate) {
      plName = `<span style="color:var(--amber)">✦ New: ${esc(g.pendingCreate.name)}</span>
                <span style="font-size:12px;color:var(--txt2);margin-left:6px">${CATEGORY_LABELS[g.pendingCreate.plCategory] || ''} · ${g.pendingCreate.plType}</span>`;
    } else {
      const plAcct = S.plAccounts.find(p => p.id === g.plId);
      plName = plAcct ? (plAcct.subLabel ? `${esc(plAcct.name)} → ${plAcct.subLabel}` : esc(plAcct.name)) : (g.plId || '(PL account)');
    }
    let rowsHtml;
    if (g.fieldMappings) {
      rowsHtml = g.fieldMappings.map(fm => {
        const a = S.monarchAccounts.find(ac => ac.id === fm.monarchId);
        const fieldLabel = esc(WRITE_FIELD_LABELS[fm.writeField] || fm.writeField);
        return `<div class="merge-card-item merge-card-item--fields">
          <span class="merge-card-item-name">${a ? esc(a.name) : esc(fm.monarchId)}</span>
          <span class="merge-field-label">→ ${fieldLabel}</span>
          <span class="merge-card-item-bal">${a ? fmt(a.balance) : ''}</span>
        </div>`;
      }).join('');
    } else {
      rowsHtml = S.monarchAccounts.filter(a => g.monarchIds.includes(a.id))
        .map(a => `<div class="merge-card-item"><span class="merge-card-item-name">${esc(a.name)}</span><span class="merge-card-item-bal">${fmt(a.balance)}</span></div>`).join('');
    }
    return `<div class="merge-card merge-card--account" data-group-id="${g.id}">
      <div class="merge-card-header">
        <span class="merge-type-badge merge-type-badge--account">⇄ Account</span>
        <span class="merge-card-title">${plName}</span>
      </div>
      <div class="merge-card-body">${rowsHtml}</div>
      <div class="merge-card-actions">
        <button class="btn btn-sm btn-ghost edit-merge" data-id="${g.id}">Edit</button>
        <button class="btn btn-sm btn-danger remove-merge" data-id="${g.id}">Remove</button>
      </div>
    </div>`;
  });

  const budgetCards = S.budgetMergeGroups.map(g => {
    const cats = g.categoryIds.map(cid => {
      const b = S.monarchBudgets.find(x => x.categoryId === cid);
      return b ? { name: b.categoryName, amount: resolveMonarchAmount(b, getGroupCategoryOverride(g, cid)) } : { name: cid.slice(0,8), amount: 0 };
    });
    const total = cats.reduce((s, c) => s + c.amount, 0);
    const hasMappings = g.mappings?.length > 0;

    // Determine expense label for the card header
    let expenseLabel = '';
    if (hasMappings) {
      const firstMapping = g.mappings[0];
      if (firstMapping.expenseId === '__CREATE_NEW__') {
        expenseLabel = `<span style="color:var(--amber)">✦ New: ${esc(firstMapping.newExpenseName || 'expense')}</span>`;
      } else {
        const exp = S.plExpenses.find(e => e.expenseId === firstMapping.expenseId);
        expenseLabel = `<span style="color:var(--txt2);font-size:12px">${esc(exp?.expenseName || firstMapping.expenseId.slice(0,8))}</span>`;
      }
    }

    return `<div class="merge-card merge-card--budget" data-bmg-id="${g.id}">
      <div class="merge-card-header">
        <span class="merge-type-badge merge-type-badge--budget">💰 Budget</span>
        <span class="merge-card-title">${esc(g.name)}</span>
        ${expenseLabel ? `<span style="margin-left:auto">${expenseLabel}</span>` : `<span style="font-size:11px;color:var(--green);font-family:var(--font-mono);margin-left:auto">${fmt(total)}/mo</span>`}
      </div>
      <div class="merge-card-body">
        ${cats.map(c => `<div class="merge-card-item">
          <span class="merge-card-item-name">${esc(c.name)}</span>
          <span class="merge-card-item-bal">${fmt(c.amount)}/mo</span>
        </div>`).join('')}
      </div>
      <div class="merge-card-actions">
        <span style="font-size:12px;color:${hasMappings ? 'var(--green)' : 'var(--txt2)'}">
          ${hasMappings
            ? `✓ ${g.mappings.length} plan mapping${g.mappings.length>1?'s':''} · ${fmt(total)}/mo`
            : 'Map in Budgets tab'}
        </span>
        <button class="btn btn-sm btn-ghost edit-budget-merge" data-id="${g.id}">Edit</button>
        <button class="btn btn-sm btn-danger remove-budget-merge" data-id="${g.id}">Remove</button>
      </div>
    </div>`;
  });

  list.innerHTML = accountCards.join('') + budgetCards.join('');

}

// One-time delegated listeners for mergeGroupList — called from setupMerge()
function setupMergeGroupListDelegation() {
  const list = byId('mergeGroupList');

  list.addEventListener('click', async e => {
    const editMerge = e.target.closest('.edit-merge');
    if (editMerge) { openMergeBuilder(editMerge.dataset.id); return; }

    const removeMerge = e.target.closest('.remove-merge');
    if (removeMerge) {
      S.mergeGroups = S.mergeGroups.filter(g => g.id !== removeMerge.dataset.id);
      await save(); renderCombinedMergeList(); renderMapping();
      return;
    }

    const editBudget = e.target.closest('.edit-budget-merge');
    if (editBudget) { openBudgetMergeBuilder(editBudget.dataset.id); return; }

    const removeBudget = e.target.closest('.remove-budget-merge');
    if (removeBudget) {
      S.budgetMergeGroups = S.budgetMergeGroups.filter(g => g.id !== removeBudget.dataset.id);
      await save(); renderCombinedMergeList(); renderBudgetMapping();
      return;
    }
  });
}

function updateExpSelStyle(expSel) {
  if (!expSel) return;
  expSel.style.color = expSel.value === '__CREATE_NEW__' ? 'var(--amber)' : '';
  expSel.style.fontWeight = expSel.value === '__CREATE_NEW__' ? '600' : '';
}

// One-time delegated listeners for the budget merge builder — called from init.
// Handles: cat checklist, plan checklist, expense selector, type selector, freq row, source override.
function setupBudgetMergeBuilderDelegation() {
  // "Apply to all" toggle — shows/hides global controls, toggles per-row controls
  byId('budgetMergeSourceAllToggle').addEventListener('change', () => {
    const isAll = byId('budgetMergeSourceAllToggle').checked;
    byId('budgetMergeSourceAllControls').classList.toggle('hidden', !isAll);
    byId('budgetMergeCatList').classList.toggle('budget-merge-cat--all-mode', isAll);
    updateBudgetMergeTotal();
  });
  // Global source mode toggle
  byId('budgetMergeSourceMode').addEventListener('change', () => {
    const mv = byId('budgetMergeSourceMode').value;
    byId('budgetMergeActualOpts').classList.toggle('hidden', !mv);
    if (mv) {
      byId('budgetMergeLookback').value = mv === 'actual' ? 3 : 1;
      byId('budgetMergeAggSel').value   = 'average';
    }
    updateBudgetMergeTotal();
  });
  byId('budgetMergeLookback').addEventListener('change', () => {
    const el = byId('budgetMergeLookback');
    el.value = Math.max(1, Math.min(12, parseInt(el.value, 10) || 1));
    updateBudgetMergeTotal();
  });
  byId('budgetMergeAggSel').addEventListener('change', updateBudgetMergeTotal);

  // Category checklist — group toggles + individual checkboxes + per-row source controls
  const catList = byId('budgetMergeCatList');
  catList.addEventListener('change', async e => {
    // Per-row source mode
    const rowModeEl = e.target.closest('.bgt-mode-sel');
    if (rowModeEl) {
      const cid = rowModeEl.dataset.cid;
      const mode = rowModeEl.value;
      const row = rowModeEl.closest('.bgt-cat-item');
      if (row) row.querySelector('.bgt-actual-opts')?.classList.toggle('hidden', !mode);
      if (!mode) { delete S.budgetSourceOverrides[cid]; }
      else {
        const existing = S.budgetSourceOverrides[cid] || {};
        const defaultLb = mode === 'actual' ? 3 : 1;
        S.budgetSourceOverrides[cid] = { ...existing, mode, lookback: existing.lookback ?? defaultLb };
      }
      await save(); updateBudgetMergeTotal(); return;
    }
    // Per-row lookback
    const rowLookbackEl = e.target.closest('.bgt-lookback');
    if (rowLookbackEl) {
      const cid = rowLookbackEl.dataset.cid;
      const v = Math.max(1, Math.min(12, parseInt(rowLookbackEl.value, 10) || 3));
      rowLookbackEl.value = v;
      S.budgetSourceOverrides[cid] = { ...(S.budgetSourceOverrides[cid] || {}), lookback: v };
      await save(); updateBudgetMergeTotal(); return;
    }
    // Per-row aggregation
    const rowAggEl = e.target.closest('.bgt-agg-sel');
    if (rowAggEl) {
      const cid = rowAggEl.dataset.cid;
      S.budgetSourceOverrides[cid] = { ...(S.budgetSourceOverrides[cid] || {}), aggregation: rowAggEl.value };
      await save(); updateBudgetMergeTotal(); return;
    }

    // Group-level source toggle (merge builder checklist)
    const cgToggle = e.target.closest('.bgt-cg-toggle');
    if (cgToggle) {
      const gn = cgToggle.dataset.groupname;
      if (cgToggle.checked) {
        S.budgetGroupSourceOverrides[gn] = { mode: '', lookback: 1, aggregation: 'average' };
      } else {
        delete S.budgetGroupSourceOverrides[gn];
      }
      await save();
      const grpDiv = [...catList.querySelectorAll('.checklist-group')].find(el => el.dataset.groupname === gn);
      if (grpDiv) {
        grpDiv.classList.toggle('bgt-cg-active', cgToggle.checked);
        const header = grpDiv.querySelector('.checklist-group-header');
        const existingInline = header.querySelector('.bgt-cg-src-inline');
        if (cgToggle.checked && !existingInline) {
          const inlineDiv = document.createElement('div');
          inlineDiv.className = 'bgt-cg-src-inline';
          inlineDiv.innerHTML = `<select class="bgt-cg-mode-sel select-input" data-groupname="${esc(gn)}">
              <option value="" selected>Default</option>
              <option value="planned">Planned</option>
              <option value="actual">Actual</option>
            </select>
            <span class="bgt-actual-opts hidden">
              <input class="bgt-cg-lookback text-input" type="number" min="1" max="12"
                     value="1" data-groupname="${esc(gn)}"
                     style="width:40px;text-align:center;padding:2px 4px;font-size:11px"/>
              <span style="font-size:10px;color:var(--txt3)">mo</span>
              <select class="bgt-cg-agg-sel select-input" data-groupname="${esc(gn)}">
                <option value="average" selected>avg</option>
                <option value="median">median</option>
              </select>
            </span>`;
          header.querySelector('.bgt-cg-src-toggle').insertAdjacentElement('beforebegin', inlineDiv);
        } else if (!cgToggle.checked && existingInline) {
          existingInline.remove();
        }
      }
      updateBudgetMergeTotal(); return;
    }
    // Group-level source mode (merge builder checklist)
    const cgModeEl = e.target.closest('.bgt-cg-mode-sel');
    if (cgModeEl) {
      const gn   = cgModeEl.dataset.groupname;
      const mode = cgModeEl.value;
      const ctrlsDiv = cgModeEl.closest('.bgt-cg-src-inline');
      if (ctrlsDiv) ctrlsDiv.querySelector('.bgt-actual-opts')?.classList.toggle('hidden', !mode);
      if (!S.budgetGroupSourceOverrides[gn]) S.budgetGroupSourceOverrides[gn] = {};
      const defaultLookback = mode === 'actual' ? 3 : 1;
      S.budgetGroupSourceOverrides[gn] = { ...S.budgetGroupSourceOverrides[gn], mode, lookback: defaultLookback };
      const lb = ctrlsDiv?.querySelector('.bgt-cg-lookback');
      if (lb) lb.value = defaultLookback;
      await save(); updateBudgetMergeTotal(); return;
    }
    // Group-level lookback (merge builder checklist)
    const cgLookbackEl = e.target.closest('.bgt-cg-lookback');
    if (cgLookbackEl) {
      const gn = cgLookbackEl.dataset.groupname;
      const v  = Math.max(1, Math.min(12, parseInt(cgLookbackEl.value, 10) || 1));
      cgLookbackEl.value = v;
      if (!S.budgetGroupSourceOverrides[gn]) S.budgetGroupSourceOverrides[gn] = {};
      S.budgetGroupSourceOverrides[gn] = { ...S.budgetGroupSourceOverrides[gn], lookback: v };
      await save(); updateBudgetMergeTotal(); return;
    }
    // Group-level aggregation (merge builder checklist)
    const cgAggEl = e.target.closest('.bgt-cg-agg-sel');
    if (cgAggEl) {
      const gn = cgAggEl.dataset.groupname;
      if (!S.budgetGroupSourceOverrides[gn]) S.budgetGroupSourceOverrides[gn] = {};
      S.budgetGroupSourceOverrides[gn] = { ...S.budgetGroupSourceOverrides[gn], aggregation: cgAggEl.value };
      await save(); updateBudgetMergeTotal(); return;
    }

    const toggle = e.target.closest('.group-toggle-cb');
    if (toggle) {
      const gi    = toggle.dataset.group;
      const group = catList.querySelector(`.checklist-group[data-group="${gi}"]`);
      group?.querySelectorAll('input[type=checkbox]:not(.group-toggle-cb):not(.bgt-cg-toggle):not(:disabled)').forEach(cb => {
        cb.checked = toggle.checked;
      });
      updateBudgetMergeTotal();
      return;
    }
    if (e.target.type === 'checkbox') {
      updateBudgetMergeTotal();
      // Sync group toggle state
      catList.querySelectorAll('.checklist-group').forEach(grp => {
        const tog = grp.querySelector('.group-toggle-cb');
        if (!tog || tog.disabled) return;
        const cbs = [...grp.querySelectorAll('input[type=checkbox]:not(.group-toggle-cb):not(.bgt-cg-toggle):not(:disabled)')];
        tog.checked = cbs.length > 0 && cbs.every(c => c.checked);
      });
    }
  });

  // Collapse/expand group button — state persists across builder opens via _collapsedBudgetGroups
  catList.addEventListener('click', e => {
    const btn = e.target.closest('.group-collapse-btn');
    if (!btn) return;
    const gi = btn.dataset.group;
    const group = catList.querySelector(`.checklist-group[data-group="${gi}"]`);
    if (!group) return;
    const collapsed = group.classList.toggle('collapsed');
    btn.textContent = collapsed ? '▸' : '▾';
    const groupName = group.querySelector('.checklist-group-name')?.textContent;
    if (groupName) { collapsed ? _collapsedBudgetGroups.add(groupName) : _collapsedBudgetGroups.delete(groupName); }
  });

  // Plan checklist — select-all + individual plan checkboxes
  const planList = byId('budgetMergePlanList');
  planList.addEventListener('change', e => {
    if (e.target.id === 'planSelectAll') {
      planList.querySelectorAll('.bmg-plan-cb').forEach(cb => { cb.checked = e.target.checked; });
    } else if (e.target.classList.contains('bmg-plan-cb')) {
      const all       = [...planList.querySelectorAll('.bmg-plan-cb')];
      const selectAll = planList.querySelector('#planSelectAll');
      if (selectAll) selectAll.checked = all.every(c => c.checked);
    }
  });

  // Expense type selector
  byId('budgetMergeNewExpenseType').addEventListener('change', () => {
    renderExpFreqControl(byId('budgetMergeNewExpenseType').value, 'budgetMergeFreqRow', 'budgetMergeNewExpenseFreq', null);
    updateBudgetMergeTotal();
  });

  // Frequency selector (dynamically rendered into container — delegate on container)
  byId('budgetMergeFreqRow').addEventListener('change', () => updateBudgetMergeTotal());

  // Expense dropdown
  const expSel = byId('budgetMergeExpenseSel');
  expSel.addEventListener('change', () => {
    const val         = expSel.value;
    const isNew       = val === '__CREATE_NEW__';
    const newExpForm  = byId('budgetMergeNewExpenseForm');
    newExpForm.classList.toggle('hidden', !isNew);
    const expenseId   = isNew ? '__CREATE_NEW__' : val.split('::')[1];
    const fn = byId('budgetMergeBuilder')._renderPlanList;
    if (fn) fn(expenseId || val);
    updateExpSelStyle(expSel);
    updateBudgetMergeTotal(); // re-compute total using new expense's frequency
  });
  // Clear amber color while open so other options don't inherit it, restore on close.
  // Both mousedown (mouse users) and focus (keyboard users) must clear it.
  function clearExpSelAmber() { expSel.style.color = ''; expSel.style.fontWeight = ''; }
  expSel.addEventListener('mousedown', clearExpSelAmber);
  expSel.addEventListener('focus',     clearExpSelAmber);
  expSel.addEventListener('blur',      () => updateExpSelStyle(expSel));
}

function openMergeBuilder(id = null) {
  editingMergeId = id;
  const group = id ? S.mergeGroups.find(g => g.id === id) : null;
  byId('mergeBuilderTitle').textContent = id ? 'Edit Merge Group' : 'New Merge Group';

  const isCreate = group?.pendingCreate != null;

  // Populate PL target dropdown
  byId('mergePlTarget').innerHTML = buildPlOptionHTML(isCreate ? '__CREATE_NEW__' : (group?.plId || ''), false);
  // If editing a pending-create group, force the value to __CREATE_NEW__
  if (isCreate) byId('mergePlTarget').value = '__CREATE_NEW__';

  // Show/hide creation sub-form
  const createForm = byId('mergeCreateForm');
  if (isCreate) {
    createForm.classList.remove('hidden');
    // Pre-fill creation fields
    byId('mergeCreateName').value = group.pendingCreate.name || '';
    byId('mergeCreateCategory').value = group.pendingCreate.plCategory || 'asset';
    onMergeCreateCategoryChange(); // populates type options
    if (group.pendingCreate.plType) byId('mergeCreateType').value = group.pendingCreate.plType;
    // Restore manual values
    const mv = group.pendingCreate.manualValues || {};
    if (byId('mergeValAmount'))       byId('mergeValAmount').value       = mv.amount       || '';
    if (byId('mergeValBalance'))      byId('mergeValBalance').value      = mv.balance      || '';
    if (byId('mergeValInitialValue')) byId('mergeValInitialValue').value = mv.initialValue || '';
  } else {
    createForm.classList.add('hidden');
  }

  // Render source list (with pre-selected fieldMappings if editing)
  refreshMergeSourceList(group?.fieldMappings || null);

  // For regular groups, also check the right boxes
  if (!isCreate && group?.monarchIds) {
    group.monarchIds.forEach(mid => {
      const cb = byId('mergeSourceList').querySelector(`input[value="${mid}"]`);
      if (cb) cb.checked = true;
    });
  }

  byId('mergeBuilder').classList.remove('hidden');
  byId('panel-merge').classList.add('builder-open');
  updateMergePreview();
}

function closeMergeBuilder() {
  byId('mergeBuilder').classList.add('hidden');
  byId('panel-merge').classList.remove('builder-open');
  byId('mergeFloatingBar').classList.add('hidden');
  const ft = byId('mergeFloatingTotal');
  if (ft) ft.textContent = '';
  editingMergeId = null;
  byId('mergeSourceList').querySelectorAll('input[type=checkbox]').forEach(cb => { cb.checked = false; });
  const hdr = byId('mergeTotalDisplay');
  if (hdr) hdr.textContent = '';
}

function updateMergePreview() {
  const checked = [...document.querySelectorAll('#mergeSourceList input[type=checkbox]:checked')].map(c => c.value);
  const total   = S.monarchAccounts.filter(a => checked.includes(a.id)).reduce((s, a) => s + (a.balance || 0), 0);
  const txt = checked.length ? fmt(total) : '';
  const hdr = byId('mergeTotalDisplay');
  if (hdr) hdr.textContent = txt;
  const floating = byId('mergeFloatingTotal');
  if (floating) floating.textContent = txt;
}

async function saveMergeGroup() {
  const plId       = byId('mergePlTarget').value;
  const isCreate   = plId === '__CREATE_NEW__';
  const monarchIds = [...document.querySelectorAll('#mergeSourceList input[type=checkbox]:checked')].map(c => c.value);

  if (!plId && !isCreate) { toast('Choose a PL account', 'err'); return; }
  if (!monarchIds.length)  { toast('Select at least one Monarch account', 'err'); return; }

  let entry;

  if (isCreate) {
    const name     = byId('mergeCreateName').value.trim();
    const cat      = byId('mergeCreateCategory').value;
    const typeRow  = byId('mergeCreateTypeRow');
    const plType   = typeRow.classList.contains('hidden') ? cat
                   : byId('mergeCreateType').value;
    if (!name) { toast('Enter a name for the new PL account', 'err'); return; }

    const isMultiField = mergeIsMultiField();
    let fieldMappings = null;

    if (isMultiField) {
      // Each checked account has its own field selector
      fieldMappings = monarchIds.map(mid => {
        const sel = document.querySelector(`#mergeSourceList .merge-field-sel[data-monarch-id="${mid}"]`);
        const acct = S.monarchAccounts.find(a => a.id === mid);
        return {
          monarchId:  mid,
          writeField: sel ? sel.value : 'amount',
          balance:    acct?.balance ?? 0,
        };
      });
    }

    const mergePaymentStatus = byId('mergeCreatePaymentStatus')?.value || 'financed';
    entry = {
      id: editingMergeId || `mg_${Date.now()}`,
      plId:          null,
      plRealId:      null,
      plField:       isMultiField ? null : 'amount',
      monarchIds,
      fieldMappings,
      pendingCreate: {
        name, plCategory: cat, plType,
        paymentStatus: mergePaymentStatus,
        manualValues: {
          amount:       parseFloat(byId('mergeValAmount')?.value       || '0') || 0,
          balance:      parseFloat(byId('mergeValBalance')?.value      || '0') || 0,
          initialValue: parseFloat(byId('mergeValInitialValue')?.value || '0') || 0,
        },
      },
    };
  } else {
    const { realId: plRealId } = resolvePlMapping(plId);
    const isMultiField = mergeIsMultiField();
    let fieldMappings = null;
    if (isMultiField) {
      fieldMappings = monarchIds.map(mid => {
        const sel = document.querySelector(`#mergeSourceList .merge-field-sel[data-monarch-id="${mid}"]`);
        const acct = S.monarchAccounts.find(a => a.id === mid);
        return {
          monarchId:  mid,
          writeField: sel ? sel.value : 'amount',
          balance:    acct?.balance ?? 0,
        };
      });
    }
    const plField = isMultiField ? null : resolvePlMapping(plId).field;
    entry = { id: editingMergeId || `mg_${Date.now()}`, plId, plRealId, plField, monarchIds, fieldMappings };
  }

  if (editingMergeId) {
    const idx = S.mergeGroups.findIndex(g => g.id === editingMergeId);
    if (idx > -1) S.mergeGroups[idx] = entry;
  } else { S.mergeGroups.push(entry); }

  await save();
  closeMergeBuilder();
  renderMerge();
  renderMapping();
  toast('Merge group saved!', 'ok');
}

// ── BUDGET MERGE GROUPS ──────────────────────────────────────
let editingBudgetMergeId = null;
// Persists collapse state across builder opens within the same popup session (keyed by group name).
const _collapsedBudgetGroups = new Set();
// Persists collapse state for budget mapping tab groups across re-renders (keyed by group name).
const _collapsedBudgetMappingGroups = new Set();

function renderBudgetMerges() {
  renderCombinedMergeList();
}

// Module-level so setupBudgetMergeBuilderDelegation can call it before builder opens.
function updateBudgetMergeTotal() {
  const list = byId('budgetMergeCatList');
  if (!list) return;
  // Determine effective override for each checked category based on the toggle state
  const isAllMode = byId('budgetMergeSourceAllToggle')?.checked !== false;
  let allModeOverride = null;
  if (isAllMode) {
    const mv = byId('budgetMergeSourceMode')?.value || '';
    allModeOverride = mv ? {
      mode: mv,
      lookback: Math.max(1, parseInt(byId('budgetMergeLookback')?.value, 10) || 3),
      aggregation: byId('budgetMergeAggSel')?.value || 'average',
    } : null;
  }
  // Update individual row balance displays and accumulate total
  const totalMonthly = [...list.querySelectorAll('input[type=checkbox]:not(.group-toggle-cb):not(.bgt-cg-toggle)')].reduce((sum, cb) => {
    const b = S.monarchBudgets.find(x => x.categoryId === cb.value);
    if (!b) return sum;
    const override = isAllMode ? allModeOverride : getEffectiveBudgetOverride(cb.value);
    const amt = resolveMonarchAmount(b, override);
    // Update the balance span for this row
    const balSpan = cb.closest('.checklist-item')?.querySelector('.checklist-item-bal');
    if (balSpan) balSpan.textContent = `${fmt(amt)}/mo`;
    return cb.checked ? sum + amt : sum;
  }, 0);

  // Determine frequency: existing expense → use its frequency; new expense → use selector
  const expSel = byId('budgetMergeExpenseSel');
  const expVal = expSel?.value || '';
  let freq = 'monthly';

  if (expVal && expVal !== '__CREATE_NEW__' && expVal !== '') {
    // Existing expense selected — find its frequency from plExpenses
    const expenseId = expVal.includes('::') ? expVal.split('::')[1] : expVal;
    const planId    = expVal.includes('::') ? expVal.split('::')[0] : null;
    const expDef    = planId
      ? S.plExpenses.find(e => e.planId === planId && e.expenseId === expenseId)
      : S.plExpenses.find(e => e.expenseId === expenseId);
    freq = expDef?.frequency || 'monthly';
  } else {
    // New expense or none — use type/freq selectors
    const expType   = byId('budgetMergeNewExpenseType')?.value || 'living-expenses';
    const freqSel   = byId('budgetMergeNewExpenseFreq');
    const fixedFreq = EXP_FIXED_FREQ[expType];
    freq = fixedFreq || freqSel?.value || 'monthly';
  }

  let displayAmt, freqSuffix;
  if (freq === 'once') {
    displayAmt = totalMonthly;
    freqSuffix = ' ⚠';
  } else {
    const opt  = PL_FREQ_OPTIONS.find(o => o.value === freq);
    displayAmt = opt ? opt.convert(totalMonthly) : totalMonthly;
    const shortFreq = { 'yearly-lump-sum':'yr', yearly:'yr', quarterly:'qtr', monthly:'mo', 'bi-weekly':'2wk', weekly:'wk', daily:'day' }[freq] || freq;
    freqSuffix = `/${shortFreq}`;
  }

  const txt = totalMonthly > 0 ? fmt(displayAmt) + freqSuffix : '';
  const el = byId('budgetMergeTotalDisplay');
  if (el) el.textContent = txt;
  const floating = byId('budgetMergeFloatingTotal');
  if (floating) floating.textContent = txt;
}

function openBudgetMergeBuilder(id = null) {
  editingBudgetMergeId = id;
  const group = id ? S.budgetMergeGroups.find(g => g.id === id) : null;
  byId('budgetMergeBuilderTitle').textContent = id ? 'Edit Budget Merge' : 'New Budget Merge';
  byId('budgetMergeName').value = group?.name || '';

  // Populate category checklist
  const thisGroupIds = new Set(group?.categoryIds || []);
  const otherMergedIds = new Set(
    S.budgetMergeGroups
      .filter(g => g.id !== id)
      .flatMap(g => g.categoryIds)
  );
  const list = byId('budgetMergeCatList');

  // Group by groupName for the checklist
  const cgOrder = [];
  const cgGrouped = {};
  for (const b of S.monarchBudgets) {
    const gn = b.groupName || 'Other';
    if (!cgGrouped[gn]) { cgGrouped[gn] = []; cgOrder.push(gn); }
    cgGrouped[gn].push(b);
  }

  list.innerHTML = cgOrder.map((groupName, gi) => {
    const groupItems = cgGrouped[groupName];
    const allDisabled = groupItems.every(b => otherMergedIds.has(b.categoryId));
    const allChecked  = groupItems.filter(b => !otherMergedIds.has(b.categoryId)).every(b => thisGroupIds.has(b.categoryId));
    const items = groupItems.map(b => {
      const disabled = otherMergedIds.has(b.categoryId);
      const checked  = thisGroupIds.has(b.categoryId);
      const catOvr   = S.budgetSourceOverrides[b.categoryId] || null;
      const ovrMode  = catOvr?.mode || '';
      const defaultLookback = ovrMode === 'actual' ? 3 : 1;
      const ovrLookback = catOvr?.lookback ?? defaultLookback;
      const ovrAgg   = catOvr?.aggregation || 'average';
      return `<div class="bgt-cat-item">
        <label class="checklist-item${disabled ? ' disabled' : ''}">
          <input type="checkbox" value="${esc(b.categoryId)}" ${checked ? 'checked' : ''} ${disabled ? 'disabled' : ''}/>
          <span class="checklist-item-name">${esc(b.categoryName)}</span>
          <span class="checklist-item-bal">${fmt(resolveMonarchAmount(b, catOvr))}/mo</span>
        </label>
        <div class="bgt-cat-source-override">
          <select class="bgt-mode-sel select-input" data-cid="${esc(b.categoryId)}">
            <option value="" ${!ovrMode ? 'selected' : ''}>Default</option>
            <option value="planned" ${ovrMode === 'planned' ? 'selected' : ''}>Planned</option>
            <option value="actual" ${ovrMode === 'actual' ? 'selected' : ''}>Actual</option>
          </select>
          <span class="bgt-actual-opts ${!ovrMode ? 'hidden' : ''}">
            <input class="bgt-lookback text-input" type="number" min="1" max="12"
                   value="${ovrLookback}" data-cid="${esc(b.categoryId)}"
                   style="width:40px;text-align:center;padding:2px 4px;font-size:11px"/>
            <span style="font-size:10px;color:var(--txt3)">mo</span>
            <select class="bgt-agg-sel select-input" data-cid="${esc(b.categoryId)}">
              <option value="average" ${ovrAgg === 'average' ? 'selected' : ''}>avg</option>
              <option value="median" ${ovrAgg === 'median' ? 'selected' : ''}>median</option>
            </select>
          </span>
        </div>
      </div>`;
    }).join('');
    const isCollapsed = _collapsedBudgetGroups.has(groupName);
    const cgOvr = S.budgetGroupSourceOverrides[groupName] || null;
    const cgActive = !!cgOvr;
    const cgMode = cgOvr?.mode || '';
    const cgDefaultLookback = cgMode === 'actual' ? 3 : 1;
    const cgLookback = cgOvr?.lookback ?? cgDefaultLookback;
    const cgAgg = cgOvr?.aggregation || 'average';
    const cgInlineControls = cgActive ? `
          <div class="bgt-cg-src-inline">
            <select class="bgt-cg-mode-sel select-input" data-groupname="${esc(groupName)}">
              <option value="" ${!cgMode ? 'selected' : ''}>Default</option>
              <option value="planned" ${cgMode === 'planned' ? 'selected' : ''}>Planned</option>
              <option value="actual" ${cgMode === 'actual' ? 'selected' : ''}>Actual</option>
            </select>
            <span class="bgt-actual-opts ${!cgMode ? 'hidden' : ''}">
              <input class="bgt-cg-lookback text-input" type="number" min="1" max="12"
                     value="${cgLookback}" data-groupname="${esc(groupName)}"
                     style="width:40px;text-align:center;padding:2px 4px;font-size:11px"/>
              <span style="font-size:10px;color:var(--txt3)">mo</span>
              <select class="bgt-cg-agg-sel select-input" data-groupname="${esc(groupName)}">
                <option value="average" ${cgAgg === 'average' ? 'selected' : ''}>avg</option>
                <option value="median" ${cgAgg === 'median' ? 'selected' : ''}>median</option>
              </select>
            </span>
          </div>` : '';
    return `<div class="checklist-group${isCollapsed ? ' collapsed' : ''}${cgActive ? ' bgt-cg-active' : ''}" data-group="${gi}" data-groupname="${esc(groupName)}">
      <div class="checklist-group-header">
        <label class="checklist-group-toggle" title="Select all / deselect all">
          <input type="checkbox" class="group-toggle-cb" data-group="${gi}" ${allChecked && !allDisabled ? 'checked' : ''} ${allDisabled ? 'disabled' : ''}/>
        </label>
        <span class="checklist-group-name">${esc(groupName)}</span>
        ${cgInlineControls}
        <label class="bgt-cg-src-toggle" title="Apply one source setting to this entire group">
          <input type="checkbox" class="bgt-cg-toggle" data-groupname="${esc(groupName)}" ${cgActive ? 'checked' : ''}/>
          <span>Group source</span>
        </label>
        <button type="button" class="group-collapse-btn" data-group="${gi}" title="Collapse / expand">${isCollapsed ? '▸' : '▾'}</button>
      </div>
      <div class="checklist-group-items">
      ${items}
      </div>
    </div>`;
  }).join('');

  // syncGroupToggles: keeps group-toggle checkboxes in sync with individual items
  function syncGroupToggles() {
    list.querySelectorAll('.checklist-group').forEach(group => {
      const toggle = group.querySelector('.group-toggle-cb');
      if (!toggle || toggle.disabled) return;
      const cbs = [...group.querySelectorAll('input[type=checkbox]:not(.group-toggle-cb):not(.bgt-cg-toggle):not(:disabled)')];
      toggle.checked = cbs.length > 0 && cbs.every(cb => cb.checked);
    });
  }

  // Populate source override controls
  // sourceModeAll === false means per-category was explicitly chosen; undefined/true = all mode
  const isAllMode = group?.sourceModeAll !== false;
  const existingSrcOverride = group?.sourceOverride || null;
  const allToggle = byId('budgetMergeSourceAllToggle');
  allToggle.checked = isAllMode;
  byId('budgetMergeSourceAllControls').classList.toggle('hidden', !isAllMode);
  byId('budgetMergeCatList').classList.toggle('budget-merge-cat--all-mode', isAllMode);
  const srcModeEl = byId('budgetMergeSourceMode');
  srcModeEl.value = existingSrcOverride?.mode || '';
  byId('budgetMergeActualOpts').classList.toggle('hidden', !srcModeEl.value);
  // Always reset to mode-appropriate defaults, then apply saved values if present
  const initMode = srcModeEl.value;
  byId('budgetMergeLookback').value = existingSrcOverride?.lookback ?? (initMode === 'actual' ? 3 : 1);
  byId('budgetMergeAggSel').value   = existingSrcOverride?.aggregation || 'average';

  // Delegation is handled by setupBudgetMergeBuilderDelegation() — called once at init
  updateBudgetMergeTotal();

  // Populate expense dropdown — reuse buildExpenseOptionHTML for consistent styling,
  // then mark the current expenseId as selected afterward
  const expSel = byId('budgetMergeExpenseSel');
  const currentExpenseId = group?.mappings?.[0]?.expenseId || '';
  const isCreatingNew = currentExpenseId === '__CREATE_NEW__';
  expSel.innerHTML = buildExpenseOptionHTML();
  // Mark the current selection
  if (isCreatingNew) {
    expSel.value = '__CREATE_NEW__';
  } else if (currentExpenseId) {
    // buildExpenseOptionHTML uses planId::expenseId keys; find the matching option by expenseId
    const matchingOpt = [...expSel.options].find(o => o.value.endsWith('::' + currentExpenseId));
    if (matchingOpt) matchingOpt.selected = true;
  }

  // Pre-fill new expense form if editing a create-new entry
  const newExpenseForm = byId('budgetMergeNewExpenseForm');
  if (isCreatingNew) {
    newExpenseForm.classList.remove('hidden');
    byId('budgetMergeNewExpenseName').value = group?.mappings?.[0]?.newExpenseName || '';
    const savedType = group?.mappings?.[0]?.newExpType || 'living-expenses';
    byId('budgetMergeNewExpenseType').value = savedType;
    const savedFreq = group?.mappings?.[0]?.newFrequency;
    if (savedFreq) byId('budgetMergeNewExpenseFreq').value = savedFreq;
    renderExpFreqControl(savedType, 'budgetMergeFreqRow', 'budgetMergeNewExpenseFreq', savedFreq);
  } else {
    newExpenseForm.classList.add('hidden');
    renderExpFreqControl('living-expenses', 'budgetMergeFreqRow', 'budgetMergeNewExpenseFreq', null);
  }

  // Type/freqRow/expSel/planList delegation handled by setupBudgetMergeBuilderDelegation()
  updateExpSelStyle(expSel);

  // Render plan checklist — also used by the delegated expSel listener via renderBudgetMergePlanList()
  function renderPlanList(expenseId) {
    const section = byId('budgetMergePlanSection');
    const planList = byId('budgetMergePlanList');
    if (!expenseId) { section.classList.add('hidden'); return; }
    section.classList.remove('hidden');

    const isNew = expenseId === '__CREATE_NEW__';
    const allPlanIds = [...new Set(S.plExpenses.map(e => e.planId))];
    const checkedPlanIds = new Set((group?.mappings || [])
      .filter(m => m.expenseId === expenseId)
      .map(m => m.planId));

    const allChecked = allPlanIds.length > 0 && allPlanIds.every(id => checkedPlanIds.has(id));

    const rows = allPlanIds.map(planId => {
      const planEntry = S.plExpenses.find(e => e.planId === planId);
      const planName = planEntry?.planName || planId.slice(0, 8);
      const existsInPlan = isNew ? false : S.plExpenses.some(e => e.planId === planId && e.expenseId === expenseId);
      const hint = (isNew || !existsInPlan) ? ' <span style="color:var(--amber);font-size:11px;font-weight:700">(new)</span>' : '';
      const checked = checkedPlanIds.has(planId);
      return `<label class="checklist-item">
        <input type="checkbox" class="bmg-plan-cb" value="${esc(planId)}" data-exists="${existsInPlan}" ${checked ? 'checked' : ''}/>
        <span class="checklist-item-name">${esc(planName)}${hint}</span>
      </label>`;
    }).join('');

    planList.innerHTML = `
      <div class="checklist-group-header" style="background:var(--bg2);border-bottom:1px solid var(--border)">
        <label class="checklist-group-toggle" title="Select all plans / deselect all">
          <input type="checkbox" id="planSelectAll" ${allChecked ? 'checked' : ''}/>
        </label>
        <span class="checklist-group-name" style="font-size:12px;color:var(--txt2)">All plans</span>
      </div>
      ${rows}`;

    // planList delegation handled by setupBudgetMergeBuilderDelegation()
  }

  // Expose renderPlanList so the delegated expSel handler can call it
  // by storing it on the builder element for retrieval
  byId('budgetMergeBuilder')._renderPlanList = renderPlanList;

  const initialVal = expSel.value;
  const initialExpId = isCreatingNew ? '__CREATE_NEW__' : (initialVal.includes('::') ? initialVal.split('::')[1] : initialVal);
  renderPlanList(initialExpId);

  byId('budgetMergeBuilder').classList.remove('hidden');
  byId('budgetMergeFloatingBar').classList.remove('hidden');
  byId('panel-merge').classList.add('builder-open');
}

function closeBudgetMergeBuilder() {
  byId('budgetMergeBuilder').classList.add('hidden');
  byId('panel-merge').classList.remove('builder-open');
  const el = byId('budgetMergeTotalDisplay');
  if (el) el.textContent = '';
  byId('budgetMergeFloatingBar').classList.add('hidden');
  byId('budgetMergeFloatingTotal').textContent = '';
  byId('budgetMergePlanList').innerHTML = '';
  byId('budgetMergePlanSection').classList.add('hidden');
  byId('budgetMergeNewExpenseForm').classList.add('hidden');
  byId('budgetMergeNewExpenseName').value = '';
  const freqEl = byId('budgetMergeNewExpenseFreq');
  if (freqEl) freqEl.value = 'monthly';
  editingBudgetMergeId = null;
}

async function saveBudgetMergeGroup() {
  const name = byId('budgetMergeName').value.trim();
  if (!name) { toast('Enter a name for this merge group', 'err'); return; }
  const checked = [...byId('budgetMergeCatList').querySelectorAll('input[type=checkbox]:checked:not(.group-toggle-cb):not(.bgt-cg-toggle)')];
  const categoryIds = checked.map(c => c.value);
  if (categoryIds.length < 2) { toast('Select at least 2 categories to merge', 'err'); return; }

  // Guard against cross-group duplicates — if another group claims any of these
  // categories, remove them from that group (this group takes ownership).
  const otherGroups = S.budgetMergeGroups.filter(g => g.id !== editingBudgetMergeId);
  const categoryIdSet = new Set(categoryIds);
  let stolen = 0;
  for (const g of otherGroups) {
    const before = g.categoryIds.length;
    g.categoryIds = g.categoryIds.filter(cid => !categoryIdSet.has(cid));
    stolen += before - g.categoryIds.length;
  }
  // Remove any groups left with fewer than 2 categories (no longer valid merges)
  S.budgetMergeGroups = S.budgetMergeGroups.filter(g =>
    g.id === editingBudgetMergeId || g.categoryIds.length >= 2
  );
  if (stolen > 0) {
    toast(`${stolen} categor${stolen > 1 ? 'ies' : 'y'} moved to this group from another`, 'ok');
  }

  // Declare existing first — needed for stable UUID reuse on create-new expense edits
  const existing = editingBudgetMergeId
    ? S.budgetMergeGroups.find(g => g.id === editingBudgetMergeId)
    : null;

  // Build mappings from expense + plan selections
  const rawExpVal = byId('budgetMergeExpenseSel').value;
  const expenseId = rawExpVal === '__CREATE_NEW__' ? '__CREATE_NEW__' : (rawExpVal.includes('::') ? rawExpVal.split('::')[1] : rawExpVal);
  const isNewExpense = expenseId === '__CREATE_NEW__';
  let mappings = [];
  if (expenseId) {
    if (isNewExpense) {
      const newName = byId('budgetMergeNewExpenseName').value.trim();
      if (!newName) { toast('Enter a name for the new expense', 'err'); return; }
      const { expType, frequency: newFreq, expTypeParams } = collectExpTypeParams('budgetMerge');
      const newExpenseId = existing?.mappings?.[0]?.expenseId === '__CREATE_NEW__'
        ? (existing.mappings[0].newExpenseId || crypto.randomUUID())
        : crypto.randomUUID();
      const planCbs = [...byId('budgetMergePlanList').querySelectorAll('.bmg-plan-cb:checked')];
      mappings = planCbs.map(cb => ({
        planId:           cb.value,
        expenseId:        '__CREATE_NEW__',
        newExpenseId,
        newExpenseName:   newName,
        newFrequency:     newFreq,
        newExpType:       expType,
        newExpTypeParams: expTypeParams,
        createIfMissing:  true,
      }));
    } else {
      const planCbs = [...byId('budgetMergePlanList').querySelectorAll('.bmg-plan-cb:checked')];
      mappings = planCbs.map(cb => ({
        planId: cb.value,
        expenseId,
        createIfMissing: cb.dataset.exists === 'false',
      }));
    }
  }

  const sourceModeAll = byId('budgetMergeSourceAllToggle').checked;
  let sourceOverride = null;
  if (sourceModeAll) {
    const srcModeVal = byId('budgetMergeSourceMode').value;
    sourceOverride = srcModeVal ? {
      mode:        srcModeVal,
      lookback:    srcModeVal === 'actual' ? (Math.max(1, parseInt(byId('budgetMergeLookback').value, 10) || 3)) : 3,
      aggregation: srcModeVal === 'actual' ? (byId('budgetMergeAggSel').value || 'average') : 'average',
    } : null;
  }

  const entry = {
    id:           existing?.id || crypto.randomUUID(),
    name,
    categoryIds,
    mappings,
    sourceModeAll,
    sourceOverride,
  };

  if (editingBudgetMergeId) {
    const idx = S.budgetMergeGroups.findIndex(g => g.id === editingBudgetMergeId);
    if (idx > -1) S.budgetMergeGroups[idx] = entry;
  } else {
    S.budgetMergeGroups.push(entry);
  }
  await save();
  closeBudgetMergeBuilder();
  renderBudgetMerges();
  renderBudgetMapping();
  toast('Budget merge group saved!', 'ok');
}

// ── BUDGET MAPPING tab ───────────────────────────────────────
function setupBudgetMapping() {
  // Search filter
  byId('budgetSearch').addEventListener('input', renderBudgetMapping);

  // Save button
  async function doSaveBudget() {
    await save();
    const status = byId('budgetSaveStatus');
    if (status) { status.textContent = '✓ Saved'; setTimeout(() => { status.textContent = ''; }, 3500); }
    toast('Budget mappings saved!', 'ok');
  }
  byId('saveBudgetMappings').addEventListener('click', doSaveBudget);
  byId('floatingBudgetSaveBtn').addEventListener('click', doSaveBudget);

  // ── Create-new expense inline form ───────────────────────────
  let _newExpenseContext = null; // {cid, bmgId} — which category/merge is pending

  function showNewExpenseForm(context) {
    _newExpenseContext = context;
    byId('budgetMapNewExpenseName').value = '';
    byId('budgetMapNewExpenseType').value = 'living-expenses';
    renderExpFreqControl('living-expenses', 'budgetMapFreqRow', 'budgetMapNewExpenseFreq', null);
    byId('budgetMapNewExpenseContext').textContent =
      context.cid
        ? `For: ${S.monarchBudgets.find(b => b.categoryId === context.cid)?.categoryName || context.cid}`
        : `For merge group: ${S.budgetMergeGroups.find(g => g.id === context.bmgId)?.name || ''}`;

    // Populate plan checklist
    const planList = byId('budgetMapNewExpensePlanList');
    const allPlanIds = [...new Set(S.plExpenses.map(e => e.planId))];
    planList.innerHTML = allPlanIds.map(planId => {
      const planName = S.plExpenses.find(e => e.planId === planId)?.planName || planId.slice(0, 8);
      return `<label class="checklist-item">
        <input type="checkbox" class="new-exp-plan-cb" value="${esc(planId)}" checked/>
        <span class="checklist-item-name">${esc(planName)}</span>
      </label>`;
    }).join('');

    byId('budgetMapNewExpenseForm').classList.remove('hidden');
    byId('budgetMapNewExpenseName').focus();
    renderExpFreqControl(byId('budgetMapNewExpenseType').value, 'budgetMapFreqRow', 'budgetMapNewExpenseFreq', null);
  }

  byId('budgetMapNewExpenseType').addEventListener('change', () => {
    renderExpFreqControl(byId('budgetMapNewExpenseType').value, 'budgetMapFreqRow', 'budgetMapNewExpenseFreq', null);
  });

  function hideNewExpenseForm() {
    byId('budgetMapNewExpenseForm').classList.add('hidden');
    _newExpenseContext = null;
    renderBudgetMapping();
  }

  byId('budgetMapCancelNewExpense').addEventListener('click', hideNewExpenseForm);

  byId('budgetMapConfirmNewExpense').addEventListener('click', async () => {
    if (!_newExpenseContext) return;
    const name = byId('budgetMapNewExpenseName').value.trim();
    if (!name) { toast('Enter a name for the new expense', 'err'); return; }

    const { expType, frequency, expTypeParams } = collectExpTypeParams('budgetMap');

    const planCbs = [...byId('budgetMapNewExpensePlanList').querySelectorAll('.new-exp-plan-cb:checked')];
    if (!planCbs.length) { toast('Select at least one plan', 'err'); return; }

    const newExpenseId = crypto.randomUUID();
    const { cid, bmgId } = _newExpenseContext;

    const mapping = (planId) => ({
      planId, expenseId: '__CREATE_NEW__',
      newExpenseId, newExpenseName: name,
      newFrequency: frequency,
      newExpType: expType,
      newExpTypeParams: expTypeParams,
      createIfMissing: true,
    });

    if (cid) {
      if (!S.budgetMappings[cid]) S.budgetMappings[cid] = [];
      planCbs.forEach(cb => S.budgetMappings[cid].push(mapping(cb.value)));
    } else if (bmgId) {
      const g = S.budgetMergeGroups.find(x => x.id === bmgId);
      if (g) {
        if (!g.mappings) g.mappings = [];
        planCbs.forEach(cb => g.mappings.push(mapping(cb.value)));
      }
    }

    await save();
    hideNewExpenseForm();
    toast(`New expense "${name}" will be created on next sync`, 'ok');
  });

  // Delegate __CREATE_NEW__ selection to show the form
  // This is called from renderBudgetMapping's change listeners via a custom event
  document.addEventListener('budget-expense-create-new', e => {
    showNewExpenseForm(e.detail);
  });

  // Show floating save when legend scrolls out of view
  const legendEl = document.querySelector('#sub-panel-budgets .mapping-legend');
  const floatBar2 = byId('floatingBudgetSaveBar');
  document.addEventListener('scroll', () => {
    if (!floatBar2 || !legendEl) return;
    floatBar2.classList.toggle('hidden', legendEl.getBoundingClientRect().bottom > 0);
  }, { passive: true });
}

// ── Expense type form helpers ─────────────────────────────────

// Types with fixed, non-user-selectable frequency
// Types with fixed, non-user-selectable frequency
const EXP_FIXED_FREQ = {
  'debt':    'monthly',
  'wedding': 'once',
  'other':   'once',
};

// Suggested default frequency per type (used to pre-select the selector)
const EXP_DEFAULT_FREQ = {
  'living-expenses':   'monthly',
  'health-care':       'yearly-lump-sum',
  'charity':           'yearly-lump-sum',
  'dependent-support': 'monthly',
};

// All PL frequency options with display label and Monarch monthly → PL period conversion
const PL_FREQ_OPTIONS = [
  { value: 'yearly-lump-sum', label: 'Once Per Year', hint: '×12',     convert: m => Math.round(m * 12) },
  { value: 'yearly',          label: 'Yearly',          hint: '×12',     convert: m => Math.round(m * 12) },
  { value: 'quarterly',      label: 'Quarterly',      hint: '×3',      convert: m => Math.round(m * 3) },
  { value: 'monthly',        label: 'Monthly',        hint: '×1',      convert: m => m },
  { value: 'bi-weekly',      label: 'Bi-Weekly',      hint: '×12÷26',  convert: m => Math.round(m * 12 / 26 * 100) / 100 },
  { value: 'weekly',         label: 'Weekly',         hint: '×12÷52',  convert: m => Math.round(m * 12 / 52 * 100) / 100 },
  { value: 'daily',          label: 'Daily',          hint: '×12÷365', convert: m => Math.round(m * 12 / 365 * 100) / 100 },
];

// Render the correct frequency control into a container div.
// Fixed types get a read-only chip; variable types get the full frequency selector.
// containerId: id of the <div> to render into
// selId: id to give the <select>
// currentValue: previously saved frequency (used when editing)
function renderExpFreqControl(expType, containerId, selId, currentValue) {
  const container = byId(containerId);
  if (!container) return;
  const fixed = EXP_FIXED_FREQ[expType];
  if (fixed) {
    const label = fixed === 'once' ? 'One-time' : fixed.charAt(0).toUpperCase() + fixed.slice(1);
    container.innerHTML = `<span style="font-size:12px;color:var(--txt2);padding:3px 0;display:block">${label} <span style="color:var(--txt3);font-size:11px">(fixed)</span></span>`;
  } else {
    const def = currentValue || EXP_DEFAULT_FREQ[expType] || 'monthly';
    const opts = PL_FREQ_OPTIONS.map(o =>
      `<option value="${o.value}" ${o.value === def ? 'selected' : ''}>${o.label} — Monarch/mo ${o.hint}</option>`
    ).join('');
    container.innerHTML = `<select id="${selId}" class="select-input" style="font-size:12px;padding:5px 24px 5px 8px">${opts}</select>`;
  }
}

// Collect type + frequency from a simplified form. prefix: 'budgetMap' or 'budgetMerge'
function collectExpTypeParams(prefix) {
  const expType   = byId(`${prefix}NewExpenseType`)?.value || 'living-expenses';
  const fixed     = EXP_FIXED_FREQ[expType];
  const frequency = fixed || byId(`${prefix}NewExpenseFreq`)?.value || EXP_DEFAULT_FREQ[expType] || 'monthly';
  return { expType, frequency, expTypeParams: {} };
}

function buildExpenseOptionHTML(excludeKeys = new Set()) {
  let html = '<option value="">— select expense —</option>';
  html += '<option value="__CREATE_NEW__" class="create-opt">✦ Create new expense…</option>';
  if (!S.plExpenses.length) return html;
  const byPlan = {};
  for (const e of S.plExpenses) {
    if (!byPlan[e.planId]) byPlan[e.planId] = { name: e.planName, items: [] };
    byPlan[e.planId].items.push(e);
  }
  for (const { name, items } of Object.values(byPlan)) {
    const available = items.filter(e => !excludeKeys.has(e.planId + '::' + e.expenseId));
    if (!available.length) continue;
    html += `<optgroup label="${esc(name)}">`;
    for (const e of available) {
      const key = e.planId + '::' + e.expenseId;
      const freqLabel = e.frequency === 'monthly' ? '/mo' : (e.frequency === 'yearly' || e.frequency === 'yearly-lump-sum') ? '/yr' : `/${e.frequency}`;
      html += `<option value="${esc(key)}">${esc(e.expenseName)} (${fmt(e.amount)}${freqLabel})</option>`;
    }
    html += '</optgroup>';
  }
  return html;
}

// Renders the budget merge group rows for the budget mapping tab
function renderBudgetMergeRows(globalMappedKeys, q = '') {
  if (!S.budgetMergeGroups.length) return '';

  const allPlanIds = [...new Set(S.plExpenses.map(e => e.planId))];
  const allPlans = allPlanIds.map(id => {
    const e = S.plExpenses.find(x => x.planId === id);
    return { planId: id, planName: e?.planName || id };
  });

  return S.budgetMergeGroups.filter(g =>
    !q || g.name.toLowerCase().includes(q) ||
    g.categoryIds.some(cid => {
      const b = S.monarchBudgets.find(x => x.categoryId === cid);
      return b?.categoryName.toLowerCase().includes(q);
    })
  ).map(g => {
    const cats = g.categoryIds.map(cid => S.monarchBudgets.find(b => b.categoryId === cid)).filter(Boolean);
    const totalMonthly = cats.reduce((s, c) => s + resolveMonarchAmount(c, getGroupCategoryOverride(g, c.categoryId)), 0);
    const targets = g.mappings || [];
    const mappedKeys = new Set(targets.map(t => t.planId + '::' + t.expenseId));

    const mappedPlansByExpense = {};
    targets.forEach(t => {
      if (!mappedPlansByExpense[t.expenseId]) mappedPlansByExpense[t.expenseId] = new Set();
      mappedPlansByExpense[t.expenseId].add(t.planId);
    });

    const targetsHtml = targets.map((t, ti) => {
      const isCreateNew = t.expenseId === '__CREATE_NEW__';
      const exp = isCreateNew ? null : S.plExpenses.find(e => e.planId === t.planId && e.expenseId === t.expenseId);
      const expDef = exp || (!isCreateNew && S.plExpenses.find(e => e.expenseId === t.expenseId));
      const planName = exp?.planName
        || S.plExpenses.find(e => e.planId === t.planId)?.planName
        || (isCreateNew ? 'New expense' : 'Unknown plan');
      const expName = isCreateNew
        ? (t.newExpenseName || 'New expense')
        : (expDef?.expenseName || t.expenseName || 'Unknown expense');
      const syncExpDef = isCreateNew
        ? { frequency: t.newFrequency || 'monthly', expType: t.newExpType || 'living-expenses' }
        : expDef;
      const { syncAmount, freqLabel, warn: syncWarn } = computeBudgetSync(syncExpDef, totalMonthly);
      const badge = isCreateNew
        ? ' <span style="color:var(--amber);font-size:11px;font-weight:700">new</span>'
        : t.createIfMissing
        ? ' <span style="color:var(--amber);font-size:11px">new</span>'
        : '';

      let applyHtml = '';
      if (ti === 0) {
        const alreadyMapped = mappedPlansByExpense[t.expenseId] || new Set();
        const remainingPlans = allPlans.filter(p => !alreadyMapped.has(p.planId));
        if (remainingPlans.length) {
          const opts = `<option value="__ALL__">⚡ All remaining plans</option>
            ${remainingPlans.map(p => {
              const existsInPlan = S.plExpenses.some(x => x.planId === p.planId && x.expenseId === t.expenseId);
              return `<option value="${esc(p.planId)}">${esc(p.planName)}${existsInPlan ? '' : ' (new)'}</option>`;
            }).join('')}`;
          applyHtml = `<select class="budget-plan-apply" data-bmg-id="${esc(g.id)}" data-expense-id="${esc(t.expenseId)}" multiple size="${Math.min(remainingPlans.length + 1, 4)}">
            ${opts}
          </select>`;
        }
      }

      return `<div class="budget-target-row" data-bmg-id="${esc(g.id)}" data-ti="${ti}">
        <span class="budget-target-label">
          <span class="budget-target-plan">${esc(planName)}</span>
          <span class="budget-target-expense">→ ${esc(expName)}${badge} <span class="budget-target-freq">${fmt(syncAmount)}${freqLabel}</span>${syncWarn ? '<span class="sync-warn-icon" title="' + esc(syncWarn) + '" style="cursor:help;color:var(--amber);font-size:11px;margin-left:3px">⚠</span>' : ''}${t.lastSynced ? `<span class="sync-check-badge" title="Last synced ${syncedAgo(t.lastSynced)}">✓</span>` : ''}</span>
        </span>
        ${applyHtml}
        <button class="btn btn-sm btn-ghost budget-remove-target" data-bmg-id="${esc(g.id)}" data-ti="${ti}" title="Remove">✕</button>
      </div>`;
    }).join('');

    const showDropdown = targets.length === 0;

    return `<div class="budget-row budget-row--merge">
      <div class="budget-row-header">
        <span class="budget-source-name">🔀 ${esc(g.name)}</span>
        <span class="budget-source-amount">${fmt(totalMonthly)}<span style="color:var(--txt3);font-size:10px;margin-left:3px">${g.sourceModeAll === false ? 'per-cat' : budgetSourceLabel(g.sourceOverride || null)}</span></span>
      </div>
      <div style="font-size:12px;color:var(--txt2);padding:0 0 4px 0">${cats.map(c => esc(c.categoryName)).join(' + ')}</div>
      ${targets.length ? `<div class="budget-targets">${targetsHtml}</div>` : ''}
      ${showDropdown ? `<div class="budget-add-row">
        <select class="budget-expense-sel" data-bmg-id="${esc(g.id)}">
          ${buildExpenseOptionHTML(globalMappedKeys)}
        </select>
      </div>` : ''}
    </div>`;
  }).join('');
}

function renderBudgetMapping() {
  const list = byId('budgetMappingList');
  if (!S.monarchBudgets.length) {
    list.innerHTML = '<div class="empty-state">Fetch Monarch accounts in Setup to load budget categories</div>';
    return;
  }
  if (!S.plExpenses.length) {
    list.innerHTML = '<div class="empty-state">Load PL Accounts in Setup to load plan expenses</div>';
    return;
  }

  // Build global mapped keys covering both individual mappings AND merge group mappings
  const globalMappedKeys = new Set([
    ...Object.values(S.budgetMappings).flat(),
    ...S.budgetMergeGroups.flatMap(g => g.mappings || []),
  ].map(t => t.planId + '::' + t.expenseId));

  // Render budget merge groups first, then individual categories
  const q = (byId('budgetSearch')?.value || '').toLowerCase();
  const mergeRowsHtml = renderBudgetMergeRows(globalMappedKeys, q);

  // Filter out categories already in a merge group, then apply search
  const mergedCategoryIds = new Set(S.budgetMergeGroups.flatMap(g => g.categoryIds));

  const filteredBudgets = S.monarchBudgets
    .filter(b => !mergedCategoryIds.has(b.categoryId))
    .filter(b => !q || b.categoryName.toLowerCase().includes(q) || b.groupName.toLowerCase().includes(q));

  // Group by groupName, preserving order of first appearance
  const groupOrder = [];
  const grouped = {};
  for (const b of filteredBudgets) {
    const gn = b.groupName || 'Other';
    if (!grouped[gn]) { grouped[gn] = []; groupOrder.push(gn); }
    grouped[gn].push(b);
  }

  // All available plans (used inside the map below)
  const allPlanIds = [...new Set(S.plExpenses.map(e => e.planId))];
  const allPlans = allPlanIds.map(id => {
    const e = S.plExpenses.find(x => x.planId === id);
    return { planId: id, planName: e?.planName || id };
  });

  const individualRowsHtml = groupOrder.map(groupName => {
    const rows = grouped[groupName].map(b => {
    const targets = S.budgetMappings[b.categoryId] || [];
    const mappedKeys = new Set(targets.map(t => t.planId + '::' + t.expenseId));
    const mappedPlansByExpense = {};
    targets.forEach(t => {
      if (!mappedPlansByExpense[t.expenseId]) mappedPlansByExpense[t.expenseId] = new Set();
      mappedPlansByExpense[t.expenseId].add(t.planId);
    });

    const catOverride = getEffectiveBudgetOverride(b.categoryId);
    const resolvedAmount = resolveMonarchAmount(b, catOverride);
    const srcLabel = budgetSourceLabel(catOverride);
    const overrideMode = catOverride?.mode || '';
    const defaultLookback = overrideMode === 'actual' ? 3 : 1;
    const overrideLookback = catOverride?.lookback ?? defaultLookback;
    const overrideAgg = catOverride?.aggregation || 'average';

    const targetsHtml = targets.map((t, ti) => {
      const isCreateNew = t.expenseId === '__CREATE_NEW__';
      const exp = isCreateNew ? null : S.plExpenses.find(e => e.planId === t.planId && e.expenseId === t.expenseId);
      const expDef = exp || (!isCreateNew && S.plExpenses.find(e => e.expenseId === t.expenseId));
      const planName = exp?.planName || S.plExpenses.find(e => e.planId === t.planId)?.planName || 'Unknown plan';
      const expName = isCreateNew
        ? (t.newExpenseName || 'New expense')
        : (expDef?.expenseName || t.expenseName || 'Unknown expense');
      const syncExpDef = isCreateNew
        ? { frequency: t.newFrequency || 'monthly', expType: t.newExpType || 'living-expenses' }
        : expDef;
      const { syncAmount, freqLabel, warn: syncWarn } = computeBudgetSync(syncExpDef, resolvedAmount);
      const badge = isCreateNew
        ? ' <span style="color:var(--amber);font-size:11px;font-weight:700">new</span>'
        : t.createIfMissing
        ? ' <span style="color:var(--amber);font-size:11px">new</span>'
        : '';
      let applyHtml = '';
      if (ti === 0) {
        const alreadyMapped = mappedPlansByExpense[t.expenseId] || new Set();
        const remainingPlans = allPlans.filter(p => !alreadyMapped.has(p.planId));
        if (remainingPlans.length) {
          const opts = `<option value="__ALL__">⚡ All remaining plans</option>
            ${remainingPlans.map(p => {
              const existsInPlan = S.plExpenses.some(x => x.planId === p.planId && x.expenseId === t.expenseId);
              return `<option value="${esc(p.planId)}">${esc(p.planName)}${existsInPlan ? '' : ' (new)'}</option>`;
            }).join('')}`;
          applyHtml = `<select class="budget-plan-apply" data-cid="${esc(b.categoryId)}" data-expense-id="${esc(t.expenseId)}" multiple size="${Math.min(remainingPlans.length + 1, 4)}">
            ${opts}
          </select>`;
        }
      }
      return `<div class="budget-target-row" data-cid="${esc(b.categoryId)}" data-ti="${ti}">
        <span class="budget-target-label">
          <span class="budget-target-plan">${esc(planName)}</span>
          <span class="budget-target-expense">→ ${esc(expName)}${badge} <span class="budget-target-freq">${fmt(syncAmount)}${freqLabel}</span>${syncWarn ? '<span class="sync-warn-icon" title="' + esc(syncWarn) + '" style="cursor:help;color:var(--amber);font-size:11px;margin-left:3px">⚠</span>' : ''}${t.lastSynced ? `<span class="sync-check-badge" title="Last synced ${syncedAgo(t.lastSynced)}">✓</span>` : ''}</span>
        </span>
        ${applyHtml}
        <button class="btn btn-sm btn-ghost budget-remove-target" data-cid="${esc(b.categoryId)}" data-ti="${ti}" title="Remove">✕</button>
      </div>`;
    }).join('');

    const actualOptsHidden = !overrideMode ? 'hidden' : '';
    const sourceControlsHtml = `<div class="budget-source-override">
      <span class="budget-override-label">Source:</span>
      <select class="bgt-mode-sel select-input" data-cid="${esc(b.categoryId)}">
        <option value="" ${!overrideMode ? 'selected' : ''}>Default</option>
        <option value="planned" ${overrideMode === 'planned' ? 'selected' : ''}>Planned</option>
        <option value="actual" ${overrideMode === 'actual' ? 'selected' : ''}>Actual</option>
      </select>
      <span class="bgt-actual-opts ${actualOptsHidden}">
        <input class="bgt-lookback text-input" type="number" min="1" max="12"
               value="${overrideLookback}" data-cid="${esc(b.categoryId)}"
               style="width:40px;text-align:center;padding:2px 4px;font-size:11px"/>
        <span style="font-size:10px;color:var(--txt3)">mo</span>
        <select class="bgt-agg-sel select-input" data-cid="${esc(b.categoryId)}">
          <option value="average" ${overrideAgg === 'average' ? 'selected' : ''}>avg</option>
          <option value="median" ${overrideAgg === 'median' ? 'selected' : ''}>median</option>
        </select>
      </span>
    </div>`;

    return `<div class="budget-row">
      <div class="budget-row-header">
        <span class="budget-source-name">${esc(b.categoryName)}</span>
        <span class="budget-source-amount">${fmt(resolvedAmount)}<span style="color:var(--txt3);font-size:10px;margin-left:3px">${srcLabel}</span></span>
      </div>
      ${sourceControlsHtml}
      ${targets.length ? `<div class="budget-targets">${targetsHtml}</div>` : ''}
      ${!targets.length ? `<div class="budget-add-row">
        <select class="budget-expense-sel" data-cid="${esc(b.categoryId)}">
          ${buildExpenseOptionHTML(globalMappedKeys)}
        </select>
      </div>` : ''}
    </div>`;
    }).join('');

    const grpOvr     = S.budgetGroupSourceOverrides[groupName] || null;
    const grpActive  = !!grpOvr;
    const grpMode    = grpOvr?.mode || '';
    const grpDefaultLookback = grpMode === 'actual' ? 3 : 1;
    const grpLookback = grpOvr?.lookback ?? grpDefaultLookback;
    const grpAgg     = grpOvr?.aggregation || 'average';
    const grpActualOptsHidden = !grpMode ? 'hidden' : '';
    const grpInlineControls = grpActive ? `
        <div class="budget-group-src-inline">
          <select class="bgt-group-mode-sel select-input" data-group="${esc(groupName)}">
            <option value="" ${!grpMode ? 'selected' : ''}>Default</option>
            <option value="planned" ${grpMode === 'planned' ? 'selected' : ''}>Planned</option>
            <option value="actual" ${grpMode === 'actual' ? 'selected' : ''}>Actual</option>
          </select>
          <span class="bgt-actual-opts ${grpActualOptsHidden}">
            <input class="bgt-group-lookback text-input" type="number" min="1" max="12"
                   value="${grpLookback}" data-group="${esc(groupName)}"
                   style="width:40px;text-align:center;padding:2px 4px;font-size:11px"/>
            <span style="font-size:10px;color:var(--txt3)">mo</span>
            <select class="bgt-group-agg-sel select-input" data-group="${esc(groupName)}">
              <option value="average" ${grpAgg === 'average' ? 'selected' : ''}>avg</option>
              <option value="median" ${grpAgg === 'median' ? 'selected' : ''}>median</option>
            </select>
          </span>
        </div>` : '';
    const isMappingGroupCollapsed = _collapsedBudgetMappingGroups.has(groupName);
    return `<div class="budget-group${grpActive ? ' budget-group--override' : ''}${isMappingGroupCollapsed ? ' collapsed' : ''}" data-group="${esc(groupName)}">
      <div class="budget-group-header">
        <span class="budget-group-name">${esc(groupName)}</span>
        ${grpInlineControls}
        <label class="budget-group-src-toggle" title="Apply one source setting to this entire group">
          <input type="checkbox" class="bgt-group-toggle" data-group="${esc(groupName)}" ${grpActive ? 'checked' : ''}/>
          <span>Group source</span>
        </label>
        <button type="button" class="budget-group-collapse-btn" data-group="${esc(groupName)}" title="Collapse / expand">${isMappingGroupCollapsed ? '▸' : '▾'}</button>
      </div>
      <div class="budget-group-body">
        ${rows}
      </div>
    </div>`;
  }).join('');

  list.innerHTML = mergeRowsHtml + individualRowsHtml;

}

// One-time delegated listeners for budgetMappingList — called from setupBudgetMapping()
function setupBudgetMappingListDelegation() {
  const list = byId('budgetMappingList');

  list.addEventListener('change', async e => {
    // Group-level source toggle
    const groupToggle = e.target.closest('.bgt-group-toggle');
    if (groupToggle) {
      const gn = groupToggle.dataset.group;
      if (groupToggle.checked) {
        S.budgetGroupSourceOverrides[gn] = { mode: '', lookback: 1, aggregation: 'average' };
      } else {
        delete S.budgetGroupSourceOverrides[gn];
      }
      await save(); renderBudgetMapping(); return;
    }

    // Group-level source mode
    const grpModeEl = e.target.closest('.bgt-group-mode-sel');
    if (grpModeEl) {
      const gn   = grpModeEl.dataset.group;
      const mode = grpModeEl.value;
      const row  = grpModeEl.closest('.budget-group-src-inline');
      if (row) row.querySelector('.bgt-actual-opts')?.classList.toggle('hidden', !mode);
      if (!S.budgetGroupSourceOverrides[gn]) S.budgetGroupSourceOverrides[gn] = {};
      const defaultLookback = mode === 'actual' ? 3 : 1;
      S.budgetGroupSourceOverrides[gn] = { ...S.budgetGroupSourceOverrides[gn], mode, lookback: defaultLookback };
      await save(); renderBudgetMapping(); return;
    }

    // Group-level lookback
    const grpLookbackEl = e.target.closest('.bgt-group-lookback');
    if (grpLookbackEl) {
      const gn = grpLookbackEl.dataset.group;
      const v  = Math.max(1, Math.min(12, parseInt(grpLookbackEl.value, 10) || 1));
      grpLookbackEl.value = v;
      if (!S.budgetGroupSourceOverrides[gn]) S.budgetGroupSourceOverrides[gn] = {};
      S.budgetGroupSourceOverrides[gn] = { ...S.budgetGroupSourceOverrides[gn], lookback: v };
      await save(); renderBudgetMapping(); return;
    }

    // Group-level aggregation
    const grpAggEl = e.target.closest('.bgt-group-agg-sel');
    if (grpAggEl) {
      const gn = grpAggEl.dataset.group;
      if (!S.budgetGroupSourceOverrides[gn]) S.budgetGroupSourceOverrides[gn] = {};
      S.budgetGroupSourceOverrides[gn] = { ...S.budgetGroupSourceOverrides[gn], aggregation: grpAggEl.value };
      await save(); renderBudgetMapping(); return;
    }

    // Per-category source mode override
    const modeEl = e.target.closest('.bgt-mode-sel');
    if (modeEl) {
      const cid = modeEl.dataset.cid;
      const mode = modeEl.value;
      // Toggle actual options visibility immediately in the DOM before re-render
      const row = modeEl.closest('.budget-row');
      if (row) row.querySelector('.bgt-actual-opts')?.classList.toggle('hidden', !mode);
      if (!mode) {
        delete S.budgetSourceOverrides[cid];
      } else {
        const existing = S.budgetSourceOverrides[cid] || {};
        const defaultLb = mode === 'actual' ? 3 : 1;
        S.budgetSourceOverrides[cid] = { ...existing, mode, lookback: existing.lookback ?? defaultLb };
      }
      await save();
      renderBudgetMapping();
      return;
    }

    // Per-category lookback override
    const lookbackEl = e.target.closest('.bgt-lookback');
    if (lookbackEl) {
      const cid = lookbackEl.dataset.cid;
      const v = Math.max(1, Math.min(12, parseInt(lookbackEl.value, 10) || 3));
      lookbackEl.value = v;
      S.budgetSourceOverrides[cid] = { ...(S.budgetSourceOverrides[cid] || {}), lookback: v };
      await save();
      renderBudgetMapping();
      return;
    }

    // Per-category aggregation override
    const aggEl = e.target.closest('.bgt-agg-sel');
    if (aggEl) {
      const cid = aggEl.dataset.cid;
      S.budgetSourceOverrides[cid] = { ...(S.budgetSourceOverrides[cid] || {}), aggregation: aggEl.value };
      await save();
      renderBudgetMapping();
      return;
    }

    // Expense selector — commits immediately on single selection
    const expSel = e.target.closest('.budget-expense-sel');
    if (expSel) {
      const key   = expSel.value;
      if (!key) return;
      const bmgId = expSel.dataset.bmgId;
      const cid   = expSel.dataset.cid;
      if (key === '__CREATE_NEW__') {
        expSel.value = '';
        document.dispatchEvent(new CustomEvent('budget-expense-create-new', { detail: { cid: cid || null, bmgId: bmgId || null } }));
        return;
      }
      const [planId, expenseId] = key.split('::');
      if (bmgId) {
        const g = S.budgetMergeGroups.find(x => x.id === bmgId);
        if (!g) return;
        if (!g.mappings) g.mappings = [];
        if (g.mappings.some(t => t.planId === planId && t.expenseId === expenseId)) return;
        g.mappings.push({ planId, expenseId });
      } else {
        if (!S.budgetMappings[cid]) S.budgetMappings[cid] = [];
        if (S.budgetMappings[cid].some(t => t.planId === planId && t.expenseId === expenseId)) return;
        S.budgetMappings[cid].push({ planId, expenseId });
      }
      await save();
      const status = byId('budgetSaveStatus');
      if (status) { status.textContent = '✓ Saved'; setTimeout(() => { status.textContent = ''; }, 3500); }
      renderBudgetMapping();
      return;
    }

    // Apply-to-plans multi-select
    const planApply = e.target.closest('.budget-plan-apply');
    if (planApply) {
      const expenseId = planApply.dataset.expenseId;
      const bmgId     = planApply.dataset.bmgId;
      const cid       = planApply.dataset.cid;
      let targetArr;
      if (bmgId) {
        const g = S.budgetMergeGroups.find(x => x.id === bmgId);
        if (!g) return;
        if (!g.mappings) g.mappings = [];
        targetArr = g.mappings;
      } else {
        if (!S.budgetMappings[cid]) S.budgetMappings[cid] = [];
        targetArr = S.budgetMappings[cid];
      }
      const allPlanIds    = [...new Set(S.plExpenses.map(e => e.planId))];
      const alreadyMapped = new Set(targetArr.filter(t => t.expenseId === expenseId).map(t => t.planId));
      const selectedPlanIds = [];
      for (const opt of planApply.selectedOptions) {
        if (opt.value === '__ALL__') {
          allPlanIds.filter(pid => !alreadyMapped.has(pid)).forEach(pid => selectedPlanIds.push(pid));
        } else if (opt.value) {
          selectedPlanIds.push(opt.value);
        }
      }
      let added = false;
      for (const planId of selectedPlanIds) {
        if (targetArr.some(t => t.planId === planId && t.expenseId === expenseId)) continue;
        const existsInPlan = S.plExpenses.some(e => e.planId === planId && e.expenseId === expenseId);
        targetArr.push({ planId, expenseId, createIfMissing: !existsInPlan });
        added = true;
      }
      if (added) {
        await save();
        const status = byId('budgetSaveStatus');
        if (status) { status.textContent = '✓ Saved'; setTimeout(() => { status.textContent = ''; }, 3500); }
        renderBudgetMapping();
      }
      return;
    }
  });

  list.addEventListener('click', async e => {
    const collapseBtn = e.target.closest('.budget-group-collapse-btn');
    if (collapseBtn) {
      const gn = collapseBtn.dataset.group;
      const grpDiv = list.querySelector(`.budget-group[data-group="${CSS.escape(gn)}"]`);
      if (!grpDiv) return;
      const collapsed = grpDiv.classList.toggle('collapsed');
      collapseBtn.textContent = collapsed ? '▸' : '▾';
      collapsed ? _collapsedBudgetMappingGroups.add(gn) : _collapsedBudgetMappingGroups.delete(gn);
      return;
    }

    const removeBtn = e.target.closest('.budget-remove-target');
    if (removeBtn) {
      const ti    = parseInt(removeBtn.dataset.ti, 10);
      const bmgId = removeBtn.dataset.bmgId;
      const cid   = removeBtn.dataset.cid;
      if (bmgId) {
        const g = S.budgetMergeGroups.find(x => x.id === bmgId);
        if (g?.mappings) g.mappings.splice(ti, 1);
      } else if (S.budgetMappings[cid]) {
        S.budgetMappings[cid].splice(ti, 1);
        if (!S.budgetMappings[cid].length) delete S.budgetMappings[cid];
      }
      await save();
      renderBudgetMapping();
    }
  });
}

// ── SYNC tab ─────────────────────────────────────────────────
let syncPlan = [];

function setupSync() {
  byId('previewSyncBtn').addEventListener('click', buildSyncPlan);
  byId('runSyncBtn').addEventListener('click', runSync);
  byId('floatingPushBtn').addEventListener('click', runSync);

  // Mirror disabled state of runSyncBtn to floatingPushBtn
  const observer = new MutationObserver(() => {
    const disabled = byId('runSyncBtn').disabled;
    const btn = byId('floatingPushBtn');
    btn.disabled = disabled;
    // Hide the bar whenever the action button becomes disabled
    if (disabled) byId('floatingPushBar').classList.add('hidden');
  });
  observer.observe(byId('runSyncBtn'), { attributes: true, attributeFilter: ['disabled'] });

  // Show floating push bar when sync list has scrolled past the legend
  const syncPreviewEl = byId('syncPreviewList');
  const floatPushBar = byId('floatingPushBar');
  const floatPushBtn = byId('floatingPushBtn');
  syncPreviewEl.addEventListener('scroll', () => {
    const scrolled = syncPreviewEl.scrollTop > 10;
    floatPushBar.classList.toggle('hidden', !scrolled || floatPushBtn.disabled);
  }, { passive: true });
  // Also update MutationObserver to sync disabled AND hidden state
  // (handled below by updating the existing observer)
}

function buildSyncPlan() {
  // Save current mappings so Preview and Push always work from persisted state
  save().then(() => {
    const status = byId('mappingSaveStatus');
    if (status) { status.textContent = '✓ Saved'; setTimeout(() => { status.textContent = ''; }, 3500); }
  });
  syncPlan = [];
  const mergedIds = new Set(S.mergeGroups.flatMap(g => g.monarchIds));

  // 1. Pending creations — skip any that have a recorded failure (require user to retry)
  for (const [monarchId, creation] of Object.entries(S.pendingCreations)) {
    const ma = S.monarchAccounts.find(a => a.id === monarchId);
    if (!ma) continue;
    if (creation.failedAt) continue; // failed — user must retry or clear from Mapping tab
    syncPlan.push({
      type: 'create', monarchId, plCategory: creation.plCategory, plType: creation.plType,
      plLabel: `${creation.name} (new ${CATEGORY_LABELS[creation.plCategory]})`,
      monarchNames: [ma.name], writeAmount: creation.balance,
      name: creation.name, status: 'pending',
    });
  }

  // 2. Merge groups
  for (const g of S.mergeGroups) {
    if (g.pendingCreate) {
      // CREATE-MERGE: create a new PL asset and map each Monarch account to a specific field
      if (g.fieldMappings) {
        // Multi-field (e.g. House: current value + loan balance)
        const monarchNames = g.fieldMappings.map(fm => {
          const a = S.monarchAccounts.find(ac => ac.id === fm.monarchId);
          return a ? a.name : fm.monarchId;
        });
        syncPlan.push({
          type:          'create-merge-multi',
          mergeGroupId:  g.id,
          pendingCreate: g.pendingCreate,
          fieldMappings: g.fieldMappings,
          plLabel:       `${g.pendingCreate.name} (new ${CATEGORY_LABELS[g.pendingCreate.plCategory]})`,
          monarchNames,
          writeAmount:   g.fieldMappings.reduce((s, fm) => s + (fm.balance || 0), 0),
          status:        'pending',
        });
      } else {
        // Single-field create-merge (savings/investment/debt)
        const sources = S.monarchAccounts.filter(a => g.monarchIds.includes(a.id));
        const total   = sources.reduce((s, a) => s + (a.balance || 0), 0);
        syncPlan.push({
          type:          'create-merge-single',
          mergeGroupId:  g.id,
          pendingCreate: g.pendingCreate,
          monarchIds:    g.monarchIds,
          plLabel:       `${g.pendingCreate.name} (new ${CATEGORY_LABELS[g.pendingCreate.plCategory]})`,
          monarchNames:  sources.map(a => a.name),
          writeAmount:   total,
          status:        'pending',
        });
      }
    } else {
      // UPDATE existing PL account
      if (!g.plRealId) continue;

      // Existing asset with field-level mappings (multi-field merge into existing asset)
      if (g.fieldMappings && !g.plField) {
        const monarchNames = g.fieldMappings.map(fm => {
          const a = S.monarchAccounts.find(ac => ac.id === fm.monarchId);
          return a ? a.name : fm.monarchId;
        });
        const plAcct = S.plAccounts.find(p => p.id === g.plId);
        const plName = plAcct?.name || g.plRealId;
        syncPlan.push({
          type:          'merge-multi',
          mergeGroupId:  g.id,
          plRealId:      g.plRealId,
          fieldMappings: g.fieldMappings,
          plLabel:       plName,
          plCategory:    'asset',
          monarchNames,
          writeAmount:   g.fieldMappings.reduce((s, fm) => s + (fm.balance || 0), 0),
          status:        'pending',
        });
        continue;
      }

      // Single-field merge into existing account
      if (!g.plField) continue;

      const sources = S.monarchAccounts.filter(a => g.monarchIds.includes(a.id));
      const total   = sources.reduce((s, a) => s + (a.balance || 0), 0);
      const plAcct  = S.plAccounts.find(p => p.id === g.plId);
      const plLabel = plAcct ? (plAcct.subLabel ? `${plAcct.name} → ${plAcct.subLabel}` : plAcct.name) : g.plRealId;
      const mergeField = correctPlField(g.plRealId, g.plField);
      syncPlan.push({
        type: 'merge', plRealId: g.plRealId, plField: mergeField, plLabel,
        plCategory: plAcct?.category || null,
        monarchNames: sources.map(a => a.name), writeAmount: total, status: 'pending',
      });
    }
  }

  // 3. One-to-one mappings
  for (const [monarchId, mapping] of Object.entries(S.mappings)) {
    if (mergedIds.has(monarchId) || !mapping?.plId) continue;
    const ma     = S.monarchAccounts.find(a => a.id === monarchId);
    const plAcct = S.plAccounts.find(p => p.id === mapping.plId);
    if (!ma) continue;
    const plLabel = plAcct ? (plAcct.subLabel ? `${plAcct.name} → ${plAcct.subLabel}` : plAcct.name) : mapping.plRealId;
    const correctedField = correctPlField(mapping.plRealId, mapping.plField);
    syncPlan.push({
      type: 'single', plRealId: mapping.plRealId, plField: correctedField, plLabel,
      plCategory: plAcct?.category || null,
      monarchNames: [ma.name], writeAmount: ma.balance, status: 'pending',
    });
  }

  // 4a. Budget merge groups → plan expenses (summed amounts)
  for (const g of S.budgetMergeGroups) {
    if (!g.mappings?.length) continue;
    const totalMonthly = g.categoryIds.reduce((s, cid) => {
      const b = S.monarchBudgets.find(x => x.categoryId === cid);
      return s + (b ? resolveMonarchAmount(b, getGroupCategoryOverride(g, cid)) : 0);
    }, 0);
    for (const { planId, expenseId, createIfMissing, newExpenseId, newExpenseName, newFrequency, newExpType, newExpTypeParams } of g.mappings) {
      const isNewExp = expenseId === '__CREATE_NEW__';
      const exp = isNewExp ? null : S.plExpenses.find(e => e.planId === planId && e.expenseId === expenseId);
      const expDef = exp || (!isNewExp && S.plExpenses.find(e => e.expenseId === expenseId));
      const expName = isNewExp ? (newExpenseName || 'New Expense') : (expDef?.expenseName || expenseId.slice(0,8));
      if (!isNewExp && !expDef) continue;
      const { writeField, syncAmount, freqLabel, warn } = isNewExp
        ? computeBudgetSync({ frequency: newFrequency || 'monthly', expType: newExpType || 'living-expenses' }, totalMonthly)
        : computeBudgetSync(expDef, totalMonthly);
      const planName = exp?.planName || S.plExpenses.find(e => e.planId === planId)?.planName || planId.slice(0,8);
      syncPlan.push({
        type:           'budget',
        categoryId:     g.id,
        planId, expenseId,
        writeField,
        createIfMissing: createIfMissing || isNewExp,
        newExpenseId, newExpenseName, newFrequency, newExpType, newExpTypeParams,
        plLabel:        `${planName} → ${expName}${(createIfMissing || isNewExp) ? ' (new)' : ''}`,
        warnText:       warn || null,
        monarchNames:   [g.name],
        writeAmount:    syncAmount,
        displayAmount:  fmt(syncAmount) + freqLabel,
        status:         'pending',
      });
    }
  }

  // 4b. Individual budget mappings → plan expenses
  for (const [categoryId, targets] of Object.entries(S.budgetMappings)) {
    const budget = S.monarchBudgets.find(b => b.categoryId === categoryId);
    if (!budget) continue;
    for (const { planId, expenseId, createIfMissing, newExpenseId, newExpenseName, newFrequency, newExpType, newExpTypeParams } of targets) {
      const isNewExp = expenseId === '__CREATE_NEW__';
      const exp = isNewExp ? null : S.plExpenses.find(e => e.planId === planId && e.expenseId === expenseId);
      const expDef = exp || (!isNewExp && S.plExpenses.find(e => e.expenseId === expenseId));
      const expName = isNewExp ? (newExpenseName || 'New Expense') : (expDef?.expenseName || expenseId.slice(0,8));
      if (!isNewExp && !expDef) continue;
      const catOverride = getEffectiveBudgetOverride(categoryId);
      const { writeField, syncAmount, freqLabel, warn } = isNewExp
        ? computeBudgetSync({ frequency: newFrequency || 'monthly', expType: newExpType || 'living-expenses' }, resolveMonarchAmount(budget, catOverride))
        : computeBudgetSync(expDef, resolveMonarchAmount(budget, catOverride));
      const planName = exp?.planName || S.plExpenses.find(e => e.planId === planId)?.planName || planId;
      syncPlan.push({
        type:          'budget',
        categoryId,
        planId, expenseId,
        writeField,
        createIfMissing: createIfMissing || isNewExp,
        newExpenseId, newExpenseName, newFrequency, newExpType, newExpTypeParams,
        plLabel:       `${planName} → ${expName}${(createIfMissing || isNewExp) ? ' (new)' : ''}`,
        warnText:      warn || null,
        monarchNames:  [budget.categoryName],
        writeAmount:   syncAmount,
        displayAmount: fmt(syncAmount) + freqLabel,
        status:        'pending',
      });
    }
  }

  renderSync();
  byId('runSyncBtn').disabled = syncPlan.length === 0;
  if (!syncPlan.length) toast('Nothing to sync — create mappings first', 'err');
}

function renderSync() {
  const list = byId('syncPreviewList');
  if (!syncPlan.length) {
    list.innerHTML = '<div class="empty-state">Click "Refresh Preview" to load</div>';
    byId('runSyncBtn').disabled = true;
    return;
  }
  // Field label helper — 'balance' means "Loan Balance" for assets, just "Balance" for savings/investments
  function fieldLabel(plField, plCategory) {
    if (plField === 'amount')       return plCategory === 'debt' ? 'Debt Amount' : 'Current Value';
    if (plField === 'initialValue') return 'Purchase Price';
    if (plField === 'balance')      return plCategory === 'asset' ? 'Loan Balance' : 'Balance';
    return plField || '';
  }

  list.innerHTML = syncPlan.map((item, i) => {
    const statusColor = { pending:'var(--txt2)', pushing:'var(--amber)', done:'var(--green)', failed:'var(--red)' }[item.status];
    const isBudget = item.type === 'budget';
    const isCreate = item.type === 'create' || item.type === 'create-merge-single' || item.type === 'create-merge-multi';
    const typeLabel = isCreate
      ? `✦ New ${isBudget ? 'expense' : 'account'}`
      : item.type === 'merge-multi'
      ? '→ multi-field update'
      : isBudget
      ? '💰 Budget → Expense'
      : `→ ${fieldLabel(item.plField, item.plCategory)}`;
    const itemKind = isBudget ? 'budget' : 'account';
    return `
      <div class="sync-row ${item.status}" data-item-kind="${itemKind}" data-idx="${i}">
        <div class="sync-source">
          <div class="sync-source-name">${item.monarchNames.map(esc).join(' + ')}</div>
          <div class="sync-source-ids">${typeLabel}</div>
        </div>
        <div class="sync-arrow">→</div>
        <div class="sync-dest">
          <div class="sync-dest-name">${esc(item.plLabel)}${item.warnText ? ` <span class="sync-warn-icon" title="${esc(item.warnText)}" style="cursor:help;color:var(--amber);font-size:12px" aria-label="${esc(item.warnText)}">⚠</span>` : ''}</div>
          <div class="sync-dest-bal" style="color:${statusColor}">${
            item.status === 'done'    ? '✓ done' :
            item.status === 'failed'  ? '✗ ' + esc(item.error || 'error') :
            item.status === 'pushing' ? '⟳ pushing…' :
            item.type === 'budget' ? item.displayAmount :
            (item.type === 'create-merge-multi' || item.type === 'merge-multi')
              ? item.fieldMappings.map((fm, fi) => {
                  const FL = { amount:'Current Value', balance:'Loan Balance', initialValue:'Purchase Price' };
                  const acctName = item.monarchNames[fi] || '';
                  return '<span style="font-size:12px;display:block;color:var(--txt2)">'
                    + esc(acctName) + ' → ' + (FL[fm.writeField] || fm.writeField) + ': ' + fmt(fm.balance)
                    + '</span>';
                }).join('')
              : fmt(item.writeAmount)
          }</div>
        </div>
      </div>`;
  }).join('');
}

let _syncRunning = false;

async function runSync() {
  if (!S.plApiKey) { toast('No PL API key — check Setup tab', 'err'); return; }
  if (!syncPlan.length) { buildSyncPlan(); return; }
  if (_syncRunning) { toast('Sync already in progress', 'err'); return; }

  // Disable buttons synchronously before any await to prevent double-push
  _syncRunning = true;
  byId('runSyncBtn').disabled = true;
  byId('previewSyncBtn').disabled = true;

  // Step 1: batch ALL account creations into ONE export→modify→restore.
  // This prevents the snapshot-overwrite bug where multiple PL_CREATE_ACCOUNTS calls
  // each re-export PL state, causing earlier creations to be lost.
  const creations = syncPlan.filter(i => i.type === 'create');
  const createMergeItems = syncPlan.filter(i =>
    i.type === 'create-merge-single' || i.type === 'create-merge-multi'
  );
  const allCreationItems = [...creations, ...createMergeItems];

  let ok   = 0;
  let fail = 0;

  if (allCreationItems.length) {
    allCreationItems.forEach(i => { i.status = 'pushing'; });
    renderSync();

    // Build one flat array of account specs for the service worker.
    // Each entry carries a _ref back to its sync item so we can route returned IDs.
    const accountSpecs = [];

    for (const item of creations) {
      accountSpecs.push({
        _ref:         item,
        _type:        'simple',
        monarchId:    item.monarchId,
        name:         item.name,
        balance:      item.writeAmount,
        plCategory:   item.plCategory,
        plType:       item.plType,
        writeField:   item.writeField    || null,
        paymentStatus: item.paymentStatus || null,
        manualValues:  item.manualValues  || {},
        monarchNames:  S.settings.writeNotes !== false ? item.monarchNames : [],
      });
    }

    for (const item of createMergeItems) {
      const pc = item.pendingCreate;
      if (item.type === 'create-merge-multi') {
        const firstFm = item.fieldMappings[0];
        const mergeMonarchNames = item.fieldMappings.map(fm => {
          const ma = S.monarchAccounts.find(a => a.id === fm.monarchId);
          return ma ? ma.name : fm.monarchId;
        });
        accountSpecs.push({
          _ref:         item,
          _type:        'merge-multi',
          _fieldMappings: item.fieldMappings,
          monarchId:    firstFm.monarchId,
          name:         pc.name,
          balance:      firstFm.balance,
          plCategory:   pc.plCategory,
          plType:       pc.plType,
          writeField:   firstFm.writeField,
          paymentStatus: pc.paymentStatus || 'financed',
          manualValues:  pc.manualValues  || {},
          monarchNames:  mergeMonarchNames,
        });
      } else {
        // create-merge-single
        const sources = S.monarchAccounts.filter(a => item.monarchIds.includes(a.id));
        const total   = sources.reduce((s, a) => s + (a.balance || 0), 0);
        const writeField = (pc.plCategory === 'debt') ? 'amount' : 'balance';
        accountSpecs.push({
          _ref:         item,
          _type:        'merge-single',
          monarchId:    item.monarchIds[0],
          name:         pc.name,
          balance:      total,
          plCategory:   pc.plCategory,
          plType:       pc.plType,
          writeField,
          monarchNames:  sources.map(a => a.name),
        });
      }
    }

    // Strip internal _ref/_type/_fieldMappings before sending to service worker
    const accountsPayload = accountSpecs.map(({ _ref, _type, _fieldMappings, ...rest }) => rest);

    try {
      const res = await msg('PL_CREATE_ACCOUNTS', { key: S.plApiKey, accounts: accountsPayload });
      if (res?.error) throw new Error(res.error);

      // Route returned IDs back to their sync items and update state
      const followUpUpdates = []; // merge-multi additional field updates (safe, atomic)

      for (const created of (res.created || [])) {
        const spec = accountSpecs.find(s => s.monarchId === created.monarchId);
        if (!spec) continue;
        const item = spec._ref;

        if (spec._type === 'simple') {
          let plId = created.plId;
          const pending = S.pendingCreations[created.monarchId];
          if (created.plCategory === 'asset' && pending?.writeField && pending.writeField !== 'amount') {
            plId = created.plId + '__' + pending.writeField;
          }
          const { realId, field } = resolvePlMapping(plId);
          S.mappings[created.monarchId] = { plId, plRealId: realId, plField: field };
          delete S.pendingCreations[created.monarchId];
          item.status = 'done'; ok++;

        } else if (spec._type === 'merge-single') {
          const newId = created.plId;
          const wf = spec.writeField;
          for (const mid of item.monarchIds) {
            S.mappings[mid] = { plId: newId, plRealId: newId, plField: wf };
          }
          const gIdx = S.mergeGroups.findIndex(g => g.id === item.mergeGroupId);
          if (gIdx > -1) {
            S.mergeGroups[gIdx] = { ...S.mergeGroups[gIdx], plId: newId, plRealId: newId,
                                     plField: wf, pendingCreate: null };
          }
          item.status = 'done'; ok++;

        } else if (spec._type === 'merge-multi') {
          const newId = created.plId;
          const fieldMappings = spec._fieldMappings;

          // Map each Monarch account to its synthetic field ID
          for (const fm of fieldMappings) {
            const plId = fm.writeField === 'amount' ? newId : `${newId}__${fm.writeField}`;
            const { realId, field } = resolvePlMapping(plId);
            S.mappings[fm.monarchId] = { plId, plRealId: realId, plField: field };
          }

          // Queue follow-up updateAccount calls for additional fields (safe — atomic, no export/restore)
          for (const fm of fieldMappings.slice(1)) {
            followUpUpdates.push({ plId: newId, writeField: fm.writeField, balance: fm.balance });
          }

          const gIdx = S.mergeGroups.findIndex(g => g.id === item.mergeGroupId);
          if (gIdx > -1) {
            S.mergeGroups[gIdx] = {
              ...S.mergeGroups[gIdx],
              plId: newId, plRealId: newId,
              pendingCreate: null,
              monarchIds: [],
              fieldMappings: S.mergeGroups[gIdx].fieldMappings,
            };
          }
          item.status = 'done'; ok++;
        }
      }

      // Execute follow-up field updates for merge-multi accounts
      for (const upd of followUpUpdates) {
        await msg('PL_UPDATE_ACCOUNT', {
          plAccountId: upd.plId,
          data: { [upd.writeField]: upd.balance },
          key: S.plApiKey,
        });
      }

    } catch (e) {
      // On failure, mark all creation items failed and flag pendingCreations
      for (const item of creations) {
        item.status = 'failed'; item.error = e.message;
        if (S.pendingCreations[item.monarchId]) {
          S.pendingCreations[item.monarchId].failedAt = Date.now();
          S.pendingCreations[item.monarchId].errorMsg = e.message;
        }
      }
      for (const item of createMergeItems) {
        item.status = 'failed'; item.error = e.message;
      }
      fail += allCreationItems.length;
    }

    await save();
    renderSync();
  }

  // Step 2: updates
  // (ok/fail were already declared and seeded above)

  for (let i = 0; i < syncPlan.length; i++) {
    if (['create','create-merge-single','create-merge-multi','merge-multi','budget'].includes(syncPlan[i].type)) continue;
    syncPlan[i].status = 'pushing';
    renderSync();
    try {
      const res = await msg('PL_UPDATE_ACCOUNT', {
        plAccountId: syncPlan[i].plRealId,
        data: { [syncPlan[i].plField]: syncPlan[i].writeAmount },
        key: S.plApiKey,
      });
      if (res?.error) throw new Error(res.error);
      syncPlan[i].status = 'done'; ok++;
    } catch (e) {
      syncPlan[i].status = 'failed'; syncPlan[i].error = e.message; fail++;
    }
    renderSync();
  }

  // Step 2b: merge-multi — update each field of an existing asset
  for (const item of syncPlan.filter(i => i.type === 'merge-multi')) {
    item.status = 'pushing'; renderSync();
    try {
      for (const fm of item.fieldMappings) {
        // Get fresh balance from current Monarch accounts
        const ma = S.monarchAccounts.find(a => a.id === fm.monarchId);
        const bal = ma?.balance ?? fm.balance ?? 0;
        // Collected below — written in PL_SYNC_ALL
        item._fieldUpdates = item._fieldUpdates || [];
        item._fieldUpdates.push({ plAccountId: item.plRealId, data: { [fm.writeField]: bal } });
      }
      item.status = 'done'; ok++;
    } catch (e) {
      item.status = 'failed'; item.error = e.message; fail++;
    }
    renderSync();
  }

  // Step 2b: budget — warn about one-time expenses
  const budgetItems = syncPlan.filter(i => i.type === 'budget');
  if (budgetItems.length) {
    const onceItems = budgetItems.filter(i => i.warnText);
    if (onceItems.length) {
      const names = onceItems.map(i => i.plLabel).join('\n• ');
      const confirmed = await showConfirm(
        `⚠ One-time expense${onceItems.length > 1 ? 's' : ''} mapped`,
        `${onceItems.length} mapping${onceItems.length > 1 ? 's target' : ' targets'} a one-time expense:\n\n• ${names}\n\nThe Monarch monthly amount will be written as the total cost of the event, which is likely not what you want.`,
        'Continue anyway'
      );
      if (!confirmed) {
        onceItems.forEach(i => { i.status = 'pending'; });
        const safeBudgetItems = budgetItems.filter(i => !i.warnText);
        if (!safeBudgetItems.length) {
          byId('runSyncBtn').disabled = false;
          byId('previewSyncBtn').disabled = false;
          _syncRunning = false;
          renderSync();
          toast('One-time expense sync cancelled — map these manually in PL', 'err');
          return;
        }
        budgetItems.length = 0;
        safeBudgetItems.forEach(i => budgetItems.push(i));
      }
    }
  }

  // Steps 2+2b+3 combined: one PL_SYNC_ALL call — one export, at most two restores.
  // Collect: account field updates, merge-multi field updates, notes, and expense updates.
  const allAccountUpdates = [];

  // Single-field account updates (type: 'single', 'merge')
  for (const item of syncPlan) {
    if (!['single', 'merge'].includes(item.type)) continue;
    if (item.status !== 'done') continue; // already failed in step 1
    allAccountUpdates.push({ plAccountId: item.plRealId, data: { [item.plField]: item.writeAmount } });
  }

  // Merge-multi field updates collected above
  for (const item of syncPlan.filter(i => i.type === 'merge-multi')) {
    if (item.status !== 'done') continue;
    for (const upd of (item._fieldUpdates || [])) allAccountUpdates.push(upd);
  }

  // Notes map
  const notesMap = {};
  if (S.settings.writeNotes !== false) {
    for (const item of syncPlan) {
      if (item.status === 'done' && item.monarchNames?.length && item.plRealId) {
        notesMap[item.plRealId] = item.monarchNames.join(' + ');
      }
    }
  }

  // Expense updates
  const allExpenseUpdates = budgetItems.map(i => ({
    planId:           i.planId,
    expenseId:        i.expenseId,
    amount:           i.writeAmount,
    writeField:       i.writeField || 'amount',
    createIfMissing:  i.createIfMissing || false,
    newExpenseId:     i.newExpenseId,
    newExpenseName:   i.newExpenseName,
    newFrequency:     i.newFrequency,
    newExpType:       i.newExpType || 'living-expenses',
    newExpTypeParams: i.newExpTypeParams || {},
    monarchNames:     S.settings.writeNotes !== false ? (i.monarchNames || []) : [],
  }));

  const needsSyncAll = allAccountUpdates.length || Object.keys(notesMap).length || allExpenseUpdates.length;

  if (needsSyncAll) {
    // Mark all singles/merges/merge-multis as pushing for UI
    for (const item of syncPlan) {
      if (['single','merge','merge-multi'].includes(item.type) && item.status === 'done') {
        item.status = 'pushing';
      }
    }
    budgetItems.forEach(i => { i.status = 'pushing'; });
    renderSync();

    try {
      const res = await msg('PL_SYNC_ALL', {
        key: S.plApiKey,
        accountUpdates: allAccountUpdates,
        notesMap: Object.keys(notesMap).length ? notesMap : null,
        expenseUpdates: allExpenseUpdates.length ? allExpenseUpdates : null,
      });
      if (res?.error) throw new Error(res.error);

      // Mark all as done
      for (const item of syncPlan) {
        if (['single','merge','merge-multi'].includes(item.type) && item.status === 'pushing') {
          item.status = 'done'; ok++;
        }
      }

      // Budget promotion
      const syncedAt = Date.now();
      for (const item of budgetItems) {
        if (item.status !== 'pushing') continue;
        item.status = 'done'; ok++;

        if (item.expenseId === '__CREATE_NEW__' && item.newExpenseId) {
          const realExpenseId = item.newExpenseId;
          const newKey = item.planId + '::' + realExpenseId;
          if (!S.plExpenses.some(e => e.key === newKey)) {
            S.plExpenses.push({
              planId: item.planId,
              planName: S.plExpenses.find(e => e.planId === item.planId)?.planName || item.planId,
              expenseId: realExpenseId, expenseName: item.newExpenseName || 'New Expense',
              amount: item.writeAmount, frequency: item.newFrequency || 'monthly', key: newKey,
            });
          }
          for (const targets of Object.values(S.budgetMappings)) {
            for (const t of targets) {
              if (t.expenseId === '__CREATE_NEW__' && t.newExpenseId === realExpenseId && t.planId === item.planId) {
                t.expenseId = realExpenseId; t.createIfMissing = false;
                delete t.newExpenseId; delete t.newExpenseName; delete t.newFrequency;
              }
            }
          }
          for (const g of S.budgetMergeGroups) {
            for (const t of g.mappings || []) {
              if (t.expenseId === '__CREATE_NEW__' && t.newExpenseId === realExpenseId && t.planId === item.planId) {
                t.expenseId = realExpenseId; t.createIfMissing = false;
                delete t.newExpenseId; delete t.newExpenseName; delete t.newFrequency;
              }
            }
          }
        } else if (item.createIfMissing) {
          for (const targets of Object.values(S.budgetMappings)) {
            for (const t of targets) {
              if (t.planId === item.planId && t.expenseId === item.expenseId) t.createIfMissing = false;
            }
          }
          for (const g of S.budgetMergeGroups) {
            for (const t of g.mappings || []) {
              if (t.planId === item.planId && t.expenseId === item.expenseId) t.createIfMissing = false;
            }
          }
        }

        // Stamp lastSynced
        const realExpenseId = item.expenseId === '__CREATE_NEW__' ? item.newExpenseId : item.expenseId;
        for (const targets of Object.values(S.budgetMappings)) {
          for (const t of targets) {
            if (t.planId === item.planId && (t.expenseId === realExpenseId || t.expenseId === item.expenseId)) t.lastSynced = syncedAt;
          }
        }
        for (const g of S.budgetMergeGroups) {
          for (const t of g.mappings || []) {
            if (t.planId === item.planId && (t.expenseId === realExpenseId || t.expenseId === item.expenseId)) t.lastSynced = syncedAt;
          }
        }
      }

      // Use the exported data from PL_SYNC_ALL to update local PL state (no extra round trip)
      if (res.exported) {
        S.plAccounts = extractPLAccounts(res.exported);
        S.plExpenses = extractPLExpenses(res.exported);
      }

    } catch (e) {
      for (const item of syncPlan) {
        if (['single','merge','merge-multi'].includes(item.type) && item.status === 'pushing') {
          item.status = 'failed'; item.error = e.message; fail++;
        }
      }
      budgetItems.filter(i => i.status === 'pushing').forEach(i => { i.status = 'failed'; i.error = e.message; fail++; });
    }

    await save();
    renderSync();
  }

  const hadCreations = creations.length > 0 || createMergeItems.length > 0;
  const anyCreationDone = [...creations, ...createMergeItems].some(i => i.status === 'done');

  if (hadCreations && anyCreationDone) {
    // Lock Push button — must Refresh Preview before pushing again to avoid re-creating accounts.
    syncPlan = [];
    renderSync();
    byId('runSyncBtn').disabled = true;
    byId('previewSyncBtn').disabled = false;
    _syncRunning = false;
    const failMsg = fail ? ` — ${fail} creation${fail>1?'s':''} failed, check Mapping tab to retry` : '';
    toast(`Sync: ${ok} done${failMsg} — Refresh Preview to sync again`, fail ? 'err' : 'ok');
  } else {
    byId('runSyncBtn').disabled = false;
    byId('previewSyncBtn').disabled = false;
    _syncRunning = false;
    const budgetDone = budgetItems.filter(i => i.status === 'done').length;
    const budgetNote = budgetDone ? ` (${budgetDone} budget expense${budgetDone > 1 ? 's' : ''} updated)` : '';
    toast(`Sync: ${ok} done${budgetNote}${fail ? ', ' + fail + ' failed' : ''}`, fail ? 'err' : 'ok');
  }
}

// ── SETTINGS tab ─────────────────────────────────────────────
function setupSettings() {
  // Connection mode toggle
  const cbHosted = byId('settingHostedMode');
  cbHosted.checked = S.connectionMode === 'hosted';
  cbHosted.addEventListener('change', async () => {
    S.connectionMode = cbHosted.checked ? 'hosted' : 'self-hosted';
    await save();
    applyConnectionModeUI();
    // Show/hide cloud backup section based on mode
    const cs = byId('cloudBackupSection');
    const vl = byId('cloudVersionList');
    if (cbHosted.checked) {
      cs.classList.remove('hidden');
      refreshCloudVersions();
    } else {
      cs.classList.add('hidden');
      vl.classList.add('hidden');
      byId('cvNameBar').classList.add('hidden');
    }
    updateCloudSaveState();
    toast(cbHosted.checked ? 'Switched to hosted service' : 'Switched to self-hosted proxy', 'ok');
  });

  const cbHidden = byId('settingShowHidden');
  cbHidden.checked = S.settings.showHidden;
  cbHidden.addEventListener('change', async () => {
    S.settings.showHidden = cbHidden.checked;
    await save();
    renderMapping();
    renderMerge();
  });

  const cbNotes = byId('settingWriteNotes');
  cbNotes.checked = S.settings.writeNotes !== false; // default true
  cbNotes.addEventListener('change', async () => {
    S.settings.writeNotes = cbNotes.checked;
    await save();
    toast(cbNotes.checked ? 'Notes will be written on sync' : 'Notes disabled', 'ok');
  });

  // MFA settings
  const cbMfa        = byId('settingMfaEnabled');
  const mfaSecretEl  = byId('settingMfaSecret');
  const mfaSecretRow = byId('mfaSecretRow');
  const mfaRefreshBtn = byId('btnRefreshMonarchSession');

  cbMfa.checked     = S.settings.monarchMfaEnabled || false;
  mfaSecretEl.value = S.settings.monarchMfaSecret  || '';
  mfaSecretRow.classList.toggle('active', cbMfa.checked);
  mfaRefreshBtn.classList.toggle('active', cbMfa.checked);

  cbMfa.addEventListener('change', async () => {
    S.settings.monarchMfaEnabled = cbMfa.checked;
    await save();
    mfaSecretRow.classList.toggle('active', cbMfa.checked);
    mfaRefreshBtn.classList.toggle('active', cbMfa.checked);
    if (cbMfa.checked) mfaSecretEl.focus();
    toast(cbMfa.checked ? 'MFA enabled' : 'MFA disabled', 'ok');
  });

  byId('toggleMfaSecret').addEventListener('click', () => {
    mfaSecretEl.type = mfaSecretEl.type === 'password' ? 'text' : 'password';
  });

  mfaSecretEl.addEventListener('change', async () => {
    S.settings.monarchMfaSecret = mfaSecretEl.value.trim().toUpperCase();
    mfaSecretEl.value = S.settings.monarchMfaSecret;
    await save();
  });

  byId('btnRefreshMonarchSession').addEventListener('click', async () => {
    const statusEl = byId('statusMfaRefresh');
    statusEl.style.color = 'var(--txt3)';
    statusEl.textContent = 'Refreshing…';
    try {
      await msg('MONARCH_REFRESH_SESSION');
      statusEl.style.color = 'var(--green)';
      statusEl.textContent = 'Session refreshed';
    } catch (e) {
      statusEl.style.color = 'var(--red)';
      statusEl.textContent = e.message || 'Refresh failed';
    }
  });

  // Budget value source settings
  const budgetModeEl    = byId('settingBudgetMode');
  const budgetLookbackEl = byId('settingBudgetLookback');
  const budgetAggEl     = byId('settingBudgetAggregation');
  const budgetActualOpts = byId('budgetActualOptions');

  const bs = S.settings.budgetSource || {};
  budgetModeEl.value     = bs.mode        || 'planned';
  const defaultSettingsLookback = (bs.mode || 'planned') === 'actual' ? 3 : 1;
  budgetLookbackEl.value = bs.lookback    || defaultSettingsLookback;
  budgetAggEl.value      = bs.aggregation || 'average';

  function syncBudgetSourceUI() {
    // Both planned and actual modes support lookback + aggregation, so always show
    budgetActualOpts.classList.remove('hidden');
  }
  syncBudgetSourceUI();

  budgetModeEl.addEventListener('change', async () => {
    const newMode = budgetModeEl.value;
    // Reset lookback to mode-appropriate default when switching modes
    const modeDefaultLookback = newMode === 'actual' ? 3 : 1;
    if (!S.settings.budgetSource?.lookback) budgetLookbackEl.value = modeDefaultLookback;
    S.settings.budgetSource = { ...(S.settings.budgetSource || {}), mode: newMode };
    syncBudgetSourceUI();
    await save();
    renderBudgetMapping(); renderBudgetMerges(); renderMerge();
    toast(newMode === 'planned' ? 'Using planned budget amounts' : 'Using actual spending amounts', 'ok');
  });

  budgetLookbackEl.addEventListener('change', async () => {
    const v = Math.max(1, Math.min(12, parseInt(budgetLookbackEl.value, 10) || 3));
    budgetLookbackEl.value = v;
    S.settings.budgetSource = { ...(S.settings.budgetSource || {}), lookback: v };
    await save();
    renderBudgetMapping(); renderBudgetMerges(); renderMerge();
  });

  budgetAggEl.addEventListener('change', async () => {
    S.settings.budgetSource = { ...(S.settings.budgetSource || {}), aggregation: budgetAggEl.value };
    await save();
    renderBudgetMapping(); renderBudgetMerges(); renderMerge();
  });

  // ── Backup & Restore ──────────────────────────────────────
  function getBackupPayload() {
    return {
      _version: 1,
      _exportedAt: new Date().toISOString(),
      mappings: S.mappings,
      pendingCreations: S.pendingCreations,
      mergeGroups: S.mergeGroups,
      budgetMappings: S.budgetMappings,
      budgetMergeGroups: S.budgetMergeGroups,
      budgetSourceOverrides: S.budgetSourceOverrides,
      budgetGroupSourceOverrides: S.budgetGroupSourceOverrides,
      settings: S.settings,
    };
  }

  function applyBackupPayload(data) {
    if (data.mappings !== undefined)                  S.mappings = data.mappings;
    if (data.pendingCreations !== undefined)           S.pendingCreations = data.pendingCreations;
    if (data.mergeGroups !== undefined)                S.mergeGroups = data.mergeGroups;
    if (data.budgetMappings !== undefined)             S.budgetMappings = data.budgetMappings;
    if (data.budgetMergeGroups !== undefined)          S.budgetMergeGroups = data.budgetMergeGroups;
    if (data.budgetSourceOverrides !== undefined)      S.budgetSourceOverrides = data.budgetSourceOverrides;
    if (data.budgetGroupSourceOverrides !== undefined) S.budgetGroupSourceOverrides = data.budgetGroupSourceOverrides;
    if (data.settings !== undefined)                   S.settings = data.settings;
  }

  // Export to file
  byId('btnExportBackup').addEventListener('click', () => {
    const payload = getBackupPayload();
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `monarch-bridge-backup-${new Date().toISOString().slice(0, 10)}.json`;
    a.click();
    URL.revokeObjectURL(url);
    toast('Backup exported', 'ok');
  });

  // Import from file
  const importInput = byId('importBackupFile');
  byId('btnImportBackup').addEventListener('click', () => importInput.click());
  importInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const text = await file.text();
      const data = JSON.parse(text);
      if (!data._version) { toast('Invalid backup file', 'err'); return; }
      applyBackupPayload(data);
      await save();
      renderAll();
      toast('Backup restored', 'ok');
    } catch (err) {
      toast('Failed to read backup file', 'err');
    }
    importInput.value = '';
  });

  // Cloud backup (hosted users only)
  const cloudSection = byId('cloudBackupSection');
  const cloudSaveBtn = byId('btnCloudSave');

  function updateCloudSaveState() {
    const enabled = S.connectionMode === 'hosted' && !!S.firebaseEmail;
    cloudSaveBtn.disabled = !enabled;
    cloudSaveBtn.title = enabled ? '' : 'Sign in to your Monarch Bridge account first';
  }

  if (S.connectionMode === 'hosted') cloudSection.classList.remove('hidden');
  updateCloudSaveState();

  const versionListEl = byId('cloudVersionList');
  const nameBar = byId('cvNameBar');
  const nameInput = byId('cvNameInput');
  const nameConfirm = byId('cvNameConfirm');
  const nameCancel = byId('cvNameCancel');
  let cloudVersions = [];
  let _nameAction = null; // { type: 'save' } or { type: 'rename', versionId }

  function showNameBar(defaultName, action) {
    _nameAction = action;
    nameInput.value = defaultName;
    nameConfirm.textContent = action.type === 'save' ? 'Save' : 'Rename';
    nameBar.classList.remove('hidden');
    nameInput.focus();
    nameInput.select();
  }

  function hideNameBar() {
    nameBar.classList.add('hidden');
    _nameAction = null;
  }

  nameCancel.addEventListener('click', hideNameBar);
  nameInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); nameConfirm.click(); }
    if (e.key === 'Escape') hideNameBar();
  });

  nameConfirm.addEventListener('click', async () => {
    const name = nameInput.value.trim() || 'Untitled';
    const action = _nameAction;
    hideNameBar();
    if (!action) return;

    const statusEl = byId('cloudBackupStatus');

    if (action.type === 'save') {
      statusEl.style.color = 'var(--txt3)';
      statusEl.textContent = 'Saving…';
      try {
        await msg('FIREBASE_SAVE_BACKUP', { backup: getBackupPayload(), name });
        statusEl.style.color = 'var(--green)';
        statusEl.textContent = 'Saved ' + new Date().toLocaleTimeString();
        await refreshCloudVersions();
      } catch (err) {
        statusEl.style.color = 'var(--red)';
        statusEl.textContent = err.message || 'Save failed';
      }
    } else if (action.type === 'rename') {
      try {
        await msg('FIREBASE_UPDATE_BACKUP', { versionId: action.versionId, name });
        const v = cloudVersions.find(x => x.id === action.versionId);
        if (v) v.name = name;
        renderCloudVersions();
      } catch (err) { toast(err.message || 'Rename failed', 'err'); }
    }
  });

  function renderCloudVersions() {
    if (!cloudVersions.length) {
      versionListEl.classList.add('hidden');
      return;
    }
    versionListEl.classList.remove('hidden');
    versionListEl.innerHTML = cloudVersions.map(v => {
      const d = v.savedAt ? new Date(v.savedAt) : null;
      const dateStr = d ? d.toLocaleDateString() + ' ' + d.toLocaleTimeString() : v.id;
      const name = v.name || 'Untitled';
      const pinned = v.pinned || false;
      const pinIcon = pinned ? '📌' : '📌';
      return `<div class="cv-row ${pinned ? 'pinned' : ''}" data-vid="${esc(v.id)}">
        <span class="cv-name" title="${esc(name)}">${esc(name)}</span>
        <span class="cv-date">${esc(dateStr)}</span>
        <span class="cv-actions">
          <button class="cv-btn pin-btn ${pinned ? 'pin-active' : ''}" data-action="pin" title="${pinned ? 'Unpin' : 'Pin (prevents auto-delete)'}">${pinIcon}</button>
          <button class="cv-btn" data-action="rename" title="Rename">✏️</button>
          <button class="cv-btn restore" data-action="restore" title="Restore this backup">↩</button>
          <button class="cv-btn delete" data-action="delete" title="Delete">✕</button>
        </span>
      </div>`;
    }).join('');
  }

  async function refreshCloudVersions() {
    try {
      const res = await msg('FIREBASE_LOAD_BACKUP', { listOnly: true });
      cloudVersions = res.versions || [];
      renderCloudVersions();
    } catch { /* silent */ }
  }

  // Delegated click handler for version list actions
  versionListEl.addEventListener('click', async (e) => {
    const btn = e.target.closest('[data-action]');
    if (!btn) return;
    const row = btn.closest('.cv-row');
    const vid = row?.dataset.vid;
    if (!vid) return;
    const action = btn.dataset.action;
    const statusEl = byId('cloudBackupStatus');

    if (action === 'restore') {
      statusEl.style.color = 'var(--txt3)';
      statusEl.textContent = 'Loading…';
      try {
        const res = await msg('FIREBASE_LOAD_BACKUP', { versionId: vid });
        if (!res.backup) { statusEl.style.color = 'var(--amber)'; statusEl.textContent = 'Not found'; return; }
        applyBackupPayload(res.backup);
        await save();
        renderAll();
        statusEl.style.color = 'var(--green)';
        statusEl.textContent = 'Restored';
        toast('Cloud backup restored', 'ok');
      } catch (err) {
        statusEl.style.color = 'var(--red)';
        statusEl.textContent = err.message || 'Restore failed';
      }
    } else if (action === 'pin') {
      const v = cloudVersions.find(x => x.id === vid);
      const newPinned = !(v?.pinned);
      const pinnedCount = cloudVersions.filter(x => x.pinned).length;
      if (newPinned && pinnedCount >= 2) { toast('Maximum 2 pinned backups', 'err'); return; }
      try {
        await msg('FIREBASE_UPDATE_BACKUP', { versionId: vid, pinned: newPinned });
        if (v) v.pinned = newPinned;
        renderCloudVersions();
      } catch (err) { toast(err.message || 'Update failed', 'err'); }
    } else if (action === 'rename') {
      const v = cloudVersions.find(x => x.id === vid);
      showNameBar(v?.name || '', { type: 'rename', versionId: vid });
    } else if (action === 'delete') {
      const v = cloudVersions.find(x => x.id === vid);
      if (v?.pinned) { toast('Unpin before deleting', 'err'); return; }
      try {
        await msg('FIREBASE_DELETE_BACKUP', { versionId: vid });
        cloudVersions = cloudVersions.filter(x => x.id !== vid);
        renderCloudVersions();
        toast('Backup deleted', 'ok');
      } catch (err) { toast(err.message || 'Delete failed', 'err'); }
    }
  });

  if (S.connectionMode === 'hosted') refreshCloudVersions();

  // React to sign-in / sign-out while the popup is open
  document.addEventListener('firebase-auth-changed', () => {
    const hosted = S.connectionMode === 'hosted';
    const cs = byId('cloudBackupSection');
    const vl = byId('cloudVersionList');
    if (hosted && S.firebaseEmail) {
      cs.classList.remove('hidden');
      refreshCloudVersions();
    } else {
      vl.classList.add('hidden');
      byId('cvNameBar').classList.add('hidden');
    }
    updateCloudSaveState();
  });

  byId('btnCloudSave').addEventListener('click', () => {
    showNameBar('Backup ' + new Date().toLocaleDateString(), { type: 'save' });
  });
}

// ── Render all ───────────────────────────────────────────────
function renderAll() {
  renderMapping();
  renderMerge();
  renderBudgetMerges();
  renderBudgetMapping();
}

// ── Utilities ────────────────────────────────────────────────
// Sends a message to the service worker. Retries once if the SW was killed and needs to restart.
async function msg(type, extra = {}) {
  try {
    return await chrome.runtime.sendMessage({ type, ...extra });
  } catch (e) {
    if (e?.message?.includes('Could not establish connection')) {
      await new Promise(r => setTimeout(r, 150));
      return chrome.runtime.sendMessage({ type, ...extra });
    }
    throw e;
  }
}
function byId(id) { return document.getElementById(id); }
function syncedAgo(ts) {
  if (!ts) return '';
  const secs = Math.floor((Date.now() - ts) / 1000);
  if (secs <  60)  return 'just now';
  if (secs < 3600) return `${Math.floor(secs / 60)}m ago`;
  if (secs < 86400) return `${Math.floor(secs / 3600)}h ago`;
  return `${Math.floor(secs / 86400)}d ago`;
}
function esc(s) {
  return String(s || '').replace(/[&<>"']/g, c =>
    ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[c])
  );
}
function fmt(v) {
  return new Intl.NumberFormat('en-US', { style:'currency', currency:'USD',
    minimumFractionDigits:0, maximumFractionDigits:2 }).format(v || 0);
}
function status(id, text, cls = '') {
  const el = byId(id); el.textContent = text; el.className = 'status-line ' + cls;
}
// Promise-based confirmation dialog that renders inside the extension window.
// Returns true if the user clicks OK, false if they cancel.
function showConfirm(title, body, okLabel = 'Continue') {
  return new Promise(resolve => {
    const modal  = byId('confirmModal');
    byId('confirmModalTitle').textContent = title;
    byId('confirmModalBody').textContent  = body;
    byId('confirmModalOk').textContent    = okLabel;
    modal.classList.remove('hidden');

    function finish(result) {
      modal.classList.add('hidden');
      okBtn.removeEventListener('click', onOk);
      cancelBtn.removeEventListener('click', onCancel);
      resolve(result);
    }
    function onOk()     { finish(true);  }
    function onCancel() { finish(false); }

    const okBtn     = byId('confirmModalOk');
    const cancelBtn = byId('confirmModalCancel');
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', onCancel);
  });
}

function toast(text, cls = '') {
  const el = byId('toast');
  el.textContent = text; el.className = 'toast show ' + cls;
  clearTimeout(el._t);
  el._t = setTimeout(() => { el.className = 'toast'; }, 2800);
}
function setLoading(id, on) {
  const el = byId(id);
  if (on) { el._orig = el.innerHTML; el.innerHTML = '<span class="spin"></span>'; el.disabled = true; }
  else    { el.innerHTML = el._orig || el.innerHTML; el.disabled = false; }
}
