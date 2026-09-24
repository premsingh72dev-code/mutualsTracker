/**
 * ==========================================================================
 * MUTUAL FUND ANALYTICS & DYNAMIC BENCHMARK ENGINE
 * Clean, High-Readability Client-Side Controller (FastAPI Integrated)
 * ==========================================================================
 */

(function () {
  'use strict';

  // ---------------------------------------------------------------------
  // Base path the app is served from. Injected by the server template, so
  // the same code works at a domain root (""), behind a sub-path reverse
  // proxy ("/mfa"), or on any cloud host - no rebuild, no hardcoded origin.
  // ---------------------------------------------------------------------
  const APP_BASE = (window.APP_BASE || '').replace(/\/+$/, '');
  const apiUrl = (path) => `${APP_BASE}${path}`;

  // Application State
  const state = {
    funds: [],
    rawRollingFunds: [],
    summary: {},
    categories: [],
    categoryStats: {},
    distribution: [],
    columnMapping: {},
    
    // Filters & Benchmarks
    searchQuery: '',
    categoryFilter: 'all',
    perfFilter: 'all',
    highlightMode: true,
    benchmarkValue: 0.62,
    autoAvgSharpe: 0.62,
    
    // Table Sorting & Pagination
    sortColumn: 'rank',
    sortDirection: 'asc',
    currentPage: 1,
    pageSize: 50,
    
    // Tab Navigation
    activeTab: 'tab-best-funds',

    // Multi-Ratio Screener & Industry Presets
    activePreset: 'all',
    ratioFilters: {
      sharpeMin: null,
      sharpeMax: null,
      stdMax: null,
      treynorMin: null,
      infoMin: null,
      rollingMin: null,
      aumMin: null
    },

    // Curated Fund Selection Basket
    selectedFundsMap: new Map(),

    // Best Mutual Funds Screener
    bestFundsBenchmarkBasis: 'category', // 'category' or 'portfolio'
    bestFundsCategoryFilter: 'all',
    bestFundsShortlist: [],
    bestFundsChampion: null,

    // Server-side session persistence (UUID assigned by the database)
    sessionUid: null,
    lastMatchReport: null,

    // Signed-in account
    currentUser: null
  };

  // Chart Registry
  const chartRegistry = {};

  // Helpers
  const $ = id => document.getElementById(id);
  const $$ = sel => document.querySelectorAll(sel);

  // Escapes untrusted text (fund names/categories from uploaded Excel files, user-typed
  // form fields, record names, etc.) before it's interpolated into innerHTML templates.
  const ESC_MAP = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
  function escapeHtml(str) {
    if (str === null || str === undefined) return '';
    return String(str).replace(/[&<>"']/g, ch => ESC_MAP[ch]);
  }

  // ==========================================
  // CATEGORY CODE -> HUMAN-READABLE LABEL
  // Short codes like "EQ-THE-BUS-CYC" or "EQ-DIV-Y" aren't self-explanatory. This maps
  // every category code actually seen in the sample Risk-Ratios/Rolling-Returns sheets to
  // a plain-English name. The raw code (e.g. "EQ-MID") is what filtering/grouping still
  // uses internally — this only changes what's shown on screen.
  // ==========================================
  const CATEGORY_LABELS = {
    'EQ-BANK': 'Equity: Banking & Financial Services',
    'EQ-CONS': 'Equity: Consumption',
    'EQ-CONTRA': 'Equity: Contra',
    'EQ-DIV-Y': 'Equity: Dividend Yield',
    'EQ-ELSS': 'Equity: ELSS (Tax Saver)',
    'EQ-ENERGY': 'Equity: Energy & Power',
    'EQ-FLX': 'Equity: Flexi Cap',
    'EQ-FMCG': 'Equity: FMCG',
    'EQ-FOCUS': 'Equity: Focused',
    'EQ-INFRA': 'Equity: Infrastructure',
    'EQ-INTL': 'Equity: International',
    'EQ-IT': 'Equity: Technology (IT)',
    'EQ-L-MC': 'Equity: Large & Mid Cap',
    'EQ-LC': 'Equity: Large Cap',
    'EQ-MID': 'Equity: Mid Cap',
    'EQ-MLC': 'Equity: Multi Cap',
    'EQ-MNC': 'Equity: MNC',
    'EQ-PHARMA': 'Equity: Pharma & Healthcare',
    'EQ-PSU': 'Equity: PSU',
    'EQ-SC': 'Equity: Small Cap',
    'EQ-THE-ACT-MOM': 'Equity: Thematic – Momentum',
    'EQ-THE-BUS-CYC': 'Equity: Thematic – Business Cycle',
    'EQ-THE-ESG': 'Equity: Thematic – ESG',
    'EQ-THE-INNOV': 'Equity: Thematic – Innovation',
    'EQ-THE-MANU': 'Equity: Thematic – Manufacturing',
    'EQ-THE-MS': 'Equity: Thematic – Manufacturing & Services',
    'EQ-THE-O': 'Equity: Thematic – Other',
    'EQ-THE-Q': 'Equity: Thematic – Quant',
    'EQ-THE-SPE-OPPOR': 'Equity: Thematic – Special Opportunities',
    'EQ-THE-TRANSPORT': 'Equity: Thematic – Transport & Logistics',
    'EQ-VALUE': 'Equity: Value',
    // Common Hybrid & Solution-Oriented codes (SEBI classification). If an uploaded sheet
    // uses different exact codes, prettifyCategoryCode() below still produces a readable
    // label automatically instead of showing the raw code as-is.
    'HY-CONSERV': 'Hybrid: Conservative',
    'HY-BAL': 'Hybrid: Balanced',
    'HY-AGGR': 'Hybrid: Aggressive',
    'HY-DYN': 'Hybrid: Dynamic Asset Allocation / Balanced Advantage',
    'HY-MULTI': 'Hybrid: Multi Asset Allocation',
    'HY-ARBIT': 'Hybrid: Arbitrage',
    'HY-EQSAV': 'Hybrid: Equity Savings',
    'SO-RETIRE': 'Solution Oriented: Retirement Fund',
    'SO-CHILD': "Solution Oriented: Children's Fund"
  };

  const CATEGORY_PREFIX_LABELS = {
    'EQ': 'Equity', 'HY': 'Hybrid', 'SO': 'Solution Oriented', 'DT': 'Debt', 'DEBT': 'Debt', 'OTH': 'Other'
  };
  const CATEGORY_SEGMENT_LABELS = {
    'THE': 'Thematic', 'DIV': 'Dividend', 'BAL': 'Balanced', 'AGGR': 'Aggressive',
    'CONSERV': 'Conservative', 'ARBIT': 'Arbitrage', 'MULTI': 'Multi Asset', 'DYN': 'Dynamic Allocation',
    'RETIRE': 'Retirement', 'CHILD': "Children's", 'SAV': 'Savings', 'O': 'Other'
  };
  const CATEGORY_SEGMENT_KEEP_AS_IS = new Set(['ELSS', 'FMCG', 'PSU', 'MNC', 'IT', 'ESG']);

  // Best-effort readable label for a category code that isn't in CATEGORY_LABELS —
  // expands the recognised prefix/segment abbreviations and title-cases the rest, so an
  // unfamiliar code degrades to something more readable rather than a raw string.
  function prettifyCategoryCode(code) {
    const parts = String(code).split('-').filter(Boolean);
    if (parts.length === 0) return code;
    const [prefix, ...rest] = parts;
    const prefixLabel = CATEGORY_PREFIX_LABELS[prefix.toUpperCase()] || prefix;
    const restLabels = rest.map(seg => {
      const upper = seg.toUpperCase();
      if (CATEGORY_SEGMENT_KEEP_AS_IS.has(upper)) return upper;
      if (CATEGORY_SEGMENT_LABELS[upper]) return CATEGORY_SEGMENT_LABELS[upper];
      return seg.charAt(0).toUpperCase() + seg.slice(1).toLowerCase();
    });
    return restLabels.length > 0 ? `${prefixLabel}: ${restLabels.join(' ')}` : prefixLabel;
  }

  function getCategoryLabel(code) {
    if (!code) return 'General';
    return CATEGORY_LABELS[code] || prettifyCategoryCode(code);
  }

  // ==========================================
  // AUTHENTICATION
  // Sign-in only — there is no public signup. Accounts are provisioned by an admin with
  // `python main.py create-user`. The token lives in an httpOnly cookie set by the
  // server, so it is never exposed to JavaScript or kept in browser storage.
  // ==========================================
  function showAuthError(msg) {
    const el = $('auth-error');
    if (!el) return;
    el.textContent = msg || '';
    el.style.display = msg ? 'block' : 'none';
  }

  // Shown when the database has no accounts yet, so the sign-in page isn't a dead end.
  function showSetupHintIfNeeded(info) {
    const footnote = $('auth-footnote');
    if (!footnote || !info || info.user_count !== 0) return;
    footnote.insertAdjacentHTML('beforebegin', `
      <div class="auth-setup-note">
        <strong>No accounts exist yet.</strong> Create the first one on the server:
        <code>python main.py create-user you@example.com "YourPassword" "Your Name"</code>
      </div>
    `);
  }

  function showAuthScreen() {
    const authScreen = $('auth-screen');
    const appContainer = $('app-container');
    if (authScreen) authScreen.style.display = 'flex';
    if (appContainer) appContainer.style.display = 'none';
    const email = $('auth-email');
    if (email) email.focus();
  }

  function showApp(user) {
    const authScreen = $('auth-screen');
    const appContainer = $('app-container');
    if (authScreen) authScreen.style.display = 'none';
    if (appContainer) appContainer.style.display = 'block';

    state.currentUser = user;
    const avatar = $('user-avatar');
    const emailEl = $('user-email');
    if (avatar) avatar.textContent = (user.name || user.email || '?').trim().charAt(0);
    if (emailEl) {
      emailEl.textContent = user.email || '';
      emailEl.title = user.name ? `${user.name} · ${user.email}` : (user.email || '');
    }
  }

  async function submitAuth() {
    const email = ($('auth-email').value || '').trim();
    const password = $('auth-password').value || '';

    if (!email || !password) {
      showAuthError('Enter your email and password.');
      return;
    }

    const submit = $('auth-submit');
    if (submit) { submit.disabled = true; submit.textContent = 'Signing in…'; }
    showAuthError('');

    try {
      const resp = await fetch(apiUrl('/api/auth/login'), {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email, password })
      });
      const res = await resp.json();
      if (!resp.ok) throw new Error(res.detail || 'Sign in failed.');

      showApp(res.user);
      $('auth-password').value = '';
      await restoreSessionFromServer();
      updateSavedRecordsBadge();
      checkDatabaseStatus();
    } catch (err) {
      showAuthError(err.message);
    } finally {
      if (submit) {
        submit.disabled = false;
        submit.textContent = 'Sign In';
      }
    }
  }

  function initAuth() {
    const form = $('auth-form');
    if (form) form.addEventListener('submit', e => { e.preventDefault(); submitAuth(); });
    const submit = $('auth-submit');
    if (submit) submit.addEventListener('click', submitAuth);

    // Show/hide password — useful on phones where typing is error-prone.
    const toggle = $('auth-toggle-password');
    const pwd = $('auth-password');
    if (toggle && pwd) {
      toggle.addEventListener('click', () => {
        const showing = pwd.type === 'text';
        pwd.type = showing ? 'password' : 'text';
        toggle.textContent = showing ? '👁' : '🙈';
        toggle.setAttribute('aria-label', showing ? 'Show password' : 'Hide password');
        pwd.focus();
      });
    }

    const btnSignout = $('btn-signout');
    if (btnSignout) {
      btnSignout.addEventListener('click', async () => {
        try { await fetch(apiUrl('/api/auth/logout'), { method: 'POST' }); } catch (e) { /* ignore */ }
        window.location.reload();
      });
    }
  }

  // Decides between the login screen and the dashboard on every page load.
  async function bootstrapAuth() {
    try {
      const resp = await fetch(apiUrl('/api/auth/me'));
      const res = await resp.json();
      if (res.status === 'success' && res.user) {
        showApp(res.user);
        await restoreSessionFromServer();
        updateSavedRecordsBadge();
        checkDatabaseStatus();
        return;
      }
      if (res.status === 'error') {
        showAuthScreen();
        showAuthError(res.detail || 'Database unreachable.');
        return;
      }
      showAuthScreen();
      showSetupHintIfNeeded(res);
      return;
    } catch (err) {
      showAuthScreen();
      showAuthError('Cannot reach the server.');
    }
  }

  document.addEventListener('DOMContentLoaded', () => {
    initAuth();
    initDropZone();
    initBenchmarkControls();
    initRatioFilterControls();
    initBasketControls();
    initTableControls();
    initTabNavigation();
    initActionButtons();
    initRecordManager();
    initBestFundsControls();
    initTableScrollShadows();
    initUploadPanelToggle();
    // Auth decides what to show; it restores the session only once signed in.
    bootstrapAuth();
  });

  // ==========================================
  // LIVE SESSION PERSISTENCE (server-side, UUID-keyed)
  // The active dataset and the UI state around it (basket, filters, benchmark, sort,
  // active tab) are written to the database, not browser storage — so the session is
  // restored on reload, and from any browser, keyed by the uid the server hands back.
  // ==========================================
  let sessionSaveTimer = null;
  let sessionRestoring = false;

  function collectSessionState() {
    return {
      files_summary: (state.currentFilesMeta || []).map(f => f.filename).join(' + '),
      data: state.currentDataPayload || null,
      files_meta: state.currentFilesMeta || [],
      match_report: state.lastMatchReport || null,
      ui: {
        searchQuery: state.searchQuery,
        categoryFilter: state.categoryFilter,
        perfFilter: state.perfFilter,
        highlightMode: state.highlightMode,
        benchmarkValue: state.benchmarkValue,
        activePreset: state.activePreset,
        ratioFilters: state.ratioFilters,
        activeTab: state.activeTab,
        pageSize: state.pageSize,
        currentPage: state.currentPage,
        sortColumn: state.sortColumn,
        sortDirection: state.sortDirection,
        basket: Array.from(state.selectedFundsMap.keys()),
        bestFundsBenchmarkBasis: state.bestFundsBenchmarkBasis,
        bestFundsCategoryFilter: state.bestFundsCategoryFilter
      }
    };
  }

  function setSessionStatus(text, tone) {
    const el = $('session-status');
    if (!el) return;
    el.textContent = text;
    el.className = `session-status ${tone || ''}`.trim();
    el.style.display = text ? 'inline-flex' : 'none';
  }

  // Debounced so a burst of filter clicks results in one write.
  function scheduleSessionSave() {
    if (sessionRestoring) return;
    if (!state.currentDataPayload) return;
    if (sessionSaveTimer) clearTimeout(sessionSaveTimer);
    sessionSaveTimer = setTimeout(saveSessionToServer, 1200);
  }

  async function saveSessionToServer() {
    if (!state.currentDataPayload) return;
    setSessionStatus('Saving…', 'saving');
    try {
      const resp = await fetch(apiUrl('/api/session'), {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ uid: state.sessionUid || null, state: collectSessionState() })
      });
      if (resp.status === 401) {
        // Token expired while the tab was open — send them back to sign in rather than
        // letting edits pile up against a session that can no longer be saved.
        setSessionStatus('Signed out — sign in to keep saving', 'error');
        showAuthScreen();
        showAuthError('Your session expired. Please sign in again.');
        return;
      }
      if (!resp.ok) throw new Error('Session save failed');
      const res = await resp.json();
      state.sessionUid = res.uid;
      setSessionStatus(`Saved to database · ${res.updated_at}`, 'saved');
    } catch (err) {
      setSessionStatus('Not saved — database unreachable', 'error');
    }
  }

  async function restoreSessionFromServer() {
    try {
      const resp = await fetch(apiUrl('/api/session/latest'));
      if (!resp.ok) return;
      const res = await resp.json();
      if (res.status !== 'success' || !res.state || !res.state.data) return;

      sessionRestoring = true;
      state.sessionUid = res.uid;
      loadAnalysisData(res.state.data, res.state.files_meta || []);
      renderMatchReport(res.state.match_report);
      applySessionUiState(res.state.ui || {});
      setSessionStatus(`Restored from database · ${res.updated_at}`, 'saved');
      showToast(`Restored your last session (${res.total_funds} funds) from the database.`, 'info');
    } catch (err) {
      // No session yet, or server unreachable — start clean.
    } finally {
      sessionRestoring = false;
    }
  }

  function applySessionUiState(ui) {
    if (!ui) return;

    if (typeof ui.benchmarkValue === 'number') state.benchmarkValue = ui.benchmarkValue;
    if (typeof ui.searchQuery === 'string') state.searchQuery = ui.searchQuery;
    if (ui.categoryFilter) state.categoryFilter = ui.categoryFilter;
    if (ui.perfFilter) state.perfFilter = ui.perfFilter;
    if (typeof ui.highlightMode === 'boolean') state.highlightMode = ui.highlightMode;
    if (ui.activePreset) state.activePreset = ui.activePreset;
    if (ui.ratioFilters) state.ratioFilters = ui.ratioFilters;
    if (ui.pageSize) state.pageSize = ui.pageSize;
    if (ui.currentPage) state.currentPage = ui.currentPage;
    if (ui.sortColumn) state.sortColumn = ui.sortColumn;
    if (ui.sortDirection) state.sortDirection = ui.sortDirection;
    if (ui.bestFundsBenchmarkBasis) state.bestFundsBenchmarkBasis = ui.bestFundsBenchmarkBasis;
    if (ui.bestFundsCategoryFilter) state.bestFundsCategoryFilter = ui.bestFundsCategoryFilter;

    // Rebuild the basket from the saved fund names
    if (Array.isArray(ui.basket) && ui.basket.length > 0) {
      state.selectedFundsMap.clear();
      const byName = new Map((state.funds || []).map(f => [f.name, f]));
      ui.basket.forEach(name => {
        const fund = byName.get(name);
        if (fund) state.selectedFundsMap.set(name, fund);
      });
    }

    // Reflect restored values in the actual controls
    const searchInput = $('table-search');
    if (searchInput) searchInput.value = state.searchQuery || '';
    const catSel = $('filter-category');
    if (catSel) catSel.value = state.categoryFilter;
    const perfSel = $('filter-performance');
    if (perfSel) perfSel.value = state.perfFilter;
    const pageSel = $('page-size-select');
    if (pageSel) pageSel.value = String(state.pageSize);
    const bestCatSel = $('best-funds-cat-filter');
    if (bestCatSel) bestCatSel.value = state.bestFundsCategoryFilter;
    const highlightBtn = $('toggle-highlight');
    if (highlightBtn) highlightBtn.classList.toggle('active', !!state.highlightMode);

    // The slider's own "input" handler is what normally keeps the cutoff label and the
    // Auto/Top-25%/Alpha buttons in sync — setting .value in code fires no event, so the
    // restore has to bring those along itself.
    if (typeof state.benchmarkValue === 'number') {
      const slider = $('benchmark-slider');
      if (slider) slider.value = state.benchmarkValue;
      const thresholdLabel = $('current-threshold-label');
      if (thresholdLabel) thresholdLabel.textContent = state.benchmarkValue.toFixed(2);

      const isAuto = Math.abs(state.benchmarkValue - (state.autoAvgSharpe || 0)) < 1e-9;
      $$('.preset-btn').forEach(btn => {
        const type = btn.dataset.preset;
        const matches = type === 'auto'
          ? isAuto
          : Math.abs(parseFloat(type) - state.benchmarkValue) < 1e-9;
        btn.classList.toggle('active', matches);
      });
    }

    $$('.ratio-preset-chip').forEach(chip => {
      chip.classList.toggle('active', chip.dataset.preset === state.activePreset);
    });
    $$('#best-benchmark-toggle .btn-toggle-pill').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.basis === state.bestFundsBenchmarkBasis);
    });

    switchToTab(ui.activeTab || 'tab-best-funds');

    recomputeThresholds();
    renderBasketBar();
    renderAllViews();
  }

  // ==========================================
  // UPLOAD PANEL COLLAPSE
  // The panel only folds away once BOTH sheets are in — collapsing it after just the
  // Risk-Ratios upload would hide the Rolling-Returns card before it's been used.
  // ==========================================
  const ROLLING_FIELDS = ['rolling_1y', 'rolling_2y', 'rolling_3y', 'rolling_5y'];

  function setUploadPanelCollapsed(collapsed) {
    const panel = $('upload-panel');
    const btn = $('btn-toggle-upload-panel');
    const label = $('btn-toggle-upload-label');
    if (!panel) return;
    panel.style.display = collapsed ? 'none' : 'grid';
    if (btn) btn.classList.toggle('active', !collapsed);
    if (label) label.textContent = collapsed ? '⚙ Change Files' : '✕ Hide Upload Panel';
  }

  function hasRollingData() {
    if (selectedRollingFile) return true;
    if (state.rawRollingFunds && state.rawRollingFunds.length > 0) return true;
    return (state.funds || []).some(f => ROLLING_FIELDS.some(k => typeof f[k] === 'number'));
  }

  // Collapse only when there's risk data AND rolling data; otherwise keep the panel
  // open so the still-missing sheet stays one click away.
  function updateUploadPanelVisibility() {
    const hasRisk = (state.funds || []).length > 0;
    setUploadPanelCollapsed(hasRisk && hasRollingData());
  }

  function initUploadPanelToggle() {
    const btn = $('btn-toggle-upload-panel');
    if (btn) {
      btn.addEventListener('click', () => {
        const panel = $('upload-panel');
        const isCollapsed = panel && panel.style.display === 'none';
        setUploadPanelCollapsed(!isCollapsed);
      });
    }

    // Shortcut straight to the rolling-returns file picker so extra category sheets
    // can be added without expanding the whole panel first.
    const btnAddRolling = $('btn-add-rolling-inline');
    if (btnAddRolling) {
      btnAddRolling.addEventListener('click', () => {
        const fileRolling = $('file-rolling');
        if (fileRolling) fileRolling.click();
      });
    }

    // Wipes the stored session so the next load starts from the empty state. Saved
    // Analysis snapshots are untouched — this only clears the auto-saved working session.
    const btnClearSession = $('btn-clear-session');
    if (btnClearSession) {
      btnClearSession.addEventListener('click', async () => {
        const ok = confirm(
          'Start fresh?\n\nThis clears the auto-saved working session (current dataset, ' +
          'basket and filters) from the database.\n\nYour saved Analysis snapshots under ' +
          '"Records" are NOT affected.'
        );
        if (!ok) return;
        try {
          if (state.sessionUid) {
            await fetch(apiUrl(`/api/session/${state.sessionUid}`), { method: 'DELETE' });
          }
          state.sessionUid = null;
          window.location.reload();
        } catch (err) {
          showToast(`Could not clear session: ${err.message}`, 'error');
        }
      });
    }

    // Toggle match report banner from toolbar
    const btnToggleMatch = $('btn-toggle-match-report');
    if (btnToggleMatch) {
      btnToggleMatch.addEventListener('click', () => {
        const el = $('match-report-notice');
        if (!el || !state.lastMatchReport) return;
        if (el.style.display === 'none') {
          state.matchReportDismissed = false;
          renderMatchReport(state.lastMatchReport);
        } else {
          state.matchReportDismissed = true;
          renderMatchReport(state.lastMatchReport);
        }
      });
    }
  }

  // ==========================================
  // HORIZONTAL SCROLL AFFORDANCE FOR WIDE TABLES
  // ==========================================
  function initTableScrollShadows() {
    const containers = $$('.table-responsive, .tabs-nav-bar');
    if (!containers.length) return;

    const update = el => {
      const scrollable = el.scrollWidth > el.clientWidth + 2;
      el.classList.toggle('can-scroll-right', scrollable && el.scrollLeft < el.scrollWidth - el.clientWidth - 2);
      el.classList.toggle('can-scroll-left', scrollable && el.scrollLeft > 2);
    };

    containers.forEach(el => {
      update(el);
      el.addEventListener('scroll', () => update(el), { passive: true });
      if (window.ResizeObserver) {
        const ro = new ResizeObserver(() => update(el));
        ro.observe(el);
        const table = el.querySelector('table');
        if (table) ro.observe(table);
      }
    });

    window.addEventListener('resize', () => containers.forEach(update));
  }

  // Dual Upload State
  let selectedRiskFile = null;
  let selectedRollingFile = null;

  // ==========================================
  // DROPZONE & DUAL FILE UPLOAD
  // ==========================================
  function initDropZone() {
    const dropRisk = $('drop-zone-risk');
    const dropRolling = $('drop-zone-rolling');
    const fileRisk = $('file-risk');
    const fileRolling = $('file-rolling');
    const btnBrowseRisk = $('btn-browse-risk');
    const btnBrowseRolling = $('btn-browse-rolling');
    const btnAddMoreRolling = $('btn-add-more-rolling');
    const btnResetRolling = $('btn-reset-rolling');
    const btnEmptyRisk = $('btn-empty-risk');
    const btnEmptyRolling = $('btn-empty-rolling');

    const triggerRisk = () => fileRisk && fileRisk.click();
    const triggerRolling = () => fileRolling && fileRolling.click();

    if (btnBrowseRisk) btnBrowseRisk.addEventListener('click', e => { e.stopPropagation(); triggerRisk(); });
    if (btnBrowseRolling) btnBrowseRolling.addEventListener('click', e => { e.stopPropagation(); triggerRolling(); });
    if (btnAddMoreRolling) btnAddMoreRolling.addEventListener('click', e => { e.stopPropagation(); triggerRolling(); });
    if (btnResetRolling) btnResetRolling.addEventListener('click', e => { e.stopPropagation(); resetRollingReturns(); });
    if (btnEmptyRisk) btnEmptyRisk.addEventListener('click', triggerRisk);
    if (btnEmptyRolling) btnEmptyRolling.addEventListener('click', triggerRolling);

    // Risk Dropzone
    if (dropRisk && fileRisk) {
      setupDropHandlers(dropRisk, files => {
        if (files.length > 0) handleFileSelection(files[0], 'risk');
      });
      fileRisk.addEventListener('change', e => {
        if (e.target.files && e.target.files.length > 0) handleFileSelection(e.target.files[0], 'risk');
      });
    }

    // Rolling Dropzone
    if (dropRolling && fileRolling) {
      setupDropHandlers(dropRolling, files => {
        if (files.length > 0) handleFileSelection(files[0], 'rolling');
      });
      fileRolling.addEventListener('change', e => {
        if (e.target.files && e.target.files.length > 0) handleFileSelection(e.target.files[0], 'rolling');
      });
    }

  }

  function setupDropHandlers(zone, onDropFiles) {
    ['dragenter', 'dragover'].forEach(name => {
      zone.addEventListener(name, e => {
        e.preventDefault();
        zone.classList.add('dragover');
      }, false);
    });

    ['dragleave', 'drop'].forEach(name => {
      zone.addEventListener(name, e => {
        e.preventDefault();
        zone.classList.remove('dragover');
      }, false);
    });

    zone.addEventListener('drop', e => {
      const files = e.dataTransfer.files;
      if (files && files.length > 0) {
        // If user dropped 2 files at once, assign risk and rolling
        if (files.length >= 2) {
          handleFileSelection(files[0], 'risk');
          handleFileSelection(files[1], 'rolling');
        } else {
          onDropFiles(files);
        }
      }
    });
  }

  function handleFileSelection(file, type) {
    if (!file) return;

    // Strict file type check
    const validExts = ['.xlsx', '.xls', '.xlsm', '.csv'];
    const fileName = file.name.toLowerCase();
    const isValid = validExts.some(ext => fileName.endsWith(ext));

    if (!isValid) {
      showToast(`Invalid file '${file.name}'. Please upload an Excel spreadsheet (.xlsx, .xls) or CSV.`, 'error');
      return;
    }

    if (type === 'risk') {
      selectedRiskFile = file;
      const st = $('status-risk');
      if (st) {
        st.textContent = `✓ ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
        st.style.color = '#34d399';
      }
      $('drop-zone-risk').classList.add('loaded-card');
      executeUpload();
    } else if (type === 'rolling') {
      selectedRollingFile = file;
      const st = $('status-rolling');
      if (st) {
        st.textContent = `✓ ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
        st.style.color = '#34d399';
      }
      $('drop-zone-rolling').classList.add('loaded-card');

      // If funds are already loaded in the active session, sequentially merge this category rolling returns file!
      if (state.funds && state.funds.length > 0) {
        executeCategoryRollingUpload(file);
      } else {
        executeUpload();
      }
    }
  }

  async function executeCategoryRollingUpload(file) {
    if (!file) return;
    if (!state.funds || state.funds.length === 0) {
      executeUpload();
      return;
    }

    const formData = new FormData();
    formData.append('rolling_file', file);
    formData.append('current_funds_json', JSON.stringify(state.funds));
    if (state.benchmarkValue) {
      formData.append('benchmark', state.benchmarkValue);
    }

    showToast(`Merging category rolling returns from '${file.name}'...`, 'info');

    try {
      const response = await fetch(apiUrl('/api/upload-rolling-category'), {
        method: 'POST',
        body: formData
      });

      if (!response.ok) {
        const err = await response.json();
        throw new Error(err.detail || 'Failed to merge rolling returns');
      }

      const res = await response.json();

      // Update state without wiping existing funds or previous rolling metrics
      state.funds = res.data.funds;
      state.summary = res.data.summary;
      state.categories = res.data.categories;
      state.categoryStats = res.data.category_stats;
      state.distribution = res.data.distribution;

      // Track uploaded rolling sheet files
      if (!state.uploadedRollingSheets) state.uploadedRollingSheets = [];
      state.uploadedRollingSheets.push({
        filename: res.filename,
        matched_count: res.matched_count,
        categories: res.matched_categories
      });

      // Update verification metadata
      if (!state.currentFilesMeta) state.currentFilesMeta = [];
      if (!state.currentFilesMeta.some(m => m.filename === res.filename)) {
        state.currentFilesMeta.push({
          filename: res.filename,
          active_sheet: 'Rolling Returns',
          row_count: res.row_count || res.matched_count
        });
      }

      // Update raw rolling funds
      if (!state.rawRollingFunds) state.rawRollingFunds = [];
      if (res.rolling_funds && res.rolling_funds.length > 0) {
        const existingNames = new Set(state.rawRollingFunds.map(f => f.name));
        for (const rf of res.rolling_funds) {
          if (!existingNames.has(rf.name)) {
            state.rawRollingFunds.push(rf);
            existingNames.add(rf.name);
          }
        }
      }

      // Update UI components
      const stRolling = $('status-rolling');
      if (stRolling) {
        const mergedCount = res.matched_count || res.total_with_rolling || (res.rolling_funds && res.rolling_funds.length) || 'Data';
        stRolling.textContent = `${mergedCount} funds matched`;
        stRolling.style.color = 'var(--emerald)';
      }
      const dropRolling = $('drop-zone-rolling');
      if (dropRolling) dropRolling.classList.add('loaded-card');
      const btnAddMore = $('btn-add-more-rolling');
      if (btnAddMore) btnAddMore.style.display = 'inline-flex';
      const btnReset = $('btn-reset-rolling');
      if (btnReset) btnReset.style.display = 'inline-flex';

      renderVerificationBanner(state.currentFilesMeta);
      renderMatchReport(res.match_report);
      updateUploadPanelVisibility();
      populateCategoryFilter();
      recomputeThresholds();
      renderAllViews();

      showToast(res.message, 'success');
    } catch (err) {
      showToast(`Rolling Upload Error: ${err.message}`, 'error');
    }
  }

  function resetRollingReturns() {
    if (!state.funds || state.funds.length === 0) return;

    state.funds.forEach(f => {
      f.rolling_1y = null;
      f.rolling_2y = null;
      f.rolling_3y = null;
      f.rolling_5y = null;
      f.rolling_avg = null;
    });

    state.rawRollingFunds = [];
    state.uploadedRollingSheets = [];
    selectedRollingFile = null;

    const fileRolling = $('file-rolling');
    if (fileRolling) fileRolling.value = '';

    const stRolling = $('status-rolling');
    if (stRolling) {
      stRolling.textContent = 'No file selected';
      stRolling.style.color = 'var(--text-muted)';
    }

    const dropRolling = $('drop-zone-rolling');
    if (dropRolling) dropRolling.classList.remove('loaded-card');

    const btnAddMore = $('btn-add-more-rolling');
    if (btnAddMore) btnAddMore.style.display = 'none';

    const btnReset = $('btn-reset-rolling');
    if (btnReset) btnReset.style.display = 'none';

    // Rolling data is gone again — reopen the panel so the sheet can be re-added.
    updateUploadPanelVisibility();
    renderAllViews();
    showToast('Rolling returns cleared. Risk ratios and rankings preserved.', 'info');
  }

  async function executeUpload() {
    if (!selectedRiskFile && !selectedRollingFile) {
      showToast('Please select Risk Ratios and/or Rolling Returns Excel file.', 'error');
      return;
    }

    const formData = new FormData();

    if (selectedRiskFile) {
      formData.append('risk_file', selectedRiskFile);
    }
    if (selectedRollingFile) {
      formData.append('rolling_file', selectedRollingFile);
    }

    showToast(`Reading and validating uploaded Excel sheet(s)...`, 'info');

    try {
      const response = await fetch(apiUrl('/api/upload'), {
        method: 'POST',
        body: formData
      });

      if (!response.ok) {
        const err = await response.json();
        throw new Error(err.detail || 'Failed to process Excel files');
      }

      const res = await response.json();
      loadAnalysisData(res.data, res.files_meta);
      showToast(`100% verified: Parsed ${res.data.summary.total_funds} funds directly from Excel!`, 'success');
    } catch (err) {
      showToast(`Excel Upload Error: ${err.message}`, 'error');
    }
  }

  function initActionButtons() {
    const btnExport = $('btn-export-excel');
    if (btnExport) {
      btnExport.addEventListener('click', async () => {
        if (!state.funds || state.funds.length === 0) {
          showToast('No data to export. Please upload your Excel files first.', 'error');
          return;
        }

        showToast('Generating formatted Excel report with rankings & Sharpe comparison...', 'info');
        try {
          const resp = await fetch(apiUrl('/api/export'), {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              funds: state.funds,
              summary: state.summary
            })
          });

          if (!resp.ok) throw new Error('Failed to generate export');

          const blob = await resp.blob();
          const url = window.URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = `Mutual_Fund_Analytics_${new Date().toISOString().slice(0,10)}.xlsx`;
          document.body.appendChild(a);
          a.click();
          a.remove();
          showToast('Excel report downloaded!', 'success');
        } catch (err) {
          showToast(`Export Error: ${err.message}`, 'error');
        }
      });
    }
  }

  // ==========================================
  // DATA PARSER & DISPATCHER
  // ==========================================
  function loadAnalysisData(data, filesMeta) {
    state.currentDataPayload = data;
    state.currentFilesMeta = filesMeta;

    state.funds = data.funds || [];
    state.rawRollingFunds = data.rolling_funds || [];
    state.summary = data.summary || {};
    state.categories = data.categories || [];
    state.categoryStats = data.category_stats || {};
    state.distribution = data.distribution || [];
    state.columnMapping = data.column_mapping || {};

    // Use the full-precision average (not the 2dp-rounded display value) as the actual
    // benchmark, so funds sitting right at the rounding boundary aren't misclassified.
    state.autoAvgSharpe = (typeof state.summary.avg_sharpe_exact === 'number') ? state.summary.avg_sharpe_exact : (state.summary.avg_sharpe || 0.62);
    state.benchmarkValue = state.autoAvgSharpe;

    // Show workspace & verification banner
    $('empty-state').style.display = 'none';
    $('kpi-section').style.display = 'block';
    $('benchmark-section').style.display = 'flex';
    $('dashboard-main').style.display = 'block';

    // Show Save Analysis Button
    const btnSave = $('btn-save-record');
    if (btnSave) btnSave.style.display = 'inline-flex';

    // Populate Excel Data Verification Banner
    renderVerificationBanner(filesMeta);
    renderMatchReport(data.match_report);
    updateUploadPanelVisibility();

    populateCategoryFilter();
    populateSchemaInspector();
    recomputeThresholds();
    renderAllViews();
    switchToTab(state.activeTab || 'tab-best-funds');
  }

  function renderVerificationBanner(filesMeta) {
    const banner = $('excel-verification-banner');
    const details = $('verify-details-text');
    const chipsContainer = $('verify-meta-chips');

    if (!banner) return;
    banner.style.display = 'flex';

    if (filesMeta && filesMeta.length > 0) {
      const names = filesMeta.map(f => `<strong>${escapeHtml(f.filename)}</strong> (${f.row_count} rows)`).join(' + ');
      details.innerHTML = `All ${state.summary.total_funds} funds parsed exclusively from uploaded Excel: ${names}`;

      chipsContainer.innerHTML = filesMeta.map(f => `
        <span class="verify-chip">
          <span>📊 ${escapeHtml(f.filename)}</span>
          <span class="verify-chip-sheet">Sheet: ${escapeHtml(f.active_sheet)}</span>
          <span class="verify-chip-count">${f.row_count} rows</span>
        </span>
      `).join('');
    } else {
      details.innerHTML = `All ${state.summary.total_funds} funds active in dashboard session.`;
      chipsContainer.innerHTML = '';
    }
  }

  // Shows how the Rolling-Returns scheme names lined up against the Risk-Ratios names,
  // so an unmatched row is visible rather than silently missing from the table.
  function renderMatchReport(report) {
    state.lastMatchReport = report || null;
    const el = $('match-report-notice');
    const toggleBtn = $('btn-toggle-match-report');
    const toggleLabel = $('btn-toggle-match-report-label');
    if (!el) return;

    if (!report || !report.rolling_rows) {
      el.style.display = 'none';
      el.innerHTML = '';
      if (toggleBtn) toggleBtn.style.display = 'none';
      return;
    }

    const unmatched = report.unmatched || 0;
    const matched = report.matched || 0;
    const viaSignature = report.matched_by_signature || 0;
    const ambiguous = (report.ambiguous_names || []).length;
    const isWarning = (unmatched > 0 || ambiguous > 0);

    const updateToolbarBtn = (visible) => {
      if (!toggleBtn) return;
      toggleBtn.style.display = 'inline-flex';
      const icon = isWarning ? '⚠️' : '✓';
      if (toggleLabel) {
        toggleLabel.textContent = visible
          ? `✕ Hide Match Report`
          : `${icon} Match Report (${matched}/${report.rolling_rows})`;
      }
    };

    if (state.matchReportDismissed) {
      el.style.display = 'none';
      updateToolbarBtn(false);
      return;
    }

    el.style.display = 'flex';
    updateToolbarBtn(true);

    if (!isWarning) {
      el.className = 'match-report-notice match-ok';
      el.innerHTML = `
        <span class="match-icon">✓</span>
        <div class="match-report-body">
          <strong>All ${report.rolling_rows} rolling-returns rows matched a fund</strong>
          ${viaSignature > 0 ? ` (${viaSignature} matched on a name variation such as "Reg"/"Regular" or "Gr"/"Growth")` : ''} —
          their rolling returns are showing in the table.
        </div>
        <div class="match-report-actions">
          <button type="button" class="btn-hide-match-notice" id="btn-hide-match-report" title="Hide this notice">
            ✕ Hide
          </button>
        </div>
      `;
    } else {
      const list = (report.unmatched_names || [])
        .map(n => `<li>${escapeHtml(n)}</li>`).join('');
      const ambiguousList = (report.ambiguous_names || [])
        .map(n => `<li>${escapeHtml(n)}</li>`).join('');

      el.className = 'match-report-notice match-warn';
      el.innerHTML = `
        <span class="match-icon">⚠️</span>
        <div class="match-report-body">
          <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:12px;">
            <div>
              <strong>${matched} of ${report.rolling_rows} rolling-returns rows matched a fund in the Risk-Ratios sheet</strong>
              ${viaSignature > 0 ? ` (${viaSignature} via a name variation)` : ''}.
              ${unmatched > 0 ? `The ${unmatched} below had no counterpart, so they were added as their own rows
                (Sharpe / Info Ratio / Treynor will show "-" for them):` : ''}
            </div>
            <div class="match-report-actions">
              <button type="button" class="btn-hide-match-notice" id="btn-hide-match-report" title="Hide this notice">
                ✕ Hide
              </button>
            </div>
          </div>
          ${unmatched > 0 ? `
            <ul class="match-name-list" id="match-unmatched-list">${list}</ul>
            <div>
              <button type="button" class="btn-toggle-match-list" id="btn-toggle-unmatched-list" data-collapsed="false">
                ▲ Hide fund list
              </button>
            </div>
          ` : ''}
          ${ambiguous > 0 ? `
            <div style="margin-top:8px;">
              ${ambiguous} name(s) were too ambiguous to match safely
              (more than one candidate shares the same name signature), so they were left unmatched:
              <ul class="match-name-list">${ambiguousList}</ul>
            </div>
          ` : ''}
        </div>
      `;
    }

    const btnHide = $('btn-hide-match-report');
    if (btnHide) {
      btnHide.addEventListener('click', () => {
        state.matchReportDismissed = true;
        el.style.display = 'none';
        updateToolbarBtn(false);
      });
    }

    const btnToggleList = $('btn-toggle-unmatched-list');
    const unmatchedList = $('match-unmatched-list');
    if (btnToggleList && unmatchedList) {
      btnToggleList.addEventListener('click', () => {
        const isCollapsed = btnToggleList.dataset.collapsed === 'true';
        if (isCollapsed) {
          unmatchedList.style.display = 'block';
          btnToggleList.dataset.collapsed = 'false';
          btnToggleList.innerHTML = '▲ Hide fund list';
        } else {
          unmatchedList.style.display = 'none';
          btnToggleList.dataset.collapsed = 'true';
          btnToggleList.innerHTML = `▼ Show ${unmatched} funds`;
        }
      });
    }
  }

  // ==========================================
  // SAVED RECORDS & SESSION MANAGER
  // ==========================================
  function initRecordManager() {
    const btnSave = $('btn-save-record');
    const btnOpenRecords = $('btn-open-records');
    const modalSave = $('modal-save-record');
    const modalRecords = $('modal-saved-records');
    const btnCloseSave = $('btn-close-save-modal');
    const btnCancelSave = $('btn-cancel-save');
    const btnConfirmSave = $('btn-confirm-save');
    const btnCloseRecords = $('btn-close-records-modal');
    const btnCloseRecordsBottom = $('btn-close-records-bottom');

    // Close Modals
    const closeAllModals = () => {
      if (modalSave) modalSave.style.display = 'none';
      if (modalRecords) modalRecords.style.display = 'none';
    };

    if (btnCloseSave) btnCloseSave.addEventListener('click', closeAllModals);
    if (btnCancelSave) btnCancelSave.addEventListener('click', closeAllModals);
    if (btnCloseRecords) btnCloseRecords.addEventListener('click', closeAllModals);
    if (btnCloseRecordsBottom) btnCloseRecordsBottom.addEventListener('click', closeAllModals);

    [modalSave, modalRecords].forEach(m => {
      if (m) {
        m.addEventListener('click', e => {
          if (e.target === m) closeAllModals();
        });
      }
    });

    // Open Save Modal
    if (btnSave) {
      btnSave.addEventListener('click', () => {
        if (!state.currentDataPayload || !state.funds || state.funds.length === 0) {
          showToast('No active mutual fund dataset to save. Please upload Excel files first.', 'error');
          return;
        }

        const now = new Date();
        const dateStr = now.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
        const nameInput = $('save-record-name');
        if (nameInput) nameInput.value = `Portfolio Analysis - ${dateStr}`;

        // Populate Preview Grid
        const sum = state.summary || {};
        $('save-preview-funds').textContent = sum.total_funds || state.funds.length;
        $('save-preview-sharpe').textContent = (sum.avg_sharpe || 0).toFixed(2);
        $('save-preview-above').textContent = `${sum.above_avg_count || 0} (${sum.outperformance_rate || 0}%)`;
        $('save-preview-aum').textContent = sum.total_aum ? '₹' + Math.round(sum.total_aum).toLocaleString('en-IN') + ' Cr' : '₹0 Cr';

        modalSave.style.display = 'flex';
        if (nameInput) nameInput.focus();
      });
    }

    // Confirm Save
    if (btnConfirmSave) {
      btnConfirmSave.addEventListener('click', async () => {
        const nameInput = $('save-record-name');
        const name = (nameInput ? nameInput.value.trim() : '') || 'Saved Portfolio Analysis';

        let filesSummary = '';
        if (state.currentFilesMeta && state.currentFilesMeta.length > 0) {
          filesSummary = state.currentFilesMeta.map(f => f.filename).join(' + ');
        } else if (selectedRiskFile || selectedRollingFile) {
          filesSummary = [selectedRiskFile?.name, selectedRollingFile?.name].filter(Boolean).join(' + ');
        }

        showToast('Saving analysis snapshot to database...', 'info');

        const payloadToSave = {
          funds: state.funds,
          summary: state.summary,
          categories: state.categories,
          category_stats: state.categoryStats || {},
          distribution: state.distribution || [],
          column_mapping: state.columnMapping || {},
          rolling_funds: state.rawRollingFunds || []
        };

        try {
          const resp = await fetch(apiUrl('/api/records/save'), {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              name: name,
              files_summary: filesSummary,
              data: payloadToSave
            })
          });

          if (!resp.ok) {
            const err = await resp.json();
            throw new Error(err.detail || 'Failed to save record');
          }

          const res = await resp.json();
          closeAllModals();
          updateSavedRecordsBadge();
          showToast(`Snapshot '${name}' saved successfully!`, 'success');
        } catch (err) {
          showToast(`Save Error: ${err.message}`, 'error');
        }
      });
    }

    // Open Saved Records History Modal
    if (btnOpenRecords) {
      btnOpenRecords.addEventListener('click', () => {
        loadSavedRecordsList();
        modalRecords.style.display = 'flex';
      });
    }

    // Badge count and DB status are fetched after sign-in (see bootstrapAuth) — calling
    // them here would just 401 while the login screen is still up.
  }

  async function checkDatabaseStatus() {
    try {
      const resp = await fetch(apiUrl('/api/db/status'));
      if (!resp.ok) return;
      const status = await resp.json();

      const badge = $('mongo-status-badge');
      const text = $('mongo-status-text');
      const tag = $('compass-status-tag');
      const details = $('compass-details-text');

      if (status.status === 'connected') {
        if (badge) {
          badge.classList.remove('status-offline');
          badge.title = 'Server MongoDB connected';
        }
        if (text) text.textContent = 'Storage: Online';
        if (tag) {
          tag.textContent = 'Server Connected';
          tag.classList.remove('tag-offline');
        }
        if (details) {
          details.textContent = 'Connected to server MongoDB. Saved analyses and working sessions are stored on the server.';
        }
      } else {
        if (badge) {
          badge.classList.add('status-offline');
          badge.title = 'Server MongoDB unreachable';
        }
        if (text) text.textContent = 'Storage: Offline';
        if (tag) {
          tag.textContent = 'Server Offline';
          tag.classList.add('tag-offline');
        }
        if (details) {
          details.textContent = 'Server MongoDB is unreachable. Changes cannot be saved until it reconnects.';
        }
      }
    } catch (e) {
      // silent
    }
  }

  async function updateSavedRecordsBadge() {
    try {
      const resp = await fetch(apiUrl('/api/records'));
      if (!resp.ok) return;
      const res = await resp.json();
      const badge = $('saved-records-badge');
      if (badge) badge.textContent = res.count || 0;
    } catch (e) {
      // silent
    }
  }

  async function loadSavedRecordsList() {
    const container = $('records-list-container');
    if (!container) return;

    container.innerHTML = '<div style="text-align:center;padding:24px;color:var(--text-muted);">Loading saved sessions...</div>';

    // Refresh database status in modal
    checkDatabaseStatus();

    try {
      const resp = await fetch(apiUrl('/api/records'));
      if (!resp.ok) throw new Error('Failed to load records');
      const res = await resp.json();

      const badge = $('saved-records-badge');
      if (badge) badge.textContent = res.count || 0;

      if (!res.records || res.records.length === 0) {
        container.innerHTML = `
          <div class="record-empty-state">
            <div class="record-empty-icon">📭</div>
            <h4>No Saved Records Yet</h4>
            <p>Upload your Excel spreadsheets and click <strong>"Save Analysis"</strong> to save snapshots for future reference.</p>
          </div>
        `;
        return;
      }

      container.innerHTML = res.records.map(r => `
        <div class="saved-record-card" id="record-card-${r.id}">
          <div class="record-main-info">
            <div class="record-title-row">
              <span class="record-title">${escapeHtml(r.name)}</span>
              <span class="record-date-badge">${escapeHtml(r.created_at)}</span>
              <span class="record-date-badge" style="background:rgba(37,99,235,0.08);color:#1d4ed8;">✓ ${escapeHtml(r.storage_engine) || 'Saved'}</span>
            </div>
            <div class="record-chips-row">
              <span class="record-stat-chip"><strong>${r.total_funds}</strong> Funds</span>
              <span class="record-stat-chip record-chip-blue">🎯 Avg Sharpe: <strong>${r.avg_sharpe.toFixed(2)}</strong></span>
              <span class="record-stat-chip record-chip-emerald">🚀 <strong>${r.above_avg_count}</strong> ≥ Avg (${r.outperformance_rate}%)</span>
              ${r.total_aum ? `<span class="record-stat-chip">💰 ₹${Math.round(r.total_aum).toLocaleString('en-IN')} Cr</span>` : ''}
              ${r.files_summary ? `<span class="record-stat-chip" title="Source files">📄 ${escapeHtml(r.files_summary)}</span>` : ''}
            </div>
          </div>
          <div class="record-actions">
            <button type="button" class="btn btn-sm btn-primary btn-load-record" data-id="${r.id}" title="Load this session into dashboard">
              <span>📂 Load in Dashboard</span>
            </button>
            <button type="button" class="btn btn-sm btn-secondary btn-delete-record" data-id="${r.id}" title="Delete this saved session">
              <span>🗑</span>
            </button>
          </div>
        </div>
      `).join('');

      // Attach Listeners
      container.querySelectorAll('.btn-load-record').forEach(btn => {
        btn.addEventListener('click', async () => {
          const recId = btn.dataset.id;
          await restoreSavedRecord(recId);
        });
      });

      container.querySelectorAll('.btn-delete-record').forEach(btn => {
        btn.addEventListener('click', async () => {
          const recId = btn.dataset.id;
          if (confirm('Are you sure you want to delete this saved record?')) {
            await deleteSavedRecord(recId);
          }
        });
      });

    } catch (err) {
      container.innerHTML = `<div style="color:var(--text-danger);text-align:center;padding:20px;">Failed to load saved records: ${err.message}</div>`;
    }
  }

  async function restoreSavedRecord(recordId) {
    showToast('Loading saved snapshot into dashboard...', 'info');
    try {
      const resp = await fetch(apiUrl(`/api/records/${recordId}`));
      if (!resp.ok) {
        const err = await resp.json();
        throw new Error(err.detail || 'Failed to load record');
      }

      const res = await resp.json();
      const meta = res.record_meta || {};

      // Close modal
      const modalRecords = $('modal-saved-records');
      if (modalRecords) modalRecords.style.display = 'none';

      // Load into dashboard
      loadAnalysisData(res.data, [{
        filename: meta.name,
        row_count: meta.total_funds,
        active_sheet: 'Saved Record'
      }]);

      showToast(`Restored snapshot: '${meta.name}' with ${meta.total_funds} mutual funds!`, 'success');
    } catch (err) {
      showToast(`Restore Error: ${err.message}`, 'error');
    }
  }

  async function deleteSavedRecord(recordId) {
    try {
      const resp = await fetch(apiUrl(`/api/records/${recordId}`), { method: 'DELETE' });
      if (!resp.ok) throw new Error('Delete failed');
      showToast('Record deleted successfully.', 'success');
      loadSavedRecordsList();
      updateSavedRecordsBadge();
    } catch (err) {
      showToast(`Delete Error: ${err.message}`, 'error');
    }
  }

  // ==========================================
  // BENCHMARK SLIDER & PRESETS
  // ==========================================
  function initBenchmarkControls() {
    const slider = $('benchmark-slider');
    const label = $('current-threshold-label');
    const presets = $$('.preset-btn');

    if (!slider) return;

    slider.addEventListener('input', e => {
      const val = parseFloat(e.target.value);
      state.benchmarkValue = val;
      if (label) label.textContent = val.toFixed(2);
      presets.forEach(p => p.classList.remove('active'));
      recomputeThresholds();
      renderAllViews();
    });

    presets.forEach(btn => {
      btn.addEventListener('click', () => {
        presets.forEach(p => p.classList.remove('active'));
        btn.classList.add('active');

        const type = btn.dataset.preset;
        state.benchmarkValue = (type === 'auto') ? state.autoAvgSharpe : parseFloat(type);

        slider.value = state.benchmarkValue;
        if (label) label.textContent = state.benchmarkValue.toFixed(2);
        recomputeThresholds();
        renderAllViews();
      });
    });
  }

  function recomputeThresholds() {
    const cut = state.benchmarkValue;
    let above = 0;
    let below = 0;

    state.funds.forEach(f => {
      if (f.sharpe_valid && f.sharpe !== null) {
        f.above_avg = f.sharpe >= cut;
        f.sharpe_diff = +(f.sharpe - cut).toFixed(2);
        if (f.above_avg) above++;
        else below++;
      } else {
        f.above_avg = false;
        f.sharpe_diff = null;
      }
    });

    state.summary.above_avg_count = above;
    state.summary.below_avg_count = below;
    state.summary.benchmark = cut;

    renderKPICards();
  }

  // ==========================================
  // KPI HUD
  // ==========================================
  function renderKPICards() {
    const s = state.summary;
    const validCount = s.valid_sharpe_count || 1;

    $('kpi-total-funds').textContent = s.total_funds || 0;
    $('kpi-cat-count').textContent = `${state.categories.length} Categories`;
    $('kpi-valid-count').textContent = `${s.valid_sharpe_count || 0} valid Sharpe`;

    $('kpi-avg-sharpe').textContent = (s.avg_sharpe !== undefined ? s.avg_sharpe.toFixed(2) : '0.00');
    $('mark-auto-avg').textContent = `Auto Avg (${s.avg_sharpe ? s.avg_sharpe.toFixed(2) : '0.62'})`;

    const abovePct = Math.round((s.above_avg_count / validCount) * 100);
    $('kpi-above-count').textContent = s.above_avg_count || 0;
    $('kpi-above-pct').textContent = `${abovePct}% of valid funds`;

    const belowPct = Math.round((s.below_avg_count / validCount) * 100);
    $('kpi-below-count').textContent = s.below_avg_count || 0;
    $('kpi-below-pct').textContent = `${belowPct}% of valid funds`;
    $('kpi-na-tag').textContent = `${s.na_count || 0} N/A data`;

    if (s.top_performer) {
      $('kpi-top-sharpe').textContent = (s.top_performer.sharpe ? s.top_performer.sharpe.toFixed(2) : 'N/A');
      $('kpi-top-name').textContent = s.top_performer.name;
    }

    $('kpi-total-aum').textContent = formatAUM(s.total_aum || 0);

    $('tab-count-all').textContent = s.total_funds || 0;
    const rollingCount = (state.rawRollingFunds && state.rawRollingFunds.length) || state.funds.filter(f => f.rolling_1y !== null).length;
    $('tab-count-rolling').textContent = rollingCount;
    $('rolling-table-count').textContent = rollingCount;
  }

  function formatAUM(val) {
    if (!val || val === 0) return '₹0 Cr';
    if (val >= 100000) return `₹${(val / 100000).toFixed(1)}L Cr`;
    if (val >= 1000) return `₹${(val / 1000).toFixed(1)}K Cr`;
    return `₹${val.toFixed(0)} Cr`;
  }

  // ==========================================
  // TAB NAVIGATION
  // ==========================================
  function switchToTab(tabId) {
    const tabBtns = $$('.tab-button');
    let matched = false;
    tabBtns.forEach(b => {
      const isTarget = b.dataset.tab === tabId;
      b.classList.toggle('active', isTarget);
      if (isTarget) matched = true;
    });
    if (!matched) return;

    $$('.tab-panel').forEach(p => p.classList.toggle('active', p.id === tabId));
    state.activeTab = tabId;

    if (tabId === 'tab-charts' || tabId === 'tab-rolling') {
      setTimeout(renderCharts, 50);
    } else if (tabId === 'tab-best-funds') {
      renderBestFundsView();
    } else if (tabId === 'tab-overview') {
      renderFundsTable();
    }
  }

  function initTabNavigation() {
    $$('.tab-button').forEach(btn => {
      btn.addEventListener('click', () => {
        switchToTab(btn.dataset.tab);
        scheduleSessionSave();
      });
    });
  }

  // ==========================================
  // TABLE FILTERING, SORTING & PAGINATION
  // ==========================================
  function initTableControls() {
    const searchInput = $('table-search');
    const clearSearch = $('clear-search');
    const catSelect = $('filter-category');
    const perfSelect = $('filter-performance');
    const toggleHighlight = $('toggle-highlight');
    const pageSizeSelect = $('page-size-select');
    const btnPrev = $('btn-prev-page');
    const btnNext = $('btn-next-page');

    if (searchInput) {
      searchInput.addEventListener('input', e => {
        state.searchQuery = e.target.value.toLowerCase().trim();
        state.currentPage = 1;
        if (clearSearch) clearSearch.style.display = state.searchQuery ? 'block' : 'none';
        renderFundsTable();
      });
    }

    if (clearSearch) {
      clearSearch.addEventListener('click', () => {
        searchInput.value = '';
        state.searchQuery = '';
        clearSearch.style.display = 'none';
        state.currentPage = 1;
        renderFundsTable();
      });
    }

    if (catSelect) {
      catSelect.addEventListener('change', e => {
        state.categoryFilter = e.target.value;
        state.currentPage = 1;
        renderFundsTable();
      });
    }

    if (perfSelect) {
      perfSelect.addEventListener('change', e => {
        state.perfFilter = e.target.value;
        state.currentPage = 1;
        renderFundsTable();
      });
    }

    if (toggleHighlight) {
      toggleHighlight.addEventListener('click', () => {
        state.highlightMode = !state.highlightMode;
        toggleHighlight.classList.toggle('active', state.highlightMode);
        renderFundsTable();
      });
    }

    if (pageSizeSelect) {
      pageSizeSelect.addEventListener('change', e => {
        state.pageSize = e.target.value === 'all' ? 999999 : parseInt(e.target.value);
        state.currentPage = 1;
        renderFundsTable();
      });
    }

    if (btnPrev) {
      btnPrev.addEventListener('click', () => {
        if (state.currentPage > 1) {
          state.currentPage--;
          renderFundsTable();
        }
      });
    }

    if (btnNext) {
      btnNext.addEventListener('click', () => {
        const filtered = getFilteredFunds();
        const maxPage = Math.ceil(filtered.length / state.pageSize);
        if (state.currentPage < maxPage) {
          state.currentPage++;
          renderFundsTable();
        }
      });
    }

    // Header Sort
    $$('#funds-table th[data-sort]').forEach(th => {
      th.addEventListener('click', () => {
        const col = th.dataset.sort;
        if (state.sortColumn === col) {
          state.sortDirection = state.sortDirection === 'asc' ? 'desc' : 'asc';
        } else {
          state.sortColumn = col;
          state.sortDirection = (col === 'name' || col === 'category') ? 'asc' : 'desc';
        }
        renderFundsTable();
      });
    });
  }

  function populateCategoryFilter() {
    const sel = $('filter-category');
    const bestSel = $('best-funds-cat-filter');
    const options = state.categories.map(cat => {
      const catNorm = cat.trim().toLowerCase();
      const count = state.funds.filter(f => (f.category || '').trim().toLowerCase() === catNorm).length;
      const safeCat = escapeHtml(cat);
      const label = escapeHtml(getCategoryLabel(cat));
      return `<option value="${safeCat}">${label} (${count})</option>`;
    }).join('');
    if (sel) {
      const prevVal = sel.value || state.categoryFilter || 'all';
      sel.innerHTML = '<option value="all">All Categories</option>' + options;
      sel.value = prevVal;
    }
    if (bestSel) {
      const prevBestVal = state.bestFundsCategoryFilter || 'all';
      bestSel.innerHTML = '<option value="all">All Categories</option>' + options;
      bestSel.value = prevBestVal;
    }
  }

  // ==========================================
  // MULTI-RATIO SCREENING & INDUSTRY PRESETS
  // ==========================================
  function initRatioFilterControls() {
    const presetChips = $$('.ratio-preset-chip');
    const btnToggleCustom = $('btn-toggle-custom-filter');
    const customDrawer = $('custom-filter-drawer');
    const customArrow = $('custom-filter-arrow');
    const btnApply = $('btn-apply-custom-filter');
    const btnReset = $('btn-reset-ratio-filters');

    // Preset Chips
    presetChips.forEach(chip => {
      chip.addEventListener('click', () => {
        presetChips.forEach(c => c.classList.remove('active'));
        chip.classList.add('active');

        const preset = chip.dataset.preset;
        state.activePreset = preset;
        state.currentPage = 1;

        // If Best Funds preset clicked, compute shortlist immediately
        if (preset === 'best-funds') {
          computeBestFundsShortlist();
        }

        // Reset manual custom filters when switching preset
        if (preset !== 'custom') {
          resetCustomInputs();
        }

        renderFundsTable();
      });
    });

    // Toggle Custom Drawer
    if (btnToggleCustom && customDrawer) {
      btnToggleCustom.addEventListener('click', () => {
        const isClosed = customDrawer.style.display === 'none';
        customDrawer.style.display = isClosed ? 'block' : 'none';
        if (customArrow) customArrow.textContent = isClosed ? '▴' : '▾';
      });
    }

    // Apply Custom Ratios
    if (btnApply) {
      btnApply.addEventListener('click', () => {
        const parseNum = id => {
          const el = $(id);
          if (!el || !el.value) return null;
          const v = parseFloat(el.value);
          return isNaN(v) ? null : v;
        };

        state.ratioFilters = {
          sharpeMin: parseNum('filter-sharpe-min'),
          sharpeMax: parseNum('filter-sharpe-max'),
          stdMax: parseNum('filter-std-max'),
          treynorMin: parseNum('filter-treynor-min'),
          infoMin: parseNum('filter-info-min'),
          rollingMin: parseNum('filter-rolling-min'),
          aumMin: parseNum('filter-aum-min')
        };

        presetChips.forEach(c => c.classList.remove('active'));
        state.activePreset = 'custom';
        state.currentPage = 1;

        renderFundsTable();
        showToast('Custom multi-ratio filters applied!', 'info');
      });
    }

    // Reset Ratio Filters
    if (btnReset) {
      btnReset.addEventListener('click', () => {
        resetCustomInputs();
        presetChips.forEach(c => c.classList.remove('active'));
        const allChip = document.querySelector('.ratio-preset-chip[data-preset="all"]');
        if (allChip) allChip.classList.add('active');
        state.activePreset = 'all';
        state.currentPage = 1;
        renderFundsTable();
        showToast('All ratio filters reset.', 'info');
      });
    }
  }

  function resetCustomInputs() {
    state.ratioFilters = {
      sharpeMin: null,
      sharpeMax: null,
      stdMax: null,
      treynorMin: null,
      infoMin: null,
      rollingMin: null,
      aumMin: null
    };

    ['filter-sharpe-min', 'filter-sharpe-max', 'filter-std-max', 'filter-treynor-min', 'filter-info-min', 'filter-rolling-min', 'filter-aum-min'].forEach(id => {
      const el = $(id);
      if (el) el.value = '';
    });
  }

  // ==========================================
  // CURATED FUND BASKET CONTROLS
  // ==========================================
  function initBasketControls() {
    const selectAllCheckbox = $('select-all-filtered');
    const btnExportBasket = $('btn-export-basket');
    const btnSaveBasket = $('btn-save-basket');
    const btnClearBasket = $('btn-clear-basket');

    // Select All Filtered Funds
    if (selectAllCheckbox) {
      selectAllCheckbox.addEventListener('change', e => {
        const checked = e.target.checked;
        const visibleFiltered = getFilteredFunds();

        visibleFiltered.forEach(f => {
          if (checked) {
            state.selectedFundsMap.set(f.name, f);
          } else {
            state.selectedFundsMap.delete(f.name);
          }
        });

        renderFundsTable();
        renderBasketBar();
      });
    }

    // Clear Basket
    if (btnClearBasket) {
      btnClearBasket.addEventListener('click', () => {
        state.selectedFundsMap.clear();
        if (selectAllCheckbox) selectAllCheckbox.checked = false;
        renderFundsTable();
        renderBasketBar();
        showToast('Basket cleared.', 'info');
      });
    }

    // Export Selected Basket to Excel
    if (btnExportBasket) {
      btnExportBasket.addEventListener('click', async () => {
        if (state.selectedFundsMap.size === 0) {
          showToast('No funds selected in the basket to export.', 'error');
          return;
        }

        const fundsArr = Array.from(state.selectedFundsMap.values());
        showToast(`Generating Excel export for ${fundsArr.length} selected funds...`, 'info');

        try {
          const resp = await fetch(apiUrl('/api/export-basket'), {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              funds: fundsArr,
              basket_name: `Curated_Basket_${fundsArr.length}_Funds`,
              filter_summary: state.activePreset === 'custom' ? 'Custom Multi-Ratio Filters' : `Preset: ${state.activePreset}`,
              benchmark: state.benchmarkValue
            })
          });

          if (!resp.ok) {
            const err = await resp.json();
            throw new Error(err.detail || 'Export failed');
          }

          const blob = await resp.blob();
          const url = window.URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = `curated_mutual_fund_basket_${fundsArr.length}_funds.xlsx`;
          document.body.appendChild(a);
          a.click();
          a.remove();
          window.URL.revokeObjectURL(url);
          showToast(`Downloaded basket report with ${fundsArr.length} funds!`, 'success');
        } catch (err) {
          showToast(`Export Error: ${err.message}`, 'error');
        }
      });
    }

    // Save Selected Basket as Snapshot in MongoDB
    if (btnSaveBasket) {
      btnSaveBasket.addEventListener('click', () => {
        if (state.selectedFundsMap.size === 0) {
          showToast('No funds selected in basket to save.', 'error');
          return;
        }

        const fundsArr = Array.from(state.selectedFundsMap.values());
        const validSharpes = fundsArr.filter(f => f.sharpe !== null && f.sharpe_valid).map(f => f.sharpe);
        const avgSharpe = validSharpes.length > 0 ? validSharpes.reduce((a, b) => a + b, 0) / validSharpes.length : 0;
        const aboveCount = validSharpes.filter(s => s >= state.benchmarkValue).length;
        const totalAUM = fundsArr.reduce((acc, f) => acc + (f.aum || 0), 0);

        const customPayload = {
          funds: fundsArr,
          summary: {
            total_funds: fundsArr.length,
            valid_sharpe_count: validSharpes.length,
            avg_sharpe: +avgSharpe.toFixed(2),
            above_avg_count: aboveCount,
            below_avg_count: validSharpes.length - aboveCount,
            outperformance_rate: validSharpes.length > 0 ? +((aboveCount / validSharpes.length) * 100).toFixed(1) : 0,
            total_aum: +totalAUM.toFixed(2),
            benchmark: state.benchmarkValue
          },
          categories: [...new Set(fundsArr.map(f => f.category))]
        };

        // Open save modal with basket name pre-populated
        const modalSave = $('modal-save-record');
        const nameInput = $('save-record-name');
        if (nameInput) nameInput.value = `Curated Basket (${fundsArr.length} Schemes) - ${new Date().toLocaleDateString('en-GB')}`;

        $('save-preview-funds').textContent = fundsArr.length;
        $('save-preview-sharpe').textContent = avgSharpe.toFixed(2);
        $('save-preview-above').textContent = `${aboveCount} (${customPayload.summary.outperformance_rate}%)`;
        $('save-preview-aum').textContent = formatAUM(totalAUM);

        // Temporarily assign currentDataPayload to this basket so confirm save saves it
        state.currentDataPayload = customPayload;

        if (modalSave) modalSave.style.display = 'flex';
        if (nameInput) nameInput.focus();
      });
    }
  }

  function renderBasketBar() {
    const bar = $('basket-toolbar');
    if (!bar) return;

    const count = state.selectedFundsMap.size;
    if (count === 0) {
      bar.style.display = 'none';
      return;
    }

    bar.style.display = 'flex';

    const funds = Array.from(state.selectedFundsMap.values());
    const validSharpes = funds.filter(f => f.sharpe !== null && f.sharpe_valid).map(f => f.sharpe);
    const avgSharpe = validSharpes.length > 0 ? (validSharpes.reduce((a, b) => a + b, 0) / validSharpes.length) : 0;
    const aboveCount = validSharpes.filter(s => s >= state.benchmarkValue).length;
    const rate = validSharpes.length > 0 ? Math.round((aboveCount / validSharpes.length) * 100) : 0;
    const totalAUM = funds.reduce((acc, f) => acc + (f.aum || 0), 0);

    $('basket-count').textContent = count;
    $('basket-avg-sharpe').textContent = avgSharpe.toFixed(2);
    $('basket-outperform-rate').textContent = `${rate}% (${aboveCount}/${validSharpes.length})`;
    $('basket-total-aum').textContent = formatAUM(totalAUM);
  }

  function getFilteredFunds() {
    return state.funds.filter(f => {
      // 1. Search query
      if (state.searchQuery) {
        const matchName = (f.name || '').toLowerCase().includes(state.searchQuery);
        const matchCat = (f.category || '').toLowerCase().includes(state.searchQuery);
        const matchCatLabel = getCategoryLabel(f.category).toLowerCase().includes(state.searchQuery);
        if (!matchName && !matchCat && !matchCatLabel) return false;
      }

      // 2. Category
      if (state.categoryFilter !== 'all' && f.category !== state.categoryFilter) {
        return false;
      }

      // 3. Performance dropdown
      if (state.perfFilter === 'above' && !f.above_avg) return false;
      if (state.perfFilter === 'below' && (f.above_avg || !f.sharpe_valid)) return false;
      if (state.perfFilter === 'valid' && !f.sharpe_valid) return false;

      // 4. Industry Ratio Presets
      if (state.activePreset === 'outperformers') {
        if (!f.above_avg) return false;
      } else if (state.activePreset === 'low-vol') {
        // Low volatility: Standard Deviation <= 16.0%
        if (f.std_dev === null || f.std_dev === undefined || f.std_dev > 16.0) return false;
      } else if (state.activePreset === 'high-treynor') {
        // High Treynor ratio >= 10.0
        if (f.treynor === null || f.treynor === undefined || f.treynor < 10.0) return false;
      } else if (state.activePreset === 'compounders') {
        // Rolling 3Y and 5Y CAGR >= 18%
        const r3 = f.rolling_3y;
        const r5 = f.rolling_5y;
        if (r3 === null || r3 === undefined || r3 < 18.0) return false;
        if (r5 !== null && r5 !== undefined && r5 < 18.0) return false;
      } else if (state.activePreset === 'giants') {
        // High AUM >= 5000 Cr
        if (f.aum === null || f.aum === undefined || f.aum < 5000) return false;
      } else if (state.activePreset === 'best-funds') {
        // Best Mutual Funds: Meets all 3 conditions (Sharpe > Avg, Info > Avg, Treynor > Avg)
        const bestNames = new Set((state.bestFundsShortlist || []).map(b => b.name));
        if (!bestNames.has(f.name)) return false;
      }

      // 5. Custom Multi-Ratio Filter Ranges
      const rf = state.ratioFilters;
      if (rf.sharpeMin !== null && (f.sharpe === null || f.sharpe < rf.sharpeMin)) return false;
      if (rf.sharpeMax !== null && (f.sharpe === null || f.sharpe > rf.sharpeMax)) return false;
      if (rf.stdMax !== null && (f.std_dev === null || f.std_dev > rf.stdMax)) return false;
      if (rf.treynorMin !== null && (f.treynor === null || f.treynor < rf.treynorMin)) return false;
      if (rf.infoMin !== null && (f.info_ratio === null || f.info_ratio < rf.infoMin)) return false;
      if (rf.rollingMin !== null) {
        const rollings = [f.rolling_1y, f.rolling_2y, f.rolling_3y, f.rolling_5y].filter(v => typeof v === 'number');
        if (rollings.length === 0 || Math.max(...rollings) < rf.rollingMin) return false;
      }
      if (rf.aumMin !== null && (f.aum === null || f.aum < rf.aumMin)) return false;

      return true;
    });
  }

  function renderFundsTable() {
    const tbody = $('funds-tbody');
    if (!tbody) return;

    // Filters, sort, pagination and basket changes all funnel through here, so this is
    // where the (debounced) persist-to-database is triggered from.
    scheduleSessionSave();

    let filtered = getFilteredFunds();

    // Update match indicator in custom filter drawer
    const matchInd = $('filter-match-indicator');
    if (matchInd) matchInd.textContent = `Matching ${filtered.length} of ${state.funds.length} funds`;

    // Sort
    const col = state.sortColumn;
    const dir = state.sortDirection === 'asc' ? 1 : -1;

    filtered.sort((a, b) => {
      const va = a[col];
      const vb = b[col];
      if ((va === null || va === undefined) && (vb === null || vb === undefined)) return 0;
      if (va === null || va === undefined) return 1;
      if (vb === null || vb === undefined) return -1;
      if (typeof va === 'string') return dir * va.localeCompare(vb);
      return dir * (va - vb);
    });

    // Pagination
    const total = filtered.length;
    const maxPage = Math.max(1, Math.ceil(total / state.pageSize));
    if (state.currentPage > maxPage) state.currentPage = maxPage;

    const start = (state.currentPage - 1) * state.pageSize;
    const end = Math.min(start + state.pageSize, total);
    const pageItems = filtered.slice(start, end);

    $('visible-count').textContent = total;
    $('total-count-label').textContent = state.funds.length;
    $('pagination-info').textContent = `Page ${state.currentPage} of ${maxPage}`;

    // Update Select All Checkbox state
    const selectAllCb = $('select-all-filtered');
    if (selectAllCb) {
      if (filtered.length > 0) {
        const allSelected = filtered.every(f => state.selectedFundsMap.has(f.name));
        const someSelected = filtered.some(f => state.selectedFundsMap.has(f.name));
        selectAllCb.checked = allSelected;
        selectAllCb.indeterminate = !allSelected && someSelected;
      } else {
        selectAllCb.checked = false;
        selectAllCb.indeterminate = false;
      }
    }

    if (pageItems.length === 0) {
      tbody.innerHTML = `<tr><td colspan="15" style="text-align:center;padding:36px;color:var(--text-muted);">No funds match the selected filter criteria. Try adjusting your ratio limits.</td></tr>`;
      return;
    }

    tbody.innerHTML = pageItems.map(f => {
      const isAbove = f.above_avg;
      const isValid = f.sharpe_valid;
      const isSelected = state.selectedFundsMap.has(f.name);

      let rowClass = isSelected ? 'row-selected' : '';
      if (state.highlightMode) {
        if (isAbove) rowClass += ' row-above-avg';
        else if (isValid) rowClass += ' row-below-avg';
      }

      // Rank Badge
      let rankBadge = `<span class="rank-badge">-</span>`;
      if (f.rank === 1) rankBadge = `<span class="rank-badge rank-gold">1</span>`;
      else if (f.rank === 2) rankBadge = `<span class="rank-badge rank-silver">2</span>`;
      else if (f.rank === 3) rankBadge = `<span class="rank-badge rank-bronze">3</span>`;
      else if (f.rank) rankBadge = `<span class="rank-badge">${f.rank}</span>`;

      // Sharpe Ratio Badge
      let sharpeBadge = `<span class="sharpe-badge sharpe-na">N/A</span>`;
      if (isValid && f.sharpe !== null) {
        const cls = isAbove ? 'sharpe-high' : 'sharpe-low';
        const arrow = isAbove ? '▲' : '▼';
        sharpeBadge = `<span class="sharpe-badge ${cls}">${arrow} ${f.sharpe.toFixed(2)}</span>`;
      }

      // Delta
      let deltaHTML = '-';
      if (f.sharpe_diff !== null && f.sharpe_diff !== undefined) {
        const cls = f.sharpe_diff >= 0 ? 'delta-pos' : 'delta-neg';
        const sign = f.sharpe_diff >= 0 ? '+' : '';
        deltaHTML = `<span class="delta-text ${cls}">${sign}${f.sharpe_diff.toFixed(2)}</span>`;
      }

      const avgRolling = f.rolling_avg !== null && f.rolling_avg !== undefined ? f.rolling_avg.toFixed(2) + '%' : '-';

      return `
        <tr class="${rowClass.trim()}">
          <td class="fund-select-col text-center">
            <input type="checkbox" class="fund-checkbox" data-name="${encodeURIComponent(f.name)}" ${isSelected ? 'checked' : ''}>
          </td>
          <td class="text-center">${rankBadge}</td>
          <td class="fund-name-cell"><span class="fund-name-text" title="${escapeHtml(f.name)}">${escapeHtml(f.name)}</span></td>
          <td><span class="category-pill" title="${escapeHtml(f.category)}">${escapeHtml(getCategoryLabel(f.category))}</span></td>
          <td class="text-right num-val" style="font-weight:600;">${f.aum !== null ? '₹' + Math.round(f.aum).toLocaleString('en-IN') : '-'}</td>
          <td class="text-center">${sharpeBadge}</td>
          <td class="text-center">${deltaHTML}</td>
          <td class="text-right num-val">${f.info_ratio !== null ? f.info_ratio.toFixed(2) : '-'}</td>
          <td class="text-right num-val">${f.treynor !== null ? f.treynor.toFixed(2) : '-'}</td>
          <td class="text-right num-val">${f.std_dev !== null ? f.std_dev.toFixed(2) + '%' : '-'}</td>
          <td class="text-right num-val" style="font-weight:700;color:${getReturnColor(f.rolling_1y)}">${f.rolling_1y !== null ? f.rolling_1y.toFixed(2) + '%' : '-'}</td>
          <td class="text-right num-val" style="font-weight:700;color:${getReturnColor(f.rolling_2y)}">${f.rolling_2y !== null ? f.rolling_2y.toFixed(2) + '%' : '-'}</td>
          <td class="text-right num-val" style="font-weight:700;color:${getReturnColor(f.rolling_3y)}">${f.rolling_3y !== null ? f.rolling_3y.toFixed(2) + '%' : '-'}</td>
          <td class="text-right num-val" style="font-weight:700;color:${getReturnColor(f.rolling_5y)}">${f.rolling_5y !== null ? f.rolling_5y.toFixed(2) + '%' : '-'}</td>
          <td class="text-right num-val" style="font-weight:700;color:${getReturnColor(f.rolling_avg)}">${avgRolling}</td>
        </tr>
      `;
    }).join('');

    // Attach Checkbox Listeners
    tbody.querySelectorAll('.fund-checkbox').forEach(cb => {
      cb.addEventListener('change', e => {
        const rawName = decodeURIComponent(cb.dataset.name);
        const fund = state.funds.find(f => f.name === rawName);
        if (fund) {
          if (cb.checked) {
            state.selectedFundsMap.set(fund.name, fund);
          } else {
            state.selectedFundsMap.delete(fund.name);
          }
          renderBasketBar();

          // Toggle row highlight
          const tr = cb.closest('tr');
          if (tr) tr.classList.toggle('row-selected', cb.checked);

          // Update header checkbox
          const visible = getFilteredFunds();
          const allSelected = visible.every(f => state.selectedFundsMap.has(f.name));
          const someSelected = visible.some(f => state.selectedFundsMap.has(f.name));
          if (selectAllCb) {
            selectAllCb.checked = allSelected;
            selectAllCb.indeterminate = !allSelected && someSelected;
          }
        }
      });
    });

    renderPaginationButtons(maxPage);
  }

  function renderPaginationButtons(maxPage) {
    const box = $('page-numbers');
    if (!box) return;

    let html = '';
    const cur = state.currentPage;

    for (let i = 1; i <= maxPage; i++) {
      if (i === 1 || i === maxPage || (i >= cur - 1 && i <= cur + 1)) {
        html += `<button class="page-num-btn ${i === cur ? 'active' : ''}" data-page="${i}">${i}</button>`;
      } else if (i === cur - 2 || i === cur + 2) {
        html += `<span style="padding:4px;color:var(--text-muted);">…</span>`;
      }
    }

    box.innerHTML = html;
    box.querySelectorAll('.page-num-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        state.currentPage = parseInt(btn.dataset.page);
        renderFundsTable();
      });
    });

    $('btn-prev-page').disabled = (cur === 1);
    $('btn-next-page').disabled = (cur === maxPage);
  }

  function getReturnColor(val) {
    if (val === null || val === undefined) return 'var(--text-muted)';
    if (val >= 25) return '#166534';
    if (val >= 20) return '#237346';
    if (val >= 15) return '#245bb2';
    if (val >= 10) return '#926014';
    return '#a33932';
  }

  // ==========================================
  // CHARTS (CHART.JS)
  // ==========================================
  function destroyChart(id) {
    if (chartRegistry[id]) {
      chartRegistry[id].destroy();
      delete chartRegistry[id];
    }
  }

  function renderCharts() {
    if (!state.funds || state.funds.length === 0) return;
    renderSharpeDistChart();
    renderTopFundsChart();
    renderCategoryChart();
    renderAumScatterChart();
    renderRollingGroupedChart();
    renderRollingVsSharpeChart();
  }

  function renderSharpeDistChart() {
    destroyChart('chart-sharpe-dist');
    const canvas = $('chart-sharpe-dist');
    if (!canvas) return;
    const ctx = canvas.getContext('2d');

    const bins = state.distribution || [];
    const labels = bins.map(b => b.label);
    const data = bins.map(b => b.count);
    const colors = bins.map(b => (b.min >= state.benchmarkValue ? 'rgba(5, 150, 105, 0.85)' : 'rgba(220, 38, 38, 0.75)'));

    chartRegistry['chart-sharpe-dist'] = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: labels,
        datasets: [{
          data: data,
          backgroundColor: colors,
          borderRadius: 4
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: '#0f172a',
            padding: 10,
            cornerRadius: 6
          }
        },
        scales: {
          x: { grid: { color: '#f1f5f9' }, ticks: { color: '#64748b', font: { size: 10, weight: '600' } } },
          y: { grid: { color: '#f1f5f9' }, ticks: { color: '#64748b', font: { size: 10 } } }
        }
      }
    });
  }

  function renderTopFundsChart() {
    destroyChart('chart-top-funds');
    const canvas = $('chart-top-funds');
    if (!canvas) return;
    const ctx = canvas.getContext('2d');

    const topFunds = [...state.funds]
      .filter(f => f.sharpe_valid && f.sharpe !== null)
      .sort((a, b) => b.sharpe - a.sharpe)
      .slice(0, 15);

    const gradient = ctx.createLinearGradient(0, 0, ctx.canvas.width, 0);
    gradient.addColorStop(0, '#2563eb');
    gradient.addColorStop(1, '#1d4ed8');

    chartRegistry['chart-top-funds'] = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: topFunds.map(f => truncText(f.name, 25)),
        datasets: [{
          data: topFunds.map(f => f.sharpe),
          backgroundColor: gradient,
          borderRadius: 4
        }]
      },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: '#0f172a',
            padding: 10,
            cornerRadius: 6,
            callbacks: {
              title: ctx => topFunds[ctx[0].dataIndex].name,
              afterLabel: ctx => `Category: ${getCategoryLabel(topFunds[ctx[0].dataIndex].category)} | AUM: ₹${topFunds[ctx[0].dataIndex].aum || '-'}`
            }
          }
        },
        scales: {
          x: { grid: { color: '#f1f5f9' }, ticks: { color: '#64748b', font: { size: 10 } } },
          y: { grid: { display: false }, ticks: { color: '#0f172a', font: { size: 10, weight: '600' } } }
        }
      }
    });
  }

  function renderCategoryChart() {
    destroyChart('chart-categories');
    const canvas = $('chart-categories');
    if (!canvas) return;
    const ctx = canvas.getContext('2d');

    const catNames = Object.keys(state.categoryStats)
      .filter(c => state.categoryStats[c].avg_sharpe !== null)
      .sort((a, b) => (state.categoryStats[b].avg_sharpe || 0) - (state.categoryStats[a].avg_sharpe || 0));
    const catLabels = catNames.map(getCategoryLabel);

    const avgs = catNames.map(c => state.categoryStats[c].avg_sharpe);
    const colors = avgs.map(v => v >= state.benchmarkValue ? 'rgba(5, 150, 105, 0.85)' : 'rgba(220, 38, 38, 0.7)');

    chartRegistry['chart-categories'] = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: catLabels,
        datasets: [{
          data: avgs,
          backgroundColor: colors,
          borderRadius: 4
        }]
      },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: '#0f172a',
            padding: 10,
            cornerRadius: 6,
            callbacks: {
              afterLabel: ctx => {
                const stat = state.categoryStats[catNames[ctx.dataIndex]];
                return `Code: ${catNames[ctx.dataIndex]} | Total Funds: ${stat.count} | Outperformers: ${stat.above_avg_count} (${stat.outperformance_rate}%) | AUM: ₹${stat.total_aum} Cr`;
              }
            }
          }
        },
        scales: {
          x: { grid: { color: '#f1f5f9' }, ticks: { color: '#64748b', font: { size: 10 } } },
          y: { grid: { display: false }, ticks: { color: '#334155', font: { size: 9, weight: '600' } } }
        }
      }
    });
  }

  function renderAumScatterChart() {
    destroyChart('chart-aum-scatter');
    const canvas = $('chart-aum-scatter');
    if (!canvas) return;
    const ctx = canvas.getContext('2d');

    const aboveData = state.funds
      .filter(f => f.sharpe_valid && f.aum !== null && f.above_avg)
      .map(f => ({ x: f.aum, y: f.sharpe, name: f.name, category: f.category }));

    const belowData = state.funds
      .filter(f => f.sharpe_valid && f.aum !== null && !f.above_avg)
      .map(f => ({ x: f.aum, y: f.sharpe, name: f.name, category: f.category }));

    chartRegistry['chart-aum-scatter'] = new Chart(ctx, {
      type: 'scatter',
      data: {
        datasets: [
          {
            label: '≥ Benchmark Sharpe',
            data: aboveData,
            backgroundColor: 'rgba(5, 150, 105, 0.8)',
            borderColor: '#059669',
            pointRadius: 5,
            pointHoverRadius: 8
          },
          {
            label: '< Benchmark Sharpe',
            data: belowData,
            backgroundColor: 'rgba(220, 38, 38, 0.65)',
            borderColor: '#dc2626',
            pointRadius: 4,
            pointHoverRadius: 7
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { labels: { color: '#475569', font: { size: 10, weight: '600' }, usePointStyle: true } },
          tooltip: {
            backgroundColor: '#0f172a',
            padding: 10,
            cornerRadius: 6,
            callbacks: {
              title: ctx => ctx[0].raw.name,
              label: ctx => [
                `AUM: ₹${ctx.raw.x.toLocaleString('en-IN')} Cr`,
                `Sharpe Ratio: ${ctx.raw.y.toFixed(2)}`,
                `Category: ${getCategoryLabel(ctx.raw.category)}`
              ]
            }
          }
        },
        scales: {
          x: {
            type: 'logarithmic',
            title: { display: true, text: 'AUM (₹ Crore, Log Scale)', color: '#64748b', font: { size: 10, weight: '600' } },
            grid: { color: '#f1f5f9' },
            ticks: { color: '#64748b', font: { size: 9 }, callback: v => v >= 1000 ? (v/1000)+'K' : v }
          },
          y: {
            title: { display: true, text: 'Sharpe Ratio', color: '#64748b', font: { size: 10, weight: '600' } },
            grid: { color: '#f1f5f9' },
            ticks: { color: '#64748b', font: { size: 9 } }
          }
        }
      }
    });
  }

  function renderRollingGroupedChart() {
    destroyChart('chart-rolling-grouped');
    const canvas = $('chart-rolling-grouped');
    if (!canvas) return;
    const ctx = canvas.getContext('2d');

    const source = (state.rawRollingFunds && state.rawRollingFunds.length > 0)
      ? state.rawRollingFunds
      : (state.funds || []).filter(f => f.rolling_1y !== null);

    const countLabel = $('rolling-table-count');
    if (countLabel) countLabel.textContent = source ? source.length : 0;

    if (!source || source.length === 0) return;

    chartRegistry['chart-rolling-grouped'] = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: source.map(f => truncText(f.name, 20)),
        datasets: [
          { label: '1Y Rolling', data: source.map(f => f.rolling_1y), backgroundColor: '#2563eb', borderRadius: 3 },
          { label: '2Y Rolling', data: source.map(f => f.rolling_2y), backgroundColor: '#059669', borderRadius: 3 },
          { label: '3Y Rolling', data: source.map(f => f.rolling_3y), backgroundColor: '#d97706', borderRadius: 3 },
          { label: '5Y Rolling', data: source.map(f => f.rolling_5y), backgroundColor: '#7c3aed', borderRadius: 3 }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { labels: { color: '#475569', font: { size: 10, weight: '600' }, usePointStyle: true } },
          tooltip: {
            backgroundColor: '#0f172a',
            padding: 10,
            cornerRadius: 6,
            callbacks: {
              title: ctx => source[ctx[0].dataIndex].name,
              label: ctx => `${ctx.dataset.label}: ${ctx.raw !== null ? ctx.raw.toFixed(2) + '%' : 'N/A'}`
            }
          }
        },
        scales: {
          x: { grid: { display: false }, ticks: { color: '#475569', font: { size: 8, weight: '500' }, maxRotation: 45 } },
          y: { grid: { color: '#f1f5f9' }, ticks: { color: '#64748b', font: { size: 9 }, callback: v => v + '%' } }
        }
      }
    });

    renderRollingTable(source);
  }

  function renderRollingVsSharpeChart() {
    destroyChart('chart-rolling-vs-sharpe');
    const canvas = $('chart-rolling-vs-sharpe');
    if (!canvas) return;
    const ctx = canvas.getContext('2d');

    const source = (state.rawRollingFunds && state.rawRollingFunds.length > 0)
      ? state.rawRollingFunds
      : (state.funds || []).filter(f => f.rolling_5y !== null && f.sharpe !== null);

    if (!source || source.length === 0) return;

    const scatterData = source
      .filter(f => f.sharpe !== null && f.rolling_5y !== null)
      .map(f => ({ x: f.sharpe, y: f.rolling_5y, name: f.name, std: f.std_dev }));

    chartRegistry['chart-rolling-vs-sharpe'] = new Chart(ctx, {
      type: 'scatter',
      data: {
        datasets: [{
          label: 'Funds (5Y Rolling Return vs Sharpe)',
          data: scatterData,
          backgroundColor: 'rgba(37, 99, 235, 0.75)',
          borderColor: '#2563eb',
          pointRadius: 6,
          pointHoverRadius: 9
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { labels: { color: '#475569', font: { size: 10, weight: '600' }, usePointStyle: true } },
          tooltip: {
            backgroundColor: '#0f172a',
            padding: 10,
            cornerRadius: 6,
            callbacks: {
              title: ctx => ctx[0].raw.name,
              label: ctx => [
                `Sharpe Ratio: ${ctx.raw.x.toFixed(2)}`,
                `5Y Rolling Return: ${ctx.raw.y.toFixed(2)}%`,
                `Std Dev: ${ctx.raw.std ? ctx.raw.std.toFixed(2) + '%' : '-'}`
              ]
            }
          }
        },
        scales: {
          x: {
            title: { display: true, text: 'Sharpe Ratio (Risk-Adjusted Return)', color: '#64748b', font: { size: 10, weight: '600' } },
            grid: { color: '#f1f5f9' },
            ticks: { color: '#64748b', font: { size: 9 } }
          },
          y: {
            title: { display: true, text: '5Y Rolling Return (%)', color: '#64748b', font: { size: 10, weight: '600' } },
            grid: { color: '#f1f5f9' },
            ticks: { color: '#64748b', font: { size: 9 }, callback: v => v + '%' }
          }
        }
      }
    });
  }

  function renderRollingTable(rollingFunds) {
    const tbody = $('rolling-tbody');
    if (!tbody) return;

    if (!rollingFunds || rollingFunds.length === 0) {
      tbody.innerHTML = '<tr><td colspan="12" style="text-align:center;padding:2rem;color:var(--text-muted);">No rolling returns loaded yet.</td></tr>';
      return;
    }

    const countLabel = $('rolling-table-count');
    if (countLabel) countLabel.textContent = rollingFunds.length;

    tbody.innerHTML = rollingFunds.map((f, i) => {
      const avg = f.rolling_avg !== null && f.rolling_avg !== undefined ? f.rolling_avg.toFixed(2) + '%' : '-';
      return `
        <tr>
          <td class="text-center"><span class="rank-badge">${i + 1}</span></td>
          <td class="fund-name-cell"><span class="fund-name-text" title="${escapeHtml(f.name)}">${escapeHtml(f.name)}</span></td>
          <td><span class="category-pill" title="${escapeHtml(f.category)}">${escapeHtml(getCategoryLabel(f.category))}</span></td>
          <td class="num-val" style="color:var(--text-muted);font-size:0.75rem;">${escapeHtml(f.inception_date) || '-'}</td>
          <td class="text-right num-val" style="font-weight:600;">${f.aum !== null ? '₹' + Math.round(f.aum).toLocaleString('en-IN') : '-'}</td>
          <td class="text-center">${f.sharpe !== null ? `<span class="sharpe-badge ${f.sharpe >= state.benchmarkValue ? 'sharpe-high' : 'sharpe-low'}">${f.sharpe.toFixed(2)}</span>` : '<span class="sharpe-badge sharpe-na">N/A</span>'}</td>
          <td class="text-right num-val">${f.std_dev !== null ? f.std_dev.toFixed(2) + '%' : '-'}</td>
          <td class="text-right num-val" style="font-weight:700;color:${getReturnColor(f.rolling_1y)}">${f.rolling_1y !== null ? f.rolling_1y.toFixed(2) + '%' : '-'}</td>
          <td class="text-right num-val" style="font-weight:700;color:${getReturnColor(f.rolling_2y)}">${f.rolling_2y !== null ? f.rolling_2y.toFixed(2) + '%' : '-'}</td>
          <td class="text-right num-val" style="font-weight:700;color:${getReturnColor(f.rolling_3y)}">${f.rolling_3y !== null ? f.rolling_3y.toFixed(2) + '%' : '-'}</td>
          <td class="text-right num-val" style="font-weight:700;color:${getReturnColor(f.rolling_5y)}">${f.rolling_5y !== null ? f.rolling_5y.toFixed(2) + '%' : '-'}</td>
          <td class="text-right num-val" style="font-weight:700;color:${getReturnColor(f.rolling_avg)}">${avg}</td>
        </tr>
      `;
    }).join('');
  }

  // ==========================================
  // SCHEMA INSPECTOR
  // ==========================================
  function populateSchemaInspector() {
    const container = $('mapping-controls');
    if (!container) return;

    const m = state.columnMapping || {};
    container.innerHTML = `
      <div class="schema-field">
        <label>Scheme Name Column</label>
        <input type="text" class="form-select" value="${escapeHtml(m.name)}" readonly>
      </div>
      <div class="schema-field">
        <label>Category Column</label>
        <input type="text" class="form-select" value="${escapeHtml(m.category)}" readonly>
      </div>
      <div class="schema-field">
        <label>AUM Column</label>
        <input type="text" class="form-select" value="${escapeHtml(m.aum)}" readonly>
      </div>
      <div class="schema-field">
        <label>Sharpe Ratio Column</label>
        <input type="text" class="form-select" value="${escapeHtml(m.sharpe)}" readonly>
      </div>
      <div class="schema-field">
        <label>Information Ratio Column</label>
        <input type="text" class="form-select" value="${escapeHtml(m.info_ratio)}" readonly>
      </div>
      <div class="schema-field">
        <label>Treynor Ratio Column</label>
        <input type="text" class="form-select" value="${escapeHtml(m.treynor)}" readonly>
      </div>
    `;
  }

  // ==========================================
  // BEST MUTUAL FUNDS FINDER (3-FACTOR ALPHA + ROLLING DOMINANCE)
  // ==========================================
  function computeBestFundsShortlist() {
    const allFunds = state.funds || [];
    if (allFunds.length === 0) {
      state.bestFundsShortlist = [];
      state.bestFundsChampion = null;
      return { shortlisted: [], champion: null, periodWinners: {}, benchmarkMetrics: {}, totalCandidates: 0 };
    }

    const basis = state.bestFundsBenchmarkBasis || 'category';
    const catFilter = state.bestFundsCategoryFilter || 'all';

    // 1. Compute Portfolio Averages (full precision — rounding here before the
    // ">" comparison below can wrongly exclude funds that sit right at the
    // boundary, e.g. a fund at 0.67 vs a true average of 0.6689 that rounds to 0.67).
    const validSharpes = allFunds.filter(f => typeof f.sharpe === 'number').map(f => f.sharpe);
    const portfolioAvgSharpe = validSharpes.length > 0 ? validSharpes.reduce((a, b) => a + b, 0) / validSharpes.length : 0.62;

    const validInfos = allFunds.filter(f => typeof f.info_ratio === 'number').map(f => f.info_ratio);
    const portfolioAvgInfo = validInfos.length > 0 ? validInfos.reduce((a, b) => a + b, 0) / validInfos.length : 0.0;

    const validTreynors = allFunds.filter(f => typeof f.treynor === 'number').map(f => f.treynor);
    const portfolioAvgTreynor = validTreynors.length > 0 ? validTreynors.reduce((a, b) => a + b, 0) / validTreynors.length : 0.0;

    // 2. Compute Category Averages (also full precision, case-insensitive normalized)
    const catAverages = {};
    (state.categories || []).forEach(cat => {
      const catNorm = (cat || '').trim().toLowerCase();
      const cFunds = allFunds.filter(f => (f.category || '').trim().toLowerCase() === catNorm);
      const cSharpes = cFunds.filter(f => typeof f.sharpe === 'number').map(f => f.sharpe);
      const cInfos = cFunds.filter(f => typeof f.info_ratio === 'number').map(f => f.info_ratio);
      const cTreynors = cFunds.filter(f => typeof f.treynor === 'number').map(f => f.treynor);

      const catObj = {
        avgSharpe: cSharpes.length > 0 ? cSharpes.reduce((a, b) => a + b, 0) / cSharpes.length : portfolioAvgSharpe,
        avgInfo: cInfos.length > 0 ? cInfos.reduce((a, b) => a + b, 0) / cInfos.length : portfolioAvgInfo,
        avgTreynor: cTreynors.length > 0 ? cTreynors.reduce((a, b) => a + b, 0) / cTreynors.length : portfolioAvgTreynor
      };
      catAverages[cat] = catObj;
      catAverages[catNorm] = catObj;
    });

    // 3. Evaluate Conditions 1, 2, 3 on candidate funds:
    const candidateFunds = catFilter === 'all'
      ? allFunds
      : allFunds.filter(f => (f.category || '').trim().toLowerCase() === catFilter.trim().toLowerCase());

    const evaluated = candidateFunds.map(f => {
      let targetAvgSharpe = portfolioAvgSharpe;
      let targetAvgInfo = portfolioAvgInfo;
      let targetAvgTreynor = portfolioAvgTreynor;

      const catNorm = (f.category || '').trim().toLowerCase();
      if (basis === 'category' && catNorm && catAverages[catNorm]) {
        targetAvgSharpe = catAverages[catNorm].avgSharpe;
        targetAvgInfo = catAverages[catNorm].avgInfo;
        targetAvgTreynor = catAverages[catNorm].avgTreynor;
      }

      const condSharpe = typeof f.sharpe === 'number' && f.sharpe > targetAvgSharpe;
      const condInfo = typeof f.info_ratio === 'number' && f.info_ratio > targetAvgInfo;
      const condTreynor = typeof f.treynor === 'number' && f.treynor > targetAvgTreynor;

      const condCount = (condSharpe ? 1 : 0) + (condInfo ? 1 : 0) + (condTreynor ? 1 : 0);
      const passesAll3 = (condSharpe && condInfo && condTreynor);

      return {
        ...f,
        targetAvgSharpe,
        targetAvgInfo,
        targetAvgTreynor,
        sharpeDelta: typeof f.sharpe === 'number' ? +(f.sharpe - targetAvgSharpe).toFixed(2) : null,
        infoDelta: typeof f.info_ratio === 'number' ? +(f.info_ratio - targetAvgInfo).toFixed(2) : null,
        treynorDelta: typeof f.treynor === 'number' ? +(f.treynor - targetAvgTreynor).toFixed(2) : null,
        condSharpe,
        condInfo,
        condTreynor,
        condCount,
        passesAll3,
        periodsWon: 0,
        wonHorizons: []
      };
    });

    // Step 5: Rolling Returns Period Dominance Comparison (1Y, 2Y, 3Y, 5Y)
    const horizons = [
      { key: 'rolling_1y', label: '1Y Rolling', short: '1Y' },
      { key: 'rolling_2y', label: '2Y Rolling', short: '2Y' },
      { key: 'rolling_3y', label: '3Y Rolling', short: '3Y' },
      { key: 'rolling_5y', label: '5Y Rolling', short: '5Y' }
    ];

    const periodWinners = {};

    horizons.forEach(h => {
      // Find highest return in this period among evaluated candidate funds
      const eligible = evaluated.filter(f => typeof f[h.key] === 'number');

      if (eligible.length > 0) {
        const maxVal = Math.max(...eligible.map(f => f[h.key]));
        const tiedWinners = eligible.filter(f => f[h.key] === maxVal);
        const winner = tiedWinners[0];
        periodWinners[h.short] = {
          horizon: h.label,
          short: h.short,
          winnerName: winner ? winner.name : '–',
          winnerCategory: winner ? winner.category : '',
          tiedCount: tiedWinners.length,
          returnVal: maxVal
        };

        // Award period win to funds matching this top return
        evaluated.forEach(sf => {
          if (typeof sf[h.key] === 'number' && sf[h.key] === maxVal) {
            sf.periodsWon += 1;
            sf.wonHorizons.push(h.short);
          }
        });
      } else {
        periodWinners[h.short] = {
          horizon: h.label,
          short: h.short,
          winnerName: 'No rolling data',
          winnerCategory: '',
          returnVal: null
        };
      }
    });

    // Strict 3 of 3 qualified outperformers
    const strict3Of3 = evaluated.filter(f => f.passesAll3);

    // Sort evaluated funds:
    // 1st by condition count (3 -> 2 -> 1 -> 0)
    // 2nd by maximum periods won (desc)
    // 3rd by average rolling return (desc)
    // 4th by Sharpe (desc)
    evaluated.sort((a, b) => {
      if (b.condCount !== a.condCount) return b.condCount - a.condCount;
      if (b.periodsWon !== a.periodsWon) return b.periodsWon - a.periodsWon;
      const bRoll = typeof b.rolling_avg === 'number' ? b.rolling_avg : -999;
      const aRoll = typeof a.rolling_avg === 'number' ? a.rolling_avg : -999;
      if (bRoll !== aRoll) return bRoll - aRoll;
      return (b.sharpe || 0) - (a.sharpe || 0);
    });

    // Shortlist determination:
    // When a specific category is chosen, display all funds in that category ranked!
    // When "all" categories are chosen, display strict 3/3 qualified funds (or all evaluated if none cleared 3/3)
    let shortlisted;
    if (catFilter !== 'all') {
      shortlisted = evaluated;
    } else {
      shortlisted = strict3Of3.length > 0 ? strict3Of3 : evaluated;
    }

    const champion = shortlisted.length > 0 ? shortlisted[0] : null;

    state.bestFundsShortlist = shortlisted;
    state.bestFundsChampion = champion;

    const catFilterNorm = catFilter.trim().toLowerCase();
    const benchmarkMetrics = {
      avgSharpe: +(basis === 'category' && catFilter !== 'all' && catAverages[catFilterNorm] ? catAverages[catFilterNorm].avgSharpe : portfolioAvgSharpe).toFixed(2),
      avgInfo: +(basis === 'category' && catFilter !== 'all' && catAverages[catFilterNorm] ? catAverages[catFilterNorm].avgInfo : portfolioAvgInfo).toFixed(2),
      avgTreynor: +(basis === 'category' && catFilter !== 'all' && catAverages[catFilterNorm] ? catAverages[catFilterNorm].avgTreynor : portfolioAvgTreynor).toFixed(2)
    };

    const rollingKeys = ['rolling_1y', 'rolling_2y', 'rolling_3y', 'rolling_5y'];
    const withRollingCount = shortlisted.filter(f => rollingKeys.some(k => typeof f[k] === 'number')).length;

    return {
      shortlisted,
      strict3Of3,
      champion,
      periodWinners,
      benchmarkMetrics,
      withRollingCount,
      totalCandidates: candidateFunds.length
    };
  }

  function renderBestFundsView() {
    const data = computeBestFundsShortlist();
    const { shortlisted, champion, periodWinners, benchmarkMetrics, withRollingCount, totalCandidates } = data;

    // 1. Update Benchmark Criteria Labels
    const sharpeValEl = $('best-crit-sharpe-val');
    const infoValEl = $('best-crit-info-val');
    const treynorValEl = $('best-crit-treynor-val');
    const shortlistCountEl = $('best-shortlist-count');
    const shortlistPctEl = $('best-shortlist-pct');
    const tabCountBestEl = $('tab-count-best');
    const bestTableCountEl = $('best-table-count');

    if (sharpeValEl) sharpeValEl.textContent = `> ${benchmarkMetrics.avgSharpe}`;
    if (infoValEl) infoValEl.textContent = `> ${benchmarkMetrics.avgInfo}`;
    if (treynorValEl) treynorValEl.textContent = `> ${benchmarkMetrics.avgTreynor}`;

    const shortCount = shortlisted.length;
    const strictCount = (data.strict3Of3 || []).length;
    const catFilter = state.bestFundsCategoryFilter || 'all';

    if (shortlistCountEl) shortlistCountEl.textContent = shortCount;
    if (tabCountBestEl) tabCountBestEl.textContent = shortCount;
    if (bestTableCountEl) bestTableCountEl.textContent = shortCount;

    if (shortlistPctEl) {
      if (catFilter !== 'all') {
        const catLabel = getCategoryLabel(catFilter);
        if (strictCount > 0) {
          shortlistPctEl.textContent = `${strictCount} of ${totalCandidates} funds cleared 3/3 benchmarks • Showing all ${totalCandidates} ranked`;
        } else {
          shortlistPctEl.textContent = `Showing all ${totalCandidates} funds in ${catLabel} ranked by performance`;
        }
      } else {
        const pct = totalCandidates > 0 ? ((shortCount / totalCandidates) * 100).toFixed(1) : 0;
        shortlistPctEl.textContent = shortCount === 0
          ? `No fund out of ${totalCandidates} meets all 3 conditions`
          : `${shortCount} of ${totalCandidates} funds (${pct}%) — 3/3 conditions met`;
      }
    }

    // 2. Render Champion Showcase Card
    const champContainer = $('champion-showcase-card');
    if (champContainer) {
      if (champion) {
        const wonText = champion.periodsWon > 0
          ? `${champion.periodsWon} of 4 Rolling Horizons Won (${champion.wonHorizons.join(', ')})`
          : (champion.passesAll3
            ? `Dominant 3-Factor Risk Score (Sharpe: ${champion.sharpe || '-'})`
            : (champion.condCount > 0
              ? `Cleared ${champion.condCount} of 3 Benchmark Ratios`
              : (champion.rolling_avg !== null ? `Top Rolling Return CAGR (${champion.rolling_avg.toFixed(2)}%)` : `Top Ranked in ${getCategoryLabel(champion.category)}`)));

        const crownTag = champion.passesAll3
          ? '#1 Champion Best Mutual Fund'
          : (catFilter !== 'all' ? `#1 Top Fund in ${escapeHtml(getCategoryLabel(champion.category))}` : '#1 Top Ranked Fund');

        const rollVals = [champion.rolling_1y, champion.rolling_2y, champion.rolling_3y, champion.rolling_5y].filter(v => typeof v === 'number');
        const bestReturnVal = champion.periodsWon > 0 && rollVals.length > 0
          ? Math.max(...rollVals).toFixed(2) + '%'
          : (typeof champion.rolling_avg === 'number' ? champion.rolling_avg.toFixed(2) + '%' : '–');

        champContainer.innerHTML = `
          <div class="champion-card-inner">
            <div class="champion-badge-row">
              <div class="champion-crown-tag">
                <span>👑</span>
                <span>${crownTag}</span>
              </div>
              <div class="champion-periods-won-badge">
                <span>★</span>
                <span>${wonText}</span>
              </div>
            </div>

            <div class="champion-main-info">
              <div>
                <h3 class="champion-fund-title">${escapeHtml(champion.name)}</h3>
                <div class="champion-meta-pills">
                  <span class="champ-meta-pill" title="${escapeHtml(champion.category)}">Category: <strong>${escapeHtml(getCategoryLabel(champion.category))}</strong></span>
                  <span class="champ-meta-pill">AUM: <strong>${champion.aum !== null ? '₹' + Math.round(champion.aum).toLocaleString('en-IN') + ' Cr' : '–'}</strong></span>
                  <span class="champ-meta-pill">Inception: <strong>${escapeHtml(champion.inception_date) || '–'}</strong></span>
                </div>
              </div>
              <div>
                <button type="button" class="btn btn-primary btn-sm" id="btn-champ-add-basket">
                  <span>🧺 Add Champion to Basket</span>
                </button>
              </div>
            </div>

            <div class="champion-metrics-grid">
              <div class="champ-stat-box">
                <span class="champ-stat-label">Sharpe Ratio</span>
                <span class="champ-stat-val gold">${champion.sharpe !== null ? champion.sharpe.toFixed(2) : '-'}</span>
                <span style="font-size:0.72rem;color:#86efac;">+${champion.sharpeDelta !== null ? champion.sharpeDelta : 0} over benchmark</span>
              </div>
              <div class="champ-stat-box">
                <span class="champ-stat-label">Information Ratio</span>
                <span class="champ-stat-val">${champion.info_ratio !== null ? champion.info_ratio.toFixed(2) : '-'}</span>
                <span style="font-size:0.72rem;color:#86efac;">+${champion.infoDelta !== null ? champion.infoDelta : 0} over benchmark</span>
              </div>
              <div class="champ-stat-box">
                <span class="champ-stat-label">Treynor Ratio</span>
                <span class="champ-stat-val">${champion.treynor !== null ? champion.treynor.toFixed(2) : '-'}</span>
                <span style="font-size:0.72rem;color:#86efac;">+${champion.treynorDelta !== null ? champion.treynorDelta : 0} over benchmark</span>
              </div>
              <div class="champ-stat-box">
                <span class="champ-stat-label">Best Rolling Horizon</span>
                <span class="champ-stat-val green">${bestReturnVal}</span>
                <span style="font-size:0.72rem;color:#cbd5e1;">Top Rolling CAGR</span>
              </div>
            </div>
          </div>
        `;

        const btnChampBasket = $('btn-champ-add-basket');
        if (btnChampBasket) {
          btnChampBasket.addEventListener('click', () => {
            state.selectedFundsMap.set(champion.name, champion);
            renderBasketBar();
            showToast(`'${champion.name}' added to your Curated Basket!`, 'success');
          });
        }
      } else {
        champContainer.innerHTML = `
          <div style="text-align:center;padding:2rem;color:#cbd5e1;">
            <span style="font-size:2rem;display:block;margin-bottom:8px;">🔍</span>
            <strong>No funds currently meet all 3 conditions in this category.</strong>
            <p style="font-size:0.82rem;color:#94a3b8;margin:6px 0 0 0;">Upload your spreadsheet or switch to 'Universe Avg' / 'All Categories' to see the best funds.</p>
          </div>
        `;
      }
    }

    // 3a. Warn when only part of the shortlist had rolling data to compete with —
    // otherwise the champion reads as "best overall" when it's really "best among
    // the few funds that happened to have rolling returns loaded".
    const coverageEl = $('rolling-coverage-notice');
    if (coverageEl) {
      const missing = shortCount - withRollingCount;
      if (shortCount > 0 && missing > 0) {
        const rollingCats = [...new Set(shortlisted
          .filter(f => ['rolling_1y', 'rolling_2y', 'rolling_3y', 'rolling_5y'].some(k => typeof f[k] === 'number'))
          .map(f => getCategoryLabel(f.category)))].sort();
        coverageEl.style.display = 'flex';
        coverageEl.innerHTML = `
          <span class="coverage-icon">⚠️</span>
          <div>
            <strong>Only ${withRollingCount} of ${shortCount} shortlisted funds have rolling-return data</strong>
            ${withRollingCount > 0 ? `(${escapeHtml(rollingCats.join(', '))})` : ''} —
            the remaining ${missing} can't compete in Step 5, so the champion below is the best
            <em>among those ${withRollingCount}</em>, not across the whole shortlist.
            Upload rolling-returns sheets for the other categories via
            <strong>⚙ Change Files → + Add Sheet</strong> for a full comparison.
          </div>
        `;
      } else {
        coverageEl.style.display = 'none';
        coverageEl.innerHTML = '';
      }
    }

    // 3. Render Period Scorecard (1Y, 2Y, 3Y, 5Y)
    const scoreGrid = $('period-cards-grid');
    if (scoreGrid) {
      const keys = ['1Y', '2Y', '3Y', '5Y'];
      scoreGrid.innerHTML = keys.map(k => {
        const pw = periodWinners[k] || { horizon: `${k} Rolling`, returnVal: null, winnerName: '–', winnerCategory: '' };
        return `
          <div class="period-card">
            <div class="period-card-top">
              <span class="period-horizon-badge">${pw.horizon}</span>
              <span class="period-winner-ret">${pw.returnVal !== null ? pw.returnVal.toFixed(2) + '%' : '-'}</span>
            </div>
            <p class="period-winner-name" title="${escapeHtml(pw.winnerName)}">🏆 ${escapeHtml(pw.winnerName)}${pw.tiedCount > 1 ? ` <span class="period-tie-tag">+${pw.tiedCount - 1} tied</span>` : ''}</p>
            <span class="period-winner-cat" title="${escapeHtml(pw.winnerCategory)}">${pw.winnerCategory ? escapeHtml(getCategoryLabel(pw.winnerCategory)) : 'Mutual Fund'}</span>
          </div>
        `;
      }).join('');
    }

    // 4. Render Leaderboard Table
    const tbody = $('best-funds-tbody');
    if (tbody) {
      if (shortlisted.length === 0) {
        tbody.innerHTML = '<tr><td colspan="13" style="text-align:center;padding:2.5rem;color:var(--text-muted);">No funds meet the multi-factor conditions. Upload more data or switch the category filter.</td></tr>';
      } else {
        tbody.innerHTML = shortlisted.map((f, i) => {
          const isChamp = (i === 0);
          const rankBadge = isChamp
            ? '<span class="rank-badge rank-1" style="background:linear-gradient(135deg,#fef08a 0%,#fde047 100%);color:#854d0e;font-weight:900;border:1px solid #facc15;box-shadow:0 1px 4px rgba(245,158,11,0.25);">👑 1</span>'
            : (i === 1
              ? '<span class="rank-badge rank-2" style="background:linear-gradient(135deg,#f1f5f9 0%,#e2e8f0 100%);color:#334155;font-weight:800;border:1px solid #cbd5e1;">🥈 2</span>'
              : (i === 2
                ? '<span class="rank-badge rank-3" style="background:linear-gradient(135deg,#ffedd5 0%,#fed7aa 100%);color:#9a3412;font-weight:800;border:1px solid #fdba74;">🥉 3</span>'
                : `<span class="rank-badge" style="background:#f8fafc;color:#64748b;font-weight:700;border:1px solid #e2e8f0;">${i + 1}</span>`));

          const rowBgStyle = isChamp
            ? 'background: linear-gradient(90deg, rgba(254, 243, 199, 0.65) 0%, rgba(254, 249, 195, 0.45) 50%, rgba(254, 252, 232, 0.25) 100%); border-left: 4px solid #f59e0b;'
            : (i === 1
              ? 'background: linear-gradient(90deg, rgba(241, 245, 249, 0.8) 0%, rgba(248, 250, 252, 0.5) 100%); border-left: 4px solid #94a3b8;'
              : (i === 2
                ? 'background: linear-gradient(90deg, rgba(255, 237, 213, 0.6) 0%, rgba(255, 247, 237, 0.35) 100%); border-left: 4px solid #f97316;'
                : `background: ${i % 2 === 0 ? '#ffffff' : '#fbfcfd'}; border-left: 4px solid transparent;`));

          const renderDelta = d => {
            if (d === null || d === undefined) return '<span style="font-size:0.7rem;color:var(--text-muted);display:block;">-</span>';
            const color = d >= 0 ? 'var(--emerald)' : 'var(--rose)';
            const sign = d >= 0 ? '+' : '';
            return `<span style="font-size:0.7rem;color:${color};display:block;font-weight:700;">(${sign}${d.toFixed(2)})</span>`;
          };
          const fmtVal = v => v !== null && v !== undefined ? v.toFixed(2) + '%' : '-';

          // Check if cell won
          const won1Y = f.wonHorizons && f.wonHorizons.includes('1Y') ? 'ret-cell-won' : '';
          const won2Y = f.wonHorizons && f.wonHorizons.includes('2Y') ? 'ret-cell-won' : '';
          const won3Y = f.wonHorizons && f.wonHorizons.includes('3Y') ? 'ret-cell-won' : '';
          const won5Y = f.wonHorizons && f.wonHorizons.includes('5Y') ? 'ret-cell-won' : '';

          const periodsBadge = f.periodsWon > 0
            ? `<span class="period-win-badge ${isChamp ? 'champ-win' : ''}">★ ${f.periodsWon} Horizon${f.periodsWon > 1 ? 's' : ''}</span>`
            : '<span style="color:var(--text-muted);font-size:0.74rem;">0</span>';

          let statusBadge;
          if (isChamp) {
            statusBadge = f.passesAll3
              ? '<span class="selection-status-badge status-champion">👑 Best Mutual Fund</span>'
              : '<span class="selection-status-badge status-champion" style="background:#fef3c7;color:#92400e;border:1px solid #fde68a;">👑 Top Performer</span>';
          } else if (f.passesAll3) {
            statusBadge = '<span class="selection-status-badge status-qualified">✓ 3/3 Ratios Passed</span>';
          } else if (f.condCount === 2) {
            statusBadge = '<span class="selection-status-badge" style="background:#e0f2fe;color:#0369a1;border:1px solid #bae6fd;padding:2px 8px;border-radius:12px;font-size:0.75rem;font-weight:600;">⚡ 2/3 Passed</span>';
          } else if (f.condCount === 1) {
            statusBadge = '<span class="selection-status-badge" style="background:#f1f5f9;color:#475569;border:1px solid #e2e8f0;padding:2px 8px;border-radius:12px;font-size:0.75rem;font-weight:600;">• 1/3 Passed</span>';
          } else if (typeof f.rolling_avg === 'number' && !isNaN(f.rolling_avg)) {
            statusBadge = '<span class="selection-status-badge" style="background:#f8fafc;color:#64748b;border:1px solid #cbd5e1;padding:2px 8px;border-radius:12px;font-size:0.75rem;font-weight:600;">🔄 Rolling Performer</span>';
          } else {
            statusBadge = '<span class="selection-status-badge" style="background:#f8fafc;color:#94a3b8;border:1px solid #e2e8f0;padding:2px 8px;border-radius:12px;font-size:0.75rem;">Candidate</span>';
          }

          const sharpeCls = f.sharpe !== null ? (f.condSharpe ? 'sharpe-high' : 'sharpe-low') : 'sharpe-na';
          const infoColor = f.info_ratio !== null ? (f.condInfo ? 'var(--emerald)' : 'var(--rose)') : 'var(--text-muted)';
          const treynorColor = f.treynor !== null ? (f.condTreynor ? 'var(--emerald)' : 'var(--rose)') : 'var(--text-muted)';

          return `
            <tr style="${rowBgStyle}">
              <td class="text-center">${rankBadge}</td>
              <td class="fund-name-cell"><span class="fund-name-text" title="${escapeHtml(f.name)}">${escapeHtml(f.name)}</span></td>
              <td><span class="category-pill" title="${escapeHtml(f.category)}">${escapeHtml(getCategoryLabel(f.category))}</span></td>
              <td class="text-right num-val" style="font-weight:600;">${f.aum !== null ? '₹' + Math.round(f.aum).toLocaleString('en-IN') : '-'}</td>
              <td class="text-center">
                <span class="sharpe-badge ${sharpeCls}">${f.sharpe !== null ? f.sharpe.toFixed(2) : '-'}</span>
                ${renderDelta(f.sharpeDelta)}
              </td>
              <td class="text-center">
                <span style="font-weight:700;color:${infoColor};">${f.info_ratio !== null ? f.info_ratio.toFixed(2) : '-'}</span>
                ${renderDelta(f.infoDelta)}
              </td>
              <td class="text-center">
                <span style="font-weight:700;color:${treynorColor};">${f.treynor !== null ? f.treynor.toFixed(2) : '-'}</span>
                ${renderDelta(f.treynorDelta)}
              </td>
              <td class="text-right num-val ${won1Y}" style="color:${getReturnColor(f.rolling_1y)}">${fmtVal(f.rolling_1y)}</td>
              <td class="text-right num-val ${won2Y}" style="color:${getReturnColor(f.rolling_2y)}">${fmtVal(f.rolling_2y)}</td>
              <td class="text-right num-val ${won3Y}" style="color:${getReturnColor(f.rolling_3y)}">${fmtVal(f.rolling_3y)}</td>
              <td class="text-right num-val ${won5Y}" style="color:${getReturnColor(f.rolling_5y)}">${fmtVal(f.rolling_5y)}</td>
              <td class="text-center">${periodsBadge}</td>
              <td class="text-center">${statusBadge}</td>
            </tr>
          `;
        }).join('');
      }
    }
  }

  function initBestFundsControls() {
    // Benchmark Basis Switcher (Category Avg vs Universe Avg)
    const toggleBtns = $$('#best-benchmark-toggle .btn-toggle-pill');
    toggleBtns.forEach(btn => {
      btn.addEventListener('click', () => {
        toggleBtns.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        state.bestFundsBenchmarkBasis = btn.dataset.basis;
        renderBestFundsView();
        showToast(`Benchmark basis switched to: ${btn.textContent}`, 'info');
      });
    });

    // Category Filter
    const catSelect = $('best-funds-cat-filter');
    if (catSelect) {
      catSelect.addEventListener('change', e => {
        state.bestFundsCategoryFilter = e.target.value;
        renderBestFundsView();
      });
    }

    // Add All Shortlisted Best Funds to Curated Basket
    const btnAddAllBasket = $('btn-add-all-best-basket');
    if (btnAddAllBasket) {
      btnAddAllBasket.addEventListener('click', () => {
        if (!state.bestFundsShortlist || state.bestFundsShortlist.length === 0) {
          showToast('No best funds to add.', 'error');
          return;
        }
        state.bestFundsShortlist.forEach(f => {
          state.selectedFundsMap.set(f.name, f);
        });
        renderBasketBar();
        showToast(`Added ${state.bestFundsShortlist.length} Best Mutual Funds to your Basket!`, 'success');
      });
    }
  }

  // ==========================================
  // ADD RECORD MODAL & LIVE STAGING
  // ==========================================
  // ==========================================
  // MASTER RENDER & TOAST
  // ==========================================
  function renderAllViews() {
    renderFundsTable();
    renderCharts();
    renderBestFundsView();
  }

  function truncText(str, maxLen) {
    if (!str) return '';
    return str.length > maxLen ? str.slice(0, maxLen) + '…' : str;
  }

  function showToast(msg, type = 'info') {
    const c = $('toast-container');
    if (!c) return;

    // Cap concurrent toasts so rapid/duplicate triggers can't pile up and block the UI
    const MAX_TOASTS = 3;
    while (c.children.length >= MAX_TOASTS) {
      c.removeChild(c.firstElementChild);
    }

    const t = document.createElement('div');
    t.className = `toast toast-${type}`;
    const icon = type === 'success' ? '✅' : (type === 'error' ? '⚠️' : 'ℹ️');
    const iconSpan = document.createElement('span');
    iconSpan.textContent = icon;
    const msgSpan = document.createElement('span');
    msgSpan.textContent = msg;
    t.appendChild(iconSpan);
    t.appendChild(msgSpan);
    c.appendChild(t);

    setTimeout(() => {
      t.style.opacity = '0';
      t.style.transform = 'translateY(10px)';
      t.style.transition = 'all 0.25s ease';
      setTimeout(() => t.remove(), 250);
    }, 3500);
  }

})();
