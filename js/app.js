/* =========================================================
   BASE 注文管理 PWA  – app.js
   ========================================================= */

// ── Config ────────────────────────────────────────────────
const SPREADSHEET_ID = '1otZ-q-pp0i6biDdlRe6iVimWCQ8rEQBFMpch6Dv3axE';
const SCOPES = 'https://www.googleapis.com/auth/spreadsheets';
const STORAGE_KEY = 'base_order_pwa';

// BASE 標準エクスポート列マッピング候補
const COLUMN_ALIASES = {
  order_id:       ['注文番号', '注文ID', 'OrderID', 'Order ID'],
  order_date:     ['注文日時', '注文日', 'OrderDate', 'Order Date'],
  payment_method: ['支払い方法', '決済方法', 'PaymentMethod'],
  payment_status: ['支払い状況', '入金状況', '支払状況', 'PaymentStatus'],
  shipping_status:['発送状況', '配送状況', '出荷状況', 'ShippingStatus'],
  last_name:      ['注文者名（姓）', '注文者姓', '氏名（姓）', 'LastName'],
  first_name:     ['注文者名（名）', '注文者名', '氏名（名）', 'FirstName'],
  customer_name:  ['注文者名', 'お名前', 'CustomerName'],
  email:          ['注文者メールアドレス', 'メールアドレス', 'Email'],
  phone:          ['注文者電話番号', '電話番号', 'Phone'],
  zip:            ['お届け先郵便番号', '郵便番号', 'Zip'],
  prefecture:     ['お届け先都道府県', '都道府県', 'Prefecture'],
  city:           ['お届け先市区町村', '市区町村', 'City'],
  address1:       ['お届け先番地', '番地', 'Address1'],
  address2:       ['お届け先建物名', '建物名', 'Address2'],
  product_name:   ['商品名', '商品', 'ProductName'],
  product_id:     ['商品番号', '商品ID', 'ProductID'],
  quantity:       ['商品個数', '数量', '個数', 'Quantity'],
  unit_price:     ['商品単価', '単価', '商品金額', 'UnitPrice'],
  discount:       ['割引金額', '割引', 'Discount'],
  shipping_fee:   ['送料', 'ShippingFee'],
  total_amount:   ['合計金額', '合計', 'Total', 'TotalAmount'],
  tracking_number:['追跡番号', 'トラッキング番号', 'TrackingNumber'],
  memo:           ['メモ', '備考', 'Note', 'Notes'],
};

const STATUS_LABELS = {
  pending:    '未発送',
  processing: '対応中',
  shipped:    '発送済',
  cancelled:  'キャンセル',
};

const SHIPPING_STATUS_MAP = {
  '未発送': 'pending',
  '発送待ち': 'pending',
  '未対応': 'pending',
  '対応中': 'processing',
  '準備中': 'processing',
  '発送済み': 'shipped',
  '発送済': 'shipped',
  '配送済み': 'shipped',
  'キャンセル': 'cancelled',
  'キャンセル済み': 'cancelled',
};

// ── State ─────────────────────────────────────────────────
const state = {
  view: 'auth',
  orders: [],
  filteredOrders: [],
  selectedOrderId: null,
  filter: 'all',
  searchQuery: '',
  isLoading: false,
  isSignedIn: false,
  userInfo: null,
  headers: [],
  columnMap: {},
  sheetName: 'Sheet1',
  lastSynced: null,
  clientId: '',
  tokenClient: null,
  accessToken: null,
  showSearch: false,
};

// ── LocalStorage helpers ───────────────────────────────────
function loadStorage() {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}');
  } catch { return {}; }
}

function saveStorage(data) {
  const current = loadStorage();
  localStorage.setItem(STORAGE_KEY, JSON.stringify({ ...current, ...data }));
}

function loadCachedOrders() {
  try {
    return JSON.parse(sessionStorage.getItem('base_orders') || 'null');
  } catch { return null; }
}

function saveCachedOrders(orders) {
  try {
    sessionStorage.setItem('base_orders', JSON.stringify(orders));
  } catch {}
}

// ── Google API ─────────────────────────────────────────────
let gapiReady = false;
let gisReady = false;

function onGapiLoaded() {
  gapi.load('client', async () => {
    await gapi.client.init({
      discoveryDocs: ['https://sheets.googleapis.com/$discovery/rest?version=v4'],
    });
    gapiReady = true;
    tryAutoSignIn();
  });
}

function onGisLoaded() {
  gisReady = true;
  tryAutoSignIn();
}

function tryAutoSignIn() {
  if (!gapiReady || !gisReady) return;
  const { clientId, accessToken } = loadStorage();
  state.clientId = clientId || '';

  if (clientId && accessToken) {
    gapi.client.setToken({ access_token: accessToken });
    initTokenClient(clientId);
    verifyToken();
  } else if (clientId) {
    initTokenClient(clientId);
    navigate('auth');
  } else {
    navigate('setup');
  }
}

function initTokenClient(clientId) {
  if (!clientId) return;
  state.tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: clientId,
    scope: SCOPES,
    callback: onTokenResponse,
  });
}

function onTokenResponse(resp) {
  if (resp.error) {
    showToast('認証エラー: ' + resp.error);
    return;
  }
  state.accessToken = resp.access_token;
  saveStorage({ accessToken: resp.access_token });
  gapi.client.setToken({ access_token: resp.access_token });
  afterSignIn();
}

async function verifyToken() {
  try {
    await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'A1',
    });
    afterSignIn();
  } catch {
    gapi.client.setToken(null);
    saveStorage({ accessToken: null });
    navigate('auth');
  }
}

function signIn() {
  const clientId = document.getElementById('input-client-id-auth')?.value.trim()
    || document.getElementById('input-client-id-settings')?.value.trim()
    || state.clientId;

  if (!clientId) {
    showToast('クライアントIDを入力してください');
    navigate('setup');
    return;
  }
  state.clientId = clientId;
  saveStorage({ clientId });
  initTokenClient(clientId);

  if (!state.tokenClient) { showToast('初期化中... もう一度お試しください'); return; }
  state.tokenClient.requestAccessToken({ prompt: 'consent' });
}

function signOut() {
  const token = gapi.client.getToken();
  if (token) google.accounts.oauth2.revoke(token.access_token);
  gapi.client.setToken(null);
  state.accessToken = null;
  state.isSignedIn = false;
  state.orders = [];
  state.userInfo = null;
  saveStorage({ accessToken: null });
  navigate('auth');
}

async function afterSignIn() {
  state.isSignedIn = true;
  navigate('orders');
  await loadOrders();
}

// ── Sheets API ─────────────────────────────────────────────
async function loadOrders(forceRefresh = false) {
  if (state.isLoading) return;

  if (!forceRefresh) {
    const cached = loadCachedOrders();
    if (cached) {
      state.orders = cached;
      applyFilter();
      renderOrdersList();
      return;
    }
  }

  state.isLoading = true;
  renderOrdersList();

  try {
    // First, get spreadsheet metadata to find actual sheet names
    const meta = await gapi.client.sheets.spreadsheets.get({
      spreadsheetId: SPREADSHEET_ID,
    });
    const sheets = meta.result.sheets;
    state.sheetName = sheets[0]?.properties?.title || 'Sheet1';

    // Read all data
    const resp = await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${state.sheetName}!A:Z`,
      valueRenderOption: 'UNFORMATTED_VALUE',
    });

    const rows = resp.result.values || [];
    if (rows.length < 2) {
      state.orders = [];
      state.isLoading = false;
      applyFilter();
      renderOrdersList();
      return;
    }

    state.headers = rows[0].map(String);
    state.columnMap = buildColumnMap(state.headers);

    const dataRows = rows.slice(1);
    const rawOrders = dataRows.map((row, i) => rowToOrder(row, i + 2)); // row index 1-based +1 for header

    // Group by order_id (multiple rows per order for multiple products)
    state.orders = groupOrders(rawOrders);
    state.lastSynced = new Date();
    saveCachedOrders(state.orders);

  } catch (err) {
    console.error('loadOrders error:', err);
    if (err.status === 401 || err.status === 403) {
      showToast('認証が切れました。再ログインしてください');
      gapi.client.setToken(null);
      saveStorage({ accessToken: null });
      navigate('auth');
    } else {
      showToast('データ取得に失敗しました');
    }
  }

  state.isLoading = false;
  applyFilter();
  renderOrdersList();
}

function buildColumnMap(headers) {
  const map = {};
  headers.forEach((h, i) => {
    const header = String(h).trim();
    for (const [field, aliases] of Object.entries(COLUMN_ALIASES)) {
      if (!map[field] && aliases.some(a => header.includes(a) || a.includes(header))) {
        map[field] = i;
      }
    }
  });
  return map;
}

function col(row, field) {
  const idx = state.columnMap[field];
  if (idx === undefined) return '';
  return row[idx] !== undefined ? String(row[idx]) : '';
}

function rowToOrder(row, rowIndex) {
  const shippingStatusRaw = col(row, 'shipping_status');
  const statusKey = SHIPPING_STATUS_MAP[shippingStatusRaw] || classifyStatus(shippingStatusRaw);

  const lastName = col(row, 'last_name');
  const firstName = col(row, 'first_name');
  const customerName = col(row, 'customer_name') || (lastName + firstName) || '(名前なし)';

  return {
    rowIndex,
    order_id: col(row, 'order_id') || `ROW${rowIndex}`,
    order_date: formatDate(col(row, 'order_date')),
    payment_method: col(row, 'payment_method'),
    payment_status: col(row, 'payment_status'),
    shipping_status_raw: shippingStatusRaw,
    status: statusKey,
    customer_name: customerName,
    email: col(row, 'email'),
    phone: col(row, 'phone'),
    zip: col(row, 'zip'),
    prefecture: col(row, 'prefecture'),
    city: col(row, 'city'),
    address1: col(row, 'address1'),
    address2: col(row, 'address2'),
    products: [{
      name: col(row, 'product_name'),
      id: col(row, 'product_id'),
      quantity: col(row, 'quantity'),
      unit_price: col(row, 'unit_price'),
    }].filter(p => p.name),
    discount: col(row, 'discount'),
    shipping_fee: col(row, 'shipping_fee'),
    total_amount: col(row, 'total_amount'),
    tracking_number: col(row, 'tracking_number'),
    memo: col(row, 'memo'),
  };
}

function groupOrders(rawOrders) {
  const map = new Map();
  for (const o of rawOrders) {
    if (map.has(o.order_id)) {
      const existing = map.get(o.order_id);
      existing.products.push(...o.products);
      existing._rows.push(o.rowIndex);
    } else {
      o._rows = [o.rowIndex];
      map.set(o.order_id, o);
    }
  }
  return Array.from(map.values());
}

function classifyStatus(raw) {
  const r = (raw || '').trim();
  if (!r || r === '未発送' || r === '発送待ち') return 'pending';
  if (r.includes('発送') || r.includes('配送')) return 'shipped';
  if (r.includes('キャンセル')) return 'cancelled';
  return 'processing';
}

function formatDate(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return String(dateStr);
  return `${d.getFullYear()}/${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getDate()).padStart(2,'0')} ${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`;
}

function formatAmount(val) {
  if (!val) return '';
  const n = parseFloat(String(val).replace(/[^\d.]/g, ''));
  if (isNaN(n)) return val;
  return '¥' + n.toLocaleString('ja-JP');
}

async function updateShippingStatus(order, newStatusLabel, trackingNumber) {
  const col_idx = state.columnMap['shipping_status'];
  const tracking_idx = state.columnMap['tracking_number'];
  if (col_idx === undefined) { showToast('発送状況の列が見つかりません'); return; }

  state.isLoading = true;
  showToast('更新中...');

  try {
    const updates = [];
    for (const rowIndex of order._rows) {
      const colLetter = indexToColumnLetter(col_idx);
      updates.push({
        range: `${state.sheetName}!${colLetter}${rowIndex}`,
        values: [[newStatusLabel]],
      });
      if (tracking_idx !== undefined && trackingNumber !== undefined) {
        const tCol = indexToColumnLetter(tracking_idx);
        updates.push({
          range: `${state.sheetName}!${tCol}${rowIndex}`,
          values: [[trackingNumber]],
        });
      }
    }

    await gapi.client.sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      resource: {
        valueInputOption: 'RAW',
        data: updates,
      },
    });

    // Update in local state
    order.shipping_status_raw = newStatusLabel;
    order.status = SHIPPING_STATUS_MAP[newStatusLabel] || classifyStatus(newStatusLabel);
    if (trackingNumber !== undefined) order.tracking_number = trackingNumber;
    saveCachedOrders(state.orders);
    applyFilter();

    showToast('更新しました');
    renderDetail(order);
  } catch (err) {
    console.error('updateShippingStatus error:', err);
    showToast('更新に失敗しました');
  }
  state.isLoading = false;
}

function indexToColumnLetter(idx) {
  let col = '';
  let n = idx + 1;
  while (n > 0) {
    col = String.fromCharCode(((n - 1) % 26) + 65) + col;
    n = Math.floor((n - 1) / 26);
  }
  return col;
}

// ── Filtering ──────────────────────────────────────────────
function applyFilter() {
  let list = state.orders;

  if (state.filter !== 'all') {
    list = list.filter(o => o.status === state.filter);
  }

  if (state.searchQuery) {
    const q = state.searchQuery.toLowerCase();
    list = list.filter(o =>
      o.order_id.toLowerCase().includes(q) ||
      o.customer_name.toLowerCase().includes(q) ||
      o.email.toLowerCase().includes(q) ||
      o.products.some(p => p.name.toLowerCase().includes(q))
    );
  }

  state.filteredOrders = list;
}

// ── Navigation ─────────────────────────────────────────────
function navigate(view, data = {}) {
  // Hide all views
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));

  const viewEl = document.getElementById(`view-${view}`);
  if (viewEl) viewEl.classList.add('active');

  // Bottom nav visibility
  const bottomNav = document.getElementById('bottom-nav');
  const showNav = ['orders', 'settings'].includes(view);
  bottomNav.classList.toggle('hidden', !showNav);

  // Update active nav item
  document.querySelectorAll('.nav-item').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.view === view);
  });

  state.view = view;

  // Render appropriate view
  if (view === 'orders')  renderOrdersList();
  if (view === 'detail' && data.order) renderDetail(data.order);
  if (view === 'settings') renderSettings();
  if (view === 'setup')   renderSetup();
  if (view === 'auth')    renderAuth();
}

// ── Render: Auth ───────────────────────────────────────────
function renderAuth() {
  const saved = loadStorage();
  const hasClientId = !!saved.clientId;

  document.getElementById('view-auth').innerHTML = `
    <div class="auth-container">
      <div class="auth-logo">
        <img src="icons/icon.svg" alt="BASE注文管理" width="80">
        <h1>BASE 注文管理</h1>
        <p>Googleアカウントでサインインして<br>注文を管理しましょう</p>
      </div>
      ${!hasClientId ? `
        <div style="width:100%">
          <label style="font-size:13px;font-weight:600;display:block;margin-bottom:6px;color:#666">
            OAuthクライアントID
          </label>
          <input id="input-client-id-auth" class="text-input" type="text"
            placeholder="xxxx.apps.googleusercontent.com"
            value="${escHtml(saved.clientId || '')}">
          <p class="text-sm text-secondary mt-8">
            <a href="#" id="link-setup-help">設定方法を確認する</a>
          </p>
        </div>
      ` : ''}
      <button id="btn-signin" class="btn btn-primary btn-large">
        Googleでサインイン
      </button>
      ${hasClientId ? `<p class="text-sm text-secondary">クライアントID: ${escHtml(saved.clientId.slice(0,20))}...</p>` : ''}
      <button class="btn btn-secondary mt-8" id="btn-goto-setup" style="width:100%">
        初期設定 / クライアントID変更
      </button>
    </div>
  `;

  document.getElementById('btn-signin').addEventListener('click', signIn);
  document.getElementById('btn-goto-setup')?.addEventListener('click', () => navigate('setup'));
  document.getElementById('link-setup-help')?.addEventListener('click', (e) => { e.preventDefault(); navigate('setup'); });
}

// ── Render: Setup ──────────────────────────────────────────
function renderSetup() {
  const saved = loadStorage();
  document.getElementById('view-setup').innerHTML = `
    <div class="app-header" style="position:sticky;top:0">
      <h1>初期設定</h1>
      ${saved.clientId ? `<button class="icon-btn" id="btn-skip-setup">×</button>` : ''}
    </div>
    <div class="scroll-content">
      <div style="padding:12px">
        <div class="detail-section">
          <div class="detail-section-title">OAuthクライアントIDの取得方法</div>
          <div class="detail-row" style="flex-direction:column;gap:8px">
            <div style="display:flex;gap:8px;align-items:flex-start">
              <span class="step-num">1</span>
              <p style="font-size:13px;color:#666;line-height:1.6">
                <a href="https://console.cloud.google.com/" target="_blank" style="color:#e63b2e">
                  Google Cloud Console
                </a> で新規プロジェクトを作成
              </p>
            </div>
            <div style="display:flex;gap:8px;align-items:flex-start">
              <span class="step-num">2</span>
              <p style="font-size:13px;color:#666;line-height:1.6">
                「APIとサービス」→「ライブラリ」で<br>
                <strong>Google Sheets API</strong> を有効化
              </p>
            </div>
            <div style="display:flex;gap:8px;align-items:flex-start">
              <span class="step-num">3</span>
              <p style="font-size:13px;color:#666;line-height:1.6">
                「認証情報」→「認証情報を作成」→<br>
                「OAuthクライアントID」→<strong>ウェブアプリケーション</strong>を選択
              </p>
            </div>
            <div style="display:flex;gap:8px;align-items:flex-start">
              <span class="step-num">4</span>
              <p style="font-size:13px;color:#666;line-height:1.6">
                「承認済みのJavaScriptオリジン」に<br>
                このアプリのURL（GitHub PagesのURL等）を追加
              </p>
            </div>
            <div style="display:flex;gap:8px;align-items:flex-start">
              <span class="step-num">5</span>
              <p style="font-size:13px;color:#666;line-height:1.6">
                発行されたクライアントIDを下に入力
              </p>
            </div>
          </div>
        </div>

        <div class="detail-section" style="margin-top:8px">
          <div class="detail-section-title">クライアントIDを入力</div>
          <div class="detail-row" style="flex-direction:column;gap:8px">
            <input id="input-client-id-settings" class="text-input" type="text"
              placeholder="xxxx.apps.googleusercontent.com"
              value="${escHtml(saved.clientId || '')}">
            <button id="btn-save-and-signin" class="btn btn-primary btn-full">
              保存してサインイン
            </button>
          </div>
        </div>

        <div class="detail-section" style="margin-top:8px">
          <div class="detail-section-title">スプレッドシートID</div>
          <div class="detail-row">
            <span class="detail-label">ID</span>
            <span class="detail-value text-sm">${SPREADSHEET_ID}</span>
          </div>
        </div>
      </div>
    </div>
  `;

  document.getElementById('btn-save-and-signin')?.addEventListener('click', () => {
    const cid = document.getElementById('input-client-id-settings')?.value.trim();
    if (!cid) { showToast('クライアントIDを入力してください'); return; }
    state.clientId = cid;
    saveStorage({ clientId: cid });
    signIn();
  });

  document.getElementById('btn-skip-setup')?.addEventListener('click', () => navigate('auth'));
}

// ── Render: Orders List ────────────────────────────────────
function renderOrdersList() {
  const list = document.getElementById('orders-list');
  if (!list) return;

  // Update filter badge counts
  const counts = { all: state.orders.length, pending: 0, processing: 0, shipped: 0, cancelled: 0 };
  state.orders.forEach(o => { if (counts[o.status] !== undefined) counts[o.status]++; });

  document.querySelectorAll('.filter-tabs .tab').forEach(tab => {
    const f = tab.dataset.filter;
    const c = counts[f] || 0;
    const label = tab.dataset.label || tab.textContent.replace(/\d+$/, '').trim();
    tab.dataset.label = label;
    tab.innerHTML = `${label}${c > 0 && f !== 'all' ? ` <span class="badge">${c}</span>` : ''}`;
    tab.classList.toggle('active', f === state.filter);
  });

  if (state.isLoading && !state.orders.length) {
    list.innerHTML = `<div class="state-box"><div class="spinner"></div><p>読み込み中...</p></div>`;
    return;
  }

  if (!state.orders.length) {
    list.innerHTML = `
      <div class="state-box">
        <div class="state-icon">📦</div>
        <p>注文データがありません<br>スプレッドシートにデータを追加してください</p>
        <button class="btn btn-secondary state-action" id="btn-reload">再読み込み</button>
      </div>`;
    document.getElementById('btn-reload')?.addEventListener('click', () => loadOrders(true));
    return;
  }

  if (!state.filteredOrders.length) {
    list.innerHTML = `
      <div class="state-box">
        <div class="state-icon">🔍</div>
        <p>該当する注文が見つかりません</p>
      </div>`;
    return;
  }

  list.innerHTML = state.filteredOrders.map(order => orderCardHTML(order)).join('');

  list.querySelectorAll('.order-card').forEach(card => {
    card.addEventListener('click', () => {
      const order = state.filteredOrders.find(o => o.order_id === card.dataset.id);
      if (order) navigate('detail', { order });
    });
  });
}

function orderCardHTML(order) {
  const productSummary = order.products.length
    ? order.products.map(p => `${p.name}${p.quantity ? ` ×${p.quantity}` : ''}`).join('、')
    : '(商品情報なし)';

  return `
    <div class="order-card" data-id="${escHtml(order.order_id)}">
      <div class="order-card-header">
        <span class="order-number">${escHtml(order.order_id)}</span>
        <span class="order-date">${escHtml(order.order_date)}</span>
      </div>
      <div class="order-name">${escHtml(order.customer_name)}</div>
      <div class="order-product">${escHtml(productSummary)}</div>
      <div class="order-card-footer">
        <span class="order-amount">${escHtml(formatAmount(order.total_amount))}</span>
        <span class="status-badge status-${order.status}">${STATUS_LABELS[order.status] || order.shipping_status_raw}</span>
      </div>
    </div>`;
}

// ── Render: Order Detail ───────────────────────────────────
function renderDetail(order) {
  state.selectedOrderId = order.order_id;
  const el = document.getElementById('view-detail');

  const address = [order.zip ? `〒${order.zip}` : '', order.prefecture, order.city, order.address1, order.address2]
    .filter(Boolean).join(' ');

  const productsHTML = order.products.length
    ? order.products.map(p => `
        <div class="detail-row">
          <span class="detail-label">商品</span>
          <span class="detail-value">${escHtml(p.name)}${p.id ? ` <span class="text-sm text-secondary">(${escHtml(p.id)})</span>` : ''}<br>
            <span class="text-sm text-secondary">${p.quantity ? `${escHtml(p.quantity)}個` : ''} ${p.unit_price ? formatAmount(p.unit_price) : ''}</span>
          </span>
        </div>`).join('')
    : `<div class="detail-row"><span class="detail-label">商品</span><span class="detail-value text-secondary">なし</span></div>`;

  el.querySelector('.detail-scroll').innerHTML = `
    <div class="detail-section">
      <div class="detail-section-title">注文情報</div>
      <div class="detail-row">
        <span class="detail-label">注文番号</span>
        <span class="detail-value" style="font-weight:700;color:#e63b2e">${escHtml(order.order_id)}</span>
      </div>
      <div class="detail-row">
        <span class="detail-label">注文日時</span>
        <span class="detail-value">${escHtml(order.order_date)}</span>
      </div>
      <div class="detail-row">
        <span class="detail-label">発送状況</span>
        <span class="detail-value"><span class="status-badge status-${order.status}">${STATUS_LABELS[order.status] || escHtml(order.shipping_status_raw)}</span></span>
      </div>
      <div class="detail-row">
        <span class="detail-label">支払い</span>
        <span class="detail-value">${escHtml(order.payment_method)} <span class="status-badge ${order.payment_status?.includes('済') ? 'status-shipped' : 'status-pending'}">${escHtml(order.payment_status)}</span></span>
      </div>
    </div>

    <div class="detail-section">
      <div class="detail-section-title">お客様情報</div>
      <div class="detail-row">
        <span class="detail-label">お名前</span>
        <span class="detail-value">${escHtml(order.customer_name)}</span>
      </div>
      ${order.email ? `<div class="detail-row"><span class="detail-label">メール</span><a class="detail-value" href="mailto:${escHtml(order.email)}" style="color:#e63b2e">${escHtml(order.email)}</a></div>` : ''}
      ${order.phone ? `<div class="detail-row"><span class="detail-label">電話</span><a class="detail-value" href="tel:${escHtml(order.phone)}" style="color:#e63b2e">${escHtml(order.phone)}</a></div>` : ''}
      ${address ? `<div class="detail-row"><span class="detail-label">住所</span><span class="detail-value">${escHtml(address)}</span></div>` : ''}
    </div>

    <div class="detail-section">
      <div class="detail-section-title">注文商品</div>
      ${productsHTML}
      ${order.discount ? `<div class="detail-row"><span class="detail-label">割引</span><span class="detail-value">-${escHtml(formatAmount(order.discount))}</span></div>` : ''}
      ${order.shipping_fee ? `<div class="detail-row"><span class="detail-label">送料</span><span class="detail-value">${escHtml(formatAmount(order.shipping_fee))}</span></div>` : ''}
      <div class="detail-row">
        <span class="detail-label">合計金額</span>
        <span class="detail-value large">${escHtml(formatAmount(order.total_amount))}</span>
      </div>
    </div>

    <div class="detail-section">
      <div class="detail-section-title">発送管理</div>
      <div class="detail-row">
        <span class="detail-label">追跡番号</span>
        <span class="detail-value">${escHtml(order.tracking_number) || '<span class="text-secondary">未入力</span>'}</span>
      </div>
      ${order.memo ? `<div class="detail-row"><span class="detail-label">メモ</span><span class="detail-value">${escHtml(order.memo)}</span></div>` : ''}
      <div class="detail-row" style="flex-direction:column;gap:8px">
        <label style="font-size:12px;color:#999;font-weight:600">追跡番号を更新</label>
        <input id="input-tracking" class="text-input" type="text"
          placeholder="追跡番号を入力"
          value="${escHtml(order.tracking_number)}">
      </div>
      <div class="detail-row" style="flex-direction:column;gap:8px">
        <label style="font-size:12px;color:#999;font-weight:600">発送状況を変更</label>
        <div style="display:flex;gap:8px">
          <select id="select-status" class="select-input" style="flex:1">
            <option value="未発送" ${order.shipping_status_raw === '未発送' ? 'selected' : ''}>未発送</option>
            <option value="対応中" ${order.shipping_status_raw === '対応中' ? 'selected' : ''}>対応中</option>
            <option value="発送済み" ${order.shipping_status_raw === '発送済み' ? 'selected' : ''}>発送済み</option>
            <option value="キャンセル" ${order.shipping_status_raw === 'キャンセル' ? 'selected' : ''}>キャンセル</option>
          </select>
          <button id="btn-update-status" class="btn btn-primary" style="height:44px;padding:0 16px">
            更新
          </button>
        </div>
      </div>
    </div>
  `;

  document.getElementById('btn-update-status')?.addEventListener('click', () => {
    const newStatus = document.getElementById('select-status')?.value;
    const tracking = document.getElementById('input-tracking')?.value.trim();
    updateShippingStatus(order, newStatus, tracking);
  });
}

// ── Render: Settings ──────────────────────────────────────
function renderSettings() {
  const saved = loadStorage();
  const el = document.getElementById('view-settings');
  const syncText = state.lastSynced
    ? state.lastSynced.toLocaleTimeString('ja-JP')
    : '未同期';

  el.querySelector('.settings-scroll').innerHTML = `
    <div class="detail-section">
      <div class="detail-section-title">アカウント</div>
      <div class="settings-account-info">
        <div class="settings-account-avatar">👤</div>
        <div>
          <div class="settings-account-name">${state.isSignedIn ? 'サインイン済み' : '未サインイン'}</div>
          <div class="settings-account-email">最終同期: ${syncText}</div>
        </div>
      </div>
      <div class="detail-row">
        <button id="btn-signout" class="btn btn-outline btn-full">サインアウト</button>
      </div>
    </div>

    <div class="detail-section">
      <div class="detail-section-title">接続設定</div>
      <div class="settings-row">
        <label>OAuthクライアントID</label>
        <input id="input-client-id-settings" class="text-input" type="text"
          placeholder="xxxx.apps.googleusercontent.com"
          value="${escHtml(saved.clientId || '')}">
        <button id="btn-update-client-id" class="btn btn-secondary btn-full mt-8">変更して再サインイン</button>
      </div>
      <div class="settings-row">
        <label>スプレッドシートID</label>
        <p>${SPREADSHEET_ID}</p>
      </div>
      <div class="settings-row">
        <label>シート名</label>
        <p>${escHtml(state.sheetName)}</p>
      </div>
    </div>

    <div class="detail-section">
      <div class="detail-section-title">データ</div>
      <div class="detail-row">
        <span class="detail-label">注文数</span>
        <span class="detail-value">${state.orders.length} 件</span>
      </div>
      <div class="detail-row">
        <button id="btn-refresh-settings" class="btn btn-secondary btn-full">データを再読み込み</button>
      </div>
    </div>

    <div class="detail-section">
      <div class="detail-section-title">初期設定ガイド</div>
      <div class="detail-row">
        <button id="btn-goto-setup-settings" class="btn btn-secondary btn-full">設定ガイドを表示</button>
      </div>
    </div>
  `;

  document.getElementById('btn-signout')?.addEventListener('click', signOut);
  document.getElementById('btn-refresh-settings')?.addEventListener('click', () => {
    sessionStorage.removeItem('base_orders');
    loadOrders(true);
    showToast('再読み込み中...');
    navigate('orders');
  });
  document.getElementById('btn-update-client-id')?.addEventListener('click', () => {
    const cid = document.getElementById('input-client-id-settings')?.value.trim();
    if (!cid) { showToast('クライアントIDを入力してください'); return; }
    state.clientId = cid;
    saveStorage({ clientId: cid, accessToken: null });
    gapi.client.setToken(null);
    initTokenClient(cid);
    state.tokenClient?.requestAccessToken({ prompt: 'consent' });
  });
  document.getElementById('btn-goto-setup-settings')?.addEventListener('click', () => navigate('setup'));
}

// ── Toast ─────────────────────────────────────────────────
let toastTimer;
function showToast(msg, duration = 2500) {
  const toast = document.getElementById('toast');
  toast.textContent = msg;
  toast.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.remove('show'), duration);
}

// ── Escape HTML ───────────────────────────────────────────
function escHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ── Event Wiring ──────────────────────────────────────────
function wireEvents() {
  // Bottom nav
  document.querySelectorAll('.nav-item').forEach(btn => {
    btn.addEventListener('click', () => navigate(btn.dataset.view));
  });

  // Back button
  document.getElementById('btn-back')?.addEventListener('click', () => navigate('orders'));

  // Refresh button
  document.getElementById('btn-refresh')?.addEventListener('click', () => {
    sessionStorage.removeItem('base_orders');
    loadOrders(true);
  });

  // Search toggle
  document.getElementById('btn-search-toggle')?.addEventListener('click', () => {
    state.showSearch = !state.showSearch;
    const bar = document.getElementById('search-bar');
    bar.classList.toggle('hidden', !state.showSearch);
    if (state.showSearch) {
      document.getElementById('search-input')?.focus();
    } else {
      state.searchQuery = '';
      if (document.getElementById('search-input')) document.getElementById('search-input').value = '';
      applyFilter();
      renderOrdersList();
    }
  });

  // Search input
  let searchTimer;
  document.getElementById('search-input')?.addEventListener('input', (e) => {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => {
      state.searchQuery = e.target.value.trim();
      applyFilter();
      renderOrdersList();
    }, 200);
  });

  // Filter tabs
  document.querySelectorAll('.filter-tabs .tab').forEach(tab => {
    tab.addEventListener('click', () => {
      state.filter = tab.dataset.filter;
      applyFilter();
      renderOrdersList();
    });
  });
}

// ── Init ──────────────────────────────────────────────────
function init() {
  // Register service worker
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./sw.js').catch(() => {});
  }

  wireEvents();

  // Show setup if no clientId; otherwise wait for API libs
  const { clientId } = loadStorage();
  if (!clientId) {
    navigate('setup');
  } else {
    navigate('auth');
  }

  // Poll for API readiness (libs load async)
  const poll = setInterval(() => {
    if (window.gapi && window.google?.accounts?.oauth2) {
      clearInterval(poll);
      onGapiLoaded();
      onGisLoaded();
    }
  }, 100);
}

// Expose callbacks for script onload attributes
window.gapiOnLoad = onGapiLoaded;
window.gisOnLoad  = onGisLoaded;

document.addEventListener('DOMContentLoaded', init);
