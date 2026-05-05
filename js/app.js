/* =========================================================
   BASE 注文管理 PWA  – app.js  (v2: 実際のシート構造対応)
   ヘッダー2行、データ3行目から
   ========================================================= */

// ── スプレッドシート列インデックス（0始まり）────────────
const C = {
  order_date:       0,   // A: 注文日
  seq_no:           1,   // B: No.
  warehouse_no:     2,   // C: 倉庫No.
  status:           4,   // E: ステータス
  base_order_id:    5,   // F: BASE注文番号
  shop_id:          7,   // H: ショップID
  shop_name:        8,   // I: ショップ名
  mail_orei:        9,   // J: メール送信(御礼)
  mail_shiire:     10,   // K: メール送信(仕入れ)
  mail_delivery:   11,   // L: メール送信(出荷通知)
  mail_review:     12,   // M: メール送信(評価依頼)
  recipient_name:  13,   // N: 送付先名
  zip:             14,   // O: 〒
  address_street:  15,   // P: 住所2
  phone:           16,   // Q: 電話番号
  addressee:       17,   // R: 宛名
  zip1:            18,   // S: 〒前半
  zip2:            19,   // T: 〒後半
  prefecture:      20,   // U: 都道府県
  city:            21,   // V: 市区町村
  address_detail:  22,   // W: 住所
  phone2:          23,   // X: 電話番号
  email:           24,   // Y: メールアドレス
  image_url:       25,   // Z: 画像リンク
  product_page:    26,   // AA: 商品ページ
  product_title:   27,   // AB: タイトル
  variation:       28,   // AC: バリエーション内容
  asin_parent:     29,   // AD: ASIN(親)
  asin_child:      30,   // AE: ASIN(子)
  link_parent:     31,   // AF: リンク(親)
  link_child:      32,   // AG: リンク(子)
  purchase_order:  33,   // AH: 仕入注文番号
  amazon_page:     34,   // AI: Ama注文ページ
  arrival_date1:   35,   // AJ: 到着予定日
  arrival_date2:   36,   // AK: 到着予定日
  arrival_date:    37,   // AL: 到着日
  unit_price:      41,   // AP: 単価
  subtotal:        42,   // AQ: 小計
  option_fee:      43,   // AR: オプション
  total_amount:    44,   // AS: 合計
  commission:      45,   // AT: 手数料
  quantity:        48,   // AW: 数量
  purchase_price:  49,   // AX: 仕入金額
  base_link:       50,   // AY: BASE
  amazon_link:     51,   // AZ: AmazonJP
  domestic_ship:   52,   // BA: 国内送料
  tracking_number: 53,   // BB: 追跡番号
  cost:            54,   // BC: コスト
  profit:          55,   // BD: 利益
  profit_rate:     56,   // BE: 利益率
};

const DATA_START_ROW = 3; // ヘッダー2行 → データは3行目から

// ── ステータス定義 ─────────────────────────────────────
const STATUS_DEF = {
  '完了':         { css: 'completed',  label: '完了' },
  '評価依頼待ち': { css: 'review',     label: '評価依頼待ち' },
  '仕入済':       { css: 'sourced',    label: '仕入済' },
  '仕入れ済':     { css: 'sourced',    label: '仕入済' },
  'キャンセル':   { css: 'cancelled',  label: 'キャンセル' },
};

const STATUS_OPTIONS = ['仕入済', '評価依頼待ち', '完了', 'キャンセル'];

const STORAGE_KEY = 'base_order_pwa';
const SPREADSHEET_ID = '1otZ-q-pp0i6biDdlRe6iVimWCQ8rEQBFMpch6Dv3axE';
const SCOPES = 'https://www.googleapis.com/auth/spreadsheets';

// ── State ──────────────────────────────────────────────
const state = {
  view: 'auth',
  orders: [],
  filtered: [],
  selectedOrder: null,
  filter: 'all',
  searchQuery: '',
  isLoading: false,
  isSignedIn: false,
  sheetName: 'Sheet1',
  lastSynced: null,
  clientId: '',
  tokenClient: null,
  showSearch: false,
};

// ── Storage ────────────────────────────────────────────
function loadStorage() {
  try { return JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}'); } catch { return {}; }
}
function saveStorage(data) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify({ ...loadStorage(), ...data }));
}
function loadCached() {
  try { return JSON.parse(sessionStorage.getItem('base_orders') || 'null'); } catch { return null; }
}
function saveCache(orders) {
  try { sessionStorage.setItem('base_orders', JSON.stringify(orders)); } catch {}
}

// ── Google API ─────────────────────────────────────────
let gapiReady = false, gisReady = false;

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
    callback: (resp) => {
      if (resp.error) { showToast('認証エラー: ' + resp.error); return; }
      state.accessToken = resp.access_token;
      saveStorage({ accessToken: resp.access_token });
      gapi.client.setToken({ access_token: resp.access_token });
      afterSignIn();
    },
  });
}

async function verifyToken() {
  try {
    await gapi.client.sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    afterSignIn();
  } catch {
    gapi.client.setToken(null);
    saveStorage({ accessToken: null });
    navigate('auth');
  }
}

function signIn() {
  const inputEl = document.getElementById('input-client-id-auth') || document.getElementById('input-client-id-settings');
  const cid = (inputEl?.value.trim()) || state.clientId;
  if (!cid) { showToast('クライアントIDを入力してください'); navigate('setup'); return; }
  state.clientId = cid;
  saveStorage({ clientId: cid });
  initTokenClient(cid);
  if (!state.tokenClient) { showToast('初期化中... 再度お試しください'); return; }
  state.tokenClient.requestAccessToken({ prompt: 'consent' });
}

function signOut() {
  const token = gapi.client.getToken();
  if (token) google.accounts.oauth2.revoke(token.access_token);
  gapi.client.setToken(null);
  state.isSignedIn = false;
  state.orders = [];
  saveStorage({ accessToken: null });
  navigate('auth');
}

async function afterSignIn() {
  state.isSignedIn = true;
  navigate('orders');
  await loadOrders();
}

// ── データ読み込み ─────────────────────────────────────
async function loadOrders(forceRefresh = false) {
  if (state.isLoading) return;

  if (!forceRefresh) {
    const cached = loadCached();
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
    // シート名を取得
    const meta = await gapi.client.sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    state.sheetName = meta.result.sheets[0]?.properties?.title || 'Sheet1';

    // データ部分のみ読み込み（3行目から）
    const resp = await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${state.sheetName}!A${DATA_START_ROW}:BG`,
      valueRenderOption: 'UNFORMATTED_VALUE',
    });

    const rows = resp.result.values || [];
    state.orders = rows
      .map((row, i) => parseRow(row, i + DATA_START_ROW))
      .filter(o => o.base_order_id || o.seq_no); // 空行除外

    state.lastSynced = new Date();
    saveCache(state.orders);

  } catch (err) {
    console.error(err);
    if (err.status === 401 || err.status === 403) {
      showToast('認証が切れました。再サインインしてください');
      gapi.client.setToken(null);
      saveStorage({ accessToken: null });
      navigate('auth');
    } else {
      showToast('読み込みに失敗しました');
    }
  }

  state.isLoading = false;
  applyFilter();
  renderOrdersList();
}

function get(row, idx) {
  return (idx < row.length && row[idx] !== undefined && row[idx] !== null)
    ? String(row[idx]).trim() : '';
}

function parseRow(row, spreadsheetRowIndex) {
  const title = get(row, C.product_title);
  const shortTitle = title.length > 50 ? title.slice(0, 50) + '…' : title;

  return {
    _row: spreadsheetRowIndex, // スプレッドシート上の行番号（更新用）
    order_date:      get(row, C.order_date),
    seq_no:          get(row, C.seq_no),
    status_raw:      get(row, C.status),
    base_order_id:   get(row, C.base_order_id),
    shop_id:         get(row, C.shop_id),
    shop_name:       get(row, C.shop_name),
    mail_orei:       get(row, C.mail_orei),
    mail_shiire:     get(row, C.mail_shiire),
    mail_delivery:   get(row, C.mail_delivery),
    mail_review:     get(row, C.mail_review),
    recipient_name:  get(row, C.recipient_name),
    zip:             get(row, C.zip),
    address_street:  get(row, C.address_street),
    phone:           get(row, C.phone),
    prefecture:      get(row, C.prefecture),
    city:            get(row, C.city),
    address_detail:  get(row, C.address_detail),
    email:           get(row, C.email),
    image_url:       get(row, C.image_url),
    product_page:    get(row, C.product_page),
    product_title:   title,
    product_title_short: shortTitle,
    variation:       get(row, C.variation),
    asin_parent:     get(row, C.asin_parent),
    asin_child:      get(row, C.asin_child),
    link_parent:     get(row, C.link_parent),
    link_child:      get(row, C.link_child),
    purchase_order:  get(row, C.purchase_order),
    amazon_page:     get(row, C.amazon_page),
    arrival_date1:   get(row, C.arrival_date1),
    arrival_date2:   get(row, C.arrival_date2),
    arrival_date:    get(row, C.arrival_date),
    unit_price:      get(row, C.unit_price),
    total_amount:    get(row, C.total_amount),
    commission:      get(row, C.commission),
    quantity:        get(row, C.quantity),
    purchase_price:  get(row, C.purchase_price),
    base_link:       get(row, C.base_link),
    amazon_link:     get(row, C.amazon_link),
    tracking_number: get(row, C.tracking_number),
    cost:            get(row, C.cost),
    profit:          get(row, C.profit),
    profit_rate:     get(row, C.profit_rate),
  };
}

// ── フィルター ─────────────────────────────────────────
function applyFilter() {
  let list = state.orders;
  if (state.filter !== 'all') {
    list = list.filter(o => o.status_raw === state.filter);
  }
  if (state.searchQuery) {
    const q = state.searchQuery.toLowerCase();
    list = list.filter(o =>
      o.seq_no.includes(q) ||
      o.base_order_id.toLowerCase().includes(q) ||
      o.recipient_name.toLowerCase().includes(q) ||
      o.shop_name.toLowerCase().includes(q) ||
      o.product_title.toLowerCase().includes(q)
    );
  }
  state.filtered = list;
}

// ── スプレッドシート更新 ───────────────────────────────
async function updateOrder(order, newStatus, trackingNumber) {
  state.isLoading = true;
  showToast('更新中...', 60000);

  try {
    const updates = [];
    const colStatus   = indexToCol(C.status);
    const colTracking = indexToCol(C.tracking_number);
    const row = order._row;

    updates.push({ range: `${state.sheetName}!${colStatus}${row}`,   values: [[newStatus]] });
    if (trackingNumber !== undefined) {
      updates.push({ range: `${state.sheetName}!${colTracking}${row}`, values: [[trackingNumber]] });
    }

    await gapi.client.sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      resource: { valueInputOption: 'RAW', data: updates },
    });

    order.status_raw = newStatus;
    if (trackingNumber !== undefined) order.tracking_number = trackingNumber;
    saveCache(state.orders);
    applyFilter();
    showToast('更新しました ✓');
    renderDetail(order);
  } catch (err) {
    console.error(err);
    showToast('更新に失敗しました');
  }
  state.isLoading = false;
}

function indexToCol(idx) {
  let col = '', n = idx + 1;
  while (n > 0) { col = String.fromCharCode(((n - 1) % 26) + 65) + col; n = Math.floor((n - 1) / 26); }
  return col;
}

// ── ナビゲーション ─────────────────────────────────────
function navigate(view, data = {}) {
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  document.getElementById(`view-${view}`)?.classList.add('active');

  const showNav = ['orders', 'settings'].includes(view);
  document.getElementById('bottom-nav').classList.toggle('hidden', !showNav);
  document.querySelectorAll('.nav-item').forEach(b => b.classList.toggle('active', b.dataset.view === view));

  state.view = view;
  if (view === 'orders')  renderOrdersList();
  if (view === 'detail' && data.order) { state.selectedOrder = data.order; renderDetail(data.order); }
  if (view === 'settings') renderSettings();
  if (view === 'setup')   renderSetup();
  if (view === 'auth')    renderAuth();
}

// ── 表示: 認証 ─────────────────────────────────────────
function renderAuth() {
  const { clientId } = loadStorage();
  document.getElementById('view-auth').innerHTML = `
    <div class="auth-container">
      <div class="auth-logo">
        <img src="icons/icon.svg" alt="" width="80">
        <h1>BASE 注文管理</h1>
        <p>Googleアカウントでサインインして<br>注文を管理しましょう</p>
      </div>
      ${!clientId ? `
        <div style="width:100%">
          <label class="input-label">OAuthクライアントID</label>
          <input id="input-client-id-auth" class="text-input" type="text"
            placeholder="xxxx.apps.googleusercontent.com" value="${escHtml(clientId || '')}">
        </div>` : ''}
      <button id="btn-signin" class="btn btn-primary btn-large">Googleでサインイン</button>
      ${clientId ? `<p class="hint-text">クライアントID設定済み</p>` : ''}
      <button id="btn-goto-setup" class="btn btn-secondary" style="width:100%">初期設定を確認する</button>
    </div>`;
  document.getElementById('btn-signin').addEventListener('click', signIn);
  document.getElementById('btn-goto-setup').addEventListener('click', () => navigate('setup'));
}

// ── 表示: セットアップ ─────────────────────────────────
function renderSetup() {
  const { clientId } = loadStorage();
  document.getElementById('view-setup').innerHTML = `
    <div class="app-header" style="position:sticky;top:0;z-index:10">
      <h1>初期設定</h1>
      ${clientId ? `<button id="btn-skip" class="icon-btn">×</button>` : ''}
    </div>
    <div class="scroll-content">
      <div style="padding:12px;display:flex;flex-direction:column;gap:10px">
        <div class="card">
          <div class="card-title">OAuthクライアントIDの取得</div>
          <div class="setup-steps">
            ${[
              ['Google Cloud Console でプロジェクトを作成', 'console.cloud.google.com'],
              ['「APIとサービス」→「ライブラリ」で <strong>Google Sheets API</strong> を有効化'],
              ['「認証情報」→「OAuthクライアントID」→<strong>ウェブアプリケーション</strong>を作成'],
              ['「承認済みのJavaScriptオリジン」に <code>https://yuumato.github.io</code> を追加'],
              ['発行されたクライアントIDを下に貼り付け'],
            ].map((s, i) => `
              <div class="setup-step">
                <span class="step-num">${i+1}</span>
                <span>${s[0]}</span>
              </div>`).join('')}
          </div>
        </div>
        <div class="card">
          <div class="card-title">クライアントIDを入力</div>
          <div style="display:flex;flex-direction:column;gap:8px;padding:12px 14px">
            <input id="input-client-id-settings" class="text-input" type="text"
              placeholder="xxxx.apps.googleusercontent.com" value="${escHtml(clientId || '')}">
            <button id="btn-save-signin" class="btn btn-primary btn-full">保存してサインイン</button>
          </div>
        </div>
      </div>
    </div>`;
  document.getElementById('btn-save-signin')?.addEventListener('click', () => {
    const cid = document.getElementById('input-client-id-settings')?.value.trim();
    if (!cid) { showToast('クライアントIDを入力してください'); return; }
    state.clientId = cid;
    saveStorage({ clientId: cid });
    signIn();
  });
  document.getElementById('btn-skip')?.addEventListener('click', () => navigate('auth'));
}

// ── 表示: 注文一覧 ─────────────────────────────────────
function renderOrdersList() {
  const listEl = document.getElementById('orders-list');
  if (!listEl) return;

  // バッジ更新
  const counts = { all: state.orders.length };
  STATUS_OPTIONS.forEach(s => { counts[s] = 0; });
  state.orders.forEach(o => { if (counts[o.status_raw] !== undefined) counts[o.status_raw]++; });

  document.querySelectorAll('.filter-tabs .tab').forEach(tab => {
    const f = tab.dataset.filter;
    const label = tab.dataset.label;
    const cnt = f === 'all' ? '' : (counts[f] > 0 ? ` <span class="badge">${counts[f]}</span>` : '');
    tab.innerHTML = label + cnt;
    tab.classList.toggle('active', f === state.filter);
  });

  if (state.isLoading && !state.orders.length) {
    listEl.innerHTML = `<div class="state-box"><div class="spinner"></div><p>読み込み中...</p></div>`;
    return;
  }
  if (!state.orders.length) {
    listEl.innerHTML = `<div class="state-box">
      <div class="state-icon">📦</div>
      <p>注文データがありません</p>
      <button class="btn btn-secondary state-action" id="btn-reload">再読み込み</button>
    </div>`;
    document.getElementById('btn-reload')?.addEventListener('click', () => loadOrders(true));
    return;
  }
  if (!state.filtered.length) {
    listEl.innerHTML = `<div class="state-box"><div class="state-icon">🔍</div><p>該当する注文がありません</p></div>`;
    return;
  }

  listEl.innerHTML = state.filtered.map(o => orderCardHTML(o)).join('');
  listEl.querySelectorAll('.order-card').forEach(card => {
    card.addEventListener('click', () => {
      const o = state.filtered.find(x => x._row === Number(card.dataset.row));
      if (o) navigate('detail', { order: o });
    });
  });
}

function statusBadge(raw) {
  const def = STATUS_DEF[raw] || { css: 'unknown', label: raw || '不明' };
  return `<span class="status-badge status-${def.css}">${def.label}</span>`;
}

function fmtAmount(v) {
  if (!v) return '';
  const n = parseFloat(String(v).replace(/[,¥,，]/g, ''));
  if (isNaN(n)) return String(v);
  return '¥' + n.toLocaleString('ja-JP');
}

function orderCardHTML(o) {
  return `
    <div class="order-card" data-row="${o._row}">
      <div class="card-row-1">
        <span class="card-no">No.${escHtml(o.seq_no)}</span>
        <span class="card-date">${escHtml(o.order_date)}</span>
        ${statusBadge(o.status_raw)}
      </div>
      <div class="card-row-2">
        <span class="card-shop">${escHtml(o.shop_name)}</span>
        <span class="card-name">${escHtml(o.recipient_name)}</span>
      </div>
      <div class="card-row-3">${escHtml(o.product_title_short)}</div>
      <div class="card-row-4">
        <span class="card-amount">${escHtml(fmtAmount(o.total_amount))}</span>
        ${o.profit ? `<span class="card-profit ${parseFloat(o.profit) >= 0 ? 'profit-pos' : 'profit-neg'}">利益 ${escHtml(fmtAmount(o.profit))} (${escHtml(o.profit_rate)})</span>` : ''}
      </div>
    </div>`;
}

// ── 表示: 注文詳細 ─────────────────────────────────────
function renderDetail(o) {
  const address = [o.prefecture, o.city, o.address_detail].filter(Boolean).join(' ');
  const zip = o.zip ? `〒${o.zip}` : '';
  const profitColor = o.profit ? (parseFloat(o.profit) >= 0 ? '#34c759' : '#ff3b30') : 'inherit';

  document.querySelector('#view-detail .detail-scroll').innerHTML = `
    ${detailSection('注文情報', [
      ['No.',         `<strong>${escHtml(o.seq_no)}</strong>`],
      ['注文日',       escHtml(o.order_date)],
      ['ステータス',   statusBadge(o.status_raw)],
      ['BASE注文番号', `<span class="mono">${escHtml(o.base_order_id)}</span>`],
      ['ショップ',     escHtml(o.shop_name)],
    ])}

    ${detailSection('お客様情報', [
      ['宛名',   escHtml(o.recipient_name)],
      ['〒・住所', escHtml([zip, address].filter(Boolean).join(' '))],
      ['電話',   o.phone ? `<a href="tel:${escHtml(o.phone)}" class="link">${escHtml(o.phone)}</a>` : ''],
      ['メール', o.email ? `<a href="mailto:${escHtml(o.email)}" class="link">${escHtml(o.email)}</a>` : ''],
    ].filter(r => r[1]))}

    ${detailSection('商品情報', [
      ['商品名',     `<span class="product-title">${escHtml(o.product_title)}</span>`],
      ['バリエーション', escHtml(o.variation)],
      ['数量',     escHtml(o.quantity)],
      ['単価',     escHtml(fmtAmount(o.unit_price))],
      ['合計',     `<strong>${escHtml(fmtAmount(o.total_amount))}</strong>`],
      ['手数料',   escHtml(fmtAmount(o.commission))],
      ['商品ページ', o.product_page ? `<a href="https://${escHtml(o.product_page)}" target="_blank" class="link">開く</a>` : ''],
    ].filter(r => r[1]))}

    ${detailSection('仕入れ情報', [
      ['仕入注文番号', escHtml(o.purchase_order)],
      ['ASIN(子)',    escHtml(o.asin_child)],
      ['仕入金額',   escHtml(fmtAmount(o.purchase_price))],
      ['到着予定日', escHtml(o.arrival_date1)],
      ['到着日',     escHtml(o.arrival_date)],
      ['Amazonリンク', o.amazon_link ? `<a href="https://${escHtml(o.amazon_link)}" target="_blank" class="link">開く</a>` : ''],
      ['Ama注文ページ', o.amazon_page ? `<a href="${escHtml(o.amazon_page)}" target="_blank" class="link">開く</a>` : ''],
    ].filter(r => r[1]))}

    ${detailSection('配送', [
      ['追跡番号', escHtml(o.tracking_number) || '<span class="hint">未入力</span>'],
      ['出荷通知メール', escHtml(o.mail_delivery)],
      ['評価依頼メール', escHtml(o.mail_review)],
    ])}

    ${o.profit ? detailSection('収益', [
      ['売上',   escHtml(fmtAmount(o.total_amount))],
      ['コスト', escHtml(fmtAmount(o.cost))],
      ['利益',   `<strong style="color:${profitColor}">${escHtml(fmtAmount(o.profit))}</strong>`],
      ['利益率', `<strong style="color:${profitColor}">${escHtml(o.profit_rate)}</strong>`],
    ]) : ''}

    <div class="card" style="margin:8px 12px">
      <div class="card-title">ステータス・追跡番号を更新</div>
      <div style="padding:12px 14px;display:flex;flex-direction:column;gap:10px">
        <div>
          <label class="input-label">追跡番号</label>
          <input id="input-tracking" class="text-input" type="text"
            placeholder="追跡番号を入力" value="${escHtml(o.tracking_number)}">
        </div>
        <div>
          <label class="input-label">ステータス</label>
          <select id="select-status" class="select-input">
            ${STATUS_OPTIONS.map(s =>
              `<option value="${s}" ${o.status_raw === s ? 'selected' : ''}>${s}</option>`
            ).join('')}
          </select>
        </div>
        <button id="btn-update" class="btn btn-primary btn-full">スプレッドシートに反映</button>
      </div>
    </div>

    <div style="height:16px"></div>
  `;

  document.getElementById('btn-update')?.addEventListener('click', () => {
    const newStatus  = document.getElementById('select-status')?.value;
    const tracking   = document.getElementById('input-tracking')?.value.trim();
    updateOrder(o, newStatus, tracking);
  });
}

function detailSection(title, rows) {
  return `
    <div class="card" style="margin:8px 12px">
      <div class="card-title">${title}</div>
      ${rows.map(([label, value]) => `
        <div class="detail-row">
          <span class="detail-label">${label}</span>
          <span class="detail-value">${value}</span>
        </div>`).join('')}
    </div>`;
}

// ── 表示: 設定 ─────────────────────────────────────────
function renderSettings() {
  const syncText = state.lastSynced ? state.lastSynced.toLocaleTimeString('ja-JP') : '未同期';
  document.querySelector('#view-settings .settings-scroll').innerHTML = `
    <div style="padding:8px 0;display:flex;flex-direction:column;gap:0">
      <div class="card" style="margin:8px 12px">
        <div class="card-title">アカウント</div>
        <div class="detail-row"><span class="detail-label">状態</span>
          <span class="detail-value">${state.isSignedIn ? '✅ サインイン済み' : '未サインイン'}</span></div>
        <div class="detail-row"><span class="detail-label">最終同期</span>
          <span class="detail-value">${syncText}</span></div>
        <div style="padding:10px 14px">
          <button id="btn-signout" class="btn btn-outline btn-full">サインアウト</button>
        </div>
      </div>

      <div class="card" style="margin:8px 12px">
        <div class="card-title">データ</div>
        <div class="detail-row"><span class="detail-label">注文件数</span>
          <span class="detail-value">${state.orders.length} 件</span></div>
        <div class="detail-row"><span class="detail-label">シート名</span>
          <span class="detail-value">${escHtml(state.sheetName)}</span></div>
        <div style="padding:10px 14px">
          <button id="btn-refresh-data" class="btn btn-secondary btn-full">データを再読み込み</button>
        </div>
      </div>

      <div class="card" style="margin:8px 12px">
        <div class="card-title">接続設定</div>
        <div style="padding:10px 14px;display:flex;flex-direction:column;gap:8px">
          <label class="input-label">OAuthクライアントID</label>
          <input id="input-cid-settings" class="text-input" type="text"
            placeholder="xxxx.apps.googleusercontent.com"
            value="${escHtml(loadStorage().clientId || '')}">
          <button id="btn-update-cid" class="btn btn-secondary btn-full">変更して再サインイン</button>
        </div>
        <div class="detail-row"><span class="detail-label">スプレッドシートID</span>
          <span class="detail-value" style="font-size:11px;word-break:break-all">${SPREADSHEET_ID}</span></div>
      </div>
    </div>`;

  document.getElementById('btn-signout')?.addEventListener('click', signOut);
  document.getElementById('btn-refresh-data')?.addEventListener('click', () => {
    sessionStorage.removeItem('base_orders');
    loadOrders(true);
    showToast('再読み込み中...');
    navigate('orders');
  });
  document.getElementById('btn-update-cid')?.addEventListener('click', () => {
    const cid = document.getElementById('input-cid-settings')?.value.trim();
    if (!cid) return;
    state.clientId = cid;
    saveStorage({ clientId: cid, accessToken: null });
    gapi.client.setToken(null);
    initTokenClient(cid);
    state.tokenClient?.requestAccessToken({ prompt: 'consent' });
  });
}

// ── Toast ──────────────────────────────────────────────
let toastTimer;
function showToast(msg, dur = 2500) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(toastTimer);
  if (dur < 60000) toastTimer = setTimeout(() => t.classList.remove('show'), dur);
}

function escHtml(s) {
  return String(s ?? '')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ── イベント配線 ───────────────────────────────────────
function wireEvents() {
  document.querySelectorAll('.nav-item').forEach(b => {
    b.addEventListener('click', () => navigate(b.dataset.view));
  });

  document.getElementById('btn-back')?.addEventListener('click', () => navigate('orders'));

  document.getElementById('btn-refresh')?.addEventListener('click', () => {
    sessionStorage.removeItem('base_orders');
    loadOrders(true);
    showToast('再読み込み中...');
  });

  document.getElementById('btn-search-toggle')?.addEventListener('click', () => {
    state.showSearch = !state.showSearch;
    document.getElementById('search-bar').classList.toggle('hidden', !state.showSearch);
    if (state.showSearch) document.getElementById('search-input')?.focus();
    else {
      state.searchQuery = '';
      if (document.getElementById('search-input')) document.getElementById('search-input').value = '';
      applyFilter(); renderOrdersList();
    }
  });

  let searchTimer;
  document.getElementById('search-input')?.addEventListener('input', e => {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => {
      state.searchQuery = e.target.value.trim();
      applyFilter(); renderOrdersList();
    }, 200);
  });

  document.querySelectorAll('.filter-tabs .tab').forEach(tab => {
    tab.addEventListener('click', () => {
      state.filter = tab.dataset.filter;
      applyFilter(); renderOrdersList();
    });
  });
}

// ── 初期化 ─────────────────────────────────────────────
function init() {
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('./sw.js').catch(() => {});

  wireEvents();

  const { clientId } = loadStorage();
  navigate(clientId ? 'auth' : 'setup');

  const poll = setInterval(() => {
    if (window.gapi && window.google?.accounts?.oauth2) {
      clearInterval(poll);
      onGapiLoaded();
      onGisLoaded();
    }
  }, 100);
}

document.addEventListener('DOMContentLoaded', init);
