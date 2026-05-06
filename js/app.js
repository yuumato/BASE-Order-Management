/* BASE 注文管理 PWA – app.js v3 (動的列検出) */

const SPREADSHEET_ID = '1otZ-q-pp0i6biDdlRe6iVimWCQ8rEQBFMpch6Dv3axE';
const SCOPES = 'https://www.googleapis.com/auth/spreadsheets';
const STORAGE_KEY = 'base_order_pwa';

const STATUS_DEF = {
  '完了':         { css: 'completed', label: '完了' },
  '評価依頼待ち': { css: 'review',    label: '評価依頼待ち' },
  '仕入済':       { css: 'sourced',   label: '仕入済' },
  '仕入れ済':     { css: 'sourced',   label: '仕入済' },
  'キャンセル':   { css: 'cancelled', label: 'キャンセル' },
};
const STATUS_OPTIONS = ['仕入済', '評価依頼待ち', '完了', 'キャンセル'];

const state = {
  view: 'auth', orders: [], filtered: [],
  selectedOrder: null, filter: 'all', searchQuery: '',
  isLoading: false, isSignedIn: false,
  sheetName: 'Sheet1', lastSynced: null,
  clientId: '', tokenClient: null,
  colMap: {}, showSearch: false,
};

// ── Storage ────────────────────────────────────────────────
function loadStorage() { try { return JSON.parse(localStorage.getItem(STORAGE_KEY)||'{}'); } catch { return {}; } }
function saveStorage(d) { localStorage.setItem(STORAGE_KEY, JSON.stringify({...loadStorage(),...d})); }
function loadCache()    { try { return JSON.parse(sessionStorage.getItem('base_orders')||'null'); } catch { return null; } }
function saveCache(o)   { try { sessionStorage.setItem('base_orders', JSON.stringify(o)); } catch {} }

// ── 動的列検出 ─────────────────────────────────────────────
function detectColumns(h1, h2) {
  // 完全一致 → 部分一致の順でフォールバック
  const find = (row, names) => {
    // 完全一致を優先
    for (const n of names) {
      const i = row.findIndex(v => v === n);
      if (i >= 0) return i;
    }
    // 部分一致（前後の空白・改行を除去して含むか）
    for (const n of names) {
      const i = row.findIndex(v => v.includes(n) || n.includes(v));
      if (i >= 0) return i;
    }
    return -1;
  };
  const nth = (row, name, n) => {
    let c = 0;
    for (let i = 0; i < row.length; i++) {
      if (row[i] === name || row[i].includes(name)) { if (c === n) return i; c++; }
    }
    return -1;
  };

  const c = {};
  // 1行目から検出
  c.order_date     = find(h1, ['注文日']);
  c.seq_no         = find(h1, ['No.', 'No']);
  c.status         = find(h1, ['ステータス']);
  c.base_order_id  = find(h1, ['BASE注文番号']);
  c.shop_id        = find(h1, ['ショップID']);
  c.shop_name      = find(h1, ['ショップ名']);
  c.image_url      = find(h1, ['画像リンク']);
  c.product_page   = find(h1, ['商品ページ']);
  c.product_title  = find(h1, ['タイトル']);        // 最初のタイトル列
  c.variation      = find(h1, ['バリエーション内容']); // 最初のバリエーション列
  c.purchase_order = find(h1, ['仕入注文番号']);
  c.amazon_page    = find(h1, ['Ama注文ページ']);
  c.arrival_date1  = find(h1, ['到着予定日']);
  c.arrival_date   = find(h1, ['到着日']);
  c.unit_price     = find(h1, ['単価']);
  c.total_amount   = find(h1, ['合計(小計+オプション)','合計']);
  c.commission     = find(h1, ['手数料']);
  c.quantity       = find(h1, ['数量']);
  c.purchase_price = find(h1, ['仕入金額']);
  c.base_link      = find(h1, ['BASE']);
  c.amazon_link    = find(h1, ['AmazonJP']);
  c.tracking       = find(h1, ['追跡番号']);
  c.cost           = find(h1, ['コスト']);
  c.profit         = find(h1, ['利益']);
  c.profit_rate    = find(h1, ['利益率']);

  // 2行目から検出
  c.mail_orei      = find(h2, ['御礼']);
  c.mail_shiire    = find(h2, ['仕入れ']);
  c.mail_delivery  = find(h2, ['出荷通知']);
  c.mail_review    = find(h2, ['評価依頼']);
  c.recipient_name = find(h2, ['送付先名']);
  c.zip            = nth(h2, '〒', 0);          // 1つ目の〒（結合形式）
  c.address_street = find(h2, ['住所2']);
  c.phone          = nth(h2, '電話番号', 0);    // 1つ目の電話番号
  c.addressee      = find(h2, ['宛名']);
  c.email          = find(h2, ['メールアドレス']);

  // 2つ目の〒の位置から都道府県・市区町村・住所詳細を算出
  const zip2 = nth(h2, '〒', 1);
  c.zip1       = zip2;
  c.zip2       = zip2 >= 0 ? zip2 + 1 : -1;
  c.prefecture = zip2 >= 0 ? zip2 + 2 : -1;
  c.city       = zip2 >= 0 ? zip2 + 3 : -1;
  c.address_detail = -1;
  if (zip2 >= 0) {
    for (let i = zip2 + 1; i < h2.length; i++) {
      if (h2[i] === '住所') { c.address_detail = i; break; }
    }
  }

  // ASIN（1行目に"ASIN"、2行目に"親""子"）
  c.asin_parent = nth(h2, '親', 0);
  c.asin_child  = nth(h2, '子', 0);
  // リンク（1行目に"リンク"、2行目に2つ目の"親""子"）
  c.link_parent = nth(h2, '親', 1);
  c.link_child  = nth(h2, '子', 1);

  return c;
}

// ── セル値取得 ─────────────────────────────────────────────
function gv(row, field) {
  const idx = state.colMap[field];
  if (idx === undefined || idx < 0 || idx >= row.length) return '';
  const v = row[idx];
  if (v === undefined || v === null) return '';
  // Google Sheets の日付シリアル値を変換（数値で30000〜55000の範囲）
  if (typeof v === 'number' && v > 30000 && v < 55000) {
    const d = new Date((v - 25569) * 86400 * 1000);
    return `${d.getUTCMonth()+1}/${d.getUTCDate()}`;
  }
  return String(v).trim();
}

function makeUrl(s) {
  if (!s) return '';
  return s.startsWith('http') ? s : 'https://' + s;
}

function fmtAmount(v) {
  if (!v) return '';
  const n = parseFloat(String(v).replace(/[,¥，,]/g, ''));
  return isNaN(n) ? String(v) : '¥' + n.toLocaleString('ja-JP');
}

function escHtml(s) {
  return String(s ?? '')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ── Google API ─────────────────────────────────────────────
let gapiReady = false, gisReady = false;

function onGapiLoaded() {
  gapi.load('client', async () => {
    await gapi.client.init({ discoveryDocs: ['https://sheets.googleapis.com/$discovery/rest?version=v4'] });
    gapiReady = true; tryAutoSignIn();
  });
}
function onGisLoaded() { gisReady = true; tryAutoSignIn(); }

function tryAutoSignIn() {
  if (!gapiReady || !gisReady) return;
  const { clientId, accessToken } = loadStorage();
  state.clientId = clientId || '';
  if (clientId && accessToken) {
    gapi.client.setToken({ access_token: accessToken });
    initTokenClient(clientId);
    verifyToken();
  } else if (clientId) {
    initTokenClient(clientId); navigate('auth');
  } else {
    navigate('setup');
  }
}

function initTokenClient(clientId) {
  if (!clientId) return;
  state.tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: clientId, scope: SCOPES,
    callback: (resp) => {
      if (resp.error) { showToast('認証エラー: ' + resp.error); return; }
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
    gapi.client.setToken(null); saveStorage({ accessToken: null }); navigate('auth');
  }
}

function signIn() {
  const el = document.getElementById('input-client-id-auth') || document.getElementById('input-client-id-settings');
  const cid = el?.value.trim() || state.clientId;
  if (!cid) { showToast('クライアントIDを入力してください'); navigate('setup'); return; }
  state.clientId = cid; saveStorage({ clientId: cid });
  initTokenClient(cid);
  state.tokenClient?.requestAccessToken({ prompt: 'consent' });
}

function signOut() {
  const token = gapi.client.getToken();
  if (token) google.accounts.oauth2.revoke(token.access_token);
  gapi.client.setToken(null);
  state.isSignedIn = false; state.orders = [];
  saveStorage({ accessToken: null }); navigate('auth');
}

async function afterSignIn() { state.isSignedIn = true; navigate('orders'); await loadOrders(); }

// ── データ読み込み ─────────────────────────────────────────
async function loadOrders(forceRefresh = false) {
  if (state.isLoading) return;
  if (!forceRefresh) {
    const cached = loadCache();
    if (cached && Object.keys(state.colMap).length > 0) {
      state.orders = cached; applyFilter(); renderOrdersList(); return;
    }
  }
  state.isLoading = true; renderOrdersList();
  try {
    const meta = await gapi.client.sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    // データは2枚目のシート（02...）にある
    const sheets = meta.result.sheets;
    state.sheetName = sheets[1]?.properties?.title || sheets[0]?.properties?.title || 'Sheet1';

    // 1行目から全読み込み（ヘッダー2行＋データ）
    const resp = await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${state.sheetName}!A1:BG`,
      valueRenderOption: 'UNFORMATTED_VALUE',
    });
    const all = resp.result.values || [];
    if (all.length < 3) { state.orders = []; state.isLoading = false; applyFilter(); renderOrdersList(); return; }

    const h1 = (all[0] || []).map(v => String(v||'').trim());
    const h2 = (all[1] || []).map(v => String(v||'').trim());
    state.colMap = detectColumns(h1, h2);

    // データは3行目（index 2）から
    // 行フィルター: 生データで空行を除外（列検出失敗でも落とさない）
    state.orders = all.slice(2)
      .filter(row => row && row.some(v => String(v ?? '').trim() !== ''))
      .map((row, i) => parseRow(row, i + 3));

    state.lastSynced = new Date();
    saveCache(state.orders);
  } catch (err) {
    console.error(err);
    if (err.status === 401 || err.status === 403) {
      showToast('認証が切れました'); gapi.client.setToken(null);
      saveStorage({ accessToken: null }); navigate('auth');
    } else { showToast('読み込みに失敗しました'); }
  }
  state.isLoading = false; applyFilter(); renderOrdersList();
}

function parseRow(row, rowIndex) {
  const title = gv(row, 'product_title');
  return {
    _row: rowIndex,
    order_date:      gv(row, 'order_date'),
    seq_no:          gv(row, 'seq_no'),
    status_raw:      gv(row, 'status'),
    base_order_id:   gv(row, 'base_order_id'),
    shop_id:         gv(row, 'shop_id'),
    shop_name:       gv(row, 'shop_name'),
    mail_delivery:   gv(row, 'mail_delivery'),
    mail_review:     gv(row, 'mail_review'),
    recipient_name:  gv(row, 'recipient_name'),
    zip:             gv(row, 'zip'),
    phone:           gv(row, 'phone'),
    addressee:       gv(row, 'addressee'),
    prefecture:      gv(row, 'prefecture'),
    city:            gv(row, 'city'),
    address_detail:  gv(row, 'address_detail'),
    address_street:  gv(row, 'address_street'),
    email:           gv(row, 'email'),
    image_url:       gv(row, 'image_url'),
    product_page:    gv(row, 'product_page'),
    product_title:   title,
    product_short:   title.length > 45 ? title.slice(0, 45) + '…' : title,
    variation:       gv(row, 'variation'),
    asin_parent:     gv(row, 'asin_parent'),
    asin_child:      gv(row, 'asin_child'),
    link_parent:     gv(row, 'link_parent'),
    link_child:      gv(row, 'link_child'),
    purchase_order:  gv(row, 'purchase_order'),
    amazon_page:     gv(row, 'amazon_page'),
    arrival_date1:   gv(row, 'arrival_date1'),
    arrival_date:    gv(row, 'arrival_date'),
    unit_price:      gv(row, 'unit_price'),
    total_amount:    gv(row, 'total_amount'),
    commission:      gv(row, 'commission'),
    quantity:        gv(row, 'quantity'),
    purchase_price:  gv(row, 'purchase_price'),
    base_link:       gv(row, 'base_link'),
    amazon_link:     gv(row, 'amazon_link'),
    tracking:        gv(row, 'tracking'),
    cost:            gv(row, 'cost'),
    profit:          gv(row, 'profit'),
    profit_rate:     gv(row, 'profit_rate'),
  };
}

// ── フィルター ─────────────────────────────────────────────
function applyFilter() {
  let list = state.orders;
  if (state.filter !== 'all') list = list.filter(o => o.status_raw === state.filter);
  if (state.searchQuery) {
    const q = state.searchQuery.toLowerCase();
    list = list.filter(o =>
      o.seq_no.includes(q) || o.base_order_id.toLowerCase().includes(q) ||
      o.recipient_name.toLowerCase().includes(q) || o.shop_name.toLowerCase().includes(q) ||
      o.product_title.toLowerCase().includes(q)
    );
  }
  state.filtered = list;
}

// ── スプレッドシート更新 ───────────────────────────────────
function colLetter(idx) {
  let s = '', n = idx + 1;
  while (n > 0) { s = String.fromCharCode(((n-1)%26)+65)+s; n=Math.floor((n-1)/26); }
  return s;
}

async function updateOrder(order, newStatus, tracking) {
  state.isLoading = true; showToast('更新中...', 60000);
  try {
    const data = [];
    const sc = state.colMap.status, tc = state.colMap.tracking, r = order._row;
    if (sc >= 0) data.push({ range: `${state.sheetName}!${colLetter(sc)}${r}`, values: [[newStatus]] });
    if (tc >= 0) data.push({ range: `${state.sheetName}!${colLetter(tc)}${r}`, values: [[tracking]] });
    await gapi.client.sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      resource: { valueInputOption: 'RAW', data },
    });
    order.status_raw = newStatus; order.tracking = tracking;
    saveCache(state.orders); applyFilter();
    showToast('更新しました ✓'); renderDetail(order);
  } catch(e) { console.error(e); showToast('更新に失敗しました'); }
  state.isLoading = false;
}

// ── ナビゲーション ─────────────────────────────────────────
function navigate(view, data={}) {
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  document.getElementById(`view-${view}`)?.classList.add('active');
  const showNav = ['orders','settings'].includes(view);
  document.getElementById('bottom-nav').classList.toggle('hidden', !showNav);
  document.querySelectorAll('.nav-item').forEach(b => b.classList.toggle('active', b.dataset.view===view));
  state.view = view;
  if (view==='orders')  renderOrdersList();
  if (view==='detail' && data.order) { state.selectedOrder=data.order; renderDetail(data.order); }
  if (view==='settings') renderSettings();
  if (view==='setup')   renderSetup();
  if (view==='auth')    renderAuth();
}

// ── Render: Auth ───────────────────────────────────────────
function renderAuth() {
  const { clientId } = loadStorage();
  document.getElementById('view-auth').innerHTML = `
    <div class="auth-container">
      <div class="auth-logo">
        <img src="icons/icon.svg" alt="" width="72">
        <h1>BASE 注文管理</h1>
        <p>Googleアカウントでサインインして<br>注文を管理しましょう</p>
      </div>
      ${!clientId ? `<div style="width:100%">
        <label class="input-label">OAuthクライアントID</label>
        <input id="input-client-id-auth" class="text-input" type="text"
          placeholder="xxxx.apps.googleusercontent.com" value="${escHtml(clientId||'')}">
      </div>` : ''}
      <button id="btn-signin" class="btn btn-primary btn-large">Googleでサインイン</button>
      ${clientId ? `<p class="hint-text">クライアントID設定済み</p>` : ''}
      <button id="btn-goto-setup" class="btn btn-secondary" style="width:100%">初期設定を確認する</button>
    </div>`;
  document.getElementById('btn-signin').addEventListener('click', signIn);
  document.getElementById('btn-goto-setup').addEventListener('click', () => navigate('setup'));
}

// ── Render: Setup ──────────────────────────────────────────
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
              'Google Cloud Console でプロジェクトを作成',
              '「APIとサービス」→「ライブラリ」で <strong>Google Sheets API</strong> を有効化',
              '「認証情報」→「OAuthクライアントID」→<strong>ウェブアプリケーション</strong>を作成',
              '「承認済みのJavaScriptオリジン」に <code>https://yuumato.github.io</code> を追加',
              '発行されたクライアントIDを下に貼り付け',
            ].map((s,i) => `<div class="setup-step"><span class="step-num">${i+1}</span><span>${s}</span></div>`).join('')}
          </div>
        </div>
        <div class="card">
          <div class="card-title">クライアントIDを入力</div>
          <div style="padding:12px 14px;display:flex;flex-direction:column;gap:8px">
            <input id="input-client-id-settings" class="text-input" type="text"
              placeholder="xxxx.apps.googleusercontent.com" value="${escHtml(clientId||'')}">
            <button id="btn-save-signin" class="btn btn-primary btn-full">保存してサインイン</button>
          </div>
        </div>
      </div>
    </div>`;
  document.getElementById('btn-save-signin')?.addEventListener('click', () => {
    const cid = document.getElementById('input-client-id-settings')?.value.trim();
    if (!cid) { showToast('クライアントIDを入力してください'); return; }
    state.clientId = cid; saveStorage({ clientId: cid }); signIn();
  });
  document.getElementById('btn-skip')?.addEventListener('click', () => navigate('auth'));
}

// ── Render: Orders List ────────────────────────────────────
function statusBadge(raw) {
  const def = STATUS_DEF[raw] || { css: 'unknown', label: raw || '不明' };
  return `<span class="status-badge status-${def.css}">${def.label}</span>`;
}

function renderOrdersList() {
  const el = document.getElementById('orders-list');
  if (!el) return;

  const counts = { all: state.orders.length };
  STATUS_OPTIONS.forEach(s => { counts[s] = 0; });
  state.orders.forEach(o => { if (counts[o.status_raw] !== undefined) counts[o.status_raw]++; });
  document.querySelectorAll('.filter-tabs .tab').forEach(tab => {
    const f = tab.dataset.filter, label = tab.dataset.label;
    const badge = f !== 'all' && counts[f] > 0 ? ` <span class="badge">${counts[f]}</span>` : '';
    tab.innerHTML = label + badge;
    tab.classList.toggle('active', f === state.filter);
  });

  if (state.isLoading && !state.orders.length) {
    el.innerHTML = `<div class="state-box"><div class="spinner"></div><p>読み込み中...</p></div>`; return;
  }
  if (!state.orders.length) {
    el.innerHTML = `<div class="state-box"><div class="state-icon">📦</div><p>注文データがありません</p>
      <button class="btn btn-secondary state-action" id="btn-reload">再読み込み</button></div>`;
    document.getElementById('btn-reload')?.addEventListener('click', () => loadOrders(true)); return;
  }
  if (!state.filtered.length) {
    el.innerHTML = `<div class="state-box"><div class="state-icon">🔍</div><p>該当する注文がありません</p></div>`; return;
  }

  el.innerHTML = state.filtered.map(o => `
    <div class="order-card" data-row="${o._row}">
      <div class="card-row-1">
        <span class="card-no">No.${escHtml(o.seq_no)}</span>
        <span class="card-date">${escHtml(o.order_date)}</span>
        ${statusBadge(o.status_raw)}
      </div>
      <div class="card-row-2">
        ${o.shop_name ? `<span class="card-shop">${escHtml(o.shop_name)}</span>` : ''}
        <span class="card-name">${escHtml(o.recipient_name)}</span>
      </div>
      <div class="card-row-3">${escHtml(o.product_short)}</div>
      <div class="card-row-4">
        <span class="card-amount">${escHtml(fmtAmount(o.total_amount))}</span>
        ${o.profit ? `<span class="card-profit ${parseFloat(o.profit)>=0?'profit-pos':'profit-neg'}">
          利益 ${escHtml(fmtAmount(o.profit))} (${escHtml(o.profit_rate)})</span>` : ''}
      </div>
    </div>`).join('');

  el.querySelectorAll('.order-card').forEach(card => {
    card.addEventListener('click', () => {
      const o = state.filtered.find(x => x._row === Number(card.dataset.row));
      if (o) navigate('detail', { order: o });
    });
  });
}

// ── Render: Detail ─────────────────────────────────────────
function section(title, rows) {
  const filtered = rows.filter(r => r && r[1]);
  if (!filtered.length) return '';
  return `<div class="card" style="margin:6px 10px">
    <div class="card-title">${title}</div>
    ${filtered.map(([label, value]) => `
      <div class="detail-row">
        <span class="detail-label">${label}</span>
        <span class="detail-value">${value}</span>
      </div>`).join('')}
  </div>`;
}

function renderDetail(o) {
  const addr = [o.prefecture, o.city, o.address_detail || o.address_street].filter(Boolean).join(' ');
  const zip = o.zip ? `〒${o.zip} ` : '';
  const profitColor = o.profit ? (parseFloat(o.profit) >= 0 ? '#34c759' : '#ff3b30') : '';

  const amazonUrl = makeUrl(o.amazon_link || o.link_child);
  const baseUrl   = makeUrl(o.base_link   || o.product_page);

  document.querySelector('#view-detail .detail-scroll').innerHTML = `

    ${o.image_url ? `
    <div class="card" style="margin:6px 10px;overflow:hidden">
      <div class="card-title">商品画像</div>
      <div style="background:#f4f4f6;text-align:center;padding:8px">
        <img src="${escHtml(o.image_url)}" alt="商品画像"
          style="max-width:100%;max-height:180px;object-fit:contain;border-radius:6px"
          onerror="this.parentElement.style.display='none'">
      </div>
    </div>` : ''}

    <div class="card" style="margin:6px 10px">
      <div class="card-title">商品タイトル</div>
      <div style="padding:10px 14px;font-size:13px;line-height:1.6;color:#111">
        ${escHtml(o.product_title) || '<span style="color:#999">なし</span>'}
        ${o.variation ? `<div style="margin-top:4px;font-size:12px;color:#666">${escHtml(o.variation)}</div>` : ''}
      </div>
    </div>

    ${section('注文情報', [
      ['No.',         `<strong>${escHtml(o.seq_no)}</strong>`],
      ['注文日',       escHtml(o.order_date)],
      ['ステータス',   statusBadge(o.status_raw)],
      ['BASE注文番号', `<span class="mono">${escHtml(o.base_order_id)}</span>`],
      ['ショップ',     escHtml(o.shop_name)],
    ])}

    ${section('お客様情報', [
      ['宛名',    escHtml(o.addressee || o.recipient_name)],
      ['〒・住所', escHtml(zip + addr)],
      ['電話',    o.phone ? `<a href="tel:${escHtml(o.phone)}" class="link">${escHtml(o.phone)}</a>` : ''],
      ['メール',  o.email ? `<a href="mailto:${escHtml(o.email)}" class="link">${escHtml(o.email)}</a>` : ''],
    ])}

    <div class="card" style="margin:6px 10px">
      <div class="card-title">仕入れURL</div>
      ${amazonUrl ? `<div class="detail-row">
        <span class="detail-label">Amazon</span>
        <a href="${escHtml(amazonUrl)}" target="_blank" class="detail-value link-btn">
          Amazon.co.jp で開く →
        </a></div>` : ''}
      ${baseUrl ? `<div class="detail-row">
        <span class="detail-label">BASE</span>
        <a href="${escHtml(baseUrl)}" target="_blank" class="detail-value link-btn">
          BASE ショップで開く →
        </a></div>` : ''}
      ${!amazonUrl && !baseUrl ? `<div class="detail-row"><span class="detail-value" style="color:#999">URLなし</span></div>` : ''}
    </div>

    ${section('仕入れ情報', [
      ['仕入注文番号', escHtml(o.purchase_order)],
      ['ASIN',       escHtml(o.asin_child || o.asin_parent)],
      ['仕入金額',   escHtml(fmtAmount(o.purchase_price))],
      ['到着予定日', escHtml(o.arrival_date1)],
      ['到着日',     escHtml(o.arrival_date)],
      ['Amazon注文', o.amazon_page ? `<a href="${escHtml(o.amazon_page)}" target="_blank" class="link">注文ページを開く →</a>` : ''],
    ])}

    ${section('金額', [
      ['単価',   escHtml(fmtAmount(o.unit_price))],
      ['数量',   escHtml(o.quantity)],
      ['合計',   `<strong>${escHtml(fmtAmount(o.total_amount))}</strong>`],
      ['手数料', escHtml(fmtAmount(o.commission))],
      ['仕入',   escHtml(fmtAmount(o.purchase_price))],
      o.profit ? ['利益', `<strong style="color:${profitColor}">${escHtml(fmtAmount(o.profit))} (${escHtml(o.profit_rate)})</strong>`] : null,
    ])}

    ${section('配送', [
      ['追跡番号',     escHtml(o.tracking) || '<span class="hint">未入力</span>'],
      ['出荷通知',     escHtml(o.mail_delivery)],
      ['評価依頼',     escHtml(o.mail_review)],
    ])}

    <div class="card" style="margin:6px 10px">
      <div class="card-title">ステータス・追跡番号を更新</div>
      <div style="padding:12px 14px;display:flex;flex-direction:column;gap:10px">
        <div>
          <label class="input-label">追跡番号</label>
          <input id="input-tracking" class="text-input" type="text"
            placeholder="追跡番号を入力" value="${escHtml(o.tracking)}">
        </div>
        <div>
          <label class="input-label">ステータス</label>
          <select id="select-status" class="select-input">
            ${STATUS_OPTIONS.map(s =>
              `<option value="${s}" ${o.status_raw===s?'selected':''}>${s}</option>`
            ).join('')}
          </select>
        </div>
        <button id="btn-update" class="btn btn-primary btn-full">スプレッドシートに反映</button>
      </div>
    </div>
    <div style="height:16px"></div>`;

  document.getElementById('btn-update')?.addEventListener('click', () => {
    updateOrder(o,
      document.getElementById('select-status')?.value,
      document.getElementById('input-tracking')?.value.trim()
    );
  });
}

// ── Render: Settings ──────────────────────────────────────
function renderSettings() {
  const sync = state.lastSynced ? state.lastSynced.toLocaleTimeString('ja-JP') : '未同期';
  document.querySelector('#view-settings .settings-scroll').innerHTML = `
    <div style="padding:6px 0">
      <div class="card" style="margin:6px 10px">
        <div class="card-title">アカウント</div>
        <div class="detail-row"><span class="detail-label">状態</span>
          <span class="detail-value">${state.isSignedIn ? '✅ サインイン済み' : '未サインイン'}</span></div>
        <div class="detail-row"><span class="detail-label">最終同期</span>
          <span class="detail-value">${sync}</span></div>
        <div style="padding:10px 14px">
          <button id="btn-signout" class="btn btn-outline btn-full">サインアウト</button>
        </div>
      </div>
      <div class="card" style="margin:6px 10px">
        <div class="card-title">データ</div>
        <div class="detail-row"><span class="detail-label">注文件数</span>
          <span class="detail-value">${state.orders.length} 件</span></div>
        <div class="detail-row"><span class="detail-label">シート名</span>
          <span class="detail-value">${escHtml(state.sheetName)}</span></div>
        <div style="padding:10px 14px">
          <button id="btn-refresh-data" class="btn btn-secondary btn-full">データを再読み込み</button>
        </div>
      </div>
      <div class="card" style="margin:6px 10px">
        <div class="card-title">接続設定</div>
        <div style="padding:10px 14px;display:flex;flex-direction:column;gap:8px">
          <label class="input-label">OAuthクライアントID</label>
          <input id="input-cid" class="text-input" type="text"
            placeholder="xxxx.apps.googleusercontent.com"
            value="${escHtml(loadStorage().clientId||'')}">
          <button id="btn-update-cid" class="btn btn-secondary btn-full">変更して再サインイン</button>
        </div>
        <div class="detail-row"><span class="detail-label">スプレッドシートID</span>
          <span class="detail-value" style="font-size:11px;word-break:break-all">${SPREADSHEET_ID}</span></div>
      </div>
      <div class="card" style="margin:6px 10px">
        <div class="card-title">列検出状況（デバッグ）</div>
        <div style="padding:10px 14px;font-size:11px;color:#555;line-height:1.8;word-break:break-all">
          ${Object.entries(state.colMap).filter(([,v])=>v>=0).map(([k,v])=>`<b>${k}</b>:${v}`).join('　') || '未検出（再読み込みしてください）'}
        </div>
      </div>
    </div>`;
  document.getElementById('btn-signout')?.addEventListener('click', signOut);
  document.getElementById('btn-refresh-data')?.addEventListener('click', () => {
    sessionStorage.removeItem('base_orders'); state.colMap = {};
    loadOrders(true); showToast('再読み込み中...'); navigate('orders');
  });
  document.getElementById('btn-update-cid')?.addEventListener('click', () => {
    const cid = document.getElementById('input-cid')?.value.trim();
    if (!cid) return;
    state.clientId = cid; saveStorage({ clientId: cid, accessToken: null });
    gapi.client.setToken(null); initTokenClient(cid);
    state.tokenClient?.requestAccessToken({ prompt: 'consent' });
  });
}

// ── Toast ──────────────────────────────────────────────────
let toastTimer;
function showToast(msg, dur=2500) {
  const t = document.getElementById('toast');
  t.textContent = msg; t.classList.add('show');
  clearTimeout(toastTimer);
  if (dur < 60000) toastTimer = setTimeout(() => t.classList.remove('show'), dur);
}

// ── イベント ───────────────────────────────────────────────
function wireEvents() {
  document.querySelectorAll('.nav-item').forEach(b =>
    b.addEventListener('click', () => navigate(b.dataset.view)));
  document.getElementById('btn-back')?.addEventListener('click', () => navigate('orders'));
  document.getElementById('btn-refresh')?.addEventListener('click', () => {
    sessionStorage.removeItem('base_orders'); state.colMap = {};
    loadOrders(true); showToast('再読み込み中...');
  });
  document.getElementById('btn-search-toggle')?.addEventListener('click', () => {
    state.showSearch = !state.showSearch;
    document.getElementById('search-bar').classList.toggle('hidden', !state.showSearch);
    if (state.showSearch) document.getElementById('search-input')?.focus();
    else { state.searchQuery=''; if(document.getElementById('search-input')) document.getElementById('search-input').value=''; applyFilter(); renderOrdersList(); }
  });
  let st;
  document.getElementById('search-input')?.addEventListener('input', e => {
    clearTimeout(st); st = setTimeout(() => { state.searchQuery=e.target.value.trim(); applyFilter(); renderOrdersList(); }, 200);
  });
  document.querySelectorAll('.filter-tabs .tab').forEach(tab =>
    tab.addEventListener('click', () => { state.filter=tab.dataset.filter; applyFilter(); renderOrdersList(); }));
}

// ── 初期化 ─────────────────────────────────────────────────
function init() {
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('./sw.js').catch(()=>{});
  wireEvents();
  navigate(loadStorage().clientId ? 'auth' : 'setup');
  const poll = setInterval(() => {
    if (window.gapi && window.google?.accounts?.oauth2) {
      clearInterval(poll); onGapiLoaded(); onGisLoaded();
    }
  }, 100);
}
document.addEventListener('DOMContentLoaded', init);
