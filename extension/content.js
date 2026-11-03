// 页面右下角的导出面板 + 自动翻页
(() => {
  const LIST_URL = 'https://channels.weixin.qq.com/platform/post/list';
  const FLAG = 'vcx_running';
  const isList = () => location.pathname.startsWith('/platform/post/list');
  const isLogin = () => /login/i.test(location.pathname);
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  const rows = new Map();
  let total = null, pageCount = 0, running = false, stopRequested = false, ui = null, finished = false, lastError = '';

  // ---------- 收数据 ----------
  window.addEventListener('message', (e) => {
    if (e.origin !== location.origin || !e.data || e.data.__vcx !== 1 || !running) return;
    const j = e.data.json;
    const d = j && j.data;
    if (!d || !Array.isArray(d.list)) {
      lastError = (j && j.errMsg) || '平台没有返回视频列表';
      pageCount++;
      return;
    }
    d.list.forEach((x) => { const id = x.objectId || x.exportId; if (id) rows.set(id, x); });
    if (typeof d.totalCount === 'number') total = d.totalCount;
    lastError = '';
    pageCount++;
    render();
  });

  // ---------- 找“下一页”：网页可能把内容放在 iframe 或 Shadow DOM 里，要一层层进去找 ----------
  function findNext(root) {
    const all = root.querySelectorAll('*');
    for (const el of all) {
      if ((el.tagName === 'A' || el.tagName === 'BUTTON') && el.textContent.trim() === '下一页' &&
          !/disabled/.test(el.className || '') && !el.disabled && el.getClientRects().length) return el;
      if (el.shadowRoot) { const f = findNext(el.shadowRoot); if (f) return f; }
      if (el.tagName === 'IFRAME') {
        try { if (el.contentDocument) { const f = findNext(el.contentDocument); if (f) return f; } } catch (e) {}
      }
    }
    return null;
  }

  async function waitForPage(before, ms) {
    const t = Date.now();
    while (pageCount === before && Date.now() - t < ms && !stopRequested) await sleep(300);
    return pageCount !== before;
  }

  // ---------- 主流程 ----------
  async function run() {
    running = true; render();
    if (!(await waitForPage(0, 30000))) return finish('没有读到视频列表。请刷新页面重试；如果一直这样，可能是视频号改版了。');
    while (!stopRequested) {
      if (lastError) return finish('平台返回了错误：' + lastError);
      if (total !== null && rows.size >= total) return finish();
      await sleep(3000 + Math.random() * 2000); // 翻页放慢一点，避免操作过于频繁
      if (stopRequested) break;
      let next = findNext(document);
      if (!next) { await sleep(3000); next = findNext(document); }
      if (!next) return finish(total !== null && rows.size < total ? '翻到最后一页了，但条数比平台显示的总数少。' : '');
      const before = pageCount;
      next.click();
      if (!(await waitForPage(before, 20000))) {
        if (stopRequested) break;
        const again = findNext(document);
        if (again) again.click();
        if (!(await waitForPage(before, 20000))) return finish('翻页后迟迟没有新数据，可能网络不好或操作被平台限制了。');
      }
    }
    finish('你手动停止了导出。');
  }

  function finish(msg = '') {
    running = false; finished = true; lastError = msg;
    sessionStorage.removeItem(FLAG);
    render();
    if (rows.size) download();
  }

  async function download() {
    setBtnBusy(true);
    try {
      const incomplete = total !== null && rows.size < total ? total : null;
      const buf = await window.VcxReport.build([...rows.values()], incomplete);
      const url = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      const a = document.createElement('a');
      a.href = url; a.download = window.VcxReport.fileName(incomplete);
      document.documentElement.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 60000);
    } catch (e) {
      lastError = '生成 Excel 失败：' + e.message;
    }
    setBtnBusy(false); render();
  }

  // ---------- 面板 ----------
  function buildUI() {
    const host = document.createElement('div');
    host.style.cssText = 'all:initial;position:fixed;right:12px;bottom:12px;z-index:2147483647';
    const sh = host.attachShadow({ mode: 'open' });
    sh.innerHTML = `<style>
      *{box-sizing:border-box;font-family:-apple-system,"PingFang SC","Microsoft YaHei",sans-serif}
      .box{width:280px;max-width:calc(100vw - 24px);max-height:calc(100vh - 24px);overflow:auto;background:#fff;color:#1f2329;border-radius:12px;box-shadow:0 6px 28px rgba(0,0,0,.2);padding:14px 16px;font-size:13px;line-height:1.55}
      .top{display:flex;justify-content:space-between;align-items:center;font-weight:600;font-size:14px}
      .x{border:0;background:none;color:#8a8f99;font-size:18px;cursor:pointer;padding:0 2px;line-height:1}
      .msg{margin:8px 0 10px;color:#4e5563;word-break:break-word}
      .err{color:#c0392b}
      .bar{height:6px;background:#eee;border-radius:3px;overflow:hidden;margin-bottom:10px}
      .bar i{display:block;height:100%;width:0;background:#fa9d3b;transition:width .3s}
      button.go{width:100%;border:0;border-radius:8px;background:#fa9d3b;color:#fff;font-size:14px;font-weight:600;padding:9px;cursor:pointer}
      button.go:disabled{opacity:.6;cursor:default}
      button.go.ghost{background:#f2f3f5;color:#4e5563}
      .pill{display:none;border:0;border-radius:20px;background:#fa9d3b;color:#fff;font-size:13px;font-weight:600;padding:9px 14px;cursor:pointer;box-shadow:0 4px 16px rgba(0,0,0,.2)}
    </style>
    <div class="box"><div class="top"><span>视频号数据导出</span><button class="x" title="收起">×</button></div>
      <div class="msg"></div><div class="err msg"></div><div class="bar"><i></i></div><button class="go"></button></div>
    <button class="pill">导出数据</button>`;
    const $ = (s) => sh.querySelector(s);
    ui = { box: $('.box'), msg: $('.msg:not(.err)'), err: $('.err'), bar: $('.bar'), fill: $('.bar i'), go: $('.go'), pill: $('.pill'), busy: false };
    $('.x').onclick = () => { ui.box.style.display = 'none'; ui.pill.style.display = 'block'; };
    ui.pill.onclick = () => { ui.box.style.display = 'block'; ui.pill.style.display = 'none'; };
    ui.go.onclick = () => {
      if (running) { stopRequested = true; ui.go.disabled = true; ui.go.textContent = '正在停止…'; return; }
      if (finished && rows.size && !ui.restart) { download(); return; }
      sessionStorage.setItem(FLAG, '1');
      if (isList()) location.reload(); else location.href = LIST_URL;
    };
    document.documentElement.appendChild(host);
  }

  function setBtnBusy(b) { if (ui) ui.busy = b; }

  function render() {
    if (!ui) return;
    const pct = total ? Math.min(100, Math.round((rows.size / total) * 100)) : 0;
    ui.fill.style.width = pct + '%';
    ui.err.textContent = lastError; ui.err.style.display = lastError ? 'block' : 'none';
    ui.go.className = 'go'; ui.go.disabled = !!ui.busy;
    if (isLogin() || sessionStorage.getItem(FLAG) && !isList() && !running) {
      ui.bar.style.display = 'none';
      ui.msg.textContent = isLogin() ? '请先用微信扫码登录。登录后可以点下面的按钮导出。' : '正在打开视频管理页…';
      ui.go.textContent = '开始导出'; ui.go.disabled = isLogin() && !sessionStorage.getItem(FLAG) ? false : true;
      if (isLogin() && sessionStorage.getItem(FLAG)) ui.msg.textContent = '请用微信扫码登录，登录后会自动开始导出。';
      return;
    }
    if (running) {
      ui.bar.style.display = 'block';
      const left = total ? Math.ceil(((total - rows.size) / 20) * 5 / 60) : null;
      ui.msg.textContent = total === null ? '正在读取第一页…' : `已获取 ${rows.size} / ${total} 条` + (left ? `，大约还要 ${left} 分钟。请停留在这个标签页，不要关闭或刷新；切走可能会变慢。` : '');
      ui.go.className = 'go ghost'; ui.go.textContent = stopRequested ? '正在停止…' : '停止并导出已获取的';
      return;
    }
    if (finished) {
      ui.bar.style.display = 'block';
      const part = total !== null && rows.size < total;
      ui.msg.textContent = rows.size ? `${part ? '只导出了部分数据：' : '导出完成：'}${rows.size}${total !== null ? ' / ' + total : ''} 条。Excel 已保存到浏览器的下载文件夹。` : '没有获取到数据。';
      ui.go.textContent = rows.size ? (ui.busy ? '正在生成…' : '再下载一次') : '重新开始';
      return;
    }
    ui.bar.style.display = 'none';
    ui.msg.textContent = '导出所有视频的标题和播放数据。数据只保存在你的电脑上。';
    ui.go.textContent = '开始导出';
  }

  // ---------- 启动 ----------
  const flagged = sessionStorage.getItem(FLAG) === '1';
  if (flagged && isList()) running = true;     // 尽早开始听，避免漏掉第一页

  function boot() {
    buildUI(); render();
    if (flagged && isList()) { run(); return; }
    // 扫码登录后可能落到别的页面：带上标记跳回视频管理页
    setInterval(() => {
      if (sessionStorage.getItem(FLAG) === '1' && !isLogin() && !isList() && !running) location.href = LIST_URL;
      else render();
    }, 2000);
  }
  if (document.body) boot(); else document.addEventListener('DOMContentLoaded', boot);
})();
