// 运行在网页自身环境里：旁听网页收到的“视频列表”数据，转交给 content.js。不发起任何额外请求。
(() => {
  if (window.__vcxHooked) return;
  window.__vcxHooked = true;
  const isList = (u) => /\/post\/post_list(\?|$)/.test(String(u || ''));
  const send = (json) => { try { window.top.postMessage({ __vcx: 1, json }, location.origin === 'null' ? '*' : location.origin); } catch (e) {} };

  const origFetch = window.fetch;
  if (origFetch) {
    window.fetch = function (...args) {
      const p = origFetch.apply(this, args);
      p.then((r) => { if (isList(r.url)) r.clone().json().then(send).catch(() => {}); }).catch(() => {});
      return p;
    };
  }

  const open = XMLHttpRequest.prototype.open;
  XMLHttpRequest.prototype.open = function (method, url, ...rest) {
    if (isList(url)) {
      this.addEventListener('load', () => {
        try {
          const t = this.responseType;
          if (t === 'json') send(this.response);
          else if (t === '' || t === 'text') send(JSON.parse(this.responseText));
        } catch (e) {}
      });
    }
    return open.call(this, method, url, ...rest);
  };
})();
