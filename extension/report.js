// 把视频列表整理成 Excel：汇总 + 全部视频
(() => {
  const FONT = 'Arial', NAVY = 'FF1F3864';
  const clean = (s) => String(s || '').replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, '');
  const num = (v) => (v === undefined || v === null || v === '' ? null : Number(v));
  const pad = (n) => String(n).padStart(2, '0');
  // Excel 单元格里的日期按 UTC 写入，所以用“本地时间的数字”构造，显示才和后台一致
  const asCell = (d) => new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds()));
  const serial = (d) => asCell(d).getTime() / 86400000 + 25569;

  function title(x) {
    const st = (x.desc && x.desc.shortTitle) || [];
    const t = clean(st[0] && st[0].shortTitle).trim();
    if (t) return t;
    const d = clean(x.desc && x.desc.description);
    const s = d.replace(/[#@]\S*/g, '').replace(/\s+/g, ' ').trim();
    return s || d.replace(/\s+/g, ' ').trim() || '（无标题）';
  }

  function toRow(x) {
    const t = new Date(x.createTime * 1000);
    const media = (x.desc && x.desc.media && x.desc.media[0]) || {};
    const rate = num(x.fullPlayRate), info = x.disableInfo || {}, note = [];
    if (rate !== null && rate > 1) note.push('完播率超过100%，是平台返回的原始值，疑似异常');
    if (rate === null) note.push('平台未返回完播率等指标');
    if ((info.disableShowWording || '').includes('限制传播')) note.push('此视频已限制传播');
    if (!media.videoPlayLen) note.push('平台未返回时长');
    return {
      title: title(x), t, year: t.getFullYear(), vis: x.visibleType === 3 ? '仅自己可见' : '公开',
      len: media.videoPlayLen || null, read: num(x.readCount) || 0, yest: num(x.yesterdayReadCount),
      heart: num(x.likeCount) || 0, thumb: num(x.favCount) || 0, cmt: num(x.commentCount) || 0, fwd: num(x.forwardCount) || 0,
      fol: num(x.followCount) || 0, unf: num(x.cancelFollowCount) || 0, rate, avg: num(x.avgPlayTimeSec), flip: num(x.fastFlipRate),
      desc: clean(x.desc && x.desc.description), note: note.join('；'),
    };
  }

  const sum = (a, k) => a.reduce((s, x) => s + x[k], 0);
  function median(a) {
    const s = [...a].sort((p, q) => p - q), m = s.length >> 1;
    return s.length ? (s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2) : 0;
  }

  async function build(items, incompleteTotal) {
    const rows = items.map(toRow).sort((a, b) => b.t - a.t);
    const n = rows.length, tot = sum(rows, 'read');
    const wb = new window.ExcelJS.Workbook();
    const s = wb.addWorksheet('汇总', { views: [{ showGridLines: false }] });
    const d = wb.addWorksheet('全部视频', { views: [{ state: 'frozen', xSplit: 2, ySplit: 1 }] });
    const base = { name: FONT, size: 10 };
    const headStyle = (c) => {
      c.font = { ...base, bold: true, color: { argb: 'FFFFFFFF' } };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } };
      c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    };

    // ---- 全部视频 ----
    const F = { n: '#,##0', p: '0.0%', d: '0.0', t: 'yyyy-mm-dd hh:mm' };
    const cols = [
      ['序号', 6], ['标题', 48], ['发布时间', 16, F.t, 'c'], ['年份', 7, null, 'c'], ['可见范围', 11, null, 'c'], ['时长(秒)', 8, F.n], ['播放量', 12, F.n],
      ['昨日播放', 10, F.n], ['爱心', 10, F.n], ['拇指赞', 10, F.n], ['评论', 9, F.n], ['转发', 10, F.n], ['新增关注', 10, F.n], ['取消关注', 10, F.n],
      ['完播率', 9, F.p], ['平均播放时长(秒)', 11, F.d], ['快速划走率', 10, F.p], ['完整文案', 60], ['备注', 30],
    ];
    d.columns = cols.map(([, width]) => ({ width }));
    d.getRow(1).values = cols.map((c) => c[0]); d.getRow(1).height = 32; d.getRow(1).eachCell(headStyle);
    rows.forEach((x, i) => {
      const r = d.getRow(i + 2);
      r.values = [i + 1, x.title, asCell(x.t), x.year, x.vis, x.len, x.read, x.yest, x.heart, x.thumb, x.cmt, x.fwd, x.fol, x.unf, x.rate, x.avg, x.flip, x.desc, x.note];
      cols.forEach((c, j) => {
        const cell = r.getCell(j + 1);
        cell.font = base;
        if (c[2]) cell.numFmt = c[2];
        cell.alignment = { vertical: 'middle', ...(c[3] === 'c' || j === 0 ? { horizontal: 'center' } : {}) };
      });
    });
    d.autoFilter = `A1:S${n + 1}`;
    const rg = (c) => `全部视频!$${c}$2:$${c}$${n + 1}`;

    // ---- 汇总 ----
    s.columns = [{ width: 2 }, { width: 44 }, ...Array(8).fill({ width: 13 })];
    const put = (ref, v, opts = {}) => {
      const c = s.getCell(ref); c.value = v;
      c.font = { ...base, ...(opts.font || {}) };
      if (opts.fmt) c.numFmt = opts.fmt;
      c.alignment = { vertical: 'middle', ...(opts.align || {}) };
      return c;
    };
    const f = (ref, formula, result, fmt, align) => put(ref, { formula, result }, { fmt, align });
    const section = (r, text, last) => {
      s.mergeCells(`B${r}:${last}${r}`);
      const c = put(`B${r}`, text, { font: { bold: true, size: 12, color: { argb: NAVY } } });
      c.border = { bottom: { style: 'medium', color: { argb: NAVY } } };
    };
    const header = (r, list) => list.forEach((h, i) => { const c = s.getCell(r, 2 + i); c.value = h; headStyle(c); });

    const now = new Date();
    put('B1', '视频号数据统计', { font: { bold: true, size: 16, color: { argb: NAVY } } });
    put('B2', `数据时间：${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())} · 来源：视频号助手「视频管理」`, { font: { color: { argb: 'FF595959' } } });
    let r = 3;
    if (incompleteTotal) { put('B3', `注意：这份数据不完整，只有 ${n} / ${incompleteTotal} 条。`, { font: { bold: true, color: { argb: 'FFC00000' } } }); }
    r = 4; section(r, '总览', 'E'); r++;
    const totalRow = r + 1;
    [
      ['视频总数', `COUNTA(${rg('B')})`, n], ['累计播放量', `SUM(${rg('G')})`, tot],
      ['单条平均播放', `AVERAGE(${rg('G')})`, n ? tot / n : 0], ['单条播放中位数', `MEDIAN(${rg('G')})`, median(rows.map((x) => x.read))],
      ['累计爱心', `SUM(${rg('I')})`, sum(rows, 'heart')], ['累计拇指赞', `SUM(${rg('J')})`, sum(rows, 'thumb')],
      ['累计评论', `SUM(${rg('K')})`, sum(rows, 'cmt')], ['累计转发', `SUM(${rg('L')})`, sum(rows, 'fwd')],
      ['累计新增关注', `SUM(${rg('M')})`, sum(rows, 'fol')], ['累计取消关注', `SUM(${rg('N')})`, sum(rows, 'unf')],
    ].forEach(([l, fm, v]) => { put(`B${r}`, l); f(`C${r}`, fm, v, F.n); r++; });

    const group = (name, first, col, key, keys) => {
      r++; section(r, name, 'G'); r++; header(r, [first, '视频数', '播放量', '播放占比', '单条平均播放', '累计转发']); r++;
      keys.forEach((k) => {
        const sel = rows.filter((x) => x[key] === k), v = sum(sel, 'read');
        put(`B${r}`, k, { align: { horizontal: 'left' } });
        f(`C${r}`, `COUNTIFS(${rg(col)},$B${r})`, sel.length, F.n);
        f(`D${r}`, `SUMIFS(${rg('G')},${rg(col)},$B${r})`, v, F.n);
        f(`E${r}`, `IF($C$${totalRow}=0,0,D${r}/$C$${totalRow})`, tot ? v / tot : 0, F.p);
        f(`F${r}`, `IF(C${r}=0,0,D${r}/C${r})`, sel.length ? v / sel.length : 0, F.n);
        f(`G${r}`, `SUMIFS(${rg('L')},${rg(col)},$B${r})`, sum(sel, 'fwd'), F.n);
        r++;
      });
    };
    group('按可见范围', '可见范围', 'E', 'vis', ['公开', '仅自己可见']);
    group('按发布年份', '年份', 'D', 'year', [...new Set(rows.map((x) => x.year))].sort((a, b) => b - a));

    // 播放量排行：同播放量时公式取不到第二条，所以排行里的明细直接写数值
    const top = [...rows].sort((a, b) => b.read - a.read).slice(0, 20);
    r++; section(r, `播放量 Top ${top.length}`, 'H'); r++; header(r, ['标题', '播放量', '发布时间', '可见范围', '爱心', '拇指赞', '转发']); r++;
    top.forEach((x) => {
      put(`B${r}`, x.title); put(`C${r}`, x.read, { fmt: F.n }); put(`D${r}`, asCell(x.t), { fmt: 'yyyy-mm-dd', align: { horizontal: 'center' } });
      put(`E${r}`, x.vis, { align: { horizontal: 'center' } }); put(`F${r}`, x.heart, { fmt: F.n }); put(`G${r}`, x.thumb, { fmt: F.n }); put(`H${r}`, x.fwd, { fmt: F.n });
      r++;
    });

    const priv = rows.filter((x) => x.vis === '仅自己可见'), odd = rows.filter((x) => x.rate !== null && x.rate > 1).length;
    r++; section(r, '说明', 'H'); r++;
    const notes = [
      ['爱心、拇指赞对应后台列表页的两个图标（接口字段 likeCount、favCount），所以没有叫“收藏”。', 20],
      ['标题：优先用发布时填写的短标题；没填的，用文案去掉 #话题 和 @提及 后的文字。完整原文在「全部视频」的「完整文案」列。', 32],
      ['昨日播放、完播率、平均播放时长、快速划走率、新增关注、取消关注不在后台列表页显示，名称按接口字段名翻译，未和后台逐项核对。', 32],
    ];
    if (priv.length) notes.push([`累计数据包含 ${priv.length} 条“仅自己可见”的视频（共 ${sum(priv, 'read').toLocaleString('en-US')} 播放）。只看公开视频，可在「全部视频」按「可见范围」筛选。`, 32]);
    if (odd) notes.push([`有 ${odd} 条视频的完播率超过 100%（备注列已标出），是平台返回的原始值，做完播率分析时建议剔除。`, 20]);
    notes.forEach(([t, h]) => {
      s.mergeCells(`B${r}:H${r}`);
      put(`B${r}`, t, { font: { color: { argb: 'FF595959' } }, align: { wrapText: true, vertical: 'top' } });
      s.getRow(r).height = h; r++;
    });
    return wb.xlsx.writeBuffer();
  }

  function fileName(incomplete) {
    const d = new Date();
    return `视频号视频数据_${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${incomplete ? '_不完整' : ''}.xlsx`;
  }
  window.VcxReport = { build, fileName };
})();
