// Build-your-box motion layer.
// The "Any 4 for $20" banner sits under the menu carousel and turns into the box
// as tarts are added: each tart photo flies into the next slot, the slot squashes on
// landing, and the fourth tart seals the box. It only listens to the existing + / -
// buttons, and re-syncs from the card quantities so edits made in the bag stay correct.
(function boot() {
  const first = document.querySelector('.tt-mpop img');
  const src = first && first.getAttribute('src');
  if (!document.querySelector('.tt-mtrk') || !src || src.includes('{{')) return setTimeout(boot, 100);

  const SPRING = 'cubic-bezier(.34,1.56,.64,1)', OUT = 'cubic-bezier(.22,1,.36,1)';
  const reduce = matchMedia('(prefers-reduced-motion: reduce)').matches;
  const menu = document.getElementById('tt-menu');
  const promo = menu.children[1], track = document.querySelector('.tt-mtrk');
  if (!promo || !track) return;

  let box = [];      // tarts in the unfinished box: { img, price }
  let sealed = [];   // finished boxes, each an array of 4 tarts
  let pending = 0;   // tarts currently in the air
  let sealing = false;
  const BOX_PRICE = 20, MAX_ROWS = 2;
  const saved = b => Math.max(0, b.reduce((a, t) => a + t.price, 0) - BOX_PRICE);
  const money = n => '$' + (Math.round(n * 100) / 100);

  const css = document.createElement('style');
  css.textContent = `
  .tt-mcard button { transition: transform .25s ${SPRING}; }
  .tt-mcard button:active { transform: scale(.82); }
  .tt-step { animation: ttPop .4s ${SPRING}; }
  .bx-bump { animation: bxBump .45s ${SPRING}; }
  @keyframes bxBump { 30% { transform: scale(1.18); } }
  .bx-squash { animation: bxSquash .5s ${SPRING}; }
  @keyframes bxSquash { 30% { transform: scale(1.2,.84); } 60% { transform: scale(.94,1.06); } }
  .bx-ring { position:absolute; inset:0; border-radius:50%; border:2px solid var(--accent); pointer-events:none; animation: bxRing .6s ease-out forwards; }
  @keyframes bxRing { from { transform:scale(.6); opacity:.9 } to { transform:scale(1.9); opacity:0 } }
  .bx-slot { width:100%; aspect-ratio:1; flex:none; border-radius:50%; position:relative; background-repeat:no-repeat; background-size:contain; background-position:center; }
  .bx-slot.em { border:2px dashed color-mix(in srgb, var(--accent) 45%, transparent); }
  .bx-slot.next { animation: bxBreathe 1.6s ease-in-out infinite; }
  @keyframes bxBreathe { 50% { transform:scale(1.1); } }
  .bx-slot.full { filter:none; }
  .bx-txt { font-family:var(--font-body); font-weight:700; font-size:15px; color:var(--accent); line-height:1.25; white-space:nowrap; }
  .bx-txt small { display:block; font-weight:500; font-size:11.5px; opacity:.8; }
  .bx-num { display:inline-block; }
  .bx-num.roll { animation: bxRoll .4s ${OUT}; }
  @keyframes bxRoll { from { transform:translateY(60%); opacity:0 } }
  .bx-crumb { position:fixed; width:7px; height:7px; border-radius:2px; pointer-events:none; z-index:200; }
  .bx-d { display:grid !important; position:sticky !important; bottom:12px; z-index:30; overflow:hidden; margin:18px 0 0 !important; box-shadow:none; transition: background-color .5s, color .4s; }
  .bx-d > * { grid-area:1/1; transition: opacity .35s, transform .55s ${OUT}; }
  .bx-d .bx-d-box { display:flex; flex-direction:column; gap:8px; opacity:0; transform:translateY(12px); pointer-events:none; cursor:pointer; }
  .bx-d.show .bx-d-promo { opacity:0; transform:translateY(-12px); }
  .bx-d.show .bx-d-box { opacity:1; transform:none; pointer-events:auto; }
  .bx-slots { display:grid; grid-template-columns:repeat(4, minmax(0, 36px)); gap:6px; margin-left:auto; flex:1 1 auto; min-width:0; justify-content:end; }
  .bx-act { display:flex; align-items:center; gap:12px; min-width:0; }
  .bx-done { display:flex; flex-direction:column; gap:6px; }
  .bx-done:empty { display:none; }
  .bx-drow { display:flex; align-items:center; gap:6px; padding-top:6px; border-top:1px dashed color-mix(in srgb, var(--accent) 30%, transparent); font-family:var(--font-body); font-size:12.5px; font-weight:700; color:var(--accent); white-space:nowrap; overflow:hidden; }
  .bx-drow .bx-mini { display:flex; gap:2px; flex:none; }
  .bx-drow .bx-mini i { width:16px; height:16px; border-radius:50%; background-size:contain; background-position:center; background-repeat:no-repeat; }
  .bx-drow .lbl { flex:1; text-align:left; min-width:0; overflow:hidden; text-overflow:ellipsis; }
  .bx-drow .sv { flex:none; padding:2px 8px; border-radius:999px; background:var(--accent); color:var(--accent-text); font-size:11px; }
  .bx-drow.in { animation: bxRowIn .55s ${SPRING}; }
  @keyframes bxRowIn { from { max-height:0; opacity:0; transform:translateY(14px) scale(.96); } to { max-height:36px; opacity:1; } }
  .bx-d.win { background:var(--accent) !important; }
  .bx-d.win .bx-drow { color:var(--accent-text); border-color:rgba(255,255,255,.3); }
  .bx-d.win .bx-drow .sv { background:var(--accent-text); color:var(--accent); }
  .bx-d.win .bx-txt { color:var(--accent-text); }
  .bx-d.win .bx-slot.em { border-color:rgba(255,255,255,.5); }
  @media (prefers-reduced-motion: reduce) { .bx-slot.next { animation:none } .bx-d, .bx-d > * { transition:none } }
  `;
  document.head.appendChild(css);

  // Move the promo banner under the carousel and give it a second face: the box.
  const inner = document.createElement('div'); inner.className = 'bx-d-promo';
  while (promo.firstChild) inner.appendChild(promo.firstChild);
  promo.appendChild(inner);
  const face = document.createElement('div'); face.className = 'bx-d-box';
  face.innerHTML = '<div class="bx-act"><div class="bx-txt"></div><div class="bx-slots"></div></div><div class="bx-done"></div>';
  promo.appendChild(face);
  promo.classList.add('bx-d');
  track.after(promo);
  const txtEl = face.querySelector('.bx-txt'), slotsEl = face.querySelector('.bx-slots'), doneEl = face.querySelector('.bx-done');

  const bagBtn = () => document.querySelector('nav button:last-child');
  const isZh = () => /购物车/.test((bagBtn() || {}).innerText || '');
  const T = (zh, en) => isZh() ? zh : en;
  face.addEventListener('click', () => bagBtn() && bagBtn().click());

  let lastLeft = null;
  function render(animIdx, animRow) {
    slotsEl.innerHTML = '';
    for (let i = 0; i < 4; i++) {
      const el = document.createElement('div');
      el.className = 'bx-slot ' + (box[i] ? 'full' : 'em' + (i === box.length ? ' next' : ''));
      if (box[i]) el.style.backgroundImage = box[i].img;
      slotsEl.appendChild(el);
    }
    // sealed boxes collapse into compact rows; older ones fold into one summary line
    doneEl.innerHTML = '';
    const hidden = Math.max(0, sealed.length - MAX_ROWS);
    if (hidden) {
      const r = document.createElement('div'); r.className = 'bx-drow';
      const sv = sealed.slice(0, hidden).reduce((a, b) => a + saved(b), 0);
      r.innerHTML = '<span class="lbl">' + T('另外 ' + hidden + ' 盒已完成', hidden + ' earlier box' + (hidden > 1 ? 'es' : '')) + '</span><span class="sv">' + T('已省 ', 'Saved ') + money(sv) + '</span>';
      doneEl.appendChild(r);
    }
    sealed.forEach((bx, n) => {
      if (n < hidden) return;
      const r = document.createElement('div'); r.className = 'bx-drow' + (animRow && n === sealed.length - 1 ? ' in' : '');
      r.innerHTML = '<span class="bx-mini"></span>'
        + '<span class="lbl">' + T('第 ' + (n + 1) + ' 盒 \u00b7 ', 'Box ' + (n + 1) + ' \u00b7 ') + money(Math.min(BOX_PRICE, bx.reduce((a, t) => a + t.price, 0))) + '</span>'
        + '<span class="sv">' + T('已省 ', 'Saved ') + money(saved(bx)) + '</span>';
      bx.forEach(t => { const i = document.createElement('i'); i.style.backgroundImage = t.img; r.firstChild.appendChild(i); });
      doneEl.appendChild(r);
    });
    const left = 4 - box.length;
    const changed = left !== lastLeft; lastLeft = left;
    let html;
    if (box.length === 4) html = T('这盒已满！', 'Box complete!') + '<small>' + T('已省 ', 'You saved ') + money(saved(box)) + '</small>';
    else html = T('任选四颗 $20', 'Any 4 for $20') + (sealed.length && !box.length ? '<small>' + T('开始下一盒', 'Start your next box') + '</small>' : '');
    txtEl.innerHTML = html;
    promo.classList.toggle('show', box.length > 0 || sealed.length > 0 || pending > 0);
    if (animIdx != null && slotsEl.children[animIdx]) {
      const el = slotsEl.children[animIdx];
      el.classList.add('bx-squash');
      const r = document.createElement('div'); r.className = 'bx-ring'; el.appendChild(r);
    }
  }

  function fly(fromEl, img, toEl, cb) {
    let ended = false;
    const finish = () => { if (ended) return; ended = true; f && f.remove(); cb(); };
    if (reduce || !fromEl || !toEl) { var f = null; return finish(); }
    const a = fromEl.getBoundingClientRect(), b = toEl.getBoundingClientRect();
    var f = document.createElement('div');
    Object.assign(f.style, { position:'fixed', left:a.left + 'px', top:a.top + 'px', width:a.width + 'px', height:a.height + 'px', zIndex:150, pointerEvents:'none',
      backgroundImage:img, backgroundSize:'contain', backgroundPosition:'center', backgroundRepeat:'no-repeat', filter:'drop-shadow(0 10px 12px rgba(60,35,5,.35))' });
    document.body.appendChild(f);
    const dx = b.left + b.width / 2 - (a.left + a.width / 2), dy = b.top + b.height / 2 - (a.top + a.height / 2), sc = b.width / a.width;
    const lift = dy < 0 ? -40 : -110;
    f.animate([
      { transform:'translate(0,0) scale(1) rotate(0)' },
      { transform:'translate(' + dx * .45 + 'px,' + (dy * .45 + lift) + 'px) scale(1.15) rotate(160deg)', offset:.45 },
      { transform:'translate(' + dx + 'px,' + dy + 'px) scale(' + sc + ') rotate(330deg)' }
    ], { duration:650, easing:'cubic-bezier(.45,0,.25,1)' }).onfinish = finish;
    setTimeout(finish, 900); // background tabs throttle animations; never leave the box waiting
  }

  function crumbs(el) {
    if (reduce) return;
    const r = el.getBoundingClientRect(), x = r.left + r.width / 2, y = r.top + r.height / 2;
    const cols = ['#C0792A', '#E7A93A', '#EABE83', '#D3A263', '#F2E3C6'];
    for (let i = 0; i < 26; i++) {
      const d = document.createElement('div'); d.className = 'bx-crumb'; d.style.background = cols[i % 5]; d.style.left = x + 'px'; d.style.top = y + 'px';
      document.body.appendChild(d);
      const ang = Math.random() * Math.PI * 2, rad = 50 + Math.random() * 90;
      d.animate([{ transform:'translate(0,0)', opacity:1 }, { transform:'translate(' + Math.cos(ang) * rad + 'px,' + (Math.sin(ang) * rad - 40) + 'px) rotate(' + Math.random() * 500 + 'deg)', opacity:0 }],
        { duration:800 + Math.random() * 400, easing:'cubic-bezier(.2,.7,.4,1)' }).onfinish = () => d.remove();
    }
  }

  function bumpBag() { const b = bagBtn(); if (b) { b.classList.remove('bx-bump'); void b.offsetWidth; b.classList.add('bx-bump'); } }

  function seal() {
    sealing = true;
    navigator.vibrate && navigator.vibrate([10, 40, 18]);
    promo.classList.add('win');
    crumbs(slotsEl);
    // hold the full box for a beat, then collapse it into a row and open a fresh one
    setTimeout(() => {
      sealed.push(box); box = [];
      promo.classList.remove('win');
      render(null, true);
      sealing = false;
    }, 1100);
  }

  function priceOf(card) { const m = card.innerText.match(/\$\s*(\d+(?:\.\d+)?)/); return m ? parseFloat(m[1]) : 6; }
  function unitOf(card) { const im = card.querySelector('.tt-mpop img'); return { img: 'url("' + (im.currentSrc || im.src) + '")', price: priceOf(card) }; }

  function add(card) {
    const thumb = card.querySelector('.tt-mpop img'); if (!thumb) return;
    const unit = unitOf(card);
    if (sealing) { setTimeout(() => add(card), 1200); return; }
    const idx = box.length;
    navigator.vibrate && navigator.vibrate(8);
    const pop = card.querySelector('.tt-mpop');
    if (pop && !reduce) pop.animate([{ transform:'scale(1)' }, { transform:'scale(.9) rotate(-6deg)' }, { transform:'scale(1.04) rotate(3deg)' }, { transform:'scale(1)' }], { duration:480, easing:SPRING });
    pending++;
    const wasShown = promo.classList.contains('show');
    setTimeout(() => {
      render();
      fly(thumb, unit.img, slotsEl.children[idx], () => {
        pending--; box.push(unit); render(idx); bumpBag();
        if (box.length === 4) seal();
      });
    }, wasShown ? 0 : 250);
  }

  function remove() {
    if (box.length) box.pop();
    else if (sealed.length) { box = sealed.pop(); box.pop(); }
    render();
  }

  // The card steppers are the source of truth. If the total drifts (bag edits, language switch), rebuild.
  let lastZh = isZh();
  function sync() {
    if (pending || sealing) return;
    if (isZh() !== lastZh) { lastZh = isZh(); render(); }
    const units = [];
    document.querySelectorAll('.tt-mcard').forEach(c => {
      const q = parseInt((c.querySelector('.tt-step span') || {}).textContent, 10) || 0;
      if (!q) return;
      const u = unitOf(c);
      for (let i = 0; i < q; i++) units.push(u);
    });
    if (units.length === sealed.length * 4 + box.length) return;
    sealed = [];
    for (let i = 0; i + 4 <= units.length; i += 4) sealed.push(units.slice(i, i + 4));
    box = units.slice(sealed.length * 4);
    render();
  }

  document.addEventListener('click', e => {
    const btn = e.target.closest('.tt-mcard button[aria-label="add"], .tt-mcard button[aria-label="remove"]');
    if (!btn) return;
    if (btn.getAttribute('aria-label') === 'add') add(btn.closest('.tt-mcard')); else remove();
  }, true);
  setInterval(sync, 500);
  render();
})();
