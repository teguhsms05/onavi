/**
 * Arc (uiarc.dev) untuk Onavi: port vanilla JS dari komponen gratis Arc. Gaya: static/css/arc.css.
 *
 *   Arc.enhance(root)            tombol −/+ (number-field) pada input angka bulat (bukan step="any")
 *   Arc.donut(el, options)       donut-chart: nilai aktif di tengah, legend menyorot & menyembunyikan segmen
 *   Arc.line(el, options)        line-chart: toggle seri, crosshair + tooltip, keyboard ← →
 *   Arc.bar(el, options)         bar-chart: angka utama = rata-rata saat diam, batang yang digeser saat menjelajah
 *
 * Geometri segmen donut (sector) mengikuti donut-chart Arc: celah antarsegmen sama lebarnya dari tepi dalam ke luar.
 */
(function () {
  "use strict";

  const NS = "http://www.w3.org/2000/svg";
  const TAU = Math.PI * 2;
  const REDUCED = window.matchMedia("(prefers-reduced-motion: reduce)");
  const CSS_D = window.CSS && CSS.supports && CSS.supports("d", 'path("M0 0")');

  function h(tag, attrs, parent, svg) {
    const node = svg ? document.createElementNS(NS, tag) : document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) {
      if (k === "text") node.textContent = v;
      else if (k === "style" && typeof v === "object") Object.assign(node.style, v);
      else node.setAttribute(k, v);
    }
    if (parent) parent.appendChild(node);
    return node;
  }
  const s = (tag, attrs, parent) => h(tag, attrs, parent, true);
  const n2 = (v) => (Math.round(v * 100) / 100).toString();
  const crisp = (v) => Math.round(v) + 0.5; // garis 1px tepat di piksel
  const setPath = (node, d, animate) => {
    if (CSS_D && animate) node.style.d = `path("${d}")`;
    else { node.style.removeProperty("d"); node.setAttribute("d", d); }
  };
  const pct = (v) => (v > 0 && v < 0.1 ? "<0,1" : v.toLocaleString("id-ID", { maximumFractionDigits: 1 })) + "%";
  const ease = (t) => 1 - Math.pow(1 - t, 3);

  function tween(ms, step) {
    if (REDUCED.matches) { step(1); return; }
    const t0 = performance.now();
    const frame = (now) => {
      const t = Math.min(1, (now - t0) / ms);
      step(ease(t));
      if (t < 1) requestAnimationFrame(frame);
    };
    requestAnimationFrame(frame);
  }

  /* Skala sumbu: 0 .. puncak yang dibulatkan, 4 garis bantu */
  function niceTop(max) {
    if (!(max > 0)) return 1;
    const raw = max / 4, mag = Math.pow(10, Math.floor(Math.log10(raw))), f = raw / mag;
    const step = (f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10) * mag;
    return step * 4;
  }

  /* ═══ Number field (−/+) ═══ */
  function enhance(root) {
    (root || document).querySelectorAll('input[type="number"]:not([step="any"]):not([data-arc-nf])').forEach((input) => {
      input.setAttribute("data-arc-nf", "");
      let group = input.parentElement;
      if (!group.classList.contains("input-group")) {
        group = h("div", { class: "input-group" });
        input.parentNode.insertBefore(group, input);
        group.appendChild(input);
      }
      group.classList.add("arc-nf");
      const minus = h("button", { type: "button", class: "arc-step", tabindex: "-1", "aria-label": "Kurangi", text: "−" });
      const plus = h("button", { type: "button", class: "arc-step", tabindex: "-1", "aria-label": "Tambah", text: "+" });
      group.insertBefore(minus, group.firstChild);
      group.appendChild(plus);
      const sync = () => {
        const v = parseFloat(input.value), min = parseFloat(input.min), max = parseFloat(input.max);
        minus.disabled = !isNaN(min) && !isNaN(v) && v <= min;
        plus.disabled = !isNaN(max) && !isNaN(v) && v >= max;
      };
      const step = (dir) => {
        if (input.value === "") input.value = input.min || 0;
        else if (dir > 0) input.stepUp(); else input.stepDown();
        input.dispatchEvent(new Event("input", { bubbles: true }));
        input.dispatchEvent(new Event("change", { bubbles: true }));
        sync();
      };
      minus.addEventListener("click", () => step(-1));
      plus.addEventListener("click", () => step(1));
      input.addEventListener("input", sync);
      sync();
    });
  }

  /* ═══ Donut ═══ */
  const LIFT = 2, MARGIN = 7, GAP = 2, CORNER = 0; // ujung segmen lurus seperti referensi; celah 2 px tetap memisahkan warna
  const OTHER_COLOR = "color-mix(in oklab, var(--foreground) 26%, var(--surface))";

  // Sektor cincin antara dua putaran (searah jarum jam dari atas), sudut membulat; port dari sector() Arc.
  function sector(c, R, r, t0, t1, gap, corner) {
    const turns = t1 - t0;
    if (turns <= 1e-6) return "";
    const at = (rad, a) => `${n2(c + rad * Math.cos(a))} ${n2(c + rad * Math.sin(a))}`;
    if (turns >= 1 - 1e-6) return `M${at(R, -Math.PI / 2)}A${n2(R)} ${n2(R)} 0 1 1 ${at(R, Math.PI / 2)}A${n2(R)} ${n2(R)} 0 1 1 ${at(R, -Math.PI / 2)}ZM${at(r, -Math.PI / 2)}A${n2(r)} ${n2(r)} 0 1 0 ${at(r, Math.PI / 2)}A${n2(r)} ${n2(r)} 0 1 0 ${at(r, -Math.PI / 2)}Z`;
    const p = Math.min(gap / 2, ((1 - turns) * TAU * R) / 2);
    const a0 = t0 * TAU - Math.PI / 2, a1 = t1 * TAU - Math.PI / 2, half = (a1 - a0) / 2, sn = Math.sin(Math.min(half, Math.PI / 2));
    if (sn * R <= p + 1e-6) return "";
    const side = (t, a, sign) => `${n2(c + t * Math.cos(a) - sign * p * Math.sin(a))} ${n2(c + t * Math.sin(a) + sign * p * Math.cos(a))}`;
    const band = (R - r) / 2;
    const ro = Math.max(0, Math.min(corner, band, sn >= 0.9999 ? Infinity : (sn * R - p) / (1 + sn)));
    const dO = Math.asin(Math.min(1, (p + ro) / (R - ro))), tO = Math.sqrt(Math.max(0, (R - ro) ** 2 - (p + ro) ** 2));
    const bigO = a1 - a0 - 2 * dO > Math.PI ? 1 : 0;
    let d = `M${side(tO, a0, 1)}`;
    if (ro > 0.01) d += `A${n2(ro)} ${n2(ro)} 0 0 1 ${at(R, a0 + dO)}`;
    d += `A${n2(R)} ${n2(R)} 0 ${bigO} 1 ${at(R, a1 - dO)}`;
    if (ro > 0.01) d += `A${n2(ro)} ${n2(ro)} 0 0 1 ${side(tO, a1, -1)}`;
    if (sn * r > p + 1e-6 || sn >= 0.9999) {
      const ri = Math.max(0, Math.min(corner, band, sn >= 0.9999 ? Infinity : (sn * r - p) / (1 - sn)));
      const dI = Math.asin(Math.min(1, (p + ri) / (r + ri))), tI = Math.sqrt(Math.max(0, (r + ri) ** 2 - (p + ri) ** 2));
      const bigI = a1 - a0 - 2 * dI > Math.PI ? 1 : 0;
      d += `L${side(tI, a1, -1)}`;
      if (ri > 0.01) d += `A${n2(ri)} ${n2(ri)} 0 0 1 ${at(r, a1 - dI)}`;
      d += `A${n2(r)} ${n2(r)} 0 ${bigI} 0 ${at(r, a0 + dI)}`;
      if (ri > 0.01) d += `A${n2(ri)} ${n2(ri)} 0 0 1 ${side(tI, a0, 1)}`;
    } else d += `L${at(p / sn, (a0 + a1) / 2)}`;
    return `${d}Z`;
  }

  function donut(container, opts) {
    const format = opts.format || ((v) => v.toLocaleString("id-ID"));
    const totalLabel = opts.totalLabel || "Total";
    const size = opts.size || 240, thickness = opts.thickness || 30;
    const c = size / 2, outer = c - MARGIN, inner = Math.max(8, outer - thickness), corner = Math.min(CORNER, thickness / 4);

    // Irisan kecil (< groupBelow) digabung jadi "Lainnya" bila ada dua atau lebih, agar tidak tampil sebagai garis tipis.
    const items = opts.items.filter((d) => d.value > 0);
    const grand = items.reduce((a, d) => a + d.value, 0);
    const small = items.filter((d) => d.value / grand < (opts.groupBelow ?? 0.04));
    let slices = items;
    if (small.length >= 2) {
      slices = items.filter((d) => !small.includes(d));
      slices.push({ key: "__other", label: opts.otherLabel || "Lainnya", value: small.reduce((a, d) => a + d.value, 0),
                    color: OTHER_COLOR, members: small.map((d) => d.label).join(", ") });
    }

    container.innerHTML = "";
    const fig = h("figure", { class: "arc-chart arc-donut", style: { margin: 0 } }, container);
    const wrap = h("div", { class: "arc-donut-inner" }, fig);
    const ring = h("div", { class: "arc-ring", style: { width: size + "px" } }, wrap);
    const svg = s("svg", { viewBox: `0 0 ${size} ${size}`, role: "img", "aria-label": opts.label || "Donut chart" }, ring);
    if (!slices.length) {
      s("circle", { class: "arc-track", cx: c, cy: c, r: (outer + inner) / 2, "stroke-width": outer - inner }, svg);
    }
    const center = h("div", { class: "arc-center", "aria-hidden": "true", style: { width: `${((inner * 2 - 20) / size) * 100}%` } }, ring);
    const cLabel = h("span", { class: "arc-center-label" }, center);
    const cValue = h("span", { class: "arc-center-value" }, center);
    const cMeta = h("span", { class: "arc-center-meta" }, center);
    const legend = h("ul", { class: "arc-legend" }, wrap);
    const live = h("p", { class: "arc-sr", "aria-live": "polite" }, fig);

    const hidden = new Set();
    let active = null;
    const cur = new Map(slices.map((d) => [d.key, [0, 0]])); // sudut yang sedang digambar
    const segs = new Map(), rows = new Map();

    slices.forEach((d) => {
      const seg = s("path", { class: "arc-seg", style: `--slice: ${d.color}` }, svg);
      segs.set(d.key, seg);
      const li = h("li", {}, legend);
      const row = h("button", { type: "button", class: "arc-row", style: `--slice: ${d.color}`, "aria-pressed": "true" }, li);
      h("span", { class: "arc-dot", "aria-hidden": "true" }, row);
      const text = h("span", { class: "arc-row-text" }, row);
      h("span", { class: "arc-row-label", text: d.label, title: d.label }, text);
      if (d.members) h("span", { class: "arc-row-members", text: d.members, title: d.members }, text);
      const val = h("span", { class: "arc-sr" }, text);
      const share = h("span", { class: "arc-row-share" }, row);
      rows.set(d.key, { row, val, share });
      row.addEventListener("mouseenter", () => setActive(d.key));
      row.addEventListener("mouseleave", () => setActive(null));
      row.addEventListener("focus", () => setActive(d.key));
      row.addEventListener("blur", () => setActive(null));
      row.addEventListener("click", () => toggle(d.key));
    });

    const visible = () => slices.filter((d) => !hidden.has(d.key));
    const totalVisible = () => visible().reduce((a, d) => a + d.value, 0);
    function targets() {
      const total = totalVisible(); let t = 0; const out = new Map();
      slices.forEach((d) => {
        const span = hidden.has(d.key) || !total ? 0 : d.value / total;
        out.set(d.key, [t, t + span]); t += span;
      });
      return out;
    }
    function draw() {
      slices.forEach((d) => {
        const [t0, t1] = cur.get(d.key);
        const R = outer + (active === d.key ? LIFT : 0);
        segs.get(d.key).setAttribute("d", sector(c, R, inner, t0, t1, GAP, corner));
        segs.get(d.key).toggleAttribute("data-dim", active !== null && active !== d.key);
      });
    }
    function readout() {
      const total = totalVisible();
      const d = active && slices.find((x) => x.key === active);
      if (d && !hidden.has(d.key)) {
        cLabel.textContent = d.label; cValue.textContent = format(d.value); cMeta.textContent = pct(total ? d.value / total * 100 : 0);
      } else {
        cLabel.textContent = totalLabel; cValue.textContent = format(total); cMeta.textContent = " ";
      }
      slices.forEach((x) => {
        const r = rows.get(x.key);
        r.val.textContent = `, ${format(x.value)}`;
        r.row.title = `${x.label}: ${format(x.value)}`;
        r.share.textContent = hidden.has(x.key) ? "–" : pct(total ? x.value / total * 100 : 0);
        r.row.toggleAttribute("data-active", active === x.key);
        r.row.toggleAttribute("data-hidden", hidden.has(x.key));
        r.row.setAttribute("aria-pressed", String(!hidden.has(x.key)));
      });
    }
    function animateTo(ms) {
      const from = new Map([...cur].map(([k, v]) => [k, v.slice()])), to = targets();
      tween(ms, (e) => {
        to.forEach((t, k) => { const f = from.get(k); cur.set(k, [f[0] + (t[0] - f[0]) * e, f[1] + (t[1] - f[1]) * e]); });
        draw();
      });
    }
    function setActive(key) {
      if (key && hidden.has(key)) key = null;
      active = key; svg.toggleAttribute("data-hovering", !!key); draw(); readout();
    }
    function toggle(key) {
      if (!hidden.has(key) && visible().length === 1) return; // satu segmen terakhir tetap tampil
      hidden.has(key) ? hidden.delete(key) : hidden.add(key);
      const d = slices.find((x) => x.key === key);
      live.textContent = `${d.label} ${hidden.has(key) ? "disembunyikan" : "ditampilkan"}. ${totalLabel} ${format(totalVisible())}.`;
      if (hidden.has(key) && active === key) active = null;
      readout(); animateTo(420);
    }
    // Satu penangan pointer untuk seluruh cincin: sudut menentukan segmen, jadi melintasi celah tidak melepas sorotan.
    function hit(e) {
      const r = svg.getBoundingClientRect(), k = size / r.width;
      const x = (e.clientX - r.left) * k - c, y = (e.clientY - r.top) * k - c, rad = Math.hypot(x, y);
      if (rad < inner - 6 || rad > c) return null;
      const turn = ((Math.atan2(y, x) + Math.PI / 2) / TAU + 1) % 1;
      const d = slices.find((x2) => { const [t0, t1] = cur.get(x2.key); return turn >= t0 && turn < t1; });
      return d ? d.key : null;
    }
    svg.addEventListener("pointermove", (e) => { const k = hit(e); if (k !== active) setActive(k); });
    svg.addEventListener("pointerleave", () => setActive(null));
    svg.addEventListener("click", (e) => { const k = hit(e); if (k) setActive(k); });

    readout(); draw(); animateTo(640);
  }

  /* ═══ Line ═══ */
  function line(container, opts) {
    const format = opts.format || ((v) => v.toLocaleString("id-ID"));
    const tickFormat = opts.tickFormat || format;
    const labels = opts.labels, n = labels.length, H = opts.height || 240;
    const series = opts.series.map((d) => Object.assign({ on: true }, d));
    const hasBar = series.some((d) => d.bar); // seri {bar: true} digambar sebagai batang (chart kombinasi)

    container.innerHTML = "";
    const fig = h("figure", { class: "arc-chart arc-line", style: { margin: 0 } }, container);
    if (series.length > 1) {
      const bar = h("div", { class: "arc-toggles", role: "group", "aria-label": "Tampilkan seri" }, fig);
      series.forEach((d) => {
        const b = h("button", { type: "button", class: "arc-toggle", "aria-pressed": "true", style: `--series: ${d.color}` }, bar);
        const sw = h("span", { class: "arc-swatch", "aria-hidden": "true" }, b);
        if (d.dashed) sw.setAttribute("data-dashed", "");
        h("span", { class: "arc-toggle-label", text: d.label }, b);
        b.addEventListener("click", () => {
          if (d.on && series.filter((x) => x.on).length === 1) return;
          d.on = !d.on; b.setAttribute("aria-pressed", String(d.on)); render(true);
        });
      });
    }
    const wrap = h("div", { class: "arc-plotwrap" }, fig);
    const plot = h("div", { class: "arc-plot", tabindex: "0", role: "img", "aria-label": opts.label || "Line chart" }, wrap);
    const svg = s("svg", { height: H }, plot);
    const gridG = s("g", {}, svg);
    const baseline = s("line", { class: "arc-baseline" }, svg);
    const paths = series.map((d) => {
      if (d.bar) return { bars: s("g", { class: "arc-bars", style: `--series: ${d.color}` }, svg) };
      const area = d.area ? s("path", { class: "arc-area", style: `--series: ${d.color}` }, svg) : null;
      const p = s("path", { class: "arc-path", style: `--series: ${d.color}` }, svg);
      if (d.dashed) p.setAttribute("data-dashed", "");
      return { area, p };
    });
    const cross = s("line", { class: "arc-cross", y1: 0, y2: H }, svg);
    const dots = series.map((d) => s("circle", { class: "arc-pt", r: 4, style: `--series: ${d.color}` }, svg));
    const gutter = h("div", { class: "arc-gutter" }, wrap);
    const axis = h("div", { class: "arc-axis" }, wrap);
    const tip = h("div", { class: "arc-tip", "aria-hidden": "true" }, wrap);

    let W = 0, top = 1, idx = null;
    // Dengan batang, tiap label mendapat pita selebar W/n dan titik ada di tengah pita
    const X = (i) => (hasBar ? ((i + 0.5) / n) * W : n > 1 ? (i / (n - 1)) * W : W / 2);
    const Y = (v) => H - (v / top) * (H - 10);

    function render(animate) {
      W = plot.clientWidth; if (!W) return;
      svg.setAttribute("width", W);
      top = niceTop(Math.max(0, ...series.filter((d) => d.on).flatMap((d) => d.values)));
      gridG.innerHTML = ""; gutter.innerHTML = "";
      for (let k = 0; k <= 4; k++) {
        const v = (top / 4) * k, y = Y(v);
        if (k > 0) s("line", { class: "arc-grid", x1: 0, x2: W, y1: crisp(y), y2: crisp(y) }, gridG);
        h("span", { text: tickFormat(v) }, h("div", { class: "arc-tick", style: { top: y + "px" } }, gutter));
      }
      baseline.setAttribute("x1", 0); baseline.setAttribute("x2", W);
      baseline.setAttribute("y1", H - 0.5); baseline.setAttribute("y2", H - 0.5);
      const barsOn = series.filter((d) => d.bar && d.on);
      const bw = (W / n) * 0.6 / Math.max(1, barsOn.length);
      series.forEach((d, si) => {
        if (d.bar) {
          paths[si].bars.innerHTML = "";
          paths[si].bars.toggleAttribute("data-hidden", !d.on);
          const k = Math.max(0, barsOn.indexOf(d));
          d.values.forEach((v, i) => s("rect", { x: n2(X(i) - (bw * barsOn.length) / 2 + k * bw), y: n2(Y(v)),
            width: n2(Math.max(1, bw - 2)), height: n2(Math.max(0, H - Y(v))) }, paths[si].bars));
          return;
        }
        const pts = d.values.map((v, i) => `${n2(X(i))} ${n2(Y(v))}`);
        const ld = "M" + pts.join("L");
        setPath(paths[si].p, ld, animate);
        paths[si].p.toggleAttribute("data-hidden", !d.on);
        if (paths[si].area) {
          setPath(paths[si].area, `${ld}L${n2(X(n - 1))} ${H}L${n2(X(0))} ${H}Z`, animate);
          paths[si].area.toggleAttribute("data-hidden", !d.on);
        }
      });
      axis.innerHTML = "";
      const room = Math.max(2, Math.floor(W / 90)), every = Math.max(1, Math.ceil((n - 1) / (room - 1)));
      for (let i = 0; i < n; i += every) {
        const x = X(i), lab = h("span", { class: "arc-axis-label", text: labels[i] }, axis);
        lab.style.left = x + "px";
        lab.style.transform = hasBar ? "translateX(-50%)" : i === 0 ? "none" : i >= n - 1 - every / 2 ? "translateX(-100%)" : "translateX(-50%)";
      }
      if (idx !== null) show(idx);
    }
    function show(i) {
      idx = Math.max(0, Math.min(n - 1, i));
      const x = X(idx);
      cross.setAttribute("x1", crisp(x)); cross.setAttribute("x2", crisp(x));
      tip.innerHTML = "";
      h("p", { class: "arc-tip-title", text: labels[idx] }, tip);
      series.forEach((d, si) => {
        dots[si].setAttribute("cx", n2(x)); dots[si].setAttribute("cy", n2(Y(d.values[idx])));
        dots[si].toggleAttribute("data-hidden", !d.on);
        if (!d.on) return;
        const row = h("p", { class: "arc-tip-row" }, tip);
        const sw = h("span", { class: "arc-tip-swatch", style: `--series: ${d.color}` }, row);
        if (d.dashed) sw.setAttribute("data-dashed", "");
        h("span", { class: "arc-tip-name", text: d.label }, row);
        h("span", { class: "arc-tip-value", text: format(d.values[idx]) }, row);
      });
      const tw = tip.offsetWidth;
      tip.style.transform = `translate(${x + 14 + tw > W ? x - 14 - tw : x + 14}px, 8px)`;
      wrap.setAttribute("data-scrubbing", "");
      plot.setAttribute("aria-label", `${opts.label || "Line chart"}. ${labels[idx]}: ` +
        series.filter((d) => d.on).map((d) => `${d.label} ${format(d.values[idx])}`).join(", "));
    }
    function hide() { idx = null; wrap.removeAttribute("data-scrubbing"); }
    const indexAt = (e) => {
      const r = plot.getBoundingClientRect(), f = (e.clientX - r.left) / r.width;
      return hasBar ? Math.floor(f * n) : Math.round(f * (n - 1));
    };
    plot.addEventListener("pointermove", (e) => show(indexAt(e)));
    plot.addEventListener("pointerdown", (e) => show(indexAt(e)));
    plot.addEventListener("pointerleave", hide);
    plot.addEventListener("blur", hide);
    plot.addEventListener("keydown", (e) => {
      const moves = { ArrowLeft: -1, ArrowRight: 1 };
      if (e.key in moves) { e.preventDefault(); show((idx ?? n - 1) + (idx === null ? 0 : moves[e.key])); }
      else if (e.key === "Home") { e.preventDefault(); show(0); }
      else if (e.key === "End") { e.preventDefault(); show(n - 1); }
      else if (e.key === "Escape") hide();
    });
    new ResizeObserver(() => render(false)).observe(plot);
    render(false);
  }

  /* ═══ Bar ═══ */
  function bar(container, opts) {
    const format = opts.format || ((v) => v.toLocaleString("id-ID"));
    const tickFormat = opts.tickFormat || format;
    const labels = opts.labels, values = opts.values, n = values.length, H = opts.height || 200;
    const mean = n ? values.reduce((a, v) => a + v, 0) / n : 0;

    container.innerHTML = "";
    const fig = h("figure", { class: "arc-chart arc-bar", style: { margin: 0 } }, container);
    const head = h("div", { class: "arc-head", "aria-live": "polite" }, fig);
    const kind = h("span", { class: "arc-kind" }, head);
    const value = h("span", { class: "arc-value" }, head);
    const when = h("span", { class: "arc-when" }, head);
    const wrap = h("div", { class: "arc-plotwrap" }, fig);
    const plot = h("div", { class: "arc-plot" }, wrap);
    const svg = s("svg", { height: H, style: "overflow: hidden" }, plot);
    const gutter = h("div", { class: "arc-gutter" }, wrap);
    const axis = h("div", { class: "arc-axis" }, wrap);
    const scrub = h("div", { class: "arc-scrub", tabindex: "0", role: "img", "aria-label": opts.label || "Bar chart" }, plot);

    let W = 0, idx = null, rects = [];
    const top = niceTop(Math.max(0, ...values, mean));
    const Y = (v) => H - (v / top) * (H - 10);

    function readout() {
      if (idx === null) {
        kind.textContent = opts.kind || "Rata-rata"; value.textContent = format(mean);
        when.textContent = opts.restWhen || " ";
      } else {
        kind.textContent = labels[idx]; value.textContent = format(values[idx]);
        when.textContent = (opts.detail && opts.detail(idx)) || " ";
      }
      rects.forEach((r, i) => r.toggleAttribute("data-active", i === idx));
      idx === null ? wrap.removeAttribute("data-scrubbing") : wrap.setAttribute("data-scrubbing", "");
    }
    function render() {
      W = plot.clientWidth; if (!W) return;
      svg.setAttribute("width", W); svg.innerHTML = ""; gutter.innerHTML = ""; axis.innerHTML = "";
      for (let k = 1; k <= 4; k++) {
        const y = Y((top / 4) * k);
        s("line", { class: "arc-grid", x1: 0, x2: W, y1: crisp(y), y2: crisp(y) }, svg);
      }
      // label sumbu yang terlalu dekat dengan label rata-rata disembunyikan agar tidak bertumpuk
      [0, top / 2, top].filter((v) => n < 2 || Math.abs(Y(v) - Y(mean)) >= 16)
        .forEach((v) => h("span", { text: tickFormat(v) }, h("div", { class: "arc-tick", style: { top: Y(v) + "px" } }, gutter)));
      const slot = W / Math.max(1, n), bw = Math.min(slot * 0.62, 44);
      rects = values.map((v, i) => {
        const y = Y(v);
        return s("rect", { class: "arc-b", x: n2(i * slot + (slot - bw) / 2), y: n2(y), width: n2(bw), height: n2(Math.max(0, H - y) + 4), rx: 4 }, svg);
      });
      s("line", { class: "arc-baseline", x1: 0, x2: W, y1: H - 0.5, y2: H - 0.5 }, svg);
      if (n > 1) {
        s("line", { class: "arc-mean", x1: 0, x2: W, y1: crisp(Y(mean)), y2: crisp(Y(mean)) }, svg);
        h("span", { text: tickFormat(mean) }, h("div", { class: "arc-tick", "data-mean": "", style: { top: Y(mean) + "px" } }, gutter));
      }
      const every = Math.max(1, Math.ceil(72 / slot)); // label berselang bila tidak muat utuh
      labels.forEach((t, i) => {
        if (i % every) return;
        const lab = h("span", { class: "arc-axis-label", text: t, title: t }, axis);
        Object.assign(lab.style, { maxWidth: slot * every - 4 + "px", overflow: "hidden", textOverflow: "ellipsis" });
        // di tengah batangnya, tetapi tidak keluar dari tepi grafik
        const w = lab.offsetWidth, x = Math.max(0, Math.min(W - w, i * slot + slot / 2 - w / 2));
        lab.style.left = x + "px";
      });
      readout();
    }
    const indexAt = (e) => { const r = plot.getBoundingClientRect(); return Math.max(0, Math.min(n - 1, Math.floor(((e.clientX - r.left) / r.width) * n))); };
    scrub.addEventListener("pointermove", (e) => { idx = indexAt(e); readout(); });
    scrub.addEventListener("pointerdown", (e) => { idx = indexAt(e); readout(); });
    scrub.addEventListener("pointerleave", () => { idx = null; readout(); });
    scrub.addEventListener("blur", () => { idx = null; readout(); });
    scrub.addEventListener("keydown", (e) => {
      const moves = { ArrowLeft: -1, ArrowRight: 1 };
      if (e.key in moves) { e.preventDefault(); idx = idx === null ? (moves[e.key] > 0 ? 0 : n - 1) : Math.max(0, Math.min(n - 1, idx + moves[e.key])); readout(); }
      else if (e.key === "Escape") { idx = null; readout(); }
    });
    new ResizeObserver(render).observe(plot);
    render();
  }

  /* Wilayah pengumuman bersama untuk pembaca layar */
  const live = h("div", { class: "arc-sr", role: "status", "aria-live": "polite" });
  document.body.appendChild(live);
  const announce = (text) => { live.textContent = ""; requestAnimationFrame(() => { live.textContent = text; }); };

  /* ═══ Confirm morph ═══
   * <form method="post" ...><span class="arc-cm" data-prompt="Hapus proyek X?">
   *   <span class="arc-cm-surface"><span class="arc-cm-face" data-face="idle">
   *     <button type="button" class="arc-cm-trigger" aria-label="Hapus proyek X">…</button></span></span></span></form>
   * Klik → "Hapus proyek X? Batal Hapus" di tempat → Hapus mengirim form. Esc, klik di luar, Batal,
   * atau 6 detik tanpa jawaban kembali ke ikon. Tekan dalam 280 ms pertama diabaikan (klik ganda tidak menghapus).
   */
  function confirmMorph(root) {
    (root || document).querySelectorAll(".arc-cm:not([data-ready])").forEach((cm) => {
      cm.setAttribute("data-ready", "");
      const form = cm.closest("form");
      const surface = cm.querySelector(".arc-cm-surface");
      const idle = cm.querySelector('[data-face="idle"]');
      const trigger = cm.querySelector(".arc-cm-trigger");
      const ask = h("span", { class: "arc-cm-face", "data-face": "ask", hidden: "" }, surface);
      h("span", { class: "arc-cm-prompt", text: cm.dataset.prompt || "Hapus?", title: cm.dataset.prompt || "" }, ask);
      const no = h("button", { type: "button", class: "arc-cm-no", text: "Batal" }, ask);
      const yes = h("button", { type: "button", class: "arc-cm-yes", text: cm.dataset.confirm || "Hapus" }, ask);
      const pending = h("span", { class: "arc-cm-face", "data-face": "pending", hidden: "" }, surface);
      const status = h("span", { class: "arc-cm-status" }, pending);
      h("span", { class: "arc-cm-spinner", "aria-hidden": "true" }, status);
      h("span", { text: cm.dataset.pending || "Menghapus" }, status);

      let openedAt = 0, timer = null;
      // Hanya lebar pil yang bergerak; isi tiap wajah tetap pada lebar aslinya.
      function show(face) {
        const from = surface.offsetWidth;
        [idle, ask, pending].forEach((f) => { f.hidden = f !== face; });
        cm.toggleAttribute("data-open", face !== idle);
        surface.style.width = "auto";
        const to = surface.offsetWidth;
        if (REDUCED.matches) return;
        surface.style.width = from + "px";
        void surface.offsetWidth;
        surface.style.width = to + "px";
      }
      const arm = () => { clearTimeout(timer); timer = setTimeout(() => close(false), 6000); };
      function onOutside(e) { if (!cm.contains(e.target)) close(false); }
      function onKey(e) { if (e.key === "Escape") { e.preventDefault(); close(true); } }
      function open() {
        show(ask); openedAt = performance.now();
        yes.focus({ preventScroll: true });
        announce(cm.dataset.prompt || "Hapus?");
        document.addEventListener("pointerdown", onOutside, true);
        cm.addEventListener("keydown", onKey);
        arm();
      }
      function close(refocus) {
        clearTimeout(timer);
        document.removeEventListener("pointerdown", onOutside, true);
        cm.removeEventListener("keydown", onKey);
        show(idle);
        if (refocus) trigger.focus({ preventScroll: true });
      }
      trigger.addEventListener("click", open);
      no.addEventListener("click", () => close(true));
      cm.addEventListener("mouseenter", () => clearTimeout(timer));
      cm.addEventListener("mouseleave", () => { if (!ask.hidden) arm(); });
      yes.addEventListener("keydown", (e) => { if (e.repeat) e.preventDefault(); });
      yes.addEventListener("click", () => {
        if (performance.now() - openedAt < 280) return;
        clearTimeout(timer);
        document.removeEventListener("pointerdown", onOutside, true);
        cm.removeEventListener("keydown", onKey);
        show(pending);
        announce(cm.dataset.pending || "Menghapus");
        form.submit();
      });
    });
  }

  /* ═══ Progress pada form yang lama diproses ═══
   * <form data-arc-progress="Membaca PDF dengan Gemini" data-busy-label="Membaca PDF…">
   * Saat dikirim: tombol nonaktif (tidak bisa terkirim dua kali) dan bar Arc progress muncul. Server tidak melaporkan
   * persentase, jadi isi bar adalah perkiraan yang melambat mendekati 90% dan yang tampil adalah waktu berjalan, bukan persen.
   */
  function progressForms(root) {
    (root || document).querySelectorAll("form[data-arc-progress]:not([data-ready])").forEach((form) => {
      form.setAttribute("data-ready", "");
      const btn = form.querySelector('button:not([type="button"])');
      const original = btn ? btn.innerHTML : "";
      let box = null, tick = null;
      form.addEventListener("submit", (e) => {
        if (form.hasAttribute("data-busy")) { e.preventDefault(); return; }
        form.setAttribute("data-busy", "");
        if (btn) { btn.disabled = true; btn.textContent = form.dataset.busyLabel || "Memproses…"; }
        const label = form.dataset.arcProgress;
        box = h("div", { class: form.classList.contains("row") ? "col-12" : "" });
        const bar = h("div", { class: "arc-progress", role: "progressbar", "aria-label": label, "aria-busy": "true" }, box);
        const meta = h("div", { class: "arc-progress-meta" }, bar);
        h("span", { text: label }, meta);
        const hint = h("span", { class: "arc-progress-hint", text: "0 detik" }, meta);
        const fill = h("span", { class: "arc-progress-fill" }, h("div", { class: "arc-progress-track" }, bar));
        form.appendChild(box);
        announce(label);
        const t0 = performance.now();
        const step = () => {
          const t = (performance.now() - t0) / 1000;
          const p = 90 * (1 - Math.exp(-t / 5)); // ±72% pada detik ke-8, tidak pernah menyentuh 100% sebelum selesai
          fill.style.transform = `translateX(${-(100 - p)}%)`;
          hint.textContent = `${Math.floor(t)} detik`;
        };
        requestAnimationFrame(step);
        tick = setInterval(step, 250);
      });
      // Kembali lewat tombol Back (bfcache): kembalikan form ke keadaan awal
      window.addEventListener("pageshow", (e) => {
        if (!e.persisted) return;
        clearInterval(tick); if (box) box.remove();
        form.removeAttribute("data-busy");
        if (btn) { btn.disabled = false; btn.innerHTML = original; }
      });
    });
  }

  /* ═══ File dropzone ═══
   * <input type="file" class="form-control" accept=".pdf" data-arc-dropzone="Pilih PDF kontrak">
   * Input asli tetap di form (tersembunyi), jadi file terkirim bersama form biasa. File bisa dipilih, diseret, atau ditempel.
   */
  function fileSize(bytes) {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1048576) return `${(bytes / 1024).toLocaleString("id-ID", { maximumFractionDigits: 0 })} KB`;
    return `${(bytes / 1048576).toLocaleString("id-ID", { maximumFractionDigits: 1 })} MB`;
  }
  function dropzones(root) {
    (root || document).querySelectorAll("input[type=file][data-arc-dropzone]:not([data-ready])").forEach((input) => {
      input.setAttribute("data-ready", "");
      const accept = (input.accept || "").split(",").map((a) => a.trim().toLowerCase()).filter(Boolean);
      const box = h("div", { class: "arc-fdz" });
      input.parentNode.insertBefore(box, input);
      const zone = h("div", { class: "arc-fdz-zone" }, box);
      const edge = s("svg", { class: "arc-fdz-edge", "aria-hidden": "true" }, zone);
      s("rect", { x: 0, y: 0, width: "100%", height: "100%", rx: 23.5, ry: 23.5 }, edge);
      const trigger = h("button", { type: "button", class: "arc-fdz-trigger" }, zone);
      const sheets = h("span", { class: "arc-fdz-sheets", "aria-hidden": "true" }, trigger);
      ["back", "side", "front"].forEach((k) => h("span", { class: "arc-fdz-sheet", "data-sheet": k }, sheets));
      h("span", { class: "arc-fdz-label", text: input.dataset.arcDropzone || "Pilih file" }, trigger);
      h("span", { class: "arc-fdz-desc", text: "Seret file ke sini atau pilih dari perangkat" }, trigger);
      h("span", { class: "arc-fdz-note", text: input.dataset.note || (accept.length ? accept.join(", ").toUpperCase().replace(/\./g, "") + ", satu file" : "Satu file") }, trigger);
      const error = h("p", { class: "arc-fdz-error", role: "alert", hidden: "" }, box);
      const row = h("div", { class: "arc-fdz-row", hidden: "" }, box);
      h("span", { class: "arc-fdz-icon", "aria-hidden": "true" }, row).appendChild(h("i", { class: input.dataset.icon || "bi bi-file-earmark-pdf" }));
      const text = h("div", { style: "min-width: 0" }, row);
      const name = h("div", { class: "arc-fdz-name" }, text);
      const meta = h("div", { class: "arc-fdz-meta" }, text);
      const remove = h("button", { type: "button", class: "arc-fdz-remove" }, row);
      remove.appendChild(h("i", { class: "bi bi-x-lg", "aria-hidden": "true" }));
      box.appendChild(input);
      input.classList.add("arc-sr");
      input.tabIndex = -1;
      trigger.setAttribute("aria-label", `${input.dataset.arcDropzone || "Pilih file"}. Seret file ke sini atau tekan untuk memilih.`);

      const ok = (file) => !accept.length || accept.some((a) =>
        a.startsWith(".") ? file.name.toLowerCase().endsWith(a) : a.endsWith("/*") ? file.type.startsWith(a.slice(0, -1)) : file.type === a);
      // Alasan penolakan (jenis atau ukuran), atau "" bila file diterima
      const maxBytes = parseFloat(input.dataset.maxMb) * 1048576 || Infinity;
      function reject(file) {
        if (!ok(file)) return `${file.name} bukan file ${accept.join(", ").toUpperCase().replace(/\./g, "")}. Pilih file yang sesuai.`;
        if (file.size > maxBytes) return `${file.name} berukuran ${fileSize(file.size)}, melebihi batas ${input.dataset.maxMb} MB.`;
        return "";
      }
      function render() {
        const file = input.files[0];
        row.hidden = !file;
        if (file) {
          name.textContent = file.name; name.title = file.name;
          meta.textContent = `${fileSize(file.size)} · siap dibaca`;
          remove.setAttribute("aria-label", `Hapus pilihan ${file.name}`);
          error.hidden = true;
        }
      }
      function take(files) {
        const file = files && files[0];
        if (!file) return;
        const problem = reject(file);
        if (problem) { error.textContent = problem; error.hidden = false; return; }
        const dt = new DataTransfer(); dt.items.add(file); input.files = dt.files;
        render(); announce(`${file.name} dipilih`);
      }
      trigger.addEventListener("click", () => input.click());
      input.addEventListener("change", () => { if (input.files[0] && reject(input.files[0])) { take(input.files); input.value = ""; } render(); });
      remove.addEventListener("click", () => { input.value = ""; render(); announce("Pilihan file dihapus"); trigger.focus(); });
      let depth = 0;
      zone.addEventListener("dragenter", (e) => { e.preventDefault(); depth++; zone.setAttribute("data-dragging", ""); });
      zone.addEventListener("dragover", (e) => { e.preventDefault(); e.dataTransfer.dropEffect = "copy"; });
      zone.addEventListener("dragleave", () => { if (--depth <= 0) { depth = 0; zone.removeAttribute("data-dragging"); } });
      zone.addEventListener("drop", (e) => { e.preventDefault(); depth = 0; zone.removeAttribute("data-dragging"); take(e.dataTransfer.files); });
      box.addEventListener("paste", (e) => { if (e.clipboardData && e.clipboardData.files.length) { e.preventDefault(); take(e.clipboardData.files); } });
      // Form dikirim tanpa file: pesan di dropzone, bukan balon browser pada input yang tersembunyi
      input.addEventListener("invalid", (e) => {
        e.preventDefault();
        error.textContent = input.dataset.required || "Pilih file dulu."; error.hidden = false; trigger.focus();
      });
      render();
    });
  }

  /* ═══ Tema: Terang · Gelap · Sistem (ikut prefers-color-scheme OS) ═══
   * Skrip di <head> base.html menerapkan pilihan sebelum halaman digambar; ini mengurus tombol dan perubahan OS.
   */
  function themeSwitch() {
    const root = document.documentElement;
    const OS = window.matchMedia("(prefers-color-scheme: dark)");
    const groups = document.querySelectorAll("[data-theme-switch]");
    const names = { light: "Terang", dark: "Gelap", system: "Sistem" };
    function apply(pref, announceIt) {
      const dark = pref === "dark" || (pref === "system" && OS.matches);
      root.setAttribute("data-theme", dark ? "dark" : "light");
      root.setAttribute("data-bs-theme", dark ? "dark" : "light");
      root.setAttribute("data-theme-pref", pref);
      groups.forEach((g) => {
        g.querySelectorAll("[data-theme-value]").forEach((b) => b.setAttribute("aria-pressed", String(b.dataset.themeValue === pref)));
        g._place && g._place(g.querySelector('[aria-pressed="true"]'));
      });
      if (announceIt) announce(`Tema ${names[pref].toLowerCase()}${pref === "system" ? ` (${dark ? "gelap" : "terang"})` : ""}`);
    }
    // Highlight meluncur (Arc button-group): ke tombol yang disorot, kembali ke tombol aktif saat pointer pergi
    groups.forEach((g) => {
      const thumb = h("span", { class: "theme-thumb", "aria-hidden": "true" });
      g.insertBefore(thumb, g.firstChild);
      const buttons = [...g.querySelectorAll("[data-theme-value]")];
      g._place = (btn) => {
        if (!btn) return;
        thumb.style.width = btn.offsetWidth + "px";
        thumb.style.transform = `translateX(${btn.offsetLeft}px)`;
        const i = buttons.indexOf(btn);
        buttons.forEach((b, n) => b.toggleAttribute("data-flank", n === i || n === i + 1));
      };
      const home = () => { g.removeAttribute("data-hovering"); g._place(g.querySelector('[aria-pressed="true"]')); };
      buttons.forEach((b) => b.addEventListener("pointerenter", (e) => { if (e.pointerType === "mouse") { g.setAttribute("data-hovering", ""); g._place(b); } }));
      buttons.forEach((b) => b.addEventListener("focus", () => { if (b.matches(":focus-visible")) g._place(b); }));
      g.addEventListener("pointerleave", home);
      g.addEventListener("focusout", (e) => { if (!g.contains(e.relatedTarget)) home(); });
      // posisi awal tanpa animasi
      g.setAttribute("data-instant", "");
      requestAnimationFrame(() => { home(); requestAnimationFrame(() => g.removeAttribute("data-instant")); });
    });
    groups.forEach((g) => g.addEventListener("click", (e) => {
      const b = e.target.closest("[data-theme-value]");
      if (!b) return;
      try { localStorage.setItem("theme", b.dataset.themeValue); } catch (_) { /* storage diblokir: berlaku sampai halaman ditutup */ }
      apply(b.dataset.themeValue, true);
    }));
    // Mode Sistem ikut berubah saat OS berganti terang/gelap (mis. macOS "Auto" pada malam hari)
    OS.addEventListener("change", () => { if (root.getAttribute("data-theme-pref") === "system") apply("system", false); });
    apply(root.getAttribute("data-theme-pref") || "system", false);
  }

  /* ═══ Disclosure: tombol [data-disclosure="id"] membuka/menutup panel #id ═══
   * Saat dibuka, fokus pindah ke field pertama; [data-disclosure-close] di dalam panel atau Esc menutup dan
   * mengembalikan fokus ke tombol pembuka.
   */
  function disclosures() {
    document.querySelectorAll("[data-disclosure]").forEach((btn) => {
      const panel = document.getElementById(btn.dataset.disclosure);
      if (!panel || panel.hasAttribute("data-disclosure-ready")) return;
      panel.setAttribute("data-disclosure-ready", "");
      const openers = () => document.querySelectorAll(`[data-disclosure="${panel.id}"]`);
      let returnTo = null;
      const set = (open, from) => {
        panel.hidden = !open;
        openers().forEach((b) => b.setAttribute("aria-expanded", String(open)));
        if (open) {
          returnTo = from || returnTo;
          panel.scrollIntoView({ block: "start", behavior: REDUCED.matches ? "auto" : "smooth" });
          const first = panel.querySelector("input:not([type=hidden]):not(.arc-sr), select, textarea, .arc-fdz-trigger");
          if (first) first.focus({ preventScroll: true });
        } else if (returnTo) returnTo.focus({ preventScroll: true });
      };
      openers().forEach((b) => b.addEventListener("click", () => set(panel.hidden, b)));
      const fromHash = () => { if (location.hash === "#" + panel.id && panel.hidden) set(true, openers()[0]); };
      window.addEventListener("hashchange", fromHash);
      fromHash();
      panel.querySelectorAll("[data-disclosure-close]").forEach((b) => b.addEventListener("click", () => set(false)));
      panel.addEventListener("keydown", (e) => { if (e.key === "Escape" && !e.defaultPrevented && !e.target.closest(".arc-cm")) set(false); });
    });
  }

  /* ═══ Command palette (Ctrl/⌘ + K atau "/") ═══
   * Kosong: Halaman (dari menu sidebar), Aksi cepat, Proyek terbaru, Riwayat hitungan dalam 2 kolom.
   * Mengetik: hasil berkelompok dari menu, aksi, dan /search (proyek, riwayat, sales). ↑↓ pilih, Enter buka, Esc tutup.
   */
  function commandPalette() {
    const trigger = document.querySelector("[data-cmdk]");
    if (!trigger) return;
    const isMac = /Mac|iPhone|iPad/.test(navigator.platform);
    trigger.querySelector(".cmdk-kbd").textContent = isMac ? "⌘ K" : "Ctrl K";
    const pages = [...document.querySelectorAll(".sidebar-nav .nav-item")].map((a) => ({
      label: a.querySelector(".label").textContent.trim(), icon: a.querySelector(".bi").className, url: a.getAttribute("href") }));
    const theme = (v) => () => document.querySelector(`[data-theme-value="${v}"]`)?.click();
    const actions = [
      { label: "Tambah proyek", icon: "bi bi-plus-lg", url: "/projects#add-project" },
      { label: "Tambah penjualan", icon: "bi bi-receipt", url: "/sales" },
      { label: "Hitung storage baru", icon: "bi bi-calculator", url: "/storage" },
      { label: "Tema terang", icon: "bi bi-sun", run: theme("light") },
      { label: "Tema gelap", icon: "bi bi-moon-stars", run: theme("dark") },
      { label: "Tema ikut sistem", icon: "bi bi-display", run: theme("system") },
    ];

    const backdrop = h("div", { class: "cmdk-backdrop", hidden: "" }, document.body);
    const dialog = h("div", { class: "cmdk", role: "dialog", "aria-modal": "true", "aria-label": "Pencarian" }, backdrop);
    const bar = h("div", { class: "cmdk-search" }, dialog);
    h("i", { class: "bi bi-search", "aria-hidden": "true" }, bar);
    const input = h("input", { type: "text", placeholder: `Cari halaman, proyek, riwayat… (${isMac ? "⌘" : "Ctrl"} + K)`, autocomplete: "off",
      spellcheck: "false", role: "combobox", "aria-expanded": "true", "aria-controls": "cmdk-list", "aria-autocomplete": "list" }, bar);
    h("kbd", { class: "cmdk-kbd", text: "esc" }, bar);
    const closeBtn = h("button", { type: "button", class: "cmdk-close", "aria-label": "Tutup pencarian" }, bar);
    h("i", { class: "bi bi-x-lg", "aria-hidden": "true" }, closeBtn);
    const body = h("div", { class: "cmdk-body", id: "cmdk-list", role: "listbox", "aria-label": "Hasil" }, dialog);
    h("div", { class: "cmdk-foot", text: "↑ ↓ pilih · Enter buka · Esc tutup" }, dialog);

    let items = [], active = 0, timer = null, ctrl = null, returnTo = null, recent = null;
    const esc = (t) => t.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
    function mark(text, q) {
      const i = q ? text.toLowerCase().indexOf(q.toLowerCase()) : -1;
      return i < 0 ? esc(text) : esc(text.slice(0, i)) + "<mark>" + esc(text.slice(i, i + q.length)) + "</mark>" + esc(text.slice(i + q.length));
    }
    function setActive(i) {
      if (!items.length) return;
      active = (i + items.length) % items.length;
      items.forEach((it, n) => it.el.setAttribute("aria-selected", String(n === active)));
      input.setAttribute("aria-activedescendant", items[active].el.id);
      items[active].el.scrollIntoView({ block: "nearest" });
    }
    function run(it) {
      close(false);
      if (it.run) it.run(); else location.href = it.url;
    }
    function render(groups, q) {
      body.innerHTML = ""; items = [];
      body.setAttribute("data-mode", q ? "search" : "browse");
      groups.filter((g) => g.items.length).forEach((g) => {
        const sec = h("div", { class: "cmdk-group", role: "group", "aria-label": g.title }, body);
        h("span", { class: "cmdk-group-title", text: g.title, "aria-hidden": "true" }, sec);
        g.items.forEach((it) => {
          const el = h("div", { class: "cmdk-item", role: "option", id: `cmdk-${items.length}`, "aria-selected": "false" }, sec);
          h("i", { class: it.icon || g.icon, "aria-hidden": "true" }, el);
          const copy = h("span", { class: "cmdk-copy" }, el);
          h("span", { class: "cmdk-label" }, copy).innerHTML = mark(it.label, q);
          if (it.desc) h("span", { class: "cmdk-desc", text: it.desc }, copy);
          const n = items.length;
          el.addEventListener("mousemove", () => { if (active !== n) setActive(n); });
          el.addEventListener("click", () => run(it));
          items.push({ el, ...it });
        });
      });
      if (!items.length) h("p", { class: "cmdk-empty", text: q ? `Tidak ada hasil untuk “${q}”.` : "Belum ada data." }, body);
      setActive(0);
    }
    function data(q) {
      if (ctrl) ctrl.abort();
      ctrl = new AbortController();
      return fetch(`${trigger.dataset.cmdk}?q=${encodeURIComponent(q)}`, { signal: ctrl.signal }).then((r) => r.json());
    }
    function show() {
      const q = input.value.trim();
      clearTimeout(timer);
      if (!q) {
        const draw = (d) => render([
          { title: "Halaman", items: pages }, { title: "Aksi cepat", items: actions },
          { title: "Proyek terbaru", icon: "bi bi-briefcase", items: d.projects.slice(0, 4) },
          { title: "Riwayat hitungan", icon: "bi bi-clock-history", items: d.logs.slice(0, 4) }], "");
        draw(recent || { projects: [], logs: [] });
        if (!recent) data("").then((d) => { recent = d; if (!input.value.trim()) draw(d); }).catch(() => {});
        return;
      }
      const has = (it) => it.label.toLowerCase().includes(q.toLowerCase());
      const local = [{ title: "Halaman", items: pages.filter(has) }, { title: "Aksi", items: actions.filter(has) }];
      render(local, q);
      timer = setTimeout(() => data(q).then((d) => {
        if (input.value.trim() !== q) return;
        render([...local, { title: "Proyek", icon: "bi bi-briefcase", items: d.projects },
          { title: "Riwayat hitungan", icon: "bi bi-clock-history", items: d.logs },
          { title: "Sales", icon: "bi bi-graph-up-arrow", items: d.sales }], q);
        announce(`${items.length} hasil`);
      }).catch(() => {}), 150);
    }
    function open() {
      if (!backdrop.hidden) return;
      returnTo = document.activeElement; recent = null;
      backdrop.hidden = false; document.documentElement.style.overflow = "hidden";
      input.value = ""; show(); input.focus();
    }
    function close(refocus = true) {
      if (backdrop.hidden) return;
      backdrop.hidden = true; document.documentElement.style.overflow = "";
      if (refocus && returnTo && document.contains(returnTo)) returnTo.focus({ preventScroll: true });
    }
    input.addEventListener("input", show);
    input.addEventListener("keydown", (e) => {
      if (e.key === "ArrowDown") { e.preventDefault(); setActive(active + 1); }
      else if (e.key === "ArrowUp") { e.preventDefault(); setActive(active - 1); }
      else if (e.key === "Enter" && items[active]) { e.preventDefault(); run(items[active]); }
    });
    dialog.addEventListener("keydown", (e) => {
      if (e.key === "Escape") { e.preventDefault(); close(); }
      else if (e.key === "Tab") { e.preventDefault(); (document.activeElement === input ? closeBtn : input).focus(); } // fokus tetap di dalam dialog
    });
    closeBtn.addEventListener("click", () => close());
    backdrop.addEventListener("mousedown", (e) => { if (e.target === backdrop) close(); });
    trigger.addEventListener("click", open);
    document.addEventListener("keydown", (e) => {
      const typing = e.target.closest && e.target.closest("input, textarea, select, [contenteditable]");
      if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === "k") { e.preventDefault(); backdrop.hidden ? open() : close(); }
      else if (e.key === "/" && !typing && backdrop.hidden) { e.preventDefault(); open(); }
    });
  }

  /* ═══ Expanding button group (mengikuti spesifikasi Arc expanding-button-group, versi tanpa React) ═══
   * <div class="arc-ebg" role="toolbar" aria-label="Aksi X" data-default="edit">
   *   <button type="button" class="arc-ebg-btn" data-id="edit" aria-label="Edit">
   *     <span class="arc-ebg-icon"><i class="bi bi-pencil"></i></span><span class="arc-ebg-label">Edit</span></button>
   *   <button type="submit" form="..." class="arc-ebg-btn" data-id="delete" data-tone="danger"
   *     data-confirm="Yakin hapus?" aria-label="Hapus">…</button>
   * </div>
   * Satu tombol selalu berlabel (data-default, else yang pertama terlihat). Tombol yang ditunjuk/difokus keyboard
   * mengambil alih label. Panah kiri/kanan, Home/End memindah fokus (satu tab stop). data-confirm: klik pertama
   * mengganti label jadi teks konfirmasi, klik kedua (dalam 4 detik) baru menjalankan tombol.
   */
  function expandingGroups(root) {
    (root || document).querySelectorAll(".arc-ebg:not([data-ready])").forEach((group) => {
      group.setAttribute("data-ready", "");
      const visible = () => [...group.querySelectorAll(".arc-ebg-btn")].filter((b) => !b.hidden);
      const resting = () => visible().find((b) => b.dataset.id === group.dataset.default) || visible()[0];
      let current = null, pointerInside = false;
      function expand(btn) {
        visible().forEach((b) => b.toggleAttribute("data-expanded", b === btn));
      }
      function rove(btn) {
        visible().forEach((b) => (b.tabIndex = b === btn ? 0 : -1));
        current = btn;
      }
      group.refresh = () => { const r = resting(); if (r) { expand(r); rove(current && !current.hidden ? current : r); } };
      group.refresh();
      group.querySelectorAll(".arc-ebg-btn").forEach((btn) => {  // termasuk tombol yang muncul belakangan (Simpan)
        btn.addEventListener("pointermove", (e) => { if (e.pointerType === "mouse") expand(btn); });
        btn.addEventListener("focus", () => { if (btn.matches(":focus-visible")) expand(btn); rove(btn); });
      });
      group.addEventListener("pointerenter", () => { pointerInside = true; });
      group.addEventListener("pointerleave", () => {
        pointerInside = false;
        if (!group.contains(document.activeElement) || !document.activeElement.matches(":focus-visible")) expand(resting());
      });
      group.addEventListener("focusout", (e) => { if (!group.contains(e.relatedTarget) && !pointerInside) expand(resting()); });
      group.addEventListener("keydown", (e) => {
        const list = visible(), i = list.indexOf(document.activeElement);
        if (i < 0) return;
        const to = { ArrowRight: (i + 1) % list.length, ArrowLeft: (i - 1 + list.length) % list.length, Home: 0, End: list.length - 1 }[e.key];
        if (to === undefined) return;
        e.preventDefault();
        list[to].focus();
      });
      // konfirmasi di tempat
      group.querySelectorAll(".arc-ebg-btn[data-confirm]").forEach((btn) => {
        const label = btn.querySelector(".arc-ebg-label"), text = label.textContent;
        let timer = null;
        const disarm = () => { clearTimeout(timer); btn.removeAttribute("data-armed"); label.textContent = text; };
        btn.addEventListener("click", (e) => {
          if (btn.hasAttribute("data-armed")) { clearTimeout(timer); return; }  // klik kedua: jalan
          e.preventDefault();
          btn.setAttribute("data-armed", "");
          label.textContent = btn.dataset.confirm;
          expand(btn);
          announce(btn.dataset.confirm);
          timer = setTimeout(disarm, 4000);
        });
        btn.addEventListener("blur", () => { if (!group.matches(":hover")) disarm(); });
      });
    });
  }

  window.Arc = { enhance, donut, line, bar, confirmMorph, progressForms, dropzones, expandingGroups };
  commandPalette();
  disclosures();
  themeSwitch();
  enhance(document);
  confirmMorph(document);
  dropzones(document);
  progressForms(document);
  expandingGroups(document);
})();
