/* Szógyakorló – német szókincs gyakorlása Excel-táblázatból
 * Feladattípusok: kvíz, névelő, lyukas szöveg legördülő listával,
 * beírós lyukas szöveg, mondatrendezés, illetve ezek vegyesen.
 */
(function () {
  "use strict";

  const STORE_DECK = "szogyak.deck.v1";
  const STORE_STATS = "szogyak.stats.v1";
  const STORE_SETTINGS = "szogyak.settings.v1";
  const DEFAULT_FILE = "szokincs.xlsx";

  const MODES = [
    { id: "mix", icon: "∗", title: "Vegyes", desc: "Minden feladattípus összekeverve" },
    { id: "quiz", icon: "?", title: "Kvíz", desc: "Válaszd ki a szó jelentését (németről magyarra és vissza)" },
    { id: "article", icon: "der", title: "Névelő", desc: "der, die vagy das? Csak főnevek" },
    { id: "select", icon: "▾", title: "Lyukas szöveg, listából", desc: "Válaszd ki a mondatba illő szót" },
    { id: "type", icon: "Aa", title: "Lyukas szöveg, beírással", desc: "Írd be a hiányzó szót" },
    { id: "order", icon: "↔", title: "Mondatrendezés", desc: "Rakd sorba az összekevert mondatrészeket" }
  ];
  const COUNTS = [10, 20, 30, 0]; // 0 = összes

  // ---------- Állapot ----------
  let deck = null;          // { name, words: [...], warnings: [...] }
  let stats = load(STORE_STATS, {});
  let settings = Object.assign({ mode: "mix", topic: "", count: 10, weak: true }, load(STORE_SETTINGS, {}));
  let session = null;       // aktuális gyakorlókör
  let lastResult = null;    // utolsó kör eredménye

  const app = typeof document !== "undefined" ? document.getElementById("app") : null;

  // ---------- Segédfüggvények ----------
  function load(key, fallback) {
    try { const v = localStorage.getItem(key); return v ? JSON.parse(v) : fallback; }
    catch (e) { return fallback; }
  }
  function save(key, value) {
    try { localStorage.setItem(key, JSON.stringify(value)); } catch (e) { /* tele a tárhely */ }
  }

  // DOM-építő: h("div", {class: "x", onclick: fn}, [gyerekek vagy szöveg])
  function h(tag, attrs, children) {
    const el = document.createElement(tag);
    if (attrs) {
      for (const k in attrs) {
        const v = attrs[k];
        if (v === null || v === undefined || v === false) continue;
        if (k.startsWith("on")) el.addEventListener(k.slice(2), v);
        else if (k === "class") el.className = v;
        else if (k === "text") el.textContent = v;
        else if (v === true) el.setAttribute(k, "");
        else el.setAttribute(k, v);
      }
    }
    (Array.isArray(children) ? children : [children]).forEach(c => {
      if (c === null || c === undefined || c === false) return;
      el.appendChild(typeof c === "string" || typeof c === "number" ? document.createTextNode(String(c)) : c);
    });
    return el;
  }

  function shuffle(arr) {
    const a = arr.slice();
    for (let i = a.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [a[i], a[j]] = [a[j], a[i]];
    }
    return a;
  }

  function norm(s) {
    return String(s || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]/g, "");
  }

  function toast(msg) {
    const t = document.getElementById("toast");
    t.textContent = msg;
    t.classList.add("show");
    clearTimeout(toast._t);
    toast._t = setTimeout(() => t.classList.remove("show"), 2800);
  }

  // ---------- Excel beolvasása ----------
  function detectColumn(header) {
    const n = norm(header);
    if (!n) return null;
    if (n.includes("mondat") && n.includes("magyar")) return "sentHu";
    if (n.includes("forditas")) return "sentHu";
    if (n.includes("mondat") || n.includes("satz")) return "sent";
    if (n.includes("nevelo") || n.includes("artikel")) return "art";
    if (n.includes("tobbes") || n.includes("plural")) return "pl";
    if (n.includes("szofaj") || n.includes("wortart")) return "pos";
    if (n.includes("tema") || n.includes("thema")) return "topic";
    if (n.includes("nemet") || n.includes("deutsch") || n === "wort" || n === "szo") return "de";
    if (n.includes("magyar") || n.includes("jelentes") || n.includes("ungarisch")) return "hu";
    return null;
  }

  function parseWorkbook(data, fileName) {
    if (typeof XLSX === "undefined") throw new Error("Hiányzik az Excel-olvasó könyvtár (lib/xlsx.full.min.js).");
    const wb = XLSX.read(data, { type: "array" });
    const names = wb.SheetNames.slice();
    const preferred = names.find(n => norm(n) === "szavak");
    if (preferred) { names.splice(names.indexOf(preferred), 1); names.unshift(preferred); }

    for (const name of names) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: "", raw: false });
      for (let hr = 0; hr < Math.min(rows.length, 5); hr++) {
        const map = {};
        rows[hr].forEach((cell, i) => { const k = detectColumn(cell); if (k && !(k in map)) map[k] = i; });
        if ("de" in map && "hu" in map) return buildDeck(rows.slice(hr + 1), map, hr + 2, fileName, name);
      }
    }
    throw new Error("Nem találtam olyan munkalapot, amelynek fejlécében szerepel a „Német” és a „Magyar” oszlop.");
  }

  function buildDeck(rows, map, firstRowNo, fileName, sheetName) {
    const words = [];
    const warnings = [];
    const seen = new Set();
    rows.forEach((r, i) => {
      const rowNo = firstRowNo + i;
      const get = k => (k in map ? String(r[map[k]] || "").trim() : "");
      let de = get("de"), hu = get("hu"), art = get("art").toLowerCase();
      if (!de && !hu) return; // üres sor
      if (!de || !hu) { warnings.push(rowNo + ". sor: hiányzik a német szó vagy a jelentés, kihagytam."); return; }

      const m = de.match(/^(der|die|das)\s+(.+)$/i);
      if (m && !art) { art = m[1].toLowerCase(); de = m[2].trim(); }
      if (art && !["der", "die", "das"].includes(art)) {
        warnings.push(rowNo + ". sor: a névelő („" + art + "”) csak der, die vagy das lehet, figyelmen kívül hagytam.");
        art = "";
      }

      const w = {
        id: norm(de) + "|" + norm(hu),
        de, hu, art,
        pl: get("pl"),
        pos: get("pos").toLowerCase(),
        topic: get("topic") || "Egyéb",
        sentHu: get("sentHu"),
        gap: null
      };
      if (seen.has(w.id)) { warnings.push(rowNo + ". sor: „" + de + "” kétszer szerepel, a másodikat kihagytam."); return; }
      seen.add(w.id);
      if (!w.pos && art) w.pos = "főnév";

      const sent = get("sent");
      if (sent) {
        w.gap = parseSentence(sent, de);
        if (!w.gap) warnings.push(rowNo + ". sor: a példamondatban nem találom a szót. Tedd csillagok közé, pl. Ich trinke einen *Kaffee*.");
      }
      words.push(w);
    });
    return { name: fileName, sheet: sheetName, words, warnings, loadedAt: Date.now() };
  }

  // Visszaadja: { before, answer, after, plain }
  function parseSentence(sent, word) {
    const marked = sent.match(/^(.*?)\*([^*]+)\*(.*)$/);
    if (marked) {
      const before = marked[1], answer = marked[2].trim(), after = marked[3].replace(/\*/g, "");
      return { before, answer, after, plain: (before + answer + after).replace(/\s+/g, " ").trim() };
    }
    const esc = word.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const re = new RegExp("(^|[^\\p{L}])(" + esc + ")(?=[^\\p{L}]|$)", "iu");
    const f = sent.match(re);
    if (!f) return null;
    const start = f.index + f[1].length;
    const answer = f[2];
    return {
      before: sent.slice(0, start), answer, after: sent.slice(start + answer.length),
      plain: sent.replace(/\s+/g, " ").trim()
    };
  }

  async function loadDefaultDeck() {
    const res = await fetch(DEFAULT_FILE);
    if (!res.ok) throw new Error("Nem sikerült betölteni a beépített szólistát.");
    const buf = await res.arrayBuffer();
    return parseWorkbook(new Uint8Array(buf), "Beépített mintaszókincs");
  }

  function setDeck(d) {
    deck = d;
    save(STORE_DECK, d);
    if (settings.topic && !topics().includes(settings.topic)) { settings.topic = ""; save(STORE_SETTINGS, settings); }
  }

  function topics() {
    return deck ? [...new Set(deck.words.map(w => w.topic))] : [];
  }

  // ---------- Feladatok előállítása ----------
  function eligible(type, w) {
    if (type === "quiz") return true;
    if (type === "article") return !!w.art;
    if (type === "select" || type === "type") return !!w.gap;
    if (type === "order") return !!w.gap && tokenize(w.gap.plain).length >= 3;
    return false;
  }

  function typesFor(mode) {
    return mode === "mix" ? ["quiz", "article", "select", "type", "order"] : [mode];
  }

  function poolFor(mode, topic) {
    const types = typesFor(mode);
    return deck.words.filter(w => (!topic || w.topic === topic) && types.some(t => eligible(t, w)));
  }

  function minNeeded(mode) {
    return mode === "article" ? 1 : 4;
  }

  function weight(w) {
    const s = stats[w.id];
    if (!s) return 1.5;
    return Math.max(0.3, 1 + 2 * s.wrong - 0.5 * s.right);
  }

  function pickWords(pool, n, weak) {
    if (!weak) return shuffle(pool).slice(0, n);
    // Súlyozott véletlen sorrend: a gyakran rontott szavak előbbre kerülnek
    return pool
      .map(w => ({ w, k: Math.pow(Math.random(), 1 / weight(w)) }))
      .sort((a, b) => b.k - a.k)
      .slice(0, n)
      .map(x => x.w);
  }

  function display(w) {
    return (w.art ? w.art + " " : "") + w.de;
  }

  function distractors(w, count, key, sameKind) {
    const val = key(w);
    const pool = deck.words.filter(x => x.id !== w.id && key(x) && norm(key(x)) !== norm(val));
    const same = shuffle(pool.filter(sameKind));
    const rest = shuffle(pool.filter(x => !sameKind(x)));
    const out = [];
    const used = new Set([norm(val)]);
    for (const x of same.concat(rest)) {
      const v = key(x);
      if (used.has(norm(v))) continue;
      used.add(norm(v));
      out.push(v);
      if (out.length === count) break;
    }
    return out;
  }

  function tokenize(sentence) {
    return sentence.split(/\s+/).filter(Boolean);
  }

  function makeQuestion(w, type) {
    const samePos = x => x.pos && x.pos === w.pos;
    if (type === "quiz") {
      const toHu = Math.random() < 0.6;
      if (toHu) {
        const opts = shuffle([w.hu].concat(distractors(w, 3, x => x.hu, samePos)));
        return { type, w, dir: "toHu", options: opts, answer: w.hu };
      }
      const opts = shuffle([display(w)].concat(distractors(w, 3, display, samePos)));
      return { type, w, dir: "toDe", options: opts, answer: display(w) };
    }
    if (type === "article") return { type, w, answer: w.art };
    if (type === "select") {
      const gapKey = x => (x.gap ? x.gap.answer : "");
      const opts = shuffle([w.gap.answer].concat(distractors(w, 3, gapKey, samePos)));
      return { type, w, options: opts, answer: w.gap.answer };
    }
    if (type === "type") return { type, w, answer: w.gap.answer };
    if (type === "order") {
      const tokens = tokenize(w.gap.plain);
      let mixed = shuffle(tokens.map((t, i) => ({ t, i })));
      for (let tries = 0; tries < 5 && mixed.every((x, i) => x.t === tokens[i]); tries++) mixed = shuffle(mixed);
      return { type, w, tokens, mixed, answer: tokens.join(" ") };
    }
  }

  function buildSession(words, mode) {
    const qs = words.map(w => {
      const ts = typesFor(mode).filter(t => eligible(t, w));
      return makeQuestion(w, ts[Math.floor(Math.random() * ts.length)]);
    });
    return { mode, questions: qs, index: 0, results: [], answered: false };
  }

  function record(w, ok) {
    const s = stats[w.id] || { right: 0, wrong: 0 };
    if (ok) s.right++; else s.wrong++;
    stats[w.id] = s;
    save(STORE_STATS, stats);
  }

  // ---------- Útválasztás ----------
  function route() {
    const r = location.hash.replace(/^#\/?/, "");
    document.querySelectorAll("[data-nav]").forEach(a => {
      a.classList.toggle("active", (r === "szavak") === (a.dataset.nav === "words"));
    });
    app.innerHTML = "";
    if (!deck) { app.appendChild(h("p", { class: "muted" }, "Szókincs betöltése…")); return; }
    if (r === "szavak") renderWords();
    else if (r === "gyakorlas" && session) renderQuestion();
    else if (r === "eredmeny" && lastResult) renderResult();
    else renderHome();
    window.scrollTo(0, 0);
  }

  function go(path) {
    if (location.hash === "#/" + path) route();
    else location.hash = "#/" + path;
  }

  // ---------- Főoldal ----------
  function renderHome() {
    const tps = topics();
    app.appendChild(h("h1", null, "Mit gyakorolsz ma?"));
    app.appendChild(h("p", { class: "deck-line" },
      deck.words.length + " szó, " + tps.length + " témakör. Forrás: " + deck.name));

    const modeWrap = h("div", { class: "modes", role: "group", "aria-label": "Feladattípus" });
    MODES.forEach(m => {
      const avail = poolFor(m.id, settings.topic).length >= minNeeded(m.id);
      const b = h("button", {
        class: "mode", type: "button", "aria-pressed": String(settings.mode === m.id), disabled: !avail,
        onclick: () => { settings.mode = m.id; save(STORE_SETTINGS, settings); route(); }
      }, [
        h("span", { class: "mode-icon" + (m.id === "article" ? " art der" : "") }, m.icon),
        h("span", null, [h("span", { class: "mode-title" }, m.title),
          h("span", { class: "mode-desc" }, avail ? m.desc : "Ehhez nincs elég megfelelő szó a kiválasztott témakörben")])
      ]);
      modeWrap.appendChild(b);
    });
    app.appendChild(modeWrap);

    const topicSel = h("select", {
      class: "input", id: "topic",
      onchange: e => { settings.topic = e.target.value; save(STORE_SETTINGS, settings); route(); }
    }, [h("option", { value: "" }, "Minden témakör")].concat(
      tps.map(t => h("option", { value: t, selected: settings.topic === t }, t))));
    app.appendChild(h("div", { class: "field" }, [h("label", { for: "topic" }, "Témakör"), topicSel]));

    const seg = h("div", { class: "seg", role: "group", "aria-label": "Kérdések száma" });
    COUNTS.forEach(c => seg.appendChild(h("button", {
      type: "button", "aria-pressed": String(settings.count === c),
      onclick: () => { settings.count = c; save(STORE_SETTINGS, settings); route(); }
    }, c === 0 ? "Összes" : String(c))));
    app.appendChild(h("div", { class: "field" }, [h("label", null, "Kérdések száma"), seg]));

    const cb = h("input", { type: "checkbox", checked: settings.weak,
      onchange: e => { settings.weak = e.target.checked; save(STORE_SETTINGS, settings); } });
    app.appendChild(h("label", { class: "check" }, [cb,
      h("span", null, ["Gyakrabban kérdezze, amit elrontottam",
        h("span", { class: "mode-desc" }, "A sokszor hibázott szavak előbb és többször kerülnek elő")])]));

    const pool = poolFor(settings.mode, settings.topic);
    const ok = pool.length >= minNeeded(settings.mode);
    app.appendChild(h("button", {
      class: "btn block start", type: "button", disabled: !ok,
      onclick: () => {
        const n = settings.count === 0 ? pool.length : Math.min(settings.count, pool.length);
        session = buildSession(pickWords(pool, n, settings.weak), settings.mode);
        go("gyakorlas");
      }
    }, ok ? "Indítás (" + Math.min(settings.count || pool.length, pool.length) + " kérdés)" : "Válassz másik feladattípust vagy témakört"));
  }

  // ---------- Kérdés megjelenítése ----------
  function renderQuestion() {
    const q = session.questions[session.index];
    const total = session.questions.length;

    app.appendChild(h("div", { class: "progress" }, [
      h("button", { class: "quit", type: "button", "aria-label": "Kilépés",
        onclick: () => { if (confirm("Biztosan abbahagyod ezt a kört?")) { session = null; go(""); } } }, "✕"),
      h("div", { class: "bar", role: "progressbar", "aria-valuemin": "0", "aria-valuemax": String(total), "aria-valuenow": String(session.index) },
        h("span", { style: "width:" + (session.index / total * 100) + "%" })),
      h("span", { class: "count" }, (session.index + 1) + "/" + total)
    ]));

    const card = h("section", { class: "card" });
    app.appendChild(card);
    const fb = h("div", { "aria-live": "polite" });

    const finish = (ok, given, note) => {
      if (session.answered) return;
      session.answered = true;
      record(q.w, ok);
      session.results.push({ q, ok, given });
      fb.appendChild(feedback(q, ok, note));
      const last = session.index === total - 1;
      const next = h("button", { class: "btn block", type: "button", onclick: nextQuestion }, last ? "Eredmény megtekintése" : "Következő");
      fb.appendChild(h("div", { class: "actions" }, next));
      next.focus();
    };

    ({ quiz: qQuiz, article: qArticle, select: qSelect, type: qType, order: qOrder })[q.type](q, card, finish);
    card.appendChild(fb);
    session.answered = false;
  }

  function nextQuestion() {
    if (session.index < session.questions.length - 1) {
      session.index++;
      route();
      window.scrollTo(0, 0);
    } else {
      lastResult = { results: session.results, mode: session.mode };
      session = null;
      go("eredmeny");
    }
  }

  function feedback(q, ok, note) {
    const box = h("div", { class: "feedback " + (ok ? "ok" : "bad") });
    box.appendChild(h("strong", null, ok ? "Helyes!" : "Nem egészen."));
    if (!ok) {
      if (q.type === "quiz") box.appendChild(h("div", null, "A helyes válasz: " + q.answer));
      if (q.type === "article") box.appendChild(h("div", null, "A helyes névelő: " + q.answer));
      if (q.type === "select" || q.type === "type") box.appendChild(h("div", null, "A hiányzó szó: " + q.answer));
    }
    if (q.w.gap && (q.type === "select" || q.type === "type" || q.type === "order")) {
      box.appendChild(h("div", { class: "full" }, q.w.gap.plain));
      if (q.w.sentHu && q.type !== "order") box.appendChild(h("div", { class: "note" }, q.w.sentHu));
    } else if (q.type === "article" || q.type === "quiz") {
      box.appendChild(h("div", { class: "full" }, [artSpan(q.w.art), q.w.de + " = " + q.w.hu]));
    }
    if (note) box.appendChild(h("div", { class: "note" }, note));
    return box;
  }

  function artSpan(art) {
    return art ? h("span", { class: "art " + art }, art + " ") : null;
  }

  // Kvíz
  function qQuiz(q, card, finish) {
    if (q.dir === "toHu") {
      card.appendChild(h("div", { class: "task" }, "Mit jelent magyarul?"));
      card.appendChild(h("div", { class: "word" }, [artSpan(q.w.art), q.w.de]));
    } else {
      card.appendChild(h("div", { class: "task" }, "Hogy mondod németül?"));
      card.appendChild(h("div", { class: "word" }, q.w.hu));
    }
    const wrap = h("div", { class: "options" });
    const btns = q.options.map(o => h("button", {
      class: "opt" + (q.dir === "toDe" ? " de" : ""), type: "button",
      onclick: () => {
        if (session.answered) return;
        btns.forEach(b => {
          b.disabled = true;
          if (b.textContent === q.answer) b.classList.add("correct");
        });
        const ok = o === q.answer;
        if (!ok) btns.find(b => b.textContent === o).classList.add("wrong");
        finish(ok, o);
      }
    }, o));
    btns.forEach(b => wrap.appendChild(b));
    card.appendChild(wrap);
  }

  // Névelő
  function qArticle(q, card, finish) {
    card.appendChild(h("div", { class: "task" }, "Melyik a helyes névelő?"));
    card.appendChild(h("div", { class: "word" }, [h("span", { class: "blank-art" }, "?"), q.w.de]));
    card.appendChild(h("div", { class: "meaning" }, q.w.hu + (q.w.pl ? "  (többes szám: " + q.w.pl + ")" : "")));
    const wrap = h("div", { class: "articles" });
    const btns = ["der", "die", "das"].map(a => h("button", {
      class: "art-btn " + a, type: "button",
      onclick: () => {
        if (session.answered) return;
        btns.forEach(b => { b.disabled = true; if (b.textContent !== q.answer) b.classList.add("dim"); });
        btns.find(b => b.textContent === q.answer).classList.add("picked");
        const blank = card.querySelector(".blank-art");
        blank.textContent = q.answer; blank.className = "art " + q.answer;
        finish(a === q.answer, a);
      }
    }, a));
    btns.forEach(b => wrap.appendChild(b));
    card.appendChild(wrap);
  }

  // Lyukas szöveg legördülő listával
  function qSelect(q, card, finish) {
    card.appendChild(h("div", { class: "task" }, "Válaszd ki a mondatba illő szót!"));
    const sel = h("select", { "aria-label": "Hiányzó szó" },
      [h("option", { value: "" }, "válassz…")].concat(q.options.map(o => h("option", { value: o }, o))));
    card.appendChild(h("p", { class: "sentence" }, [q.w.gap.before, sel, q.w.gap.after]));
    if (q.w.sentHu) card.appendChild(h("div", { class: "hint" }, q.w.sentHu));
    const btn = h("button", { class: "btn block", type: "button", disabled: true,
      onclick: () => {
        sel.disabled = true; btn.remove();
        finish(sel.value === q.answer, sel.value);
      } }, "Ellenőrzés");
    sel.addEventListener("change", () => { btn.disabled = !sel.value; });
    card.appendChild(h("div", { class: "actions" }, btn));
  }

  // Lyukas szöveg beírással
  function qType(q, card, finish) {
    card.appendChild(h("div", { class: "task" }, "Írd be a hiányzó szót!"));
    const inp = h("input", { type: "text", autocomplete: "off", autocapitalize: "off", spellcheck: "false",
      "aria-label": "Hiányzó szó", enterkeyhint: "done" });
    card.appendChild(h("p", { class: "sentence" }, [q.w.gap.before, inp, q.w.gap.after]));

    const base = norm(q.w.de) !== norm(q.answer) ? ", szótári alakja: " : "";
    card.appendChild(h("div", { class: "hint" }, ["Jelentése: ", h("b", null, q.w.hu),
      base ? base : "", base ? h("b", null, q.w.de) : ""]));
    if (q.w.sentHu) card.appendChild(h("div", { class: "hint" }, q.w.sentHu));

    const um = h("div", { class: "umlauts", "aria-label": "Különleges betűk" });
    ["ä", "ö", "ü", "ß"].forEach(ch => um.appendChild(h("button", { type: "button",
      onmousedown: e => e.preventDefault(),
      onclick: () => {
        const s = inp.selectionStart ?? inp.value.length, e = inp.selectionEnd ?? inp.value.length;
        inp.value = inp.value.slice(0, s) + ch + inp.value.slice(e);
        inp.focus(); inp.setSelectionRange(s + 1, s + 1);
      } }, ch)));
    card.appendChild(um);

    const check = () => {
      if (session.answered || !inp.value.trim()) return;
      inp.disabled = true; btn.remove(); um.remove();
      const r = compareTyped(inp.value, q.answer, !!q.w.art);
      finish(r.ok, inp.value.trim(), r.note);
    };
    const btn = h("button", { class: "btn block", type: "button", onclick: check }, "Ellenőrzés");
    inp.addEventListener("keydown", e => { if (e.key === "Enter") { e.preventDefault(); check(); } });
    card.appendChild(h("div", { class: "actions" }, btn));
    setTimeout(() => inp.focus(), 50);
  }

  function compareTyped(given, answer, isNoun) {
    const g = given.trim().replace(/\s+/g, " ");
    if (g === answer) return { ok: true };
    if (g.toLowerCase() === answer.toLowerCase()) {
      return { ok: true, note: "Elfogadtam, de figyelj a kis- és nagybetűkre: " + answer + "." +
        (isNoun ? " A német főnevek mindig nagybetűvel kezdődnek." : "") };
    }
    const alt = s => s.toLowerCase().replace(/ä/g, "ae").replace(/ö/g, "oe").replace(/ü/g, "ue").replace(/ß/g, "ss");
    if (alt(g) === alt(answer)) {
      return { ok: true, note: "Elfogadtam, de a helyes írásmód: " + answer + ". Az ä, ö, ü, ß gombokkal könnyen beírhatod." };
    }
    return { ok: false };
  }

  // Mondatrendezés
  function qOrder(q, card, finish) {
    card.appendChild(h("div", { class: "task" }, "Rakd helyes sorrendbe a mondatot!"));
    if (q.w.sentHu) card.appendChild(h("div", { class: "hint", style: "margin-bottom:10px" }, q.w.sentHu));
    const line = h("div", { class: "answer-line", "aria-label": "Az általad összerakott mondat" });
    const pool = h("div", { class: "tiles" });
    const chosen = [];

    const btn = h("button", { class: "btn", type: "button", disabled: true, onclick: () => {
      if (session.answered) return;
      const given = chosen.map(c => c.t).join(" ");
      const ok = given === q.answer;
      line.classList.add(ok ? "ok" : "bad");
      line.querySelectorAll("button").forEach(b => b.disabled = true);
      pool.remove(); btnRow.remove();
      finish(ok, given);
    } }, "Ellenőrzés");
    const reset = h("button", { class: "btn ghost", type: "button", onclick: () => {
      chosen.length = 0; redraw();
    } }, "Újrakezdés");
    const btnRow = h("div", { class: "btn-row" }, [reset, btn]);

    function redraw() {
      line.innerHTML = ""; pool.innerHTML = "";
      chosen.forEach((c, pos) => line.appendChild(h("button", { class: "tile", type: "button",
        onclick: () => { chosen.splice(pos, 1); redraw(); } }, c.t)));
      q.mixed.forEach(m => {
        const used = chosen.includes(m);
        pool.appendChild(h("button", { class: "tile" + (used ? " used" : ""), type: "button", tabindex: used ? "-1" : null,
          "aria-hidden": used ? "true" : null,
          onclick: () => { if (!chosen.includes(m)) { chosen.push(m); redraw(); } } }, m.t));
      });
      btn.disabled = chosen.length !== q.mixed.length;
    }
    redraw();
    card.appendChild(line);
    card.appendChild(pool);
    card.appendChild(btnRow);
  }

  // ---------- Eredmény ----------
  function renderResult() {
    const res = lastResult.results;
    const good = res.filter(r => r.ok).length;
    const pct = res.length ? Math.round(good / res.length * 100) : 0;
    const verdict = pct === 100 ? "Hibátlan kör, szép munka!"
      : pct >= 80 ? "Nagyon jó, már csak néhány szó hiányzik."
      : pct >= 50 ? "Jó úton jársz, a hibás szavakat érdemes még egyszer átnézni."
      : "Gyakorold még ezeket a szavakat, a hibás szavakat az app ezután gyakrabban kérdezi.";

    app.appendChild(h("h1", null, "A kör véget ért"));
    app.appendChild(h("div", { class: "score" }, [h("span", { class: "big" }, String(good)), h("span", { class: "of" }, "/ " + res.length + " helyes (" + pct + "%)")]));
    app.appendChild(h("p", { class: "verdict" }, verdict));

    const wrong = res.filter(r => !r.ok);
    const row = h("div", { class: "btn-row" });
    if (wrong.length) row.appendChild(h("button", { class: "btn", type: "button", onclick: () => {
      session = buildSession(shuffle(wrong.map(r => r.q.w)), lastResult.mode);
      go("gyakorlas");
    } }, "Hibák újragyakorlása (" + wrong.length + ")"));
    row.appendChild(h("button", { class: "btn " + (wrong.length ? "ghost" : ""), type: "button", onclick: () => go("") }, "Új kör"));
    app.appendChild(row);

    if (wrong.length) {
      app.appendChild(h("h2", null, "Ezeket rontottad el"));
      const ul = h("ul", { class: "mistakes" });
      wrong.forEach(r => {
        const q = r.q;
        const qText = q.w.gap && (q.type === "select" || q.type === "type" || q.type === "order") ? q.w.gap.plain : display(q.w) + " = " + q.w.hu;
        ul.appendChild(h("li", null, [
          h("div", { class: "q" }, qText),
          h("div", { class: "a" }, ["Te: " + (r.given || "(üres)") + ", helyesen: ", h("b", null, q.answer)])
        ]));
      });
      app.appendChild(ul);
    }
  }

  // ---------- Szókincs oldal ----------
  function renderWords() {
    app.appendChild(h("h1", null, "Szókincs"));

    const fileInput = h("input", {
      type: "file", hidden: true,
      accept: ".xlsx,.xls,.ods,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,application/vnd.oasis.opendocument.spreadsheet,text/csv",
      onchange: async e => {
        const f = e.target.files && e.target.files[0];
        if (!f) return;
        try {
          const d = parseWorkbook(new Uint8Array(await f.arrayBuffer()), f.name);
          if (d.words.length === 0) throw new Error("A fájlban nem találtam egyetlen szót sem.");
          setDeck(d);
          toast(d.words.length + " szó betöltve.");
          route();
        } catch (err) {
          alert("A fájlt nem sikerült betölteni.\n\n" + err.message);
        }
        fileInput.value = "";
      }
    });

    const box = h("div", { class: "import-box" }, [
      h("div", { class: "file-name" }, deck.name),
      h("div", { class: "muted small" }, deck.words.length + " szó, " +
        deck.words.filter(w => w.gap).length + " példamondattal, " + topics().length + " témakörben"),
      h("div", { class: "btn-row" }, [
        h("button", { class: "btn", type: "button", onclick: () => fileInput.click() }, "Excel-fájl betöltése"),
        h("button", { class: "btn ghost", type: "button", onclick: async () => {
          if (!confirm("Visszaállítod a beépített mintaszókincset? A most betöltött lista helyére kerül.")) return;
          try { setDeck(await loadDefaultDeck()); toast("Mintaszókincs visszaállítva."); route(); }
          catch (err) { alert(err.message); }
        } }, "Mintaszókincs")
      ]),
      fileInput
    ]);
    app.appendChild(box);

    if (deck.warnings && deck.warnings.length) {
      const list = h("ul", null, deck.warnings.slice(0, 15).map(t => h("li", null, t)));
      app.appendChild(h("div", { class: "warn" }, [
        h("strong", null, "Néhány sorral gond volt a betöltéskor:"), list,
        deck.warnings.length > 15 ? h("div", null, "…és még " + (deck.warnings.length - 15) + " figyelmeztetés.") : null
      ]));
    }

    const hasStats = Object.keys(stats).length > 0;
    if (hasStats) {
      app.appendChild(h("div", { class: "btn-row" }, h("button", { class: "btn light", type: "button", onclick: () => {
        if (!confirm("Törlöd az eddigi eredményeket (melyik szót hányszor rontottad el)?")) return;
        stats = {}; save(STORE_STATS, stats); toast("Statisztika törölve."); route();
      } }, "Statisztika törlése")));
    }

    const search = h("input", { class: "input search", type: "search", placeholder: "Keresés németül vagy magyarul", "aria-label": "Keresés" });
    app.appendChild(search);
    const listWrap = h("div");
    app.appendChild(listWrap);

    function drawList() {
      const qn = norm(search.value);
      listWrap.innerHTML = "";
      let shown = 0;
      topics().forEach(t => {
        const ws = deck.words.filter(w => w.topic === t && (!qn || norm(w.de).includes(qn) || norm(w.hu).includes(qn)));
        if (!ws.length) return;
        shown += ws.length;
        listWrap.appendChild(h("div", { class: "topic-head" }, t + " (" + ws.length + ")"));
        listWrap.appendChild(h("ul", { class: "wlist" }, ws.map(w => {
          const s = stats[w.id];
          return h("li", null, [
            h("span", null, [h("span", { class: "de" }, [artSpan(w.art), w.de]), w.pl ? h("span", { class: "pl" }, "  " + w.pl) : null]),
            h("span", { class: "hu" }, [w.hu, s && s.wrong ? h("span", { class: "stat" }, s.wrong + " hiba") : null])
          ]);
        })));
      });
      if (!shown) listWrap.appendChild(h("p", { class: "empty muted" }, "Nincs találat erre a keresésre."));
    }
    search.addEventListener("input", drawList);
    drawList();
  }

  // ---------- Indulás ----------
  async function init() {
    window.addEventListener("hashchange", route);
    deck = load(STORE_DECK, null);
    if (!deck || !deck.words || !deck.words.length) {
      route();
      try { setDeck(await loadDefaultDeck()); }
      catch (err) {
        app.innerHTML = "";
        app.appendChild(h("div", { class: "empty" }, [
          h("p", null, "A beépített szókincset nem sikerült betölteni: " + err.message),
          h("p", { class: "muted" }, "A Szókincs oldalon betölthetsz egy saját Excel-fájlt.")
        ]));
        deck = { name: "Nincs szókincs", words: [], warnings: [] };
        return;
      }
    }
    route();
  }

  // Teszteléshez (Node alatt)
  if (typeof module !== "undefined") {
    module.exports = { parseSentence, buildDeck, detectColumn, compareTyped, tokenize, norm };
  }
  if (typeof document !== "undefined" && document.getElementById("app")) init();
})();
