/* Standpoint.Bio — the Word task pane.
 *
 * Evidence from the scientist's own library, beside the sentence they are
 * drafting. Served by the app itself over TLS, so every call below is
 * SAME-ORIGIN: no host to configure, no CORS entry, no pairing code. That is
 * the whole architecture, and it is why `location.origin` is the only address
 * in this file.
 *
 * WHAT NEVER LEAVES THE MACHINE: the draft. Searching the library is local. The
 * only thing here that can reach the internet is the literature search at the
 * bottom, and it sends the KEYWORDS — picked locally, shown first, editable,
 * and only when a button is pressed.
 *
 * No framework and no build step: what is sideloaded is what is in git.
 */

'use strict';

var API = location.origin;

/* The API shape this pane was written against. The pane ships inside the app,
 * so the two cannot normally disagree — but the manifest in Word's wef/ folder
 * survives upgrades, which is the one real way to end up talking to a backend
 * that has moved on. A version string would be theatre; this is not. */
var ADDIN_CONTRACT = 1;

/* Word coalesces paragraph events internally at ~0.87/s (measured on Mac Word),
 * so this is a pause detector, not a debounce. */
var PAUSE_MS = 1500;

var state = {
  browserMode: false,        // no Word host: the textarea drives the pane
  standpointId: null,
  live: true,
  cache: new Map(),
  builtAt: null,
  pauseTimer: null,
  lastQuery: '',
  seq: 0,                  // guards against an older search landing last
  diag: [],
};

/* ── helpers ──────────────────────────────────────────────────────────────── */

function el(tag, cls, text) {
  var e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text != null) e.textContent = text;
  return e;
}
function clear(node) { while (node.firstChild) node.removeChild(node.firstChild); return node; }
function byId(id) { return document.getElementById(id); }

function log(line) {
  state.diag.push(new Date().toISOString().slice(11, 19) + '  ' + line);
  if (state.diag.length > 60) state.diag.shift();
  var d = byId('diag');
  if (d) d.textContent = state.diag.join('\n');
}

function api(path, opts) {
  return fetch(API + path, opts).then(function (r) {
    if (!r.ok) {
      var err = new Error('HTTP ' + r.status);
      err.status = r.status;
      // Which route failed, so a 404 can be read as "this backend is older
      // than the panel" rather than as a generic error. Only /addin/* and
      // /write/* carry that meaning; a 404 elsewhere is an ordinary one.
      err.addinRoute = path.indexOf('/addin/') === 0 || path.indexOf('/write/') === 0;
      throw err;
    }
    return r.status === 204 ? null : r.json();
  });
}

/* ── the one line a scientist reads ───────────────────────────────────────── */
//
// These states are mutually exclusive and each wants a DIFFERENT action from
// the reader, which is why they live in one place rather than as scattered
// flags. "Standpoint isn't running" and "your library is still being read" both
// show an empty pane and mean entirely different things.
//
// There is deliberately NO "certificate not trusted" state: an untrusted
// certificate means this page never loads at all, so a state for it could never
// be seen. Settings handles that, and addin/README.md says so.

function setState(kind, message, extraNode) {
  var box = byId('state');
  clear(box);
  box.className = 'state state--' + kind;
  var wrap = el('div', null, message);
  if (extraNode) wrap.appendChild(extraNode);
  box.appendChild(wrap);
}

function stateIndexing(percent) {
  var wrap = el('div');
  wrap.appendChild(el('b', null, 'Reading your library'));
  wrap.appendChild(document.createTextNode(percent == null ? '…' : '… ' + percent + '%'));
  wrap.appendChild(el('div', 'hint',
    'Searches are incomplete until this finishes. It runs once.'));
  if (percent != null) {
    var bar = el('div', 'bar');
    var fill = el('i');
    fill.style.width = Math.max(2, percent) + '%';
    bar.appendChild(fill);
    wrap.appendChild(bar);
  }
  var box = byId('state');
  clear(box);
  box.className = 'state state--wait';
  box.appendChild(wrap);
}

function stateNoMatches(onWeaker) {
  var wrap = el('div');
  wrap.appendChild(el('b', null, 'No strong matches'));
  wrap.appendChild(el('div', 'hint',
    'Nothing in your library matches this sentence closely. That is an answer, '
    + 'not a failure — but the weaker matches are there if you want them.'));
  var act = el('div', 'act');
  var b = el('button', null, 'Show weaker results');
  b.addEventListener('click', onWeaker);
  act.appendChild(b);
  wrap.appendChild(act);
  var box = byId('state');
  clear(box);
  box.className = 'state state--wait';
  box.appendChild(wrap);
}

/* ── cards ────────────────────────────────────────────────────────────────── */

var LABELS = {
  supports: 'Supports your standpoint',
  contradicts: 'Contradicts your standpoint',
  ambiguous: 'Mixed evidence',
  not_relevant: 'Not relevant to your standpoint',
};

function labelChip(label) {
  // null is NOT not_relevant. `not_relevant` is the majority verdict and means
  // it WAS assessed and judged not to bear on the standpoint; null means nobody
  // has judged it yet. Collapsing the two would misreport most of the library.
  var key = label || 'none';
  return el('span', 'chip chip--' + key, label ? LABELS[label] : 'Not assessed');
}

/* A task pane is about 320px wide and a retrieved chunk is ~1,100 characters —
 * roughly fifteen lines each, 3,365px of scrolling for six cards when measured.
 * Nobody reads that mid-sentence. The full text is already in the payload, so
 * expanding costs nothing and the default stays scannable. */
var PASSAGE_CHARS = 260;

function passageNode(card) {
  var full = card.text || '';
  var wrap = el('div', 'passage' + (card.below_threshold ? ' weak' : ''));
  if (full.length <= PASSAGE_CHARS) {
    wrap.textContent = full;
    return wrap;
  }
  // Cut at a word boundary: a mid-word ellipsis reads as corruption.
  var cut = full.slice(0, PASSAGE_CHARS);
  var sp = cut.lastIndexOf(' ');
  if (sp > PASSAGE_CHARS * 0.6) cut = cut.slice(0, sp);

  var short = el('span', null, cut + '… ');
  var more = el('button', 'inline-more', 'more');
  more.addEventListener('click', function () {
    wrap.textContent = full;
  });
  wrap.appendChild(short);
  wrap.appendChild(more);
  return wrap;
}

function cardNode(card) {
  var node = el('div', 'card');
  node.appendChild(labelChip(card.label));
  node.appendChild(passageNode(card));
  if (card.below_threshold) node.appendChild(el('div', 'hint', 'Weaker match'));

  if (card.file_trashed) {
    // FLAGGED, NEVER FILTERED — the rule services/synthesis.py already keeps.
    // The finding still carries its weight in the standpoint's balance, so a
    // pane that quietly dropped it would disagree with the app.
    node.appendChild(el('div', 'flag',
      '⚠ Source file "' + (card.file_name || 'unknown') + '" is in the trash. '
      + 'This conclusion still counts towards your standpoint.'));
  }

  var acts = el('div', 'acts');
  var open = el('button', null, 'Open in Standpoint');
  open.addEventListener('click', function () { openInApp(card, open); });
  acts.appendChild(open);
  if (card.doi) {
    var copy = el('button', null, 'Copy DOI');
    copy.addEventListener('click', function () { copyText(card.doi, copy); });
    acts.appendChild(copy);
  }
  node.appendChild(acts);
  return node;
}

/* Grouped by paper, IN THE RETRIEVER'S ORDER, which this must not re-sort:
 * that order is what the sealed eval scored. */
function renderCards(payload, query) {
  var box = clear(byId('results'));
  var cards = payload.cards || [];
  if (!cards.length) return;

  var groups = [];
  var byPaper = new Map();
  cards.forEach(function (c) {
    var key = c.paper_id != null ? 'p' + c.paper_id : 'f' + c.finding_id;
    var g = byPaper.get(key);
    if (!g) { g = { head: c, cards: [] }; byPaper.set(key, g); groups.push(g); }
    g.cards.push(c);
  });

  groups.forEach(function (g) {
    var head = g.head;
    var wrap = el('div', 'group');
    var h = el('div', 'group-head');
    h.appendChild(el('div', 'ptitle',
      head.title || (head.finding_id != null ? 'Your own bench conclusion' : 'Untitled')));
    var meta = [head.journal, head.year].filter(Boolean).join(' · ');
    if (meta) h.appendChild(el('div', 'pmeta', meta));
    wrap.appendChild(h);

    g.cards.forEach(function (c) { wrap.appendChild(cardNode(c)); });

    if (head.extra_from_paper > 0) {
      var more = el('button', 'more',
        '+' + head.extra_from_paper + ' more from this paper');
      more.addEventListener('click', function () {
        more.disabled = true;
        more.textContent = 'Loading…';
        search(query, { includeOverflow: true, weak: payload._weak, force: true });
      });
      wrap.appendChild(more);
    }
    box.appendChild(wrap);
  });

  // The contradiction guarantee's honest remainder. Disagreement past the cap
  // is REPORTED rather than dropped, because a silent cap looks like agreement.
  var hidden = cards[0] && cards[0].more_disagreeing;
  if (hidden) {
    var warn = el('div', 'state state--warn');
    warn.appendChild(el('div', null,
      '+' + hidden + ' more passage' + (hidden === 1 ? '' : 's')
      + ' disagree with your standpoint and are not shown.'));
    box.appendChild(warn);
  }
}

/* ── searching ────────────────────────────────────────────────────────────── */

function cacheKey(query, weak, overflow) {
  // Text, standpoint AND the index's build stamp. Without the stamp, a paper
  // added through the pane would not appear in a paragraph already looked at —
  // exactly when the scientist is checking whether the add worked. Without the
  // standpoint, the same sentence would show another standpoint's verdicts.
  return [query, state.standpointId, state.builtAt, weak ? 'w' : '',
          overflow ? 'o' : ''].join(' ');
}

function search(query, opts) {
  opts = opts || {};
  query = (query || '').trim();
  if (!query || state.standpointId == null) return Promise.resolve();

  var key = cacheKey(query, opts.weak, opts.includeOverflow);
  if (!opts.force && state.cache.has(key)) {
    applyResult(state.cache.get(key), query);
    log('cache hit');
    return Promise.resolve();
  }

  var mine = ++state.seq;
  state.lastQuery = query;
  setState('wait', 'Searching your library…');

  return api('/write/retrieve', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      text: query,
      hypothesis_id: state.standpointId,
      include_weak: !!opts.weak,
      include_overflow: !!opts.includeOverflow,
    }),
  }).then(function (res) {
    // An older search must never overwrite a newer one; typing produces
    // overlapping requests by construction.
    if (mine !== state.seq) { log('stale result dropped'); return; }
    res._weak = !!opts.weak;
    state.cache.set(key, res);
    applyResult(res, query);
  }).catch(function (err) {
    if (mine !== state.seq) return;
    handleError(err);
  });
}

function applyResult(res, query) {
  if (res.index_status === 'building') {
    clear(byId('results'));
    stateIndexing(res.percent);
    return;
  }
  if (res.suppressed) {
    clear(byId('results'));
    stateNoMatches(function () { search(query, { weak: true, force: true }); });
    return;
  }
  var n = (res.cards || []).length;
  if (!n) {
    clear(byId('results'));
    setState('wait', 'Nothing to show for this sentence.');
    return;
  }
  setState('ok', n + ' passage' + (n === 1 ? '' : 's') + ' from your library'
    + (res._weak ? ' (weaker matches)' : ''));
  renderCards(res, query);
}

function handleError(err) {
  clear(byId('results'));
  if (err && err.status === 404 && err.addinRoute) {
    // A 404 on an /addin/* route means the BACKEND PREDATES THIS PANEL — the
    // route simply does not exist in the running app. Saying "Standpoint isn't
    // running" there is both wrong and unhelpful: it plainly is running, and
    // the reader would go looking for a process that is already there.
    //
    // This happened in development and cost real confusion: a backend started
    // 26 minutes before the endpoint was added served a healthy /addin/status
    // and a 404 for /addin/standpoints, so the pane looked fine with an empty
    // picker. One is fixed by STARTING the app, the other by RESTARTING it.
    setState('warn', 'This Standpoint is older than the panel. Restart '
      + 'Standpoint — or update it, if restarting does not help.');
  } else if (err && err.status) {
    setState('bad', 'Standpoint replied with an error (HTTP ' + err.status + ').');
  } else {
    // A failed fetch from a pane that LOADED means the app stopped — not that
    // the certificate is untrusted, which would have stopped this page
    // rendering at all.
    setState('bad', 'Standpoint isn’t running on this computer. '
      + 'Open it, then press Find evidence.');
  }
  log('error: ' + (err && err.message));
}

/* ── actions on a card ────────────────────────────────────────────────────── */

function openInApp(card, button) {
  // POSTs to the backend, which validates against the database and asks
  // Electron to come forward. NOT a custom-scheme link: whether a sandboxed
  // task pane may hand a URL to the OS is unknown, and a failure there is
  // silent — the link would simply do nothing.
  var isFinding = card.finding_id != null && card.paper_id == null;
  var body = {
    kind: isFinding ? 'finding' : 'paper',
    id: isFinding ? card.finding_id : card.paper_id,
    hypothesis_id: state.standpointId,
  };
  if (body.id == null) return;
  var restore = function (text) {
    button.textContent = text;
    setTimeout(function () { button.textContent = 'Open in Standpoint'; }, 1900);
  };
  api('/addin/open', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  }).then(function () { restore('Opened'); })
    .catch(function () { restore('Couldn’t open'); });
}

function copyText(text, button) {
  var done = function (msg) {
    button.textContent = msg;
    setTimeout(function () { button.textContent = 'Copy DOI'; }, 1800);
  };
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(text)
      .then(function () { done('Copied'); })
      .catch(function () { fallbackCopy(text, done); });
  } else {
    fallbackCopy(text, done);
  }
}

function fallbackCopy(text, done) {
  // An embedded webview can refuse the clipboard. Showing the text so it can be
  // selected by hand beats a button that silently does nothing.
  setState('wait', 'Copy this: ' + text);
  done('Shown above');
}

/* ── live hints ───────────────────────────────────────────────────────────── */

function currentText() {
  // BROWSER TEST MODE: the textarea stands in for the document, so Find
  // evidence, live hints, the cards and every state can be exercised without
  // Word — on a machine that has none.
  if (state.browserMode) {
    var box = byId('browser-text');
    return Promise.resolve(box ? box.value.trim() : '');
  }
  if (typeof Word === 'undefined') return Promise.resolve('');
  return Word.run(function (ctx) {
    var sel = ctx.document.getSelection();
    sel.load('text');
    var paras = sel.paragraphs;
    paras.load('items/text');
    return ctx.sync().then(function () {
      // A selection if there is one, otherwise the paragraph the cursor is in.
      var selected = (sel.text || '').trim();
      if (selected.length > 2) return selected;
      var items = paras.items || [];
      return items.length ? (items[0].text || '').trim() : '';
    });
  }).catch(function () { return ''; });
}

function onPause() {
  if (!state.live) return;
  currentText().then(function (text) {
    if (!text || text === state.lastQuery) return;
    log('live hint, ' + text.length + ' chars');
    search(text);
  });
}

function schedulePause() {
  if (state.pauseTimer) clearTimeout(state.pauseTimer);
  state.pauseTimer = setTimeout(onPause, PAUSE_MS);
}

/* ── literature ───────────────────────────────────────────────────────────── */

function renderLiterature() {
  var box = clear(byId('lit'));

  box.appendChild(el('p', 'note',
    'Keywords are chosen on this computer from the sentence you are writing. '
    + 'Your sentence is never sent anywhere — only the words below, and only '
    + 'when you press Search.'));

  var field = el('label', 'field');
  field.appendChild(el('span', null, 'Keywords (edit before searching)'));
  var input = el('input');
  input.type = 'text';
  field.appendChild(input);
  box.appendChild(field);

  // NAMED BEFORE THE CLICK, not after. That is the difference between choosing
  // to send three words to three public databases and finding out afterwards.
  box.appendChild(el('p', 'note', 'Searches PubMed, Europe PMC and OpenAlex.'));

  var row = el('div', 'row');
  var suggest = el('button', null, 'Suggest from my sentence');
  suggest.addEventListener('click', function () { suggestKeywords(input, suggest); });
  var go = el('button', 'primary', 'Search');
  go.addEventListener('click', function () { runLiterature(input.value, go); });
  row.appendChild(suggest);
  row.appendChild(go);
  box.appendChild(row);

  var out = el('div');
  out.id = 'lit-results';
  box.appendChild(out);
}

function suggestKeywords(input, button) {
  button.disabled = true;
  currentText().then(function (text) {
    if (!text) return null;
    return api('/addin/keywords', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sentence: text }),
    }).then(function (res) {
      var kws = (res && res.keywords) || [];
      input.value = kws.join(' ');
      if (!kws.length) {
        // A sentence of pure scaffolding has nothing to search for, and saying
        // so beats an empty box with no explanation.
        clear(byId('lit-results')).textContent =
          'No distinctive terms in that sentence — try one naming the molecule, '
          + 'tissue or method, or type your own keywords.';
      }
    });
  }).catch(function () {}).then(function () { button.disabled = false; });
}

function runLiterature(terms, button) {
  terms = (terms || '').trim();
  var out = clear(byId('lit-results'));
  if (!terms) { out.textContent = 'Enter some keywords first.'; return; }
  button.disabled = true;
  out.textContent = 'Searching PubMed, Europe PMC and OpenAlex…';

  api('/literature/search?query=' + encodeURIComponent(terms) + '&retmax=10')
    .then(function (res) {
      var results = (res && res.results) || [];
      var box = clear(byId('lit-results'));
      if (!results.length) { box.textContent = 'No results for those keywords.'; return; }
      results.forEach(function (r) {
        var card = el('div', 'lit-result');
        card.appendChild(el('div', 'ptitle', r.title || 'Untitled'));
        var meta = [r.journal, r.year].filter(Boolean).join(' · ');
        if (meta) card.appendChild(el('div', 'pmeta', meta));
        // NO LABEL, deliberately. These are not in the library and have not
        // been assessed against the standpoint; a verdict here would be a
        // fabrication.
        var acts = el('div', 'acts');
        if (r.pmid) {
          var add = el('button', null, 'Add to library');
          add.addEventListener('click', function () { addPaper(r, add); });
          acts.appendChild(add);
        } else {
          // POST /literature/add REQUIRES a pmid. Europe PMC and OpenAlex both
          // return records without one, so offering a button that can only 422
          // would be a lie about what the app can do. Say why instead.
          acts.appendChild(el('span', 'hint',
            'No PubMed id \u2014 add this one from the app'));
        }
        if (r.doi) {
          var cp = el('button', null, 'Copy DOI');
          cp.addEventListener('click', function () { copyText(r.doi, cp); });
          acts.appendChild(cp);
        }
        card.appendChild(acts);
        box.appendChild(card);
      });
    })
    .catch(function (err) {
      clear(byId('lit-results')).textContent =
        'Search failed' + (err && err.status ? ' (HTTP ' + err.status + ')' : '') + '.';
    })
    .then(function () { button.disabled = false; });
}

function addPaper(result, button) {
  button.disabled = true;
  button.textContent = 'Adding…';
  api('/literature/add', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    // CARRY THE WHOLE RECORD. The search already knows the journal, year,
    // authors and pmcid, and schemas.py records that these were previously
    // "arriving and being thrown away" — a stored paper that cannot be cited
    // later is a paper the scientist has to look up twice.
    body: JSON.stringify({
      hypothesis_id: state.standpointId,
      pmid: result.pmid,
      pmcid: result.pmcid || null,
      title: result.title || '',
      authors: result.authors || null,
      year: result.year || null,
      doi: result.doi || null,
      journal: result.journal || null,
    }),
  }).then(function () {
    // INDEXING, not indexed. request_reconcile runs in the background, so the
    // paper is not searchable the instant it lands — and a search that missed
    // it would look like the add having failed.
    button.textContent = 'Added, indexing…';
    // Drop the cache so the next search sees the new paper.
    state.cache.clear();
    state.builtAt = null;
  }).catch(function (err) {
    button.disabled = false;
    button.textContent = (err && err.status === 409) ? 'Already in library'
                                                     : 'Couldn’t add';
  });
}

/* ── browser test mode ────────────────────────────────────────────────────── */

function enableBrowserMode(why) {
  state.browserMode = true;
  log('browser test mode (' + why + ')');
  var wrap = byId('browser-mode');
  if (wrap) wrap.hidden = false;
  var box = byId('browser-text');
  if (box) {
    // The same pause detector Word drives, so live hints behave identically.
    box.addEventListener('input', schedulePause);
    box.focus();
  }
  var sub = document.querySelector('.sub');
  if (sub) sub.textContent =
    'Browser test mode — the box below stands in for your Word document.';
  buildStateSwitch();
}

/* Every state, on demand. The alternative is contriving a half-built index or
 * a stale manifest to look at a message, which means nobody looks. */
var DEBUG_STATES = {
  indexing: function () { stateIndexing(46); clear(byId('results')); },
  'no-matches': function () {
    clear(byId('results'));
    stateNoMatches(function () { setState('ok', 'weaker results would load here'); });
  },
  'too-old': function () {
    setState('warn', 'This panel needs a newer Standpoint. Update the app, '
      + 'then remove and re-add the add-in in Word.');
  },
  offline: function () { handleError(new Error('Failed to fetch')); },
  'no-model': function () {
    setState('warn', 'Search is unavailable: this build has no embedding model.');
  },
  ready: function () { setState('wait', 'Press Find evidence, or start typing.'); },
};

function buildStateSwitch() {
  var row = byId('state-switch');
  if (!row) return;
  clear(row);
  row.appendChild(el('span', 'hint', 'Force a state:'));
  Object.keys(DEBUG_STATES).forEach(function (name) {
    var b = el('button', null, name);
    b.addEventListener('click', function () { DEBUG_STATES[name](); });
    row.appendChild(b);
  });
}

/* ?state=indexing — so a state can be linked to, screenshotted and reported
 * without anyone having to describe the clicks that produced it. */
function applyStateFromUrl() {
  var m = /[?&]state=([a-z-]+)/.exec(location.search);
  if (m && DEBUG_STATES[m[1]]) {
    setTimeout(DEBUG_STATES[m[1]], 300);   // after startup settles
    log('forced state: ' + m[1]);
    return true;
  }
  return false;
}

/* ── startup ──────────────────────────────────────────────────────────────── */

function loadStandpoints() {
  return api('/addin/standpoints').then(function (res) {
    var sel = clear(byId('standpoint'));
    (res.standpoints || []).forEach(function (s) {
      var o = el('option', null, s.title);
      o.value = s.id;
      sel.appendChild(o);
    });
    // Per-DOCUMENT memory, stored in the .docx: a manuscript about cartilage
    // keeps pointing at the cartilage standpoint even after the app moves on.
    var remembered = null;
    try {
      if (typeof Office !== 'undefined' && Office.context && Office.context.document
          && Office.context.document.settings) {
        remembered = Office.context.document.settings.get('standpointId');
      }
    } catch (e) { /* no document settings outside Word */ }
    var pick = remembered != null ? remembered : res.last_used;
    if (pick != null) sel.value = String(pick);
    state.standpointId = Number(sel.value) || res.last_used || null;
    log('standpoint ' + state.standpointId);
  });
}

function rememberStandpoint(id) {
  try {
    if (typeof Office !== 'undefined' && Office.context && Office.context.document
        && Office.context.document.settings) {
      Office.context.document.settings.set('standpointId', id);
      Office.context.document.settings.saveAsync(function () {});
    }
  } catch (e) { /* a remembered default is not worth an error */ }
}

function checkContract() {
  return api('/addin/status').then(function (s) {
    if (s.contract !== ADDIN_CONTRACT) {
      setState('warn', 'This panel needs a newer Standpoint. Update the app, '
        + 'then remove and re-add the add-in in Word.');
      return false;
    }
    return true;
  }).catch(function () { return true; });   // not running is handled elsewhere
}

function start() {
  byId('find').addEventListener('click', function () {
    currentText().then(function (t) {
      if (!t) { setState('wait', 'Put the cursor in a paragraph, or select a sentence.'); return; }
      search(t, { force: true });
    });
  });

  var live = byId('live');
  try {
    var saved = localStorage.getItem('standpoint.live');
    if (saved !== null) live.checked = saved === '1';
  } catch (e) { /* private mode */ }
  state.live = live.checked;
  live.addEventListener('change', function () {
    state.live = live.checked;
    try { localStorage.setItem('standpoint.live', live.checked ? '1' : '0'); } catch (e) {}
  });

  byId('standpoint').addEventListener('change', function (e) {
    state.standpointId = Number(e.target.value) || null;
    state.cache.clear();                 // verdicts differ per standpoint
    rememberStandpoint(state.standpointId);
    clear(byId('results'));
    setState('wait', 'Press Find evidence, or keep typing.');
  });

  renderLiterature();
  buildStateSwitch();
  var forced = applyStateFromUrl();

  checkContract()
    .then(function (ok) { return ok ? loadStandpoints() : null; })
    .then(function () { return api('/write/index-status'); })
    .then(function (s) {
      if (!s) return;
      state.builtAt = s.built_at;
      if (s.status === 'building') { stateIndexing(s.percent); return; }
      if (!s.model_available) {
        setState('warn', 'Search is unavailable: this build has no embedding model.');
        return;
      }
      if (!forced) setState('wait', 'Press Find evidence, or start typing.');
    })
    .catch(function (err) { if (!forced) handleError(err); });

  if (typeof Office === 'undefined' || !Office.onReady) {
    enableBrowserMode('Office.js did not load');
    return;
  }
  Office.onReady(function (info) {
    if (!info.host) {
      // Office.js loads and even calls onReady outside Office, reporting a
      // null host. That is the browser case, not a broken Word.
      enableBrowserMode('no Word host');
      return;
    }
    log('Word on ' + info.platform);
    try {
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged, schedulePause, function () {});
    } catch (e) { log('selection events unavailable'); }
    Word.run(function (ctx) {
      if (!ctx.document.onParagraphChanged) return ctx.sync();
      ctx.document.onParagraphChanged.add(function () {
        schedulePause();
        return Promise.resolve();
      });
      return ctx.sync();
    }).catch(function () { log('paragraph events unavailable'); });
  });
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', start);
} else {
  start();
}
