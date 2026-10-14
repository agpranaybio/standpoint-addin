/* Standpoint.Bio — Word add-in, Phase 2 connection and events spike.
 *
 * This pane MEASURES; it is not a feature.
 *
 * QUESTION 2 IS ANSWERED. Mac Word fires paragraph events about 0.9 times a
 * second while someone types (min 882 ms, max 1584 ms, mean 1131 over 17
 * intervals), with WordApi 1.1-1.7 available. Word coalesces internally, so
 * live hints need no debouncing of their own.
 *
 * QUESTION 1 IS WHY THIS STILL EXISTS. Mac Word refused
 * http://localhost:8000 in 7 MILLISECONDS -- against 45 ms for a request that
 * really happened, so nothing travelled and no prompt appeared. Chrome, by
 * contrast, HUNG and then worked once the user granted a local-network
 * permission.
 *
 * But that refused request was three things at once: plaintext, aimed at the
 * local machine, and cross-origin. Only the failure was measured, not which of
 * the three caused it -- and the fix differs per cause. So the probes below
 * separate them, and the pane is run from two origins to vary the third.
 *
 * It must also RUN OUTSIDE WORD. A pane that only works inside its host cannot
 * be debugged anywhere else, so with no Office host present everything except
 * the Word-specific checks still runs.
 */

'use strict';

var BACKEND = 'http://localhost:8000';
// Option B's address: the backend serving BOTH this page and its routes over
// TLS, so the two share an origin. The spike runs it from
// addin/serve_local_https.py; nothing listens here in an ordinary install.
var BACKEND_TLS = 'https://127.0.0.1:8443';
var PROBE_TIMEOUT_MS = 6000;

var state = {
  office: 'not started',
  env: {},
  probes: [],
  selection: null,
  selectionError: null,
  events: { selection: 0, paragraph: 0, stamps: [], api: [] },
  startedAt: new Date().toISOString(),
};

/* ── tiny DOM helpers ─────────────────────────────────────────────────────── */

function row(parent, key, value, cls) {
  var d = document.createElement('div');
  d.className = 'row';
  var k = document.createElement('span'); k.className = 'k'; k.textContent = key;
  var v = document.createElement('span'); v.className = 'v ' + (cls || ''); v.textContent = value;
  d.appendChild(k); d.appendChild(v); parent.appendChild(d);
  return v;
}

function clear(id) { var e = document.getElementById(id); e.innerHTML = ''; return e; }

/* Which of the two sideloaded copies is this?
 *
 * Both manifests open an identical-looking pane, so without a visible label two
 * reports can be transposed and the entire comparison is silently inverted --
 * the one mistake that would make this spike worse than not running it. The
 * answer is just the page's own origin, which is also the variable under test. */
function servedFrom() {
  // Compared against BACKEND_TLS EXACTLY. A looser test ("does the origin
  // mention localhost") called a dev server on another port the LOCAL copy and
  // claimed same-origin when probe 3 was still cross-origin -- which would have
  // mislabelled the one variable this spike is varying.
  if (location.origin === BACKEND_TLS) {
    return 'LOCAL copy — served by the app over TLS, SAME ORIGIN as probe 3';
  }
  if (location.protocol === 'file:') return 'opened from a file, not served';
  if (/^https?:\/\/(127\.0\.0\.1|localhost|\[::1\])(:|$)/.test(location.origin)) {
    return 'dev server at ' + location.origin + ' — NOT either sideloaded copy';
  }
  return 'HOSTED copy — served from ' + location.host + ', a public origin '
       + '(every probe below is cross-origin)';
}

/* ── 1 · environment ──────────────────────────────────────────────────────── */
//
// The user-agent string is the single most useful line in the whole report: it
// names the engine (WebView2 vs WebKit) and its version, which is what any
// later search about a platform difference has to start from.

function readEnv(officeInfo) {
  var e = state.env;
  e.userAgent = navigator.userAgent;
  e.origin = location.origin;
  e.secureContext = String(window.isSecureContext);
  e.officeJsLoaded = String(typeof Office !== 'undefined' && !!Office);

  // Office.js loads and even calls onReady OUTSIDE Office, with host and
  // platform both null. String(null) is "null", which is truthy -- so naively
  // stringifying here painted "no host at all" green, exactly like a healthy
  // Word. Keep the absence absent.
  if (officeInfo) {
    if (officeInfo.host != null) e.host = String(officeInfo.host);
    if (officeInfo.platform != null) e.platform = String(officeInfo.platform);
  }
  try {
    if (typeof Office !== 'undefined' && Office.context && Office.context.diagnostics) {
      var d = Office.context.diagnostics;
      if (d.version) e.officeVersion = String(d.version);
      if (d.host) e.officeHost = String(d.host);
      if (d.platform) e.officePlatform = String(d.platform);
    }
  } catch (err) { e.diagnosticsError = String(err); }

  // Which API levels this host actually offers. The event used below for
  // content changes needs WordApi 1.5; without it, only the Common API
  // selection event is available and live hints lose their best signal.
  try {
    if (typeof Office !== 'undefined' && Office.context && Office.context.requirements) {
      var sets = [];
      ['1.1', '1.2', '1.3', '1.4', '1.5', '1.6', '1.7'].forEach(function (v) {
        try {
          if (Office.context.requirements.isSetSupported('WordApi', v)) sets.push(v);
        } catch (err) { /* older hosts throw rather than return false */ }
      });
      e.wordApi = sets.length ? sets.join(', ') : '(none reported)';
    }
  } catch (err) { e.requirementsError = String(err); }

  var box = clear('env');
  row(box, 'Office.js loaded', e.officeJsLoaded, e.officeJsLoaded === 'true' ? 'ok' : 'bad');
  row(box, 'host', e.host || '(no Office host)', e.host ? 'ok' : 'warn');
  row(box, 'platform', e.platform || '—', e.platform ? 'ok' : 'idle');
  row(box, 'Office version', e.officeVersion || '—', 'idle');
  row(box, 'WordApi sets', e.wordApi || '—', e.wordApi ? 'ok' : 'idle');
  row(box, 'page origin', e.origin, 'idle');
  row(box, 'secure context', e.secureContext, e.secureContext === 'true' ? 'ok' : 'warn');
  row(box, 'user agent', e.userAgent, 'idle');
}

/* ── 2 · reaching the backend ─────────────────────────────────────────────── */
//
// WHY SEVERAL PROBES RATHER THAN ONE. From JavaScript, a blocked request, a
// CORS rejection, a rejected certificate and a refused connection are
// INDISTINGUISHABLE -- all arrive as a bare TypeError with no further detail,
// by design, so that a page cannot use fetch to port-scan a machine. One red
// cross would not say which architecture to choose.
//
// The TIMING carries information the error text does not. A real request to a
// real server costs tens of milliseconds; Mac Word's refusal came back in 7,
// which is too fast for anything to have left the machine. And a request
// awaiting a permission HANGS rather than failing, which is how Chrome's gate
// was identified. So both are reported.
//
// The control proves fetch works in this webview at all, so a failure above is
// specific to its target rather than general.

function probe(label, url, expect) {
  var started = Date.now();
  var ctrl = typeof AbortController !== 'undefined' ? new AbortController() : null;
  // A BLOCKED request can hang rather than reject, which would leave the row
  // on "running" forever and read as a crash.
  var timer = ctrl ? setTimeout(function () { ctrl.abort(); }, PROBE_TIMEOUT_MS) : null;

  var opts = { method: 'GET', cache: 'no-store' };
  if (ctrl) opts.signal = ctrl.signal;

  return fetch(url, opts).then(function (res) {
    if (timer) clearTimeout(timer);
    return res.text().then(function (body) {
      return {
        label: label, url: url, ok: res.ok, status: res.status,
        ms: Date.now() - started,
        detail: (body || '').slice(0, 120),
        expect: expect,
      };
    });
  }).catch(function (err) {
    if (timer) clearTimeout(timer);
    var aborted = err && (err.name === 'AbortError');
    return {
      label: label, url: url, ok: false, status: 0, ms: Date.now() - started,
      detail: aborted
        ? 'timed out after ' + PROBE_TIMEOUT_MS + ' ms (a blocked request can hang rather than fail)'
        : (err && err.name ? err.name : 'Error') + ': ' + (err && err.message ? err.message : String(err)),
      expect: expect,
    };
  });
}

function runProbes() {
  var box = clear('net');
  row(box, 'running…', '', 'idle');

  // FOUR probes, because the failure we are chasing had THREE causes at once
  // and only one of them was measured. Mac Word refused
  // http://localhost:8000 in 7 ms, and that request was simultaneously
  // plaintext, aimed at the local machine, and cross-origin. These four
  // separate them:
  //
  //   plaintext http  ..... rows 1 and 2
  //   secure https    ..... row 3, same destination
  //   cross-origin    ..... rows 1-3 when this page is served from Pages,
  //                         row 3 is SAME-origin when served from :8443
  //   nothing at all  ..... row 4, which must always pass
  //
  // So one run of this pane from each of the two origins tells us which rule
  // actually bit, rather than which workaround happens to survive.
  var jobs = [
    probe('http localhost:8000', BACKEND + '/health', 'what Mac Word refused'),
    probe('http 127.0.0.1:8000', 'http://127.0.0.1:8000/health', 'same thing by IP'),
    probe('https 127.0.0.1:8443', BACKEND_TLS + '/health', 'Option B'),
    // Same origin as this page, so it needs no CORS and no plaintext: if THIS
    // fails, the webview has no network and nothing above means anything.
    probe('control (own origin)', location.origin + location.pathname + '?probe=' + Date.now(), 'must succeed'),
  ];

  return Promise.all(jobs).then(function (results) {
    state.probes = results;
    var b = clear('net');
    results.forEach(function (r) {
      var good = r.ok;
      var text = good ? ('HTTP ' + r.status + ' · ' + r.ms + ' ms') : r.detail;
      row(b, r.label, text, good ? 'ok' : 'bad');
      if (good && r.detail && r.label !== 'control (own origin)') {
        row(b, '', r.detail, 'idle');
      }
    });
    buildReport();
    return results;
  });
}

/* ── 3 · the selection ────────────────────────────────────────────────────── */

function readSelection() {
  var box = clear('sel');
  if (typeof Word === 'undefined') {
    row(box, 'status', 'no Word host — skipped', 'idle');
    return Promise.resolve();
  }
  return Word.run(function (ctx) {
    var sel = ctx.document.getSelection();
    sel.load('text');
    return ctx.sync().then(function () {
      var t = (sel.text || '').replace(/\s+/g, ' ').trim();
      state.selection = t;
      var b = clear('sel');
      row(b, 'characters', String(t.length), t.length ? 'ok' : 'warn');
      row(b, 'text', t ? (t.slice(0, 80) + (t.length > 80 ? '…' : '')) : '(nothing selected)',
          t.length ? 'ok' : 'idle');
      buildReport();
    });
  }).catch(function (err) {
    state.selectionError = String(err && err.message ? err.message : err);
    var b = clear('sel');
    row(b, 'failed', state.selectionError, 'bad');
    buildReport();
  });
}

/* ── 4 · events while typing ──────────────────────────────────────────────── */
//
// TWO DIFFERENT EVENTS, because they answer different halves of the question.
//
//   DocumentSelectionChanged (Common API, supported for a long time) fires when
//   the CURSOR or selection moves. It is the fallback signal: "refresh when the
//   cursor reaches a new paragraph".
//
//   Word.Document.onParagraphChanged (WordApi 1.5) fires when a paragraph's
//   CONTENT changes -- which is what typing actually is, and therefore the
//   signal live hints want. It may not exist on an older host, so it is
//   registered defensively and its absence is reported rather than thrown.
//
// Registering only one would have produced a confident, wrong answer: a host
// where selection events are sparse while typing might still stream paragraph
// events, and vice versa.

function stamp(kind) {
  var now = Date.now();
  var s = state.events.stamps;
  var gap = s.length ? (now - s[s.length - 1].t) : null;
  s.push({ t: now, kind: kind, gap: gap });
  if (s.length > 40) s.shift();

  if (kind === 'selection') state.events.selection++;
  else state.events.paragraph++;

  document.getElementById('c-sel').textContent = String(state.events.selection);
  document.getElementById('c-par').textContent = String(state.events.paragraph);
  renderEvents();
}

function renderEvents() {
  var box = clear('ev');
  var apis = state.events.api;
  apis.forEach(function (a) { row(box, a.name, a.status, a.ok ? 'ok' : 'bad'); });

  var s = state.events.stamps;
  var gaps = s.map(function (x) { return x.gap; }).filter(function (g) { return g !== null; });
  if (gaps.length) {
    var sorted = gaps.slice().sort(function (a, b) { return a - b; });
    var median = sorted[Math.floor(sorted.length / 2)];
    row(box, 'median gap', median + ' ms', 'ok');
    row(box, 'fastest / slowest', sorted[0] + ' / ' + sorted[sorted.length - 1] + ' ms', 'idle');
    row(box, 'last 10 gaps',
        s.slice(-10).map(function (x) {
          return (x.kind === 'selection' ? 's' : 'p') + ':' + (x.gap === null ? '—' : x.gap);
        }).join('  '), 'idle');
  }
  buildReport();
}

function registerEvents() {
  // The Common API selection event.
  try {
    if (typeof Office !== 'undefined' && Office.context && Office.context.document &&
        Office.context.document.addHandlerAsync) {
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        function () { stamp('selection'); },
        function (res) {
          state.events.api.push({
            name: 'DocumentSelectionChanged',
            ok: res.status === Office.AsyncResultStatus.Succeeded,
            status: res.status === Office.AsyncResultStatus.Succeeded
              ? 'registered'
              : ('failed: ' + (res.error && res.error.message)),
          });
          renderEvents();
        });
    } else {
      state.events.api.push({ name: 'DocumentSelectionChanged', ok: false, status: 'no Office host' });
    }
  } catch (err) {
    state.events.api.push({ name: 'DocumentSelectionChanged', ok: false, status: String(err) });
  }

  // The Word content event. WordApi 1.5; absent hosts must degrade, not throw.
  if (typeof Word === 'undefined') {
    state.events.api.push({ name: 'onParagraphChanged', ok: false, status: 'no Word host' });
    renderEvents();
    return;
  }
  Word.run(function (ctx) {
    if (!ctx.document.onParagraphChanged) {
      state.events.api.push({
        name: 'onParagraphChanged', ok: false,
        status: 'not available on this host (needs WordApi 1.5)',
      });
      renderEvents();
      return ctx.sync();
    }
    ctx.document.onParagraphChanged.add(function () { stamp('paragraph'); return Promise.resolve(); });
    return ctx.sync().then(function () {
      state.events.api.push({ name: 'onParagraphChanged', ok: true, status: 'registered' });
      renderEvents();
    });
  }).catch(function (err) {
    state.events.api.push({
      name: 'onParagraphChanged', ok: false,
      status: String(err && err.message ? err.message : err),
    });
    renderEvents();
  });
}

/* ── 5 · the report ───────────────────────────────────────────────────────── */
//
// Rendered into a textarea as well as copied, because navigator.clipboard can
// be refused inside an embedded webview and a "Copy" button that silently did
// nothing would lose the whole run. Select-all in the box always works.

function buildReport() {
  var L = [];
  L.push('STANDPOINT WORD ADD-IN — PHASE 2 SPIKE');
  // First line after the title, so a pasted report can never be attributed to
  // the wrong copy.
  L.push(servedFrom());
  L.push('started ' + state.startedAt);
  L.push('');
  L.push('[environment]');
  Object.keys(state.env).forEach(function (k) { L.push('  ' + k + ': ' + state.env[k]); });
  L.push('  Office.onReady: ' + state.office);
  L.push('');
  L.push('[backend probes]  target ' + BACKEND);
  if (!state.probes.length) L.push('  (not run yet)');
  state.probes.forEach(function (p) {
    L.push('  ' + p.label + ' -> ' + p.url);
    L.push('      ' + (p.ok ? 'OK HTTP ' + p.status : 'FAILED') + ' in ' + p.ms + ' ms');
    L.push('      ' + p.detail);
  });
  L.push('');
  L.push('[selection]');
  if (state.selectionError) L.push('  error: ' + state.selectionError);
  else if (state.selection === null) L.push('  (not read)');
  else L.push('  ' + state.selection.length + ' chars: ' + state.selection.slice(0, 120));
  L.push('');
  L.push('[events]');
  state.events.api.forEach(function (a) { L.push('  ' + a.name + ': ' + a.status); });
  L.push('  selection events: ' + state.events.selection);
  L.push('  paragraph events: ' + state.events.paragraph);
  var gaps = state.events.stamps.map(function (x) { return x.gap; })
    .filter(function (g) { return g !== null; });
  if (gaps.length) {
    var sorted = gaps.slice().sort(function (a, b) { return a - b; });
    L.push('  gaps ms  median ' + sorted[Math.floor(sorted.length / 2)] +
           '  min ' + sorted[0] + '  max ' + sorted[sorted.length - 1] +
           '  n ' + sorted.length);
    L.push('  sequence: ' + state.events.stamps.slice(-20).map(function (x) {
      return (x.kind === 'selection' ? 's' : 'p') + ':' + (x.gap === null ? '0' : x.gap);
    }).join(' '));
  }
  document.getElementById('report').value = L.join('\n');
}

function copyReport() {
  var note = document.getElementById('copied');
  var text = document.getElementById('report').value;
  var done = function (msg) { note.textContent = msg; };
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(text)
      .then(function () { done('Copied. Paste it back to me.'); })
      .catch(function (err) { done('Clipboard refused (' + err.name + ') — select the text below and copy by hand.'); });
  } else {
    done('No clipboard API here — select the text below and copy by hand.');
  }
}

/* ── start ────────────────────────────────────────────────────────────────── */
//
// Everything that does not need Word runs immediately, so a host that never
// becomes ready still produces a usable report instead of a blank pane.

function start() {
  document.getElementById('again').addEventListener('click', runProbes);
  document.getElementById('copy').addEventListener('click', copyReport);
  document.getElementById('refresh').addEventListener('click', function () {
    readSelection(); buildReport();
  });

  document.getElementById('served-from').textContent = '· ' + servedFrom();

  readEnv(null);
  runProbes();
  buildReport();

  if (typeof Office === 'undefined' || !Office.onReady) {
    state.office = 'Office.js absent — running outside Word';
    readEnv(null);
    renderEvents();
    buildReport();
    return;
  }

  // A host that never calls onReady would otherwise leave this at "not
  // started" with nothing saying so.
  var settled = false;
  setTimeout(function () {
    if (!settled) {
      state.office = 'onReady did not fire within 8 s';
      buildReport();
    }
  }, 8000);

  Office.onReady(function (info) {
    settled = true;
    // onReady fires even in a plain browser, reporting a null host. Saying
    // "ready" there would read as a successful Word connection in the report.
    state.office = info.host
      ? ('ready (host=' + info.host + ', platform=' + info.platform + ')')
      : 'onReady fired with NO host — this is not Word';
    readEnv(info);
    readSelection();
    registerEvents();
    buildReport();
  });
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', start);
} else {
  start();
}
