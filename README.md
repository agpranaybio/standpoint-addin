# Standpoint.Bio — Word task pane (Phase 2 spike)

A diagnostic pane, not a feature. It exists to settle how the real add-in is
built, and desktop Word is the only place that can happen.

**One question is now closed.** Events fire while you type at about **0.9 per
second** on Mac Word, with `WordApi` 1.1–1.7 available. Word coalesces
internally, so live hints need no debouncing of their own. Measurements below.

**The open question is how the pane reaches Standpoint.** Mac Word refused
`http://localhost:8000` in **7 milliseconds** — a refusal, not a question, with
no prompt and nothing to click. Chrome, on the same URL from the same page,
*hung* and then worked once a local-network permission was granted. Different
engines, different rules.

But that refused request was **three things at once** — plaintext, aimed at the
local machine, and cross-origin — and only the failure was measured, not which
of the three caused it. They have different fixes, and one of them is much
cheaper than the other:

- if **plaintext** was the blocker, the pane can stay hosted and a fix is a
  `git push` every tester gets at once;
- if **reaching the local machine from a public page** is blocked, the app has
  to serve the pane itself and every fix needs a new installer.

So the pane now probes all of them, and is run from two origins to vary the
third. Which rule actually bit is still unknown: for Chrome it was a permission,
but WebKit may well be refusing the plaintext, and guessing between those costs
the same trip either way.

It reads your selection and nothing else, and it **cannot write to your
document** — the manifest asks for `ReadDocument`, the lowest permission that
can read a selection, so Word itself would refuse an edit.

## What you need

- **Standpoint.Bio running on the same machine.** The installed v0.2.0 build is
  enough; this spike only uses `GET /health`, which that build has. Nothing
  needs rebuilding.
- Desktop Word with a Microsoft 365 licence.

## Sideloading

Paths vary by Office version — confirm them on the machine rather than trusting
this list.

**macOS**

```sh
cp manifest.xml manifest-local.xml ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
```

Create the `wef` folder if it is not there, then restart Word. Both appear under
**Home ▸ Standpoint** — as **Diagnostics** (hosted) and **Local** (served by the
app). Copy both: running the same pane from the two origins is what separates
the causes.

**Windows**

1. Put **both** manifests in a folder and share it (right-click ▸
   *Properties ▸ Sharing ▸ Share*).
2. Word ▸ **File ▸ Options ▸ Trust Center ▸ Trust Center Settings ▸ Trusted
   Add-in Catalogs**. Paste the share path (`\\MACHINE\folder`), click **Add
   catalog**, tick **Show in Menu**, OK.
3. Restart Word. **Insert ▸ My Add-ins ▸ Shared Folder** — both
   *Standpoint (spike)* and *Standpoint (local)* should be listed.

## Running it

1. Start Standpoint, so the backend is listening on `127.0.0.1:8000`, and
   start `serve_local_https.py` (below) so `127.0.0.1:8443` answers too.
2. Open a document with a few paragraphs of real prose.
3. Open the pane. Sections 1 and 2 fill in by themselves.
4. **Select a sentence.** Section 3 should show its text.
5. **Click into a paragraph and type for about 20 seconds.** Watch the two
   counters in section 4. What matters is not that they move but *how* — the
   median gap distinguishes per-keystroke from per-word from per-paragraph.
6. Press **Copy report** and send the block back. If the clipboard is refused —
   embedded webviews sometimes do — select the text in the box and copy by hand.

## What is already measured

### Mac Word refused, in 7 milliseconds

| | Mac Word (WebKit) | Chrome 149 (Chromium) |
| --- | --- | --- |
| `http://localhost:8000` | `TypeError: Load failed` in **7 ms** | hung, then **worked after Allow** |
| `http://127.0.0.1:8000` | `TypeError: Load failed` in **7 ms** | same |
| the page itself | HTTP 200 in 45 ms | HTTP 200 in 25 ms |

**7 ms is a refusal, not a question.** The one request that really happened took
45 ms, so nothing travelled — a rule was checked and that was that. No prompt
appeared. Chrome, by contrast, *hung*: it was waiting on a local-network
permission, and granting it made the call work. Two engines, two behaviours, and
the Mac one has nothing to click.

The user agent confirms the engine: `AppleWebKit/605.1.15` with no Chrome or
Safari token — a bare embedded WKWebView.

### Events: answered, and the answer is good

Both event APIs registered, `WordApi` **1.1 through 1.7**. While typing: min
882 ms, max 1584 ms, **mean 1131 ms over 17 intervals — about 0.9 events per
second**. Word coalesces internally, so live hints need no debouncing of their
own. Reading the selection works too (`0 chars` with *no error* means the whole
chain succeeded and nothing was selected).

That criterion is closed. Everything below is about the connection.

### What has NOT been established

The refused request was **three things at once**, and only the failure was
measured:

| | plaintext `http` | public page → local machine | cross-origin |
| --- | --- | --- | --- |
| what was tested | ✗ | ✗ | ✗ |
| hosted pane → `https://127.0.0.1:8443` | fixed | **still there** | **still there** |
| locally served pane | fixed | fixed | fixed |

The three have different prices. If plaintext was the only blocker, the pane can
stay hosted and **a fix is a `git push`** that every tester gets at once. If
reaching the local machine from a public page is blocked, the app must ship the
pane and **every fix needs a new installer**. And there is direct evidence the
middle column is real: Chrome's prompt said *"access devices on your local
network"* — about the destination, not the scheme.

So the next run tests both, in one sitting.

## The run: two manifests, side by side

### 1. Start the TLS server

```sh
python addin/serve_local_https.py
```

It makes a self-signed certificate covering `localhost` **and** `127.0.0.1`,
serves the pane and a `/health` route on `https://127.0.0.1:8443`, and prints
the command to trust it — **and the command to remove it again afterwards.**
Run the removal when you are done; a spike has no business leaving a trusted
root on a machine.

**Trust the certificate before opening the pane.** An untrusted one does not
produce an error anywhere — the pane is simply blank.

While you are there, note **how many steps the trust command takes and whether
it asks for your password.** That number is this option's product cost and is
worth writing down even if everything works.

### 2. Sideload both manifests

`manifest.xml` (hosted, ribbon button **Diagnostics**) and `manifest-local.xml`
(served by the script, ribbon button **Local**). They carry different GUIDs, so
Word installs both and you can switch between them.

Each pane says at the top which copy it is, and so does the first line of its
report — the panes are identical otherwise, and two transposed reports would
invert the entire conclusion.

### 3. Read the server's log, not just the pane

The log is the instrument, because the pane cannot tell these apart — all three
arrive as a bare `TypeError`:

| in the server's log | what it means |
| --- | --- |
| `<- GET /health` | the request got through |
| `!! TLS HANDSHAKE REFUSED` | the certificate is not trusted — trust it, re-run |
| **nothing at all** | the browser never sent it: a **policy block** |

That last row is the finding this spike exists for. The middle one is a setup
problem wearing the same clothes, which is why it is called out separately.

### 4. The result, Mac, Oct 14 2026

Both panes, one sitting, same machine:

| pane served from | `http://…:8000` | `https://127.0.0.1:8443` | own origin |
| --- | --- | --- | --- |
| `https://127.0.0.1:8443` (**Local**) | FAILED 2 ms | **OK 200 · 19 ms** | OK 200 · 24 ms |
| `agpranaybio.github.io` (**Diagnostics**) | FAILED 5 ms | **FAILED 45 ms** | OK 200 · 102 ms |

**Option B works; Option A is dead on both counts.** Plaintext is refused in
2–5 ms — faster than the 45 ms a real request costs, so nothing travelled. And
the hosted pane's *secure* call failed where the local pane's identical call
succeeded in 19 ms, so it is not only plaintext that is blocked: a public origin
cannot reach this machine either.

Trusting the certificate took **one step**. WKWebView accepted it without
complaint — the pane loaded, and `/health` answered same-origin.

**Where the evidence stops.** The server log for the hosted run was not
captured; the excerpt available starts 115 s later and contains only the local
pane's own requests. The server **was** running, which rules out
connection-refused. What stays unconfirmed is whether the hosted request reached
the server and was blocked on the response, or never left the browser — and no
response header remains to try, since `Access-Control-Allow-Origin: *` and
`Access-Control-Allow-Private-Network: true` were already being sent. **Inferred
from timing, not from a logged absence.** Recorded this way because the first
conclusion drawn here was nearly written down from a log that did not cover the
window.

**Consequences.** The pane ships inside the app; a pane fix needs a new
installer rather than a `git push`, and this repo stops being the delivery
route. In exchange the pane is same-origin with the API, so there is no CORS
surface at all.

**Still owed:** the same two manifests on Windows.

## The control experiment, on both machines

Optional, and only if a result is confusing. On the Mac open **Safari** and
**Chrome**; on Windows open **Edge** and **Chrome**, and visit

```
https://agpranaybio.github.io/standpoint-addin/taskpane.html
```

Safari is WebKit and Edge/Chrome are Chromium — the same engines Word embeds.
If a plain browser on that machine behaves exactly like the pane inside Word,
the cause is the engine's policy and no add-in code will change it. Word's own
webview console is awkward to reach on Windows and effectively unavailable on
Mac, so this stands in for a debugger.

## Files

| | |
| --- | --- |
| `manifest.xml` | Loads the pane from GitHub Pages. The XML manifest, not the unified JSON one: that is Windows/M365-first and sideloading it into Mac desktop Word is not dependable. |
| `manifest-local.xml` | Loads the identical pane from `https://127.0.0.1:8443`. Its own GUID, so both install at once. |
| `serve_local_https.py` | The TLS spike server: certificate, pane, `/health`, and a log that tells a policy block apart from a refused certificate. |
| `taskpane.html` | One screen, a row per check. |
| `taskpane.js` | The checks. No framework and no build step, so what is sideloaded is what is in git. |
| `assets/icon-*.png` | Ribbon icons. |
| `check_manifest.py` | Run before publishing: `python check_manifest.py`. Checks both manifests and catches the mistakes whose only symptom in Word is a pane that silently never appears — including the two sharing a GUID, which would install only one. |

## Hosting

Authored in the private app repo under `addin/` and published to the public
`standpoint-addin` repo with `git subtree push --prefix=addin`, served by GitHub
Pages at <https://agpranaybio.github.io/standpoint-addin/>. One source of truth,
and nothing here but static files that talk to the viewer's own localhost.

## Not in this spike

The pane does **not** call `/write/retrieve` and shows no evidence cards. That
endpoint does not exist in the v0.2.0 build this runs against, so testing it
would force a release. Retrieval in the pane is Phase 4; the pairing token and
CORS allowlist are Phase 3, and must land before any tester sees this.
