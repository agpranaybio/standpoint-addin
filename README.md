# Standpoint.Bio — Word task pane (Phase 2 spike)

The Standpoint task pane for Word. Static HTML and JS — no build step, so what
is sideloaded is what is in git.

**Standpoint serves these files itself**, over TLS on `https://127.0.0.1:8443`,
and answers the pane's requests from that same address. That is not a
convenience: it is the only arrangement that works, and it is also what makes
the security model simple.

It reads your selection and nothing else, and it **cannot write to your
document** — the manifest asks for `ReadDocument`, the lowest permission that
can read a selection, so Word itself would refuse an edit.

## Why the app serves it

Measured on Mac Word, Oct 14 2026, running the identical pane from two origins
in one sitting:

| pane served from | `http://…:8000` | `https://127.0.0.1:8443` | own origin |
| --- | --- | --- | --- |
| `https://127.0.0.1:8443` (the app) | FAILED 2 ms | **OK 200 · 19 ms** | OK 200 · 24 ms |
| `agpranaybio.github.io` (hosted) | FAILED 5 ms | **FAILED 45 ms** | OK 200 · 102 ms |

Plaintext is refused in 2–5 ms — faster than the 45 ms a real request costs, so
nothing travelled; a rule was checked. And the hosted pane's **secure** call
failed where the local pane's identical call succeeded, so it is not only
plaintext that is blocked: a public origin cannot reach this machine at all.

Events were measured at the same time and passed: `WordApi` 1.1–1.7, both event
APIs register, and paragraph events arrive about **0.87 per second** while
typing. Word coalesces internally, so live hints need no debouncing of their own.

## What that bought

Because the pane comes from the app, its requests are **same-origin**. Nothing
to pair, no token, and no CORS entry — and that is what allowed the backend to
drop `allow_origins=["*"]`, which until this change let **any website the
scientist visited read their entire library**. Demonstrated from a public origin
in Chrome, before and after:

| | before | after |
| --- | --- | --- |
| `/hypothesis` | 200, 909 bytes | blocked |
| `/papers/23` | 200, 1880 bytes | blocked |
| `/project/tree?hypothesis_id=2` | 200, **21320 bytes** | blocked |

## What is verified, and what is not

**The packaged code path works** — exercised in a real packaged build (on
Linux, which uses the same asar archive and the same custom-scheme machinery as
the Mac build):

| | |
| --- | --- |
| window origin | `standpoint-app://app` |
| secure context | true |
| renderer read out of `app.asar` | React rendered, 3 stylesheets, real UI |
| API from that origin, through the guard | `GET` 200, `POST` 200, **0 refusals** |

That closes the "never verified in a packaged build" risk for the scheme, the
asar read and the allowlist.

**What is Mac-only and still unverified:** whether trusting the certificate in
the login keychain satisfies WKWebView inside Word. It needs macOS, a GUI
session and a human at a password prompt, so no CI runner and no Linux machine
can answer it. `verify_mac_trust.sh` walks it end to end — including pressing
Cancel once on purpose, and asking the *system* whether it trusts the
certificate rather than asking our own record.

## Installing it

1. **Start Standpoint.** The TLS listener and its certificate come up with the
   backend; the log line reads `[addin] task pane on https://127.0.0.1:8443/…`.
2. **Trust the certificate, once.** Generated per machine into the app's data
   dir as `addin-cert.pem`; the private key never leaves that machine and is
   never shipped. On the Mac this took one step.

   ```sh
   sudo security add-trusted-cert -d -r trustRoot \
     -k /Library/Keychains/System.keychain "<data dir>/addin-cert.pem"
   ```

   **Until it is trusted the pane is BLANK** — a rejected certificate produces
   no error text anywhere, which is the most confusing failure in this feature.
3. **Sideload `manifest-local.xml`.** On macOS copy it into
   `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/` and restart
   Word; on Windows share the folder and add it as a trusted add-in catalogue
   (*File ▸ Options ▸ Trust Center ▸ Trusted Add-in Catalogs*). It appears under
   **Home ▸ Standpoint**.

Steps 2 and 3 are what the "Install Word add-in" button will do in one flow —
see Phase 5 in the add-in plan.

## Testing it on Linux, without Word

Desktop Word does not exist on Linux, so the pane has two ways to be exercised
here.

### Browser test mode — no Word at all

Open the pane in Chrome:

```
http://127.0.0.1:8000/addin/taskpane.html
```

With no Word host, a **textarea appears and stands in for the document**. Find
evidence, live hints (type and pause), the cards, "+N more", the labels, the
literature flow and every state all behave exactly as they do in Word, because
they run the same code paths — `currentText()` simply reads the box instead of
the selection.

Plain `http` is fine for this: the app serves the pane on both doors, and a
browser needs no TLS to talk to itself. TLS is Word's requirement, not Chrome's.

### Forcing a state

Every state has a button under **Diagnostics**, and a URL:

```
…/addin/taskpane.html?state=indexing      # "Reading your library… 46%"
…/addin/taskpane.html?state=no-matches    # + "Show weaker results"
…/addin/taskpane.html?state=too-old       # contract mismatch
…/addin/taskpane.html?state=offline       # Standpoint not running
…/addin/taskpane.html?state=no-model      # a build with no embedding model
…/addin/taskpane.html?state=ready
```

States nobody can look at are states nobody checks — contriving a half-built
index to see one message is how they end up never reviewed.

### Trusting the dev certificate in Chrome on Linux

**Skip this unless you are using Word on the web.** Browser test mode above
runs on the plain door (`http://…:8000`) and needs no certificate at all. The
certificate exists because *Word* refuses a plain-http pane; Chrome does not
care.

Only needed for the **https** door (`:8443`), which is what Word on the web
loads. Chrome on Linux uses its own NSS store, not the system one:

```sh
certutil -d sql:$HOME/.pki/nssdb -A -t "C,," \
  -n "Standpoint.Bio dev" -i backend/addin-cert.pem
```

`-t "C,,"` trusts it for TLS servers only. The certificate is a CA with
`localhost`, `127.0.0.1` and `::1` in its SAN, generated per machine; the
private key never leaves it. Needs `libnss3-tools` (`certutil` is already on
this machine). **Restart Chrome afterwards.** To undo:

```sh
certutil -d sql:$HOME/.pki/nssdb -D -n "Standpoint.Bio dev"
```

Check it took:

```sh
curl https://127.0.0.1:8443/health      # no -k; a 200 means the system trusts it
```

### Word on the web

Word on the web runs in Chrome on Linux, and is the only Word available here.
`manifest-local.xml` is an XML manifest with an `https` SourceLocation, which is
the shape Word on the web accepts for sideloading.

1. Trust the certificate (above) and **restart Chrome**.
2. Visit `https://127.0.0.1:8443/addin/taskpane.html` once and confirm it loads
   with no warning. If Chrome still objects, Word on the web will fail the same
   way and the pane will be blank.
3. In Word on the web: **Home ▸ Add-ins ▸ More Add-ins ▸ My Add-ins ▸ Upload My
   Add-in**, and choose `addin/manifest-local.xml`.
4. The pane appears under **Home ▸ Standpoint**.

**Not verified end to end, and worth saying so.** Step 3 needs a Microsoft
account, which is yours, not mine — everything up to it is checked here (the
manifest validates, the TLS door serves the pane, the certificate covers both
names). Two things to expect if it does not work: Word on the web may refuse a
private-IP SourceLocation even over https, and its iframe sandbox differs from
desktop Word's. If it refuses, that is a fact about Word on the web, **not**
about the Mac desktop path measured in Phase 2 — the browser test mode above
still exercises everything except Word's own event delivery.

## Reading it

The pane leads with one line: **Connected to Standpoint**, or that Standpoint
isn't running. The probe rows that made this phase debuggable are still there,
folded under **Diagnostics**, with a **Copy report** button.

For debugging in a real browser with devtools, the app serves the same pane on
its plain door: `http://127.0.0.1:8000/addin/taskpane.html`. Word cannot use
that one — Office refuses a plain-http `SourceLocation`.

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
| `manifest.xml` | Loads the pane from GitHub Pages. **Kept only for the browser control experiment below** — measured not to work in Word. The XML manifest, not the unified JSON one: that is Windows/M365-first and sideloading it into Mac desktop Word is not dependable. |
| `manifest-local.xml` | **The real manifest.** Loads the pane from the app at `https://127.0.0.1:8443/addin/`. |
| `taskpane.html` | One screen, a row per check. |
| `taskpane.js` | The checks. No framework and no build step, so what is sideloaded is what is in git. |
| `assets/icon-*.png` | Ribbon icons. |
| `check_manifest.py` | Run before publishing: `python check_manifest.py`. Checks both manifests and catches the mistakes whose only symptom in Word is a pane that silently never appears — including the two sharing a GUID, which would install only one. |

## Hosting

Authored in the private app repo under `addin/` and published to the public
`standpoint-addin` repo with `git subtree push --prefix=addin`. **The published
copy is no longer the delivery route** — Standpoint ships these files and serves
them itself. The public copy stays for now because the control experiment above
needs a public origin to test against; it goes away once that is no longer
useful.

## Not here

- **Live hints, result cards, "no strong matches", deep links** — Phase 4. The
  pane still shows diagnostics rather than evidence.
- **The install button** — Phase 5. Until then the certificate trust and the
  sideload are the two manual steps above.
- **A pairing code or token** — dropped deliberately. The pane is same-origin
  with the API, so there is nothing to pair; a token would guard only against
  other programs on this machine, which can open the database file directly.
