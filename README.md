# Standpoint.Bio — Word task pane (Phase 2 spike)

A diagnostic pane, not a feature. It answers the two questions that decide how
the real add-in is built, and neither can be answered on a Linux dev machine:

1. **Can a pane served from public HTTPS reach `http://localhost:8000`?**
   Windows Word runs WebView2 (Chromium), Mac Word runs WebKit, and the two may
   differ on whether a plain-HTTP localhost request from an HTTPS page counts as
   mixed content. If it is blocked, the pane has to be served by the backend
   over `https://localhost` with a locally trusted certificate instead.
2. **Do events fire while you type, and in what pattern?** Per keystroke, per
   word and per paragraph are three different products. This decides live hints
   versus the fallback of refreshing when the cursor reaches a new paragraph.

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
cp manifest.xml ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
```

Create the `wef` folder if it is not there, then restart Word. The add-in
appears under **Home ▸ Standpoint ▸ Diagnostics**.

**Windows**

1. Put `manifest.xml` in a folder and share that folder (right-click ▸
   *Properties ▸ Sharing ▸ Share*).
2. Word ▸ **File ▸ Options ▸ Trust Center ▸ Trust Center Settings ▸ Trusted
   Add-in Catalogs**. Paste the share path (`\\MACHINE\folder`), click **Add
   catalog**, tick **Show in Menu**, OK.
3. Restart Word. **Insert ▸ My Add-ins ▸ Shared Folder ▸ Standpoint (spike)**.

## Running it

1. Start Standpoint, so the backend is listening on `127.0.0.1:8000`.
2. Open a document with a few paragraphs of real prose.
3. Open the pane. Sections 1 and 2 fill in by themselves.
4. **Select a sentence.** Section 3 should show its text.
5. **Click into a paragraph and type for about 20 seconds.** Watch the two
   counters in section 4. What matters is not that they move but *how* — the
   median gap distinguishes per-keystroke from per-word from per-paragraph.
6. Press **Copy report** and send the block back. If the clipboard is refused —
   embedded webviews sometimes do — select the text in the box and copy by hand.

## What Chromium already did, before Word is involved

Measured on the Linux dev machine, Chrome 149, against this very Pages URL:

| page origin | fetch to | result |
| --- | --- | --- |
| `https://agpranaybio.github.io` | `http://localhost:8000` | **hung, then timed out** |
| `https://agpranaybio.github.io` | `http://127.0.0.1:8111` | **hung, then timed out** |
| `https://agpranaybio.github.io` | itself | HTTP 200 in 25 ms |
| `http://127.0.0.1:9123` | `http://localhost:8000` | HTTP 200 in 100 ms |

A request-logging server on the target port recorded **nothing** from the
browser — not even a preflight `OPTIONS`. So the request never left Chrome.
That rules out CORS and rules out a Private Network Access preflight, both of
which would have put something in that log; and it is not classic mixed content
either, because `http://localhost` is "potentially trustworthy" under that spec
and is supposed to be exempt. What remains is Chromium gating requests from a
public page into the local network — and in an automated browser with no way to
show a permission prompt, hanging is what that looks like.

**So Option A is now in doubt on Windows too, not only on Mac**, which is the
opposite of what this phase assumed. That makes both runs below worth doing
rather than one: WebView2 inside Word is not Chrome-with-extensions and may not
apply the same gate.

**If you see the two localhost rows time out, the pane is working correctly.**
That is the measurement, not a bug. Watch for a permission prompt from Word or
the system — if one appears, say so in the report, because a prompt means
Option A survives with a one-time grant.

## The control experiment, on both machines

Do this on each machine before concluding anything. On the Mac open
**Safari** and **Chrome**; on Windows open **Edge** and **Chrome**. In each,
go to

```
https://agpranaybio.github.io/standpoint-addin/taskpane.html
```

and read section 2 in each.

Safari is WebKit and Edge/Chrome are Chromium — the same engines Word embeds.
This separates "the browser engine blocks it" from "Word blocks it": if a plain
browser on that machine behaves exactly like the pane inside Word, the cause is
the engine's policy and no amount of add-in code will change it. Word's own
webview console is awkward to reach on Windows and effectively unavailable on
Mac, so this comparison stands in for a debugger.

Note the Linux result above: a desktop Chrome already blocks this. If a browser
on your machine blocks it too but Word does **not**, that is the interesting
finding and it keeps Option A alive.

## Files

| | |
| --- | --- |
| `manifest.xml` | The XML manifest. Not the unified JSON one: that is Windows/M365-first and sideloading it into Mac desktop Word is not dependable. |
| `taskpane.html` | One screen, a row per check. |
| `taskpane.js` | The checks. No framework and no build step, so what is sideloaded is what is in git. |
| `assets/icon-*.png` | Ribbon icons. |
| `check_manifest.py` | Run before publishing: `python check_manifest.py`. Catches the mistakes whose only symptom in Word is a pane that silently never appears. |

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
