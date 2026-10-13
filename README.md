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

## The Mac control experiment

If the localhost probes fail on the Mac, do this before concluding anything. In
**Safari** and then in **Chrome**, open

```
https://agpranaybio.github.io/standpoint-addin/taskpane.html
```

and read section 2 in each.

Safari is WebKit, Chrome is Chromium — the same two engines Word embeds. If the
probes fail in Safari and succeed in Chrome, the cause is WebKit's mixed-content
policy rather than anything about Word, and the add-in must be served over
`https://localhost` instead. Word's own webview console is awkward to reach on
Windows and effectively unavailable on Mac, so this comparison stands in for a
debugger.

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
