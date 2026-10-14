#!/usr/bin/env python
"""Check the manifests before they cost a trip to a test machine.

Sideloading is manual and slow, and Word's failure mode for a bad manifest is a
pane that silently never appears -- no error, nothing to read. Every mistake
caught here is otherwise discovered on the other side of the room.

    python addin/check_manifest.py                 # every manifest here
    python addin/check_manifest.py manifest.xml    # just one

TWO MANIFESTS, ONE PANE. manifest.xml loads the pane from GitHub Pages;
manifest-local.xml loads the identical pane from https://127.0.0.1:8443, served
by the app itself. They exist to be sideloaded together and compared, which is
how the Mac refusal gets attributed to a cause rather than guessed at.

The drift this exists for: each manifest repeats its own hosting base EIGHT
times -- SourceLocation, two icon URLs, the ribbon's three images, and two more
in the overrides. Changing seven of eight leaves a manifest that loads a stale
pane or a broken icon, with nothing to say so. And with two manifests there is
now a second way to go wrong: mixing one's base into the other.
"""

from __future__ import annotations

import re
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

HERE = Path(__file__).resolve().parent
NS = {
    "o": "http://schemas.microsoft.com/office/appforoffice/1.1",
    "bt": "http://schemas.microsoft.com/office/officeappbasictypes/1.0",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
}
GUID = re.compile(r"^[0-9A-Fa-f]{8}-(?:[0-9A-Fa-f]{4}-){3}[0-9A-Fa-f]{12}$")

# Every file a manifest must point at. Listed rather than inferred: a typo'd
# URL simply stops matching any local file, and without an expected set that
# would read as "no assets referenced" and pass.
REQUIRED_ASSETS = {
    "taskpane.html",
    "assets/icon-16.png", "assets/icon-32.png",
    "assets/icon-64.png", "assets/icon-80.png",
}


def split_asset(url: str) -> tuple[str | None, str | None]:
    """Split a URL into (hosting base, path) when it names a file we ship.

    Works for either base without being told which -- the Pages one has a repo
    segment in the path, the local one does not -- by finding the longest
    trailing portion that exists on disk next to this manifest.
    """
    if "://" not in url:
        return None, None
    path = url.split("://", 1)[1].split("?")[0]
    parts = path.split("/")
    for i in range(1, len(parts)):
        rel = "/".join(parts[i:])
        if rel and (HERE / rel).is_file():
            return url[: len(url) - len(rel)], rel
    return None, None


def check_one(manifest: Path) -> list[str]:
    failures: list[str] = []

    def check(ok: bool, msg: str) -> None:
        if ok:
            print(f"  ok    {msg}")
        else:
            failures.append(msg)
            print(f"  FAIL  {msg}")

    print(f"\n{manifest.name}")
    try:
        root = ET.fromstring(manifest.read_text(encoding="utf-8"))
    except ET.ParseError as exc:
        print(f"  FAIL  not well-formed XML: {exc}")
        return [f"{manifest.name}: malformed"]

    check(root.tag == f"{{{NS['o']}}}OfficeApp", "root element is OfficeApp")
    check(root.get(f"{{{NS['xsi']}}}type") == "TaskPaneApp", 'xsi:type is "TaskPaneApp"')

    for tag in ("Id", "Version", "ProviderName", "DefaultLocale", "DisplayName",
                "Description", "IconUrl", "SupportUrl", "Hosts",
                "DefaultSettings", "Permissions"):
        check(root.find(f"o:{tag}", NS) is not None, f"<{tag}> present")

    el = root.find("o:Id", NS)
    check(el is not None and bool(GUID.match((el.text or "").strip())), "<Id> is a GUID")

    el = root.find("o:Version", NS)
    check(el is not None and bool(re.match(r"^\d+(\.\d+){1,3}$", (el.text or "").strip())),
          "<Version> is numeric dotted")

    # v0 NEVER edits the document. ReadDocument is what makes that a guarantee
    # Word enforces rather than a promise the JavaScript keeps, so widening it
    # is a product decision and must not happen by accident.
    el = root.find("o:Permissions", NS)
    check(el is not None and (el.text or "").strip() == "ReadDocument",
          "<Permissions> is ReadDocument (the pane cannot write)")

    hosts = [h.get("Name") for h in root.findall("o:Hosts/o:Host", NS)]
    check(hosts == ["Document"], f"host is Word only, got {hosts}")

    urls: list[tuple[str, str]] = []
    for el in root.iter():
        tag = el.tag.split("}")[-1]
        val = el.get("DefaultValue")
        if val and "://" in val:
            urls.append((tag, val))
        if tag == "SourceLocation" and el.text and "://" in el.text:
            urls.append((tag, el.text.strip()))
    check(bool(urls), "manifest carries URLs")

    # Office refuses an http SourceLocation, and an http icon would be mixed
    # content inside the pane's own chrome.
    bad = [f"{t}={u}" for t, u in urls if not u.startswith("https://")]
    check(not bad, f"every URL is https ({len(urls)} checked)" if not bad
          else f"non-https URLs: {bad}")

    bases, seen = set(), set()
    for _t, u in urls:
        base, rel = split_asset(u)
        if base:
            bases.add(base)
            seen.add(rel)

    check(len(bases) == 1, f"all assets share ONE hosting base, got {sorted(bases)}")
    missing = REQUIRED_ASSETS - seen
    check(not missing,
          f"all {len(REQUIRED_ASSETS)} shipped files are referenced and exist"
          if not missing else
          f"not referenced, or pointing at a file that does not exist: {sorted(missing)}")

    src = root.find("o:DefaultSettings/o:SourceLocation", NS)
    src_url = (src.get("DefaultValue") or "") if src is not None else ""
    check(src_url.endswith(".html"), "<SourceLocation> points at an .html file")
    if bases:
        check(src_url.startswith(next(iter(bases))),
              "<SourceLocation> uses the same base as the icons")

    return failures


def main() -> int:
    if len(sys.argv) > 1:
        targets = [HERE / a if not Path(a).is_absolute() else Path(a) for a in sys.argv[1:]]
    else:
        targets = sorted(HERE.glob("manifest*.xml"))
    if not targets:
        print(f"no manifests in {HERE}")
        return 1

    failures: list[str] = []
    for m in targets:
        if not m.exists():
            print(f"\n{m}\n  FAIL  does not exist")
            failures.append(str(m))
            continue
        failures += check_one(m)

    # The two manifests must NOT collide in Word. Same GUID and the second
    # sideload silently replaces the first, which would look like the local
    # copy "not working" when it was never installed.
    ids = {}
    for m in targets:
        if not m.exists():
            continue
        try:
            el = ET.fromstring(m.read_text(encoding="utf-8")).find("o:Id", NS)
            if el is not None:
                ids.setdefault((el.text or "").strip(), []).append(m.name)
        except ET.ParseError:
            pass
    clash = {k: v for k, v in ids.items() if len(v) > 1}
    if len(targets) > 1:
        print()
        if clash:
            failures.append("duplicate GUIDs")
            print(f"  FAIL  manifests share a GUID — Word would install only one: {clash}")
        else:
            print(f"  ok    each manifest has its own GUID ({len(ids)} distinct)")

    print(f"\n{'FAILED: ' + str(len(failures)) + ' problem(s) — Word would fail silently' if failures else 'PASS'}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
