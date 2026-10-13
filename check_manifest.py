#!/usr/bin/env python
"""Check manifest.xml before it costs a trip to a test machine.

Sideloading is manual and slow, and Word's failure mode for a bad manifest is a
pane that silently never appears -- no error, nothing to read. Every mistake
this catches is otherwise discovered on the other side of the room.

The drift it really exists for: the hosting URL appears EIGHT times across
SourceLocation, two icon URLs, the ribbon's three images and the overrides' two
URLs. Changing the host in seven of eight places leaves a manifest that loads
and then shows a broken icon, or loads an old pane.

    python addin/check_manifest.py          # exit 0 = pass
"""

from __future__ import annotations

import re
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

HERE = Path(__file__).resolve().parent
MANIFEST = HERE / "manifest.xml"
NS = {
    "o": "http://schemas.microsoft.com/office/appforoffice/1.1",
    "bt": "http://schemas.microsoft.com/office/officeappbasictypes/1.0",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
}
GUID = re.compile(r"^[0-9A-Fa-f]{8}-(?:[0-9A-Fa-f]{4}-){3}[0-9A-Fa-f]{12}$")

failures: list[str] = []


def check(ok: bool, msg: str) -> None:
    if ok:
        print(f"  ok    {msg}")
    else:
        failures.append(msg)
        print(f"  FAIL  {msg}")


def main() -> int:
    if not MANIFEST.exists():
        print(f"no manifest at {MANIFEST}")
        return 1
    raw = MANIFEST.read_text(encoding="utf-8")
    try:
        root = ET.fromstring(raw)
    except ET.ParseError as exc:
        print(f"  FAIL  manifest is not well-formed XML: {exc}")
        return 1
    print(f"checking {MANIFEST.name}")

    check(root.tag == f"{{{NS['o']}}}OfficeApp", "root element is OfficeApp")
    check(root.get(f"{{{NS['xsi']}}}type") == "TaskPaneApp", 'xsi:type is "TaskPaneApp"')

    # Word requires each of these; a missing one is a pane that never appears.
    for tag in ("Id", "Version", "ProviderName", "DefaultLocale", "DisplayName",
                "Description", "IconUrl", "SupportUrl", "Hosts",
                "DefaultSettings", "Permissions"):
        check(root.find(f"o:{tag}", NS) is not None, f"<{tag}> present")

    el = root.find("o:Id", NS)
    check(el is not None and bool(GUID.match((el.text or "").strip())),
          "<Id> is a GUID")

    el = root.find("o:Version", NS)
    check(el is not None and bool(re.match(r"^\d+(\.\d+){1,3}$", (el.text or "").strip())),
          "<Version> is numeric dotted")

    # v0 NEVER edits the document. ReadDocument is what makes that a guarantee
    # the host enforces rather than a promise the JavaScript keeps, so widening
    # it is a product decision and must not happen by accident.
    el = root.find("o:Permissions", NS)
    check(el is not None and (el.text or "").strip() == "ReadDocument",
          "<Permissions> is ReadDocument (the pane cannot write)")

    hosts = [h.get("Name") for h in root.findall("o:Hosts/o:Host", NS)]
    check(hosts == ["Document"], f"host is Word only, got {hosts}")

    # Every URL must be https. Office refuses http SourceLocation, and an http
    # icon would be mixed content inside the pane's own chrome.
    urls: list[tuple[str, str]] = []
    for el in root.iter():
        tag = el.tag.split("}")[-1]
        val = el.get("DefaultValue")
        if val and ("://" in val):
            urls.append((tag, val))
        if tag == "SourceLocation" and el.text and "://" in el.text:
            urls.append((tag, el.text.strip()))
    check(bool(urls), "manifest carries URLs")
    bad = [f"{t}={u}" for t, u in urls if not u.startswith("https://")]
    check(not bad, f"every URL is https ({len(urls)} checked)" if not bad
          else f"non-https URLs: {bad}")

    # THE DRIFT CHECK. All the add-in's own assets must share one origin; a
    # stray old host loads a stale pane or a missing icon, silently.
    own = [u for _t, u in urls if "github.io" in u]
    origins = {re.match(r"^(https://[^/]+/[^/]+/)", u).group(1)
               for u in own if re.match(r"^(https://[^/]+/[^/]+/)", u)}
    check(len(origins) <= 1,
          f"all hosted assets share one base, got {sorted(origins)}")

    # The files those URLs name must actually exist next to this manifest, or
    # Pages serves a 404 and Word shows an empty pane.
    missing = []
    for u in own:
        rel = u.split("/", 4)[-1].split("?")[0]
        if rel and not (HERE / rel).exists():
            missing.append(rel)
    check(not missing, f"every referenced file exists locally"
          if not missing else f"referenced but missing: {sorted(set(missing))}")

    src = root.find("o:DefaultSettings/o:SourceLocation", NS)
    check(src is not None and (src.get("DefaultValue") or "").endswith(".html"),
          "<SourceLocation> points at an .html file")

    print(f"\n{'FAILED: ' + str(len(failures)) if failures else 'PASS'}"
          + ("" if not failures else " problem(s) — Word would fail silently"))
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
