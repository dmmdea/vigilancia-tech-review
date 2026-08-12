#!/usr/bin/env python3
"""List a link-shared (public) Google Drive folder anonymously.

Usage:
    python drive_list.py <folder-url-or-id>

Prints JSON to stdout:
    {"folder_id": "...", "folder_name": "...", "entries": [
        {"id": "...", "name": "...", "kind": "folder|file|gslides|gdoc|gsheet|other",
         "href": "..."}]}

Exit codes: 0 ok · 2 folder not accessible (not public / not found) · 3 parse error.
No credentials required; works for any folder shared as "anyone with the link".
"""
import html as html_mod
import json
import re
import sys
import urllib.error
import urllib.request

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) vigilancia-tech-review/1.0"


def extract_folder_id(arg: str) -> str:
    m = re.search(r"/folders/([\w-]+)", arg)
    if m:
        return m.group(1)
    m = re.search(r"[?&]id=([\w-]+)", arg)
    if m:
        return m.group(1)
    if re.fullmatch(r"[\w-]{10,}", arg):
        return arg
    raise ValueError(f"Cannot extract a Drive folder id from: {arg}")


def classify(href: str, name: str) -> str:
    if "/drive/folders/" in href or "/folderview" in href:
        return "folder"
    if "docs.google.com/presentation/" in href:
        return "gslides"
    if "docs.google.com/document/" in href:
        return "gdoc"
    if "docs.google.com/spreadsheets/" in href:
        return "gsheet"
    if "/file/d/" in href or "drive.google.com/open" in href:
        return "file"
    return "other"


def list_folder(folder_id: str) -> dict:
    url = f"https://drive.google.com/embeddedfolderview?id={folder_id}"
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            page = resp.read().decode("utf-8", errors="replace")
            final_url = resp.geturl()
    except urllib.error.HTTPError as e:
        print(f"ERROR: folder {folder_id} not accessible (HTTP {e.code}). "
              "Is it shared as 'anyone with the link'?", file=sys.stderr)
        sys.exit(2)
    if "accounts.google.com" in final_url:
        print(f"ERROR: folder {folder_id} requires sign-in (not public).",
              file=sys.stderr)
        sys.exit(2)

    title_m = re.search(r"<title>([^<]*)</title>", page)
    folder_name = html_mod.unescape(title_m.group(1)) if title_m else ""

    if "flip-entries" not in page:
        print("ERROR: unexpected page layout from Drive (no entry container). "
              "The embeddedfolderview endpoint may have changed.", file=sys.stderr)
        sys.exit(3)

    entry_re = re.compile(
        r'<div class="flip-entry" id="entry-([\w-]+)".*?<a href="([^"]+)".*?'
        r'flip-entry-title">([^<]*)</div>',
        re.S,
    )
    entries = []
    for eid, href, title in entry_re.findall(page):
        name = html_mod.unescape(title).strip()
        href = html_mod.unescape(href)
        entries.append(
            {"id": eid, "name": name, "kind": classify(href, name), "href": href}
        )
    return {"folder_id": folder_id, "folder_name": folder_name, "entries": entries}


def main() -> None:
    if len(sys.argv) != 2:
        print(__doc__, file=sys.stderr)
        sys.exit(1)
    folder_id = extract_folder_id(sys.argv[1])
    result = list_folder(folder_id)
    json.dump(result, sys.stdout, ensure_ascii=False, indent=2)
    print()


if __name__ == "__main__":
    main()
