#!/usr/bin/env python3
"""Download a file from a link-shared Google Drive anonymously.

Usage:
    python drive_download.py <file-id> <output-path> [--kind file|gslides]

- kind=file    : binary download (pptx, pdf, ...) via drive.usercontent.google.com
- kind=gslides : Google Slides native file, exported as .pptx

Verifies the payload is a real PPTX/PDF (magic bytes), not an HTML error page.
Exit codes: 0 ok · 2 download failed / not accessible · 3 payload is not a document.
"""
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) vigilancia-tech-review/1.0"
CHUNK = 1 << 20
ID_RE = re.compile(r"[\w-]{10,}\Z")
ALLOWED_HOSTS = (".google.com", ".googleusercontent.com")


def safe_google_url(url: str) -> bool:
    p = urllib.parse.urlparse(url)
    return p.scheme == "https" and (
        p.hostname or "").endswith(ALLOWED_HOSTS)


def fetch(url: str) -> tuple[bytes, str]:
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=300) as resp:
        ctype = resp.headers.get("Content-Type", "")
        data = b""
        while True:
            chunk = resp.read(CHUNK)
            if not chunk:
                break
            data += chunk
        return data, ctype


def confirm_url_from_interstitial(page: str) -> str | None:
    """Large files get a virus-scan interstitial form; rebuild its action URL."""
    form = re.search(r'<form[^>]+action="([^"]+)"(.*?)</form>', page, re.S)
    if not form:
        return None
    action, body = form.group(1), form.group(2)
    params = dict(re.findall(r'name="([^"]+)"\s+value="([^"]*)"', body))
    return action + "?" + urllib.parse.urlencode(params)


def looks_like_document(data: bytes) -> bool:
    return data[:4] == b"PK\x03\x04" or data[:5] == b"%PDF-" or data[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def is_html(data: bytes) -> bool:
    head = data[:32].lstrip(b"\xef\xbb\xbf \t\r\n").lower()
    return head.startswith((b"<!doctype", b"<html"))


def html_title(data: bytes) -> str:
    m = re.search(rb"<title>([^<]*)</title>", data[:4096], re.I)
    return ("page title: " + m.group(1).decode("utf-8", "replace").strip()
            if m else "untitled HTML page")


def download(file_id: str, out_path: str, kind: str) -> None:
    if not ID_RE.fullmatch(file_id):
        print(f"ERROR: '{file_id}' is not a valid Drive file id.", file=sys.stderr)
        sys.exit(1)
    if kind == "gslides":
        url = f"https://docs.google.com/presentation/d/{file_id}/export/pptx"
    else:
        url = (f"https://drive.usercontent.google.com/download?id={file_id}"
               "&export=download&confirm=t")
    try:
        data, ctype = fetch(url)
        if is_html(data):
            retry = confirm_url_from_interstitial(data.decode("utf-8", "replace"))
            if retry is None:
                print(f"ERROR: Drive returned an HTML page with no download form "
                      f"for {file_id} ({html_title(data)}). Possible quota limit "
                      "or permission problem.", file=sys.stderr)
                sys.exit(2)
            retry = urllib.parse.urljoin(url, retry)  # form actions can be relative
            if not safe_google_url(retry):
                print(f"ERROR: Drive interstitial pointed to a non-Google URL, "
                      f"refusing to follow: {retry}", file=sys.stderr)
                sys.exit(2)
            data, ctype = fetch(retry)
    except urllib.error.HTTPError as e:
        print(f"ERROR: download failed for {file_id} (HTTP {e.code}).", file=sys.stderr)
        sys.exit(2)
    except OSError as e:  # URLError, resets, timeouts, incomplete reads
        print(f"ERROR: network error for {file_id}: {e}", file=sys.stderr)
        sys.exit(2)

    if not looks_like_document(data):
        detail = html_title(data) if is_html(data) else f"content-type {ctype}"
        print(f"ERROR: payload for {file_id} is not a PPTX/PDF ({detail}). "
              "File may not be public, requires sign-in, or hit a quota page.",
              file=sys.stderr)
        sys.exit(3)

    parent = os.path.dirname(os.path.abspath(out_path))
    os.makedirs(parent, exist_ok=True)
    with open(out_path, "wb") as f:
        f.write(data)
    print(f"OK {out_path} {len(data)} bytes")


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    kind = "file"
    for a in sys.argv[1:]:
        if a.startswith("--kind"):
            kind = a.split("=", 1)[1] if "=" in a else ""
    if kind not in ("file", "gslides"):
        print(f"ERROR: invalid --kind '{kind}' (use file or gslides).",
              file=sys.stderr)
        sys.exit(1)
    if len(args) != 2:
        print(__doc__, file=sys.stderr)
        sys.exit(1)
    download(args[0], args[1], kind)


if __name__ == "__main__":
    main()
