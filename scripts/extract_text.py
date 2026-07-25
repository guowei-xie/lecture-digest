#!/usr/bin/env python3
"""Extract plain text from lecture/meeting source files — stdlib only, no pip installs.

Handles the formats knowledge-review material usually arrives in:
  .docx / .pptx  -> unzipped and read from the underlying XML (works even when
                    python-docx / python-pptx aren't installed)
  .txt / .md     -> read as UTF-8
  .vtt / .srt    -> subtitle/caption transcripts, timestamps stripped

Usage:
  python3 extract_text.py <file> [<file> ...]
  python3 extract_text.py meeting.docx > transcript.txt

Prints the extracted text to stdout. With multiple files, each is prefixed by a
"===== filename =====" banner so you can tell sources apart.
"""
import re
import sys
import zipfile


def _xml_to_text(xml: str, para_tag: str) -> str:
    """Pull visible text out of Office XML, inserting a newline per paragraph."""
    # Normalize paragraph boundaries first so we don't glue sentences together.
    xml = re.sub(r"</w:p>|</a:p>", "\n", xml)
    xml = re.sub(r"<w:br[^>]*/>|<a:br[^>]*/>", "\n", xml)
    # Grab the run-text nodes (<w:t> for Word, <a:t> for PowerPoint).
    chunks = re.findall(r"<(?:w|a):t[^>]*>(.*?)</(?:w|a):t>", xml, re.S)
    text = "".join(chunks)
    # Unescape the handful of XML entities Office actually emits.
    for a, b in (("&amp;", "&"), ("&lt;", "<"), ("&gt;", ">"),
                 ("&quot;", '"'), ("&apos;", "'")):
        text = text.replace(a, b)
    # Collapse runs of blank lines the paragraph splitting introduced.
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def extract(path: str) -> str:
    low = path.lower()
    if low.endswith(".docx"):
        with zipfile.ZipFile(path) as z:
            return _xml_to_text(z.read("word/document.xml").decode("utf-8", "ignore"), "w:p")
    if low.endswith(".pptx"):
        parts = []
        with zipfile.ZipFile(path) as z:
            # Slides are numbered; sort so the deck reads in order.
            names = sorted(
                (n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)),
                key=lambda n: int(re.search(r"(\d+)", n).group(1)),
            )
            for i, n in enumerate(names, 1):
                body = _xml_to_text(z.read(n).decode("utf-8", "ignore"), "a:p")
                if body:
                    parts.append(f"--- Slide {i} ---\n{body}")
        return "\n\n".join(parts)
    if low.endswith((".vtt", ".srt")):
        with open(path, encoding="utf-8", errors="ignore") as f:
            lines = []
            for ln in f:
                s = ln.strip()
                # Drop cue numbers, timestamps, and the WEBVTT header.
                if not s or s.isdigit() or "-->" in s or s.upper().startswith("WEBVTT"):
                    continue
                lines.append(s)
            return "\n".join(lines)
    # Plain text / markdown / anything else: read as UTF-8.
    with open(path, encoding="utf-8", errors="ignore") as f:
        return f.read()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    multi = len(sys.argv) > 2
    for p in sys.argv[1:]:
        if multi:
            print(f"===== {p} =====")
        try:
            print(extract(p))
        except Exception as e:  # noqa: BLE001 — surface the reason, keep going
            print(f"[extract_text: could not read {p}: {e}]", file=sys.stderr)
        if multi:
            print()
