#!/usr/bin/env python3
"""Scan a URL for a keyword and generate a PDF report.

Usage:
    python scan_url_to_pdf.py https://example.com "keyword"
"""

from __future__ import annotations

import argparse
import datetime as dt
import gzip
import html
import re
import textwrap
import time
import zlib
from collections import deque
from dataclasses import dataclass
from html.parser import HTMLParser
from pathlib import Path
from typing import Iterable
from urllib.error import HTTPError, URLError
from urllib.parse import urldefrag, urljoin, urlparse
from urllib.request import Request, urlopen

import pandas as pd


# Use a browser-like UA by default; some sites block generic programmatic UAs.
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/123.0.0.0 Safari/537.36"
)
PAGE_WIDTH = 595
PAGE_HEIGHT = 842
MARGIN = 54
LINE_HEIGHT = 14
FONT_SIZE = 10


@dataclass
class PageMatch:
    url: str
    title: str
    count: int
    snippets: list[str]


@dataclass
class ScanError:
    url: str
    message: str


class ContentParser(HTMLParser):
    """Extract visible-ish text and links from an HTML document."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.title = ""
        self._in_title = False
        self._skip_depth = 0
        self._text_parts: list[str] = []
        self.links: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.lower()
        if tag in {"script", "style", "noscript", "svg", "canvas"}:
            self._skip_depth += 1
        if tag == "title":
            self._in_title = True
        if tag == "a":
            href = dict(attrs).get("href")
            if href:
                self.links.append(href)
                
        if tag == "img":
            alt = dict(attrs).get("alt")
            if alt:
                self._text_parts.append(alt)
        
        
        if tag in {"p", "br", "div", "section", "article", "li", "tr", "h1", "h2", "h3"}:
            self._text_parts.append("\n")

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if tag in {"script", "style", "noscript", "svg", "canvas"} and self._skip_depth:
            self._skip_depth -= 1
        if tag == "title":
            self._in_title = False

    def handle_data(self, data: str) -> None:
        if self._skip_depth:
            return
        if self._in_title:
            self.title += data.strip() + " "
        self._text_parts.append(" " + data.strip() + " ")

    @property
    def text(self) -> str:
        decoded = html.unescape(" ".join(self._text_parts))
        return re.sub(r"\s+", " ", decoded).strip()


def normalize_url(url: str) -> str:
    parsed = urlparse(url if "://" in url else f"https://{url}")
    cleaned = parsed._replace(fragment="")
    return cleaned.geturl()


def is_http_url(url: str) -> bool:
    return urlparse(url).scheme in {"http", "https"}


def is_same_site(candidate: str, base_netloc: str) -> bool:
    parsed = urlparse(candidate)
    return parsed.scheme in {"http", "https"} and parsed.netloc == base_netloc


def fetch_html(url: str, timeout: int) -> tuple[str, str]:
    request = Request(
        url,
        headers={
            "User-Agent": USER_AGENT,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9",
            "Accept-Encoding": "gzip, deflate",
            "Connection": "close",
        },
    )
    with urlopen(request, timeout=timeout) as response:
        content_type = response.headers.get("Content-Type", "")
        if "text/html" not in content_type and "application/xhtml+xml" not in content_type:
            raise ValueError(f"skipped non-HTML content: {content_type or 'unknown content type'}")
        charset = response.headers.get_content_charset() or "utf-8"
        body = response.read()
        encoding = (response.headers.get("Content-Encoding", "") or "").lower().strip()
        if encoding == "gzip":
            body = gzip.decompress(body)
        elif encoding == "deflate":
            try:
                body = zlib.decompress(body)
            except zlib.error:
                body = zlib.decompress(body, -zlib.MAX_WBITS)
    return body.decode(charset, errors="replace"), content_type


def extract_snippets(text: str, keyword: str, context_words: int = 25, max_snippets: int = 5) -> list[str]:
    words = re.findall(r"\S+", text)
    lowered_keyword = keyword.casefold()
    snippets: list[str] = []

    for index, word in enumerate(words):
        if lowered_keyword in word.casefold():
            start = max(0, index - context_words)
            end = min(len(words), index + context_words + 1)
            snippet = " ".join(words[start:end])
            snippets.append(snippet)
            if len(snippets) >= max_snippets:
                break

    return snippets


def count_keyword(text: str, keyword: str) -> int:
    return len(re.findall(re.escape(keyword), text, flags=re.IGNORECASE))




def get_keywords_from_excel(excel_path: str) -> list[str]:
    """Read latest keywords from Excel file."""

    df = pd.read_excel(excel_path)

    keywords = []

    if "Keyword" not in df.columns:
        raise ValueError("Excel file must contain 'Keyword' column")

    for value in df["Keyword"].dropna():

        parts = [
            item.strip().casefold()
            for item in str(value).split(",")
            if item.strip()
        ]

        keywords.extend(parts)

    return list(set(keywords))


def scan_site(
    start_urls: list[str],
    keywords: list[str],
    *,
    max_pages: int,
    max_depth: int,
    timeout: int,
    delay: float,
    same_site_only: bool,
) -> tuple[list[PageMatch], list[ScanError], int]:
    normalized_starts = [normalize_url(url) for url in start_urls if url]
    base_netloc = urlparse(normalized_starts[0]).netloc if normalized_starts else ""
    queue: deque[tuple[str, int]] = deque((url, 0) for url in normalized_starts)
    visited: set[str] = set()
    matches: list[PageMatch] = []
    errors: list[ScanError] = []

    while queue and len(visited) < max_pages:
        url, depth = queue.popleft()
        url = normalize_url(url)
        if url in visited:
            continue
        visited.add(url)

        try:
            raw_html, _ = fetch_html(url, timeout)
            parser = ContentParser()
            parser.feed(raw_html)
            text = parser.text

            hit_count = 0
            matched_keywords = []

            for kw in keywords:
                hits = count_keyword(text, kw)

                if hits:
                    hit_count += hits
                    matched_keywords.append(kw)

            if hit_count:

                all_snippets = []

                for kw in matched_keywords:
                    all_snippets.extend(extract_snippets(text, kw))

                matches.append(
                    PageMatch(
                        url=url,
                        title=parser.title.strip() or "(untitled)",
                        count=hit_count,
                        snippets=all_snippets[:5],
                    )
                )

            if depth < max_depth:
                for href in parser.links:
                    absolute = urljoin(url, href)
                    if not is_http_url(absolute):
                        continue
                    absolute, _ = urldefrag(absolute)
                    absolute = normalize_url(absolute)
                    if same_site_only and not is_same_site(absolute, base_netloc):
                        continue
                    if absolute not in visited:
                        queue.append((absolute, depth + 1))
        except (HTTPError, URLError, TimeoutError, ValueError, OSError) as exc:
            errors.append(ScanError(url=url, message=str(exc)))

        if delay > 0:
            time.sleep(delay)

    return matches, errors, len(visited)


def pdf_escape(text: str) -> str:
    return text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def wrap_lines(text: str, width: int = 92) -> list[str]:
    lines: list[str] = []
    for part in text.splitlines() or [""]:
        wrapped = textwrap.wrap(part, width=width, replace_whitespace=False) or [""]
        lines.extend(wrapped)
    return lines


def paginate(lines: Iterable[str], lines_per_page: int) -> list[list[str]]:
    pages: list[list[str]] = []
    current: list[str] = []
    for line in lines:
        current.append(line)
        if len(current) >= lines_per_page:
            pages.append(current)
            current = []
    if current or not pages:
        pages.append(current)
    return pages


def build_report_lines(
    *,
    start_urls: list[str],
    keywords: list[str],
    scanned_pages: int,
    matches: list[PageMatch],
    errors: list[ScanError],
) -> list[str]:
    generated_at = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    total_hits = sum(match.count for match in matches)
    lines = [
        "Keyword Scan Report",
        "",
        f"Generated: {generated_at}",
        f"Start URL: {start_urls[0] if start_urls else ''}",
        f"Keywords: {', '.join(keywords)}",
        f"Pages scanned: {scanned_pages}",
        f"Pages with matches: {len(matches)}",
        f"Total keyword hits: {total_hits}",
        "",
    ]
    if len(start_urls) > 1:
        lines.extend(["Seed URLs:", *[f"- {url}" for url in start_urls[1:]], ""])
    lines.extend(["Matches", "======="])

    if not matches:
        lines.append("No matches found.")
    else:
        for index, match in enumerate(sorted(matches, key=lambda item: item.count, reverse=True), start=1):
            lines.extend(
                [
                    "",
                    f"{index}. {match.title}",
                    f"URL: {match.url}",
                    f"Hits: {match.count}",
                    "Snippets:",
                ]
            )
            for snippet in match.snippets:
                lines.append(f"- {snippet}")

    if errors:
        lines.extend(["", "Fetch Errors", "============"])
        for error in errors[:25]:
            lines.append(f"- {error.url}: {error.message}")
        if len(errors) > 25:
            lines.append(f"... {len(errors) - 25} more errors omitted")

    return lines


def write_pdf(lines: list[str], output_path: Path) -> None:
    lines_per_page = int((PAGE_HEIGHT - (2 * MARGIN)) / LINE_HEIGHT)
    pages = paginate([line for item in lines for line in wrap_lines(item)], lines_per_page)
    objects: list[bytes] = []

    def add_object(content: bytes) -> int:
        objects.append(content)
        return len(objects)

    catalog_id = add_object(b"<< /Type /Catalog /Pages 2 0 R >>")
    pages_id = add_object(b"")
    font_id = add_object(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")
    page_ids: list[int] = []

    for page_lines in pages:
        content_lines = ["BT", f"/F1 {FONT_SIZE} Tf", f"{MARGIN} {PAGE_HEIGHT - MARGIN} Td"]
        for index, line in enumerate(page_lines):
            if index:
                content_lines.append(f"0 -{LINE_HEIGHT} Td")
            content_lines.append(f"({pdf_escape(line)}) Tj")
        content_lines.append("ET")
        stream = "\n".join(content_lines).encode("latin-1", errors="replace")
        content_id = add_object(
            b"<< /Length " + str(len(stream)).encode("ascii") + b" >>\nstream\n" + stream + b"\nendstream"
        )
        page_id = add_object(
            (
                f"<< /Type /Page /Parent {pages_id} 0 R /MediaBox [0 0 {PAGE_WIDTH} {PAGE_HEIGHT}] "
                f"/Resources << /Font << /F1 {font_id} 0 R >> >> /Contents {content_id} 0 R >>"
            ).encode("ascii")
        )
        page_ids.append(page_id)

    objects[pages_id - 1] = (
        f"<< /Type /Pages /Kids [{' '.join(f'{page_id} 0 R' for page_id in page_ids)}] "
        f"/Count {len(page_ids)} >>"
    ).encode("ascii")

    output = bytearray(b"%PDF-1.4\n")
    offsets = [0]
    for object_id, content in enumerate(objects, start=1):
        offsets.append(len(output))
        output.extend(f"{object_id} 0 obj\n".encode("ascii"))
        output.extend(content)
        output.extend(b"\nendobj\n")

    xref_offset = len(output)
    output.extend(f"xref\n0 {len(objects) + 1}\n".encode("ascii"))
    output.extend(b"0000000000 65535 f \n")
    for offset in offsets[1:]:
        output.extend(f"{offset:010d} 00000 n \n".encode("ascii"))
    output.extend(
        (
            f"trailer\n<< /Size {len(objects) + 1} /Root {catalog_id} 0 R >>\n"
            f"startxref\n{xref_offset}\n%%EOF\n"
        ).encode("ascii")
    )

    output_path.write_bytes(output)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Scan a URL and its linked pages for a keyword, then generate a PDF report."
    )
    parser.add_argument("url", help="Starting URL to scan, for example https://example.com")
    parser.add_argument("keyword", help="Keyword or phrase to search for")
    parser.add_argument("-o", "--output", default="keyword_scan_report.pdf", help="Output PDF file path")
    parser.add_argument(
        "--seed-url",
        action="append",
        default=[],
        help="Extra URL to add to the crawl queue (repeatable). Useful for sections not linked from the start page.",
    )
    parser.add_argument("--max-pages", type=int, default=25, help="Maximum pages to scan")
    parser.add_argument("--max-depth", type=int, default=1, help="How many link levels to crawl from the start URL")
    parser.add_argument("--timeout", type=int, default=15, help="HTTP timeout in seconds")
    parser.add_argument("--delay", type=float, default=0.2, help="Delay between requests in seconds")
    parser.add_argument(
        "--include-external",
        action="store_true",
        help="Also follow links outside the starting website's domain",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    start_url = normalize_url(args.url)
    seed_urls = [normalize_url(url) for url in (args.seed_url or []) if url]
    start_urls = [start_url, *seed_urls]
    output_path = Path(args.output)

    matches, errors, scanned_pages = scan_site(
        start_urls,
        [args.keyword.casefold()],
        max_pages=max(1, args.max_pages),
        max_depth=max(0, args.max_depth),
        timeout=max(1, args.timeout),
        delay=max(0, args.delay),
        same_site_only=not args.include_external,
    )
    report_lines = build_report_lines(
        start_urls=start_urls,
        keywords=[args.keyword.casefold()],
        scanned_pages=scanned_pages,
        matches=matches,
        errors=errors,
    )
    write_pdf(report_lines, output_path)

    print(f"Scanned {scanned_pages} page(s).")
    print(f"Found {sum(match.count for match in matches)} hit(s) across {len(matches)} page(s).")
    print(f"PDF report written to: {output_path.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
    
    
def run_scan_ui(url, max_pages, max_depth, output_path):
    start_urls = [normalize_url(url)]

    excel_path = "keywords.xlsx"

    keywords = get_keywords_from_excel(excel_path)

    matches, errors, scanned_pages = scan_site(
        start_urls,
        keywords,
        max_pages=max(1, max_pages),
        max_depth=max(0, max_depth),
        timeout=15,
        delay=0.2,
        same_site_only=True,
    )

    report_lines = build_report_lines(
        start_urls=start_urls,
        keywords=keywords,
        scanned_pages=scanned_pages,
        matches=matches,
        errors=errors,
    )

    write_pdf(report_lines, Path(output_path))

    return scanned_pages, len(matches)
