#!/usr/bin/env python3
"""Refresh the public discovery index without republishing article content."""

from __future__ import annotations

import email.utils
import html
import json
import re
import sys
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "data" / "latest.json"
START = datetime(2026, 8, 26, tzinfo=timezone.utc)
USER_AGENT = "RasuwaFloodSourceIndex/1.0 (+https://rasuwaflood.org/)"
QUERIES = (
    '"Rasuwa flood" OR "Bhote Koshi flood" OR "Bhotekoshi flood" when:30d',
    '"Kyirong Rasuwa flood" OR "Lhende flood" when:30d',
    '(Rasuwa OR Rasuwagadhi) (flood OR avalanche) report assessment satellite when:30d',
)
REQUIRED = re.compile(r"\b(rasuwa|rasuwagadhi|bhote\s*koshi|bhotekoshi|kyirong|lhende)\b", re.I)
HAZARD = re.compile(r"\b(flood|avalanche|landslide|glacier|disaster|mudflow|rockflow)\b", re.I)
TAG = re.compile(r"<[^>]+>")


def fetch(url: str) -> bytes:
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(request, timeout=25) as response:
        return response.read()


def clean(value: str | None, limit: int = 260) -> str:
    text = html.unescape(TAG.sub(" ", value or ""))
    text = re.sub(r"\s+", " ", text).strip()
    return text[:limit].rstrip()


def iso_date(value: str | None) -> tuple[str, float]:
    try:
        parsed = email.utils.parsedate_to_datetime(value or "").astimezone(timezone.utc)
    except (TypeError, ValueError, OverflowError):
        parsed = datetime.now(timezone.utc)
    return parsed.date().isoformat(), parsed.timestamp()


def canonical(url: str) -> str:
    parsed = urllib.parse.urlsplit(url)
    query = urllib.parse.parse_qsl(parsed.query, keep_blank_values=True)
    query = [(k, v) for k, v in query if not k.lower().startswith("utm_")]
    return urllib.parse.urlunsplit((parsed.scheme, parsed.netloc.lower(), parsed.path, urllib.parse.urlencode(query), ""))


def google_news() -> list[dict]:
    found: list[dict] = []
    for query in QUERIES:
        url = "https://news.google.com/rss/search?" + urllib.parse.urlencode(
            {"q": query, "hl": "en", "gl": "US", "ceid": "US:en"}
        )
        root = ET.fromstring(fetch(url))
        for node in root.findall("./channel/item"):
            title = clean(node.findtext("title"), 180)
            summary = clean(node.findtext("description"))
            searchable = f"{title} {summary}"
            if not (REQUIRED.search(searchable) and HAZARD.search(searchable)):
                continue
            date, timestamp = iso_date(node.findtext("pubDate"))
            if timestamp < START.timestamp():
                continue
            source_node = node.find("source")
            source = clean(source_node.text if source_node is not None else "News source", 80)
            found.append({
                "source": source,
                "date": date,
                "timestamp": timestamp,
                "title": title,
                "summary": summary or "Relevant reporting discovered by the automated source index.",
                "url": canonical(node.findtext("link") or ""),
                "kind": "News and reporting",
            })
    return found


def previous_items() -> list[dict]:
    try:
        data = json.loads(OUTPUT.read_text())
        return [item for item in data.get("items", []) if item.get("url")]
    except (FileNotFoundError, json.JSONDecodeError):
        return []


def main() -> int:
    items = previous_items()
    errors: list[str] = []
    try:
        items.extend(google_news())
    except Exception as exc:  # keep the last good index if a feed is temporarily unavailable
        errors.append(f"Google News RSS: {exc}")

    unique: dict[str, dict] = {}
    for item in items:
        url = canonical(str(item.get("url", "")))
        if not url.startswith(("http://", "https://")):
            continue
        item["url"] = url
        item.setdefault("timestamp", 0)
        item.setdefault("kind", "Source")
        unique[url] = item

    ordered = sorted(unique.values(), key=lambda item: (item.get("timestamp", 0), item.get("date", "")), reverse=True)[:60]
    for item in ordered:
        item.pop("timestamp", None)
    payload = {
        "generated_at": datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z"),
        "count": len(ordered),
        "collection_method": "Automated metadata discovery; inclusion is not verification.",
        "errors": errors,
        "items": ordered,
    }
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n")
    print(f"Wrote {len(ordered)} unique sources to {OUTPUT.relative_to(ROOT)}")
    if errors:
        print("; ".join(errors), file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
