"""Monitor a public RSS feed of freelance projects.

Fetches the feed, keeps the projects that match the chosen categories and
keywords, remembers what it has already seen in SQLite, and exports the
result to Excel and JSON.

Usage:
    python monitor.py --category Python --category "Парсинг данных" --out out
"""

from __future__ import annotations

import argparse
import html
import json
import logging
import re
import sqlite3
import sys
import time
import urllib.error
import urllib.request
import xml.etree.ElementTree as ET
from dataclasses import asdict, dataclass, field
from datetime import datetime
from email.utils import parsedate_to_datetime
from pathlib import Path

DEFAULT_FEED = "https://freelancehunt.com/projects.rss"
USER_AGENT = "rss-project-monitor/1.0"

log = logging.getLogger("monitor")

# Budget is appended to the title: "Tender BOT - 100USD".
_BUDGET_RE = re.compile(r"\s+-\s+(\d+(?:[.,]\d+)?)\s*([A-Z]{3})\s*$")
_ID_RE = re.compile(r"/(\d+)\.html")
_TAG_RE = re.compile(r"<[^>]+>")
_SPACE_RE = re.compile(r"\s+")


class FeedError(Exception):
    """The feed could not be fetched or parsed (not the same as 'no new items')."""


@dataclass
class Project:
    id: str
    title: str
    url: str
    published: str | None  # ISO 8601, or None when the feed gives no valid date
    categories: list[str] = field(default_factory=list)
    budget_amount: float | None = None  # None = budget not stated in the feed
    budget_currency: str | None = None
    description: str = ""


def fetch(url: str, retries: int = 3, timeout: float = 30, pause: float = 2.0) -> bytes:
    """Download the feed, retrying temporary failures."""
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    last_error: Exception | None = None
    for attempt in range(1, retries + 1):
        try:
            with urllib.request.urlopen(request, timeout=timeout) as response:
                return response.read()
        except urllib.error.HTTPError as exc:
            # 4xx will not fix itself (except 429), so do not hammer the server.
            if exc.code < 500 and exc.code != 429:
                raise FeedError(f"HTTP {exc.code} for {url}") from exc
            last_error = exc
        except (urllib.error.URLError, TimeoutError, OSError) as exc:
            last_error = exc
        log.warning("attempt %d/%d failed: %s", attempt, retries, last_error)
        if attempt < retries:
            time.sleep(pause * attempt)
    raise FeedError(f"could not fetch {url}: {last_error}")


def clean_text(raw: str | None) -> str:
    """Strip HTML tags and collapse whitespace."""
    if not raw:
        return ""
    return _SPACE_RE.sub(" ", _TAG_RE.sub(" ", html.unescape(raw))).strip()


def split_budget(title: str) -> tuple[str, float | None, str | None]:
    """Split 'Name - 100USD' into ('Name', 100.0, 'USD')."""
    match = _BUDGET_RE.search(title)
    if not match:
        return title.strip(), None, None
    amount = float(match.group(1).replace(",", "."))
    return title[: match.start()].strip(), amount, match.group(2)


def parse_date(raw: str | None) -> str | None:
    if not raw:
        return None
    try:
        return parsedate_to_datetime(raw).isoformat()
    except (TypeError, ValueError):
        return None


def parse_feed(xml_bytes: bytes) -> list[Project]:
    """Turn the RSS document into a list of projects."""
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError as exc:
        raise FeedError(f"feed is not valid XML: {exc}") from exc

    projects: list[Project] = []
    for item in root.iterfind("./channel/item"):
        link = (item.findtext("link") or "").split("?")[0].strip()
        id_match = _ID_RE.search(link)
        if not link or not id_match:
            log.warning("skipping item without a usable link: %r", item.findtext("title"))
            continue
        title, amount, currency = split_budget(item.findtext("title") or "")
        projects.append(
            Project(
                id=id_match.group(1),
                title=title,
                url=link,
                published=parse_date(item.findtext("pubDate")),
                categories=[c.text.strip() for c in item.findall("category") if c.text],
                budget_amount=amount,
                budget_currency=currency,
                description=clean_text(item.findtext("description")),
            )
        )
    return projects


def matches(project: Project, categories: list[str], keywords: list[str]) -> bool:
    """A project matches when it hits any chosen category OR any keyword.

    Keywords match at the start of a word. With no filters, every project matches.
    """
    if not categories and not keywords:
        return True
    wanted = {c.casefold() for c in categories}
    if wanted & {c.casefold() for c in project.categories}:
        return True
    text = f"{project.title} {project.description}".casefold()
    # Match from the start of a word: "бот" finds "боты" but not "работа".
    return any(re.search(r"(?<!\w)" + re.escape(k.casefold()), text) for k in keywords)


class SeenStore:
    """Remembers project ids between runs so each project is reported once."""

    def __init__(self, path: Path | str):
        self.conn = sqlite3.connect(path)
        self.conn.execute(
            "CREATE TABLE IF NOT EXISTS seen (id TEXT PRIMARY KEY, first_seen TEXT NOT NULL)"
        )

    def only_new(self, projects: list[Project]) -> list[Project]:
        """Return the projects not seen before and record them."""
        fresh = []
        now = datetime.now().isoformat(timespec="seconds")
        for project in projects:
            cursor = self.conn.execute(
                "INSERT OR IGNORE INTO seen (id, first_seen) VALUES (?, ?)", (project.id, now)
            )
            if cursor.rowcount:
                fresh.append(project)
        self.conn.commit()
        return fresh

    def close(self) -> None:
        self.conn.close()


def export_json(projects: list[Project], path: Path) -> None:
    path.write_text(
        json.dumps([asdict(p) for p in projects], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def export_excel(projects: list[Project], path: Path) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Font

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Projects"
    headers = ["ID", "Title", "Budget", "Currency", "Published", "Categories", "URL", "Description"]
    sheet.append(headers)
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    for p in projects:
        sheet.append(
            [
                p.id,
                p.title,
                p.budget_amount,  # empty cell when the budget is not stated
                p.budget_currency,
                p.published,
                ", ".join(p.categories),
                p.url,
                p.description,
            ]
        )
    for column, width in zip("ABCDEFGH", (10, 60, 10, 10, 26, 40, 60, 100)):
        sheet.column_dimensions[column].width = width
    sheet.freeze_panes = "A2"
    workbook.save(path)


def run(args: argparse.Namespace) -> int:
    try:
        xml_bytes = Path(args.file).read_bytes() if args.file else fetch(args.url)
        projects = parse_feed(xml_bytes)
    except (FeedError, OSError) as exc:
        log.error("feed error: %s", exc)
        return 2  # distinct from "nothing new", which is a normal exit 0

    selected = [p for p in projects if matches(p, args.category, args.keyword)]
    log.info("feed: %d projects, matched filters: %d", len(projects), len(selected))

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)

    if not args.all:
        store = SeenStore(out_dir / "seen.sqlite3")
        selected = store.only_new(selected)
        store.close()
        log.info("new since last run: %d", len(selected))

    if not selected:
        log.info("nothing to export")
        return 0

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    json_path = out_dir / f"projects-{stamp}.json"
    xlsx_path = out_dir / f"projects-{stamp}.xlsx"
    export_json(selected, json_path)
    export_excel(selected, xlsx_path)
    log.info("saved %s and %s", json_path, xlsx_path)
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("--url", default=DEFAULT_FEED, help="RSS feed URL")
    parser.add_argument("--file", help="read the feed from a local file instead of the network")
    parser.add_argument("--category", action="append", default=[], help="category to keep (repeatable)")
    parser.add_argument("--keyword", action="append", default=[], help="keyword in title/description (repeatable)")
    parser.add_argument("--out", default="out", help="output folder")
    parser.add_argument("--all", action="store_true", help="export every match, not only new ones")
    return parser


def main(argv: list[str] | None = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    return run(build_parser().parse_args(argv))


if __name__ == "__main__":
    sys.exit(main())
