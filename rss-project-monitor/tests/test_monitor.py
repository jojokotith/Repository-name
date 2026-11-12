import json
import sys
from pathlib import Path

import pytest
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import monitor  # noqa: E402

SAMPLE = Path(__file__).parent / "sample_feed.xml"


@pytest.fixture
def projects():
    return monitor.parse_feed(SAMPLE.read_bytes())


def test_parse_skips_items_without_link(projects):
    assert [p.id for p in projects] == ["1001", "1002", "1003"]


def test_parse_fields(projects):
    p = projects[0]
    assert p.title == "Parser for a product catalogue"
    assert (p.budget_amount, p.budget_currency) == (100.0, "USD")
    assert p.url == "https://example.com/project/parser-catalogue/1001.html"
    assert p.categories == ["Python", "Парсинг данных"]
    assert p.description == "Collect prices into Excel"
    assert p.published == "2026-11-12T12:58:08+02:00"


def test_missing_budget_stays_empty(projects):
    assert projects[1].budget_amount is None
    assert projects[1].budget_currency is None


def test_decimal_comma_budget_and_bad_date(projects):
    p = projects[2]
    assert (p.budget_amount, p.budget_currency) == (2500.5, "UAH")
    assert p.published is None


@pytest.mark.parametrize(
    "title, expected",
    [
        ("Fix site - 700UAH", ("Fix site", 700.0, "UAH")),
        ("Plain title", ("Plain title", None, None)),
        ("Step 1 - 2 of work", ("Step 1 - 2 of work", None, None)),
    ],
)
def test_split_budget(title, expected):
    assert monitor.split_budget(title) == expected


def test_matches_category_or_keyword(projects):
    by_category = [p.id for p in projects if monitor.matches(p, ["python"], [])]
    by_keyword = [p.id for p in projects if monitor.matches(p, [], ["бот"])]
    both = [p.id for p in projects if monitor.matches(p, ["Python"], ["бот"])]
    assert by_category == ["1001"]
    assert by_keyword == ["1003"]
    assert both == ["1001", "1003"]
    assert all(monitor.matches(p, [], []) for p in projects)


def test_keyword_matches_only_at_word_start():
    project = monitor.Project(id="1", title="Разработка сайта", url="u", published=None,
                              description="Нужна обработка заявок, работа на месяц")
    assert not monitor.matches(project, [], ["бот"])
    project.description = "Нужны боты для Telegram"
    assert monitor.matches(project, [], ["бот"])


def test_invalid_xml_raises():
    with pytest.raises(monitor.FeedError):
        monitor.parse_feed(b"<rss><channel>")


def test_seen_store_reports_each_project_once(tmp_path, projects):
    db = tmp_path / "seen.sqlite3"
    store = monitor.SeenStore(db)
    assert len(store.only_new(projects)) == 3
    assert store.only_new(projects) == []
    store.close()
    # Survives a restart.
    store = monitor.SeenStore(db)
    assert store.only_new(projects) == []
    store.close()


def test_cli_exports_then_reports_nothing_new(tmp_path):
    args = ["--file", str(SAMPLE), "--category", "Python", "--keyword", "бот", "--out", str(tmp_path)]
    assert monitor.main(args) == 0

    data = json.loads(next(tmp_path.glob("*.json")).read_text(encoding="utf-8"))
    assert [d["id"] for d in data] == ["1001", "1003"]

    sheet = load_workbook(next(tmp_path.glob("*.xlsx"))).active
    rows = list(sheet.iter_rows(values_only=True))
    assert rows[0][0] == "ID"
    assert [r[0] for r in rows[1:]] == ["1001", "1003"]
    assert rows[1][2] == 100.0

    # Second run: nothing new, no extra files.
    assert monitor.main(args) == 0
    assert len(list(tmp_path.glob("*.json"))) == 1


def test_cli_returns_2_on_feed_error(tmp_path):
    assert monitor.main(["--file", str(tmp_path / "missing.xml"), "--out", str(tmp_path)]) == 2
