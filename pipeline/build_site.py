"""Assemble the static site in site/ for hosting (e.g. Vercel).

  site/index.html        landing page: headline value and the three value items
  site/dashboard.html    process dashboard on the 10-month dataset   (data/)
  site/value-case.html   value-framing one-pager                     (data/)
  site/live.html         live Salesforce extract proof               (data/live/)
  site/data/*.csv|.xes   the event logs, downloadable
  site/decks/*.pdf       the story and tech setup decks

Usage:
  python3 pipeline/build_site.py
"""
import json
import shutil
from pathlib import Path

SITE = Path("site")
HEAD = ('<!doctype html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n'
        '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n')
PAGES = [
    ("dashboard/landing.template.html", "data", "index.html"),
    ("dashboard/template.html", "data", "dashboard.html"),
    ("dashboard/value-case.template.html", "data", "value-case.html"),
    ("dashboard/template.html", "data/live", "live.html"),
]


def embed(path):
    return json.dumps(json.loads(Path(path).read_text()), separators=(",", ":")).replace("</", "<\\/")


def render(template_path, data_dir):
    template = Path(template_path).read_text()
    body = template.replace("/*__DATA__*/null", embed(f"{data_dir}/insights.json"))
    body = body.replace("/*__LIVE__*/null", embed("data/live/insights.json"))
    # Everything before the first <div> is head material (title, fonts, scripts, styles).
    split = body.index("<div")
    return f"{HEAD}{body[:split]}</head>\n<body>\n{body[split:]}\n</body>\n</html>\n"


def main():
    SITE.mkdir(exist_ok=True)
    for template, data_dir, out in PAGES:
        (SITE / out).write_text(render(template, data_dir))
        print(f"  {out}")
    downloads = SITE / "data"
    downloads.mkdir(exist_ok=True)
    for src, name in [("data/event_log.csv", "demo_event_log.csv"), ("data/event_log.xes", "demo_event_log.xes"),
                      ("data/live/event_log.csv", "live_event_log.csv"), ("data/live/event_log.xes", "live_event_log.xes")]:
        shutil.copy(src, downloads / name)
        print(f"  data/{name}")
    decks = SITE / "decks"
    decks.mkdir(exist_ok=True)
    for pdf in sorted(Path("deck").glob("*.pdf")):
        shutil.copy(pdf, decks / pdf.name)
        print(f"  decks/{pdf.name}")
    print(f"Wrote {SITE}/")


if __name__ == "__main__":
    main()
