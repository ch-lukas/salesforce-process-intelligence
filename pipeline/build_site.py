"""Assemble the static site in site/ for hosting (e.g. Vercel).

  site/index.html        overview: value items, business case, approach, live Salesforce run
  site/dashboard.html    process analysis on the 10-month dataset    (data/)
  site/data/*.csv|.xes   the event logs, downloadable
  site/decks/*.pdf       the story and tech setup decks
  site/img/*             org screenshots used on the overview
  site/vercel.json       redirects from the retired /value-case.html and /live.html

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
]
IMAGES = {"docs/img/org/salesforce-with-url.jpg": "salesforce-org.jpg", "docs/img/org/stage-history.png": "stage-history.png"}
# Old pages now live as sections of the overview; keep their links working.
REDIRECTS = {"redirects": [
    {"source": "/value-case.html", "destination": "/#business-case", "permanent": False},
    {"source": "/live.html", "destination": "/#live-run", "permanent": False},
]}


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
    for retired in ("value-case.html", "live.html"):
        (SITE / retired).unlink(missing_ok=True)
    img = SITE / "img"
    img.mkdir(exist_ok=True)
    for src, name in IMAGES.items():
        shutil.copy(src, img / name)
        print(f"  img/{name}")
    (SITE / "vercel.json").write_text(json.dumps(REDIRECTS, indent=2) + "\n")
    print("  vercel.json")
    decks = SITE / "decks"
    decks.mkdir(exist_ok=True)
    for pdf in sorted(Path("deck").glob("*.pdf")):
        shutil.copy(pdf, decks / pdf.name)
        print(f"  decks/{pdf.name}")
    print(f"Wrote {SITE}/")


if __name__ == "__main__":
    main()
