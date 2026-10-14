"""Assemble the static site in site/ for hosting (e.g. Vercel).

  site/index.html        process dashboard on the 10-month dataset   (data/)
  site/value-case.html   value-framing one-pager                     (data/)
  site/live.html         live Salesforce extract proof               (data/live/)
  site/data/*.csv|.xes   the event logs, downloadable

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
    ("dashboard/template.html", "data", "index.html"),
    ("dashboard/value-case.template.html", "data", "value-case.html"),
    ("dashboard/template.html", "data/live", "live.html"),
]


def render(template_path, data_dir):
    insights = json.loads(Path(f"{data_dir}/insights.json").read_text())
    template = Path(template_path).read_text()
    payload = json.dumps(insights, separators=(",", ":")).replace("</", "<\\/")
    body = template.replace("/*__DATA__*/null", payload)
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
    print(f"Wrote {SITE}/")


if __name__ == "__main__":
    main()
