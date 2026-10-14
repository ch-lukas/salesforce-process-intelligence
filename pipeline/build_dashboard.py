"""Embed data/insights.json into dashboard/template.html -> dashboard/index.html."""
import json
import os
from pathlib import Path

DATA_DIR = os.environ.get("PI_DATA", "data")

insights = json.loads(Path(f"{DATA_DIR}/insights.json").read_text())
template = Path("dashboard/template.html").read_text()
html = template.replace("/*__DATA__*/null", json.dumps(insights, separators=(",", ":")).replace("</", "<\\/"))
out = os.environ.get("PI_OUT", "dashboard/index.html")
Path(out).write_text(html)
print(f"Wrote {out}")
