import re
import sys
from pathlib import Path


def main():
    root = Path(__file__).resolve().parent.parent
    js_path = root / "web" / "static" / "app.js"
    html_path = root / "web" / "static" / "index.html"
    tpl_path = root / "web" / "templates" / "index.html"

    js = js_path.read_text(encoding="utf-8")
    html = html_path.read_text(encoding="utf-8")
    tpl = tpl_path.read_text(encoding="utf-8")

    assert html == tpl, "web/static/index.html and web/templates/index.html must match"

    ids = set(re.findall(r"document\.getElementById\(['\"]([^'\"]+)['\"]\)", js))
    missing = [i for i in sorted(ids) if f'id="{i}"' not in html and f"id='{i}'" not in html]

    if missing:
        print(f"FAILED: Missing IDs in HTML ({len(missing)}): {missing}")
        sys.exit(1)

    print(
        f"SUCCESS: All {len(ids)} DOM element IDs from app.js are present in index.html and templates/index.html"
    )


if __name__ == "__main__":
    main()
