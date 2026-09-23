import sys
import json
from pathlib import Path
from docxtpl import DocxTemplate

TEMPLATE  = "templates/title_page.docx"
TEMP_ROOT = "TEMP"

def main():
    if len(sys.argv) < 2:
        print("Usage: python fill_title.py <CABINET_ID> [vars.json]", file=sys.stderr)
        print("  If vars.json is omitted, JSON is read from stdin.", file=sys.stderr)
        sys.exit(1)

    cabinet_id = sys.argv[1]

    if len(sys.argv) >= 3:
        with open(sys.argv[2], encoding="utf-8") as f:
            context = json.load(f)
    else:
        context = json.load(sys.stdin)

    out_dir = Path(TEMP_ROOT) / f"SET_{cabinet_id}"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / "title_page_filled.docx"

    doc = DocxTemplate(TEMPLATE)
    doc.render(context)
    doc.save(str(out_path))

    print(f"[fill_title] Wrote {out_path}")

if __name__ == "__main__":
    main()