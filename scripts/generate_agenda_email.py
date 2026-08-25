#!/usr/bin/env python3
"""スプレッドシート（Excel/CSV）の会議情報・議題を議事次第メールテンプレートに流し込む。

使い方:
    # Excelワークブック（シート「会議情報」「議題」を含む .xlsx）から生成
    python3 scripts/generate_agenda_email.py --workbook meeting_data.xlsx

    # CSV2ファイル（会議情報・議題を分けて管理）から生成
    python3 scripts/generate_agenda_email.py --info meeting_info.csv --items agenda_items.csv

出力先は省略時 output/agenda_<年>-<月>.html。
ブラウザで開いて全選択→コピーし、メールソフトの本文に貼り付けて送信する。

入力フォーマットは scripts/sample_data/ のサンプルを参照。
"""
import argparse
import csv
import html
import re
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_TEMPLATE = REPO_ROOT / "meeting_agenda_template.html"

REQUIRED_INFO_KEYS = ["年", "月", "日時", "場所", "議長", "記録係", "出席者"]
AGENDA_COLUMNS = ["No", "議題", "担当", "時間"]

ROW_TEMPLATE = """        <tr>
          <td align="center" style="font-size:13px; color:#1a1a1a; padding:12px 6px;{border}">{no}</td>
          <td style="font-size:13px; color:#1a1a1a; padding:12px 10px;{border}">
            {topic}
          </td>
          <td style="font-size:13px; color:#1a1a1a; padding:12px 10px;{border}">{owner}</td>
          <td align="center" style="font-size:13px; color:#1a1a1a; padding:12px 6px;{border}">{minutes}</td>
        </tr>"""
ROW_BORDER = " border-bottom:1px solid #eef0f2;"


def read_info_csv(path):
    info = {}
    with open(path, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if not row.get("項目"):
                continue
            info[row["項目"].strip()] = (row.get("内容") or "").strip()
    return info


def read_items_csv(path):
    items = []
    with open(path, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if not any((row.get(k) or "").strip() for k in AGENDA_COLUMNS):
                continue
            items.append({k: (row.get(k) or "").strip() for k in AGENDA_COLUMNS})
    return items


def read_workbook(path):
    try:
        import openpyxl
    except ImportError:
        sys.exit(
            "xlsxファイルを読むには openpyxl が必要です。`pip install openpyxl` を実行してください。"
        )
    wb = openpyxl.load_workbook(path, data_only=True)
    if "会議情報" not in wb.sheetnames:
        sys.exit('シート「会議情報」が見つかりません。')
    if "議題" not in wb.sheetnames:
        sys.exit('シート「議題」が見つかりません。')

    info_ws = wb["会議情報"]
    info = {}
    for row in info_ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        key, value = row[0], row[1] if len(row) > 1 else None
        info[str(key).strip()] = "" if value is None else str(value).strip()

    items_ws = wb["議題"]
    rows_iter = items_ws.iter_rows(min_row=1, max_row=1, values_only=True)
    header = [str(c).strip() if c is not None else "" for c in next(rows_iter)]
    items = []
    for row in items_ws.iter_rows(min_row=2, values_only=True):
        if not row or all(v is None for v in row):
            continue
        record = {
            header[i]: ("" if row[i] is None else str(row[i]).strip())
            for i in range(min(len(header), len(row)))
        }
        if not any(record.get(k) for k in AGENDA_COLUMNS):
            continue
        items.append({k: record.get(k, "") for k in AGENDA_COLUMNS})
    return info, items


def validate_info(info):
    missing = [k for k in REQUIRED_INFO_KEYS if not info.get(k)]
    if missing:
        sys.exit(f"会議情報に不足があります: {', '.join(missing)}")


def replace_field(text, key, value):
    pattern = re.compile(r"(<!--F:%s-->)(.*?)(<!--/F-->)" % re.escape(key), re.DOTALL)
    if not pattern.search(text):
        sys.exit(f"テンプレートにマーカー <!--F:{key}--> が見つかりません。")
    return pattern.sub(lambda m: m.group(1) + html.escape(value) + m.group(3), text)


def replace_rows(text, marker, rows_html):
    pattern = re.compile(
        r"(<!--%s:START-->)(.*?)(<!--%s:END-->)" % (re.escape(marker), re.escape(marker)),
        re.DOTALL,
    )
    if not pattern.search(text):
        sys.exit(f"テンプレートにマーカー <!--{marker}:START--> が見つかりません。")
    return pattern.sub(lambda m: m.group(1) + "\n" + rows_html + "\n" + m.group(3), text)


def build_agenda_rows(items):
    if not items:
        sys.exit("議題が1件もありません。")
    rows = []
    last = len(items) - 1
    for i, item in enumerate(items):
        rows.append(
            ROW_TEMPLATE.format(
                no=html.escape(item.get("No", "")),
                topic=html.escape(item.get("議題", "")),
                owner=html.escape(item.get("担当", "")),
                minutes=html.escape(item.get("時間", "")),
                border="" if i == last else ROW_BORDER,
            )
        )
    return "\n".join(rows)


def main():
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument("--workbook", type=Path, help="会議情報・議題シートを含むExcelファイル(.xlsx)")
    parser.add_argument("--info", type=Path, help="会議情報CSVファイル（--items とセットで指定）")
    parser.add_argument("--items", type=Path, help="議題CSVファイル（--info とセットで指定）")
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE, help="元になるHTMLテンプレート")
    parser.add_argument("--output", type=Path, help="出力先HTMLファイル（省略時は output/agenda_YYYY-MM.html）")
    args = parser.parse_args()

    if args.workbook:
        info, items = read_workbook(args.workbook)
    elif args.info and args.items:
        info, items = read_info_csv(args.info), read_items_csv(args.items)
    else:
        parser.error("--workbook か、--info と --items の組み合わせのいずれかを指定してください。")

    validate_info(info)

    title = info.get("件名") or f"【{info['年']}年{info['月']}月度】定例会議"

    if not args.template.exists():
        sys.exit(f"テンプレートが見つかりません: {args.template}")
    text = args.template.read_text(encoding="utf-8")
    text = replace_field(text, "TITLE", title)
    text = replace_field(text, "DATETIME", info["日時"])
    text = replace_field(text, "PLACE", info["場所"])
    text = replace_field(text, "CHAIR", info["議長"])
    text = replace_field(text, "RECORDER", info["記録係"])
    text = replace_field(text, "ATTENDEES", info["出席者"])
    text = replace_rows(text, "ROWS:AGENDA", build_agenda_rows(items))

    output = args.output
    if output is None:
        year = info["年"]
        month = str(info["月"]).zfill(2)
        output = REPO_ROOT / "output" / f"agenda_{year}-{month}.html"
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(text, encoding="utf-8")
    print(f"生成しました: {output}")


if __name__ == "__main__":
    main()
