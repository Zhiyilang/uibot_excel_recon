# -*- coding: utf-8 -*-
"""
main.py — UiBot Creator（Python 模式）入口 + 独立可跑的演示

在 UiBot Creator 里：
  1. 新建「流程」项目，语言选 Python；
  2. 把本目录作为项目根；
  3. 用 UiBot 原生「Excel 读取」命令把两张表读成二维数组 / 数据表；
  4. 把数组转成 [{'列名': '值', ...}, ...] 后调用 run()；
  5. 运行即可生成核对报告 CSV。

脱离 UiBot 也能跑：python main.py 会生成示例数据并演示完整核对流程。
"""
from __future__ import annotations

import os
import tempfile

from recon_engine import read_csv, read_xlsx, reconcile, write_report, summarize


def run(a_path: str, b_path: str, key: str,
        compare_cols=None, out_path: str = "核对报告.csv",
        xlsx: bool = False) -> dict:
    """核对两张表并写出报告。返回 reconcile() 的结果字典。"""
    read = read_xlsx if xlsx else read_csv
    a = read(a_path)
    b = read(b_path)
    result = reconcile(a, b, key, compare_cols)
    write_report(result, out_path)
    return result


# ---------------------------------------------------------------------------
# UiBot 调用示例（在 Creator 内把 Excel 读出的数据表转成行字典后调用）
# ---------------------------------------------------------------------------
def uibot_example():
    # 下面这段在 UiBot Creator 里替换成你的真实 Excel 读取：
    #
    #   ui_bot_table_a = uibot.Excel.ReadRange(文件路径A, 表名, 起始行, 结束行)
    #   rows_a = [dict(zip(表头, 行)) for 行 in ui_bot_table_a]
    #   rows_b = ... 同理
    #   result = reconcile(rows_a, rows_b, key="订单号")
    #   print(summarize(result))
    #
    # 这里给一个纯 Python 的可运行版本，方便你先把逻辑跑通：
    demo_a = "示例_本月.xlsx"
    demo_b = "示例_财务.xlsx"
    result = run(demo_a, demo_b, key="订单号", out_path="核对报告.csv", xlsx=True)
    print(summarize(result))


# ---------------------------------------------------------------------------
# 独立演示（python main.py）
# ---------------------------------------------------------------------------
def _make_sample_csv(path: str, rows):
    import csv
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["订单号", "客户", "金额", "状态"])
        for r in rows:
            w.writerow(r)


def _make_sample_xlsx(path: str, header, rows):
    """用标准库生成一个最小可用的 .xlsx，用于演示 read_xlsx()。"""
    import zipfile
    from xml.sax.saxutils import escape

    def col_letter(idx):
        s = ""
        idx += 1
        while idx:
            idx, r = divmod(idx - 1, 26)
            s = chr(65 + r) + s
        return s

    shared = []
    def sid(text):
        shared.append(text)
        return len(shared) - 1

    sheet_rows = []
    sheet_rows.append("".join(
        f'<c r="{col_letter(i)}{1}" t="s"><v>{sid(str(h))}</v></c>' for i, h in enumerate(header)))
    for ri, row in enumerate(rows, start=2):
        cells = []
        for ci, val in enumerate(row):
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                cells.append(f'<c r="{col_letter(ci)}{ri}"><v>{val}</v></c>')
            else:
                cells.append(f'<c r="{col_letter(ci)}{ri}" t="s"><v>{sid(str(val))}</v></c>')
        sheet_rows.append("".join(cells))

    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData>{"".join(f"<row r=\"{i+1}\">{c}</row>" for i, c in enumerate(sheet_rows))}</sheetData>'
        '</worksheet>'
    )
    shared_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="'
        f'{len(shared)}" uniqueCount="{len(shared)}">'
        + "".join(f'<si><t xml:space="preserve">{escape(s)}</t></si>' for s in shared) + '</sst>'
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        '</Types>'
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>'
    )
    wb = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'
    )
    wb_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
        '</Relationships>'
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("xl/workbook.xml", wb)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        z.writestr("xl/sharedStrings.xml", shared_xml)


def main():
    tmp = tempfile.mkdtemp(prefix="recon_demo_")
    a = os.path.join(tmp, "示例_本月.xlsx")
    b = os.path.join(tmp, "示例_财务.xlsx")
    out = os.path.join(tmp, "核对报告.csv")
    header = ["订单号", "客户", "金额", "状态"]
    rows_a = [
        ["A001", "张三", 100, "已付"],
        ["A002", "李四", 200, "已付"],
        ["A003", "王五", 300, "待付"],
    ]
    rows_b = [
        ["A001", "张三", 100, "已付"],
        ["A002", "李四", 250, "已付"],   # 金额不一致
        ["A004", "赵六", 400, "已付"],   # 仅 B 有
    ]
    _make_sample_xlsx(a, header, rows_a)
    _make_sample_xlsx(b, header, rows_b)
    result = run(a, b, key="订单号", out_path=out, xlsx=True)
    print("=" * 40)
    print(summarize(result))
    print("=" * 40)
    print("报告已生成：", out)
    print("\n—— 值不一致明细 ——")
    for m in result["value_mismatch"]:
        print(f"  主键 {m['key']} (A行{m['a_row']}/B行{m['b_row']}):")
        for d in m["diffs"]:
            print(f"    {d['col']}: A={d['a']}  B={d['b']}")


if __name__ == "__main__":
    main()
