# -*- coding: utf-8 -*-
"""recon_engine 单元测试（标准库 unittest，无需 pytest / openpyxl）。"""
import os
import tempfile
import unittest

import recon_engine as eng
# main 提供 _make_sample_xlsx 用于构造最小可用 xlsx
import main as demo


class TestCsv(unittest.TestCase):
    def test_read_csv_and_reconcile(self):
        tmp = tempfile.mkdtemp()
        a = os.path.join(tmp, "a.csv")
        b = os.path.join(tmp, "b.csv")
        with open(a, "w", encoding="utf-8-sig", newline="") as f:
            f.write("订单号,客户,金额\nA001,张三,100\nA002,李四,200\nA003,王五,300\n")
        with open(b, "w", encoding="utf-8-sig", newline="") as f:
            f.write("订单号,客户,金额\nA001,张三,100\nA002,李四,250\nA004,赵六,400\n")
        ra = eng.read_csv(a)
        rb = eng.read_csv(b)
        self.assertEqual(len(ra), 3)
        res = eng.reconcile(ra, rb, key="订单号")
        s = res["stats"]
        self.assertEqual(s["only_a"], 1)      # A003
        self.assertEqual(s["only_b"], 1)      # A004
        self.assertEqual(s["mismatch"], 1)    # A002 金额 200 vs 250
        self.assertEqual(s["matched"], 1)     # A001

    def test_value_mismatch_detail(self):
        tmp = tempfile.mkdtemp()
        a = os.path.join(tmp, "a.csv")
        b = os.path.join(tmp, "b.csv")
        with open(a, "w", encoding="utf-8-sig", newline="") as f:
            f.write("id,name\nX,苹果\n")
        with open(b, "w", encoding="utf-8-sig", newline="") as f:
            f.write("id,name\nX,香蕉\n")
        res = eng.reconcile(eng.read_csv(a), eng.read_csv(b), key="id")
        self.assertEqual(len(res["value_mismatch"]), 1)
        diff = res["value_mismatch"][0]["diffs"][0]
        self.assertEqual(diff["col"], "name")
        self.assertEqual(diff["a"], "苹果")
        self.assertEqual(diff["b"], "香蕉")

    def test_whitespace_normalization(self):
        tmp = tempfile.mkdtemp()
        a = os.path.join(tmp, "a.csv")
        b = os.path.join(tmp, "b.csv")
        with open(a, "w", encoding="utf-8-sig", newline="") as f:
            f.write("id,v\nK,  hello \n")
        with open(b, "w", encoding="utf-8-sig", newline="") as f:
            f.write("id,v\nk,hello\n")
        # 主键忽略大小写 + 值折叠空白 → 应判为一致
        res = eng.reconcile(eng.read_csv(a), eng.read_csv(b), key="id",
                            case_insensitive_key=True)
        self.assertEqual(res["stats"]["matched"], 1)
        self.assertEqual(res["stats"]["mismatch"], 0)


class TestXlsx(unittest.TestCase):
    def test_read_xlsx(self):
        tmp = tempfile.mkdtemp()
        p = os.path.join(tmp, "s.xlsx")
        header = ["订单号", "客户", "金额"]
        rows = [["A001", "张三", 100], ["A002", "李四", 200]]
        demo._make_sample_xlsx(p, header, rows)
        data = eng.read_xlsx(p)
        self.assertEqual(len(data), 2)
        self.assertEqual(data[0]["订单号"], "A001")
        self.assertEqual(data[1]["客户"], "李四")
        self.assertEqual(data[0]["金额"], "100")  # 数字被读成字符串，比对时归一即可

    def test_reconcile_xlsx(self):
        tmp = tempfile.mkdtemp()
        a = os.path.join(tmp, "a.xlsx")
        b = os.path.join(tmp, "b.xlsx")
        header = ["订单号", "金额"]
        demo._make_sample_xlsx(a, header, [["A001", 100], ["A002", 200]])
        demo._make_sample_xlsx(b, header, [["A001", 100], ["A003", 300]])
        res = eng.reconcile(eng.read_xlsx(a), eng.read_xlsx(b), key="订单号")
        self.assertEqual(res["stats"]["only_a"], 1)
        self.assertEqual(res["stats"]["only_b"], 1)
        self.assertEqual(res["stats"]["matched"], 1)


class TestReport(unittest.TestCase):
    def test_write_report(self):
        tmp = tempfile.mkdtemp()
        a = os.path.join(tmp, "a.csv")
        b = os.path.join(tmp, "b.csv")
        # A 多一行 A003（产生"仅A有"），A002 金额不一致（产生"值不一致"）
        with open(a, "w", encoding="utf-8-sig", newline="") as f:
            f.write("订单号,金额\nA001,100\nA002,200\nA003,300\n")
        with open(b, "w", encoding="utf-8-sig", newline="") as f:
            f.write("订单号,金额\nA001,100\nA002,250\n")
        res = eng.reconcile(eng.read_csv(a), eng.read_csv(b), key="订单号")
        self.assertEqual(res["stats"]["only_a"], 1)
        out = os.path.join(tmp, "report.csv")
        n = eng.write_report(res, out)
        self.assertGreater(n, 0)
        with open(out, encoding="utf-8-sig") as f:
            text = f.read()
        self.assertIn("核对状态", text)
        self.assertIn("值不一致", text)
        self.assertIn("仅A有", text)


if __name__ == "__main__":
    unittest.main()
