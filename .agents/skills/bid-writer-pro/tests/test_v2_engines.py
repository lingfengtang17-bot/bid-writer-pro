#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import json
import pathlib
import subprocess
import sys
import tempfile
import unittest

import docx
from docx.shared import Cm, Pt, RGBColor

SKILL_DIR = pathlib.Path(__file__).resolve().parent.parent
SCRIPTS = SKILL_DIR / "scripts"


class TestV2Engines(unittest.TestCase):
    def test_regional_matcher_guizhou_water(self):
        with tempfile.TemporaryDirectory() as tmp:
            tender = pathlib.Path(tmp) / "贵州某水库枢纽工程招标文件.txt"
            tender.write_text(
                "本项目位于贵州省遵义市，属于水利工程施工暗标评审项目。\n"
                "技术标采用A4纸张，上边距2.0cm，其余2.5cm，宋体四号，行间距固定值25磅，总页数不超过800页，不得设置目录和页码。\n",
                encoding="utf-8",
            )
            out_cfg = pathlib.Path(tmp) / "norm.json"
            proc = subprocess.run(
                [sys.executable, str(SCRIPTS / "regional_matcher.py"), "--input", str(tender), "--out", str(out_cfg)],
                capture_output=True, text=True, encoding="utf-8", check=True,
            )
            data = json.loads(proc.stdout)
            self.assertEqual(data["matched_preset_id"], "guizhou_2026")
            self.assertEqual(data["profile"]["structure"]["max_pages"], 800)
            self.assertEqual(data["profile"]["body"]["line_pt"], 25.0)
            self.assertTrue(out_cfg.exists())

    def test_blind_check_and_fix(self):
        with tempfile.TemporaryDirectory() as tmp:
            bad = pathlib.Path(tmp) / "bad_bid.docx"
            doc = docx.Document()
            doc.core_properties.author = "张三"
            doc.core_properties.comments = "XX建设有限公司内部稿"
            sec = doc.sections[0]
            sec.top_margin = Cm(3.5)
            sec.header.paragraphs[0].text = "某某水利工程有限公司技术标"
            p1 = doc.add_paragraph()
            r1 = p1.add_run("本工程施工总工期 365 天，联系手机 13800138000。")
            r1.bold = True
            r1.font.color.rgb = RGBColor(255, 0, 0)
            r1.font.size = Pt(12)
            doc.add_paragraph("")  # 空段
            doc.save(str(bad))

            fixed = pathlib.Path(tmp) / "fixed_bid.docx"
            proc = subprocess.run(
                [
                    sys.executable, str(SCRIPTS / "blind_check_engine.py"),
                    str(bad), "--preset", "guizhou_2026", "--keyword", "某某水利", "--fix", "--out", str(fixed),
                ],
                capture_output=True, text=True, encoding="utf-8", check=True,
            )
            res = json.loads(proc.stdout)
            before_ids = {i["rule_id"] for i in res["before_audit"]["issues"]}
            after_ids = {i["rule_id"] for i in res["after_audit"]["issues"]}

            self.assertIn("page.margins", before_ids)
            self.assertIn("struct.header_footer", before_ids)
            self.assertIn("struct.doc_props", before_ids)
            self.assertIn("font.effects", before_ids)
            self.assertIn("font.color", before_ids)
            self.assertIn("privacy.keyword", before_ids)

            # 格式、页眉、属性、空段等可自动修复项必须全部清零
            for auto_rule in ("page.margins", "struct.header_footer", "struct.doc_props", "font.effects", "font.color", "struct.empty_paragraph", "para.line_spacing"):
                self.assertNotIn(auto_rule, after_ids)
            # 文本里的手机号仍应拦截提示人工脱敏
            self.assertIn("privacy.keyword", after_ids)

    def test_duplicate_check(self):
        with tempfile.TemporaryDirectory() as tmp:
            f1 = pathlib.Path(tmp) / "a.txt"
            f2 = pathlib.Path(tmp) / "b.txt"
            shared = "本工程施工导流采用枯水期围堰一次性拦断河床方案，围堰设计洪水标准为十年一遇。"
            f1.write_text(shared + "\n其余独立章节内容一二三四五六七八九十。\n", encoding="utf-8")
            f2.write_text(shared + "\n另一份标段完全不同的施工安排内容甲乙丙丁。\n", encoding="utf-8")

            proc = subprocess.run(
                [sys.executable, str(SCRIPTS / "duplicate_check.py"), str(f1), str(f2)],
                capture_output=True, text=True, encoding="utf-8", check=True,
            )
            data = json.loads(proc.stdout)
            self.assertGreater(data["pairs"][0]["overall_similarity_pct"], 20.0)
            self.assertEqual(len(data["pairs"][0]["high_similarity_paragraphs"]), 1)


if __name__ == "__main__":
    unittest.main()
