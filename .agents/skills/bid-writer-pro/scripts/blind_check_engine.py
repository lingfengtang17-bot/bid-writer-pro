#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word 暗标合规检查与一键自动修复引擎 (bid-writer-pro v2.0)
基于 OpenBidKit_Yibiao (AGPL-3.0) 的暗标规则体系，使用纯 Python (python-docx + lxml) 本地运行：
- 零外部 API 依赖、不上传企业文件
- 支持按照地区规范预设（如贵州2026、黑龙江2026、内蒙、河北、四川、水利等）逐条审计 .docx
- 支持 --fix 生成脱敏合规副本 (<原名>_已修复.docx)，绝不覆盖原件
"""

import argparse
import json
import pathlib
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

SKILL_DIR = pathlib.Path(__file__).resolve().parent.parent
PRESETS_PATH = SKILL_DIR / "references" / "regional_presets.json"

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
TWIPS_PER_CM = 566.929133858


def load_profile(preset_id: str | None, config_path: str | None) -> tuple[str, dict]:
    if config_path:
        raw = json.loads(pathlib.Path(config_path).read_text(encoding="utf-8"))
        profile = raw["profile"] if "profile" in raw else raw
        return profile.get("name", config_path), profile
    data = json.loads(PRESETS_PATH.read_text(encoding="utf-8"))
    pid = preset_id or data.get("default_preset_id", "general")
    if pid not in data["presets"]:
        raise ValueError(f"未知规范预设: {pid}")
    return data["presets"][pid]["name"], data["presets"][pid]


def twips_to_cm(twips: int | None) -> float | None:
    if twips is None:
        return None
    return round(twips / TWIPS_PER_CM, 2)


def get_run_font_cn(run) -> str:
    rPr = run._r.find(qn("w:rPr"))
    if rPr is not None:
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is not None:
            east = rFonts.get(qn("w:eastAsia"))
            if east:
                return east
            ascii_f = rFonts.get(qn("w:ascii"))
            if ascii_f:
                return ascii_f
    return run.font.name or ""


def get_run_size_pt(run) -> float | None:
    if run.font.size is not None:
        return round(run.font.size.pt, 1)
    rPr = run._r.find(qn("w:rPr"))
    if rPr is not None:
        sz = rPr.find(qn("w:sz"))
        if sz is not None and sz.get(qn("w:val")):
            return round(int(sz.get(qn("w:val"))) / 2.0, 1)
    return None


def inspect_xml_features(docx_path: pathlib.Path) -> dict:
    info = {
        "has_comments": False,
        "has_revisions": False,
        "textbox_count": 0,
        "hyperlink_count": 0,
        "page_field_count": 0,
    }
    with zipfile.ZipFile(docx_path, "r") as zf:
        names = set(zf.namelist())
        if "word/comments.xml" in names:
            c_root = ET.fromstring(zf.read("word/comments.xml"))
            if len(list(c_root)) > 0:
                info["has_comments"] = True
        for xml_name in names:
            if not (xml_name.startswith("word/") and xml_name.endswith(".xml")):
                continue
            raw = zf.read(xml_name).decode("utf-8", errors="ignore")
            if "<w:ins " in raw or "<w:del " in raw or "<w:ins>" in raw or "<w:del>" in raw:
                info["has_revisions"] = True
            info["textbox_count"] += raw.count("txbxContent") // 2
            if xml_name == "word/document.xml":
                info["hyperlink_count"] += raw.count("<w:hyperlink")
            if "PAGE" in raw or "NUMPAGES" in raw:
                info["page_field_count"] += len(re.findall(r"\b(?:PAGE|NUMPAGES)\b", raw))
    return info


def inspect_app_company(docx_path: pathlib.Path) -> str:
    with zipfile.ZipFile(docx_path, "r") as zf:
        if "docProps/app.xml" not in zf.namelist():
            return ""
        root = ET.fromstring(zf.read("docProps/app.xml"))
        for el in root.iter():
            if el.tag.endswith("Company") and el.text and el.text.strip():
                return el.text.strip()
    return ""


def audit_docx(docx_path: pathlib.Path, profile: dict, extra_keywords: list[str] | None = None) -> dict:
    doc = docx.Document(str(docx_path))
    issues = []

    def add_issue(rule_id: str, severity: str, category: str, message: str, count: int = 1, samples: list | None = None, auto_fixable: bool = True):
        issues.append({
            "rule_id": rule_id,
            "severity": severity,
            "category": category,
            "message": message,
            "count": count,
            "auto_fixable": auto_fixable,
            "samples": (samples or [])[:5],
        })

    # 1. 页面与页边距
    page_cfg = profile.get("page", {})
    tol_cm = float(page_cfg.get("tolerance_margin_cm", 0.1))
    margin_mismatches = []
    header_footer_hits = []
    column_hits = []

    for idx, sec in enumerate(doc.sections, start=1):
        w_cm = twips_to_cm(sec.page_width.twips) if sec.page_width else None
        h_cm = twips_to_cm(sec.page_height.twips) if sec.page_height else None
        if page_cfg.get("paper") == "A4" and w_cm and h_cm:
            if not (abs(w_cm - 21.0) <= 0.3 and abs(h_cm - 29.7) <= 0.3):
                margin_mismatches.append(f"第{idx}节纸张 {w_cm}x{h_cm}cm（应为 A4 21.0x29.7cm）")

        checks = [
            ("上边距", sec.top_margin.twips if sec.top_margin else None, page_cfg.get("margin_top_cm")),
            ("下边距", sec.bottom_margin.twips if sec.bottom_margin else None, page_cfg.get("margin_bottom_cm")),
            ("左边距", sec.left_margin.twips if sec.left_margin else None, page_cfg.get("margin_left_cm")),
            ("右边距", sec.right_margin.twips if sec.right_margin else None, page_cfg.get("margin_right_cm")),
        ]
        for label, tw, target in checks:
            if target is not None and tw is not None:
                actual = twips_to_cm(tw)
                if abs(actual - float(target)) > tol_cm:
                    margin_mismatches.append(f"第{idx}节{label} {actual}cm（要求 {target}cm）")

        if profile.get("structure", {}).get("forbid_header_content"):
            ht = "".join(p.text.strip() for p in sec.header.paragraphs)
            if ht:
                header_footer_hits.append(f"第{idx}节页眉含内容: {ht[:30]}")
        if profile.get("structure", {}).get("forbid_footer_content"):
            ft = "".join(p.text.strip() for p in sec.footer.paragraphs)
            if ft:
                header_footer_hits.append(f"第{idx}节页脚含内容: {ft[:30]}")

        if page_cfg.get("forbid_columns"):
            cols = sec._sectPr.find(qn("w:cols"))
            if cols is not None and cols.get(qn("w:num")) and int(cols.get(qn("w:num"))) > 1:
                column_hits.append(f"第{idx}节存在多栏排版（{cols.get(qn('w:num'))}栏）")

    if margin_mismatches:
        add_issue("page.margins", "high", "page", "页面尺寸或页边距不符合规范", len(margin_mismatches), margin_mismatches, True)
    if header_footer_hits:
        add_issue("struct.header_footer", "high", "structure", "暗标禁止页眉/页脚内容", len(header_footer_hits), header_footer_hits, True)
    if column_hits:
        add_issue("page.columns", "medium", "page", "暗标禁止多栏排版", len(column_hits), column_hits, True)

    # 2. 文档属性泄密
    privacy_cfg = profile.get("privacy", {})
    if privacy_cfg.get("check_doc_props"):
        cp = doc.core_properties
        leaks = []
        for field_name in ("author", "last_modified_by", "comments", "keywords", "subject"):
            val = getattr(cp, field_name, None)
            if val and str(val).strip():
                leaks.append(f"文档属性 {field_name}={str(val).strip()[:30]}")
        comp = inspect_app_company(docx_path)
        if comp:
            leaks.append(f"文档属性 Company={comp[:30]}")
        if leaks:
            add_issue("struct.doc_props", "high", "privacy", "Word 文档属性隐藏了作者/单位等身份信息", len(leaks), leaks, True)

    # 3. XML 隐藏风险（批注、修订、文本框、超链接、页码域）
    xml_info = inspect_xml_features(docx_path)
    struct_cfg = profile.get("structure", {})
    if struct_cfg.get("check_comments") and xml_info["has_comments"]:
        add_issue("struct.comment", "high", "structure", "文档内存在批注，暗标必须清除", 1, ["存在 word/comments.xml 批注"], True)
    if struct_cfg.get("check_revisions") and xml_info["has_revisions"]:
        add_issue("struct.revision", "high", "structure", "文档存在未接受/未清除的修订痕迹", 1, ["存在 w:ins / w:del 修订标签"], True)
    if struct_cfg.get("check_textbox") and xml_info["textbox_count"] > 0:
        add_issue("struct.textbox", "medium", "structure", "检测到文本框（暗标通常禁止文本框）", xml_info["textbox_count"], [f"共 {xml_info['textbox_count']} 处文本框"], False)
    if struct_cfg.get("check_hyperlink") and xml_info["hyperlink_count"] > 0:
        add_issue("struct.hyperlink", "medium", "structure", "检测到超链接（可能暴露网址或跳转标记）", xml_info["hyperlink_count"], [f"共 {xml_info['hyperlink_count']} 处超链接"], False)
    if struct_cfg.get("page_number") == "forbid" and xml_info["page_field_count"] > 0:
        add_issue("struct.page_number_text", "high", "structure", "暗标禁止出现页码域代码", xml_info["page_field_count"], ["检测到 PAGE/NUMPAGES 域"], True)

    # 4. 段落与文字样式 + 敏感身份内容扫描
    body_cfg = profile.get("body", {})
    target_font = body_cfg.get("font_cn")
    target_size = body_cfg.get("size_pt")
    tol_pt = float(body_cfg.get("tolerance_pt", 0.5))
    target_line_pt = body_cfg.get("line_pt")
    tol_line_pt = float(body_cfg.get("tolerance_line_pt", 1.0))

    empty_paras = []
    font_samples = []
    size_samples = []
    line_samples = []
    effect_samples = []
    color_samples = []
    toc_hits = []
    identity_hits = []

    identity_patterns = []
    if privacy_cfg.get("check_phone"):
        identity_patterns.append(("手机号", re.compile(r"(?<!\d)1[3-9]\d{9}(?!\d)")))
    if privacy_cfg.get("check_id_card"):
        identity_patterns.append(("身份证号", re.compile(r"(?<!\d)\d{17}[\dXx](?!\d)")))
    if privacy_cfg.get("check_credit_code"):
        identity_patterns.append(("统一社会信用代码", re.compile(r"(?<![A-Z0-9])9[1-9][0-9A-Z]{16}(?![A-Z0-9])")))
    if privacy_cfg.get("check_email"):
        identity_patterns.append(("邮箱", re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")))
    if privacy_cfg.get("check_url"):
        identity_patterns.append(("网址", re.compile(r"https?://[^\s]+|www\.[^\s]+", re.I)))
    if privacy_cfg.get("check_company_suffix"):
        identity_patterns.append(("疑似单位名称", re.compile(r"[\u4e00-\u9fa5]{2,20}(?:有限责任公司|股份有限公司|有限公司|集团公司|工程局|设计研究院|设计院)")))
    for kw in (extra_keywords or []):
        if kw.strip():
            identity_patterns.append((f"指定敏感词({kw})", re.compile(re.escape(kw.strip()))))

    for p_idx, p in enumerate(doc.paragraphs, start=1):
        txt = p.text.strip()
        style_name = p.style.name if p.style else ""
        is_heading = style_name.startswith("Heading")

        if struct_cfg.get("check_empty_paragraphs") and not txt:
            has_drawing = bool(p._p.findall(f".//{{{NS_W}}}drawing"))
            has_br = bool(p._p.findall(f".//{{{NS_W}}}br"))
            if not has_drawing and not has_br:
                empty_paras.append(f"第{p_idx}段为空段")

        if struct_cfg.get("forbid_toc") and txt in ("目录", "目 录"):
            toc_hits.append(f"第{p_idx}段出现目录标题")

        # 身份信息扫描
        if txt:
            for label, pat in identity_patterns:
                for m in pat.finditer(txt):
                    identity_hits.append(f"第{p_idx}段[{label}]: {m.group(0)[:30]}")

        if not txt:
            continue

        # 行距检查（针对正文段落）
        if not is_heading and body_cfg.get("line_rule") == "exact" and target_line_pt:
            ls = p.paragraph_format.line_spacing
            if ls is None:
                line_samples.append(f"第{p_idx}段未显式设置固定行距: {txt[:18]}")
            else:
                try:
                    actual_pt = round(ls.pt, 1) if hasattr(ls, "pt") else None
                    if actual_pt is not None and abs(actual_pt - float(target_line_pt)) > tol_line_pt:
                        line_samples.append(f"第{p_idx}段行距 {actual_pt}磅（要求 {target_line_pt}磅）: {txt[:16]}")
                except Exception:
                    pass

        for run in p.runs:
            if not run.text.strip():
                continue
            if not is_heading and target_font:
                fcn = get_run_font_cn(run)
                if fcn and target_font not in fcn:
                    font_samples.append(f"第{p_idx}段字体「{fcn}」（要求 {target_font}）: {run.text[:16]}")
            if not is_heading and target_size:
                spt = get_run_size_pt(run)
                if spt is not None and abs(spt - float(target_size)) > tol_pt:
                    size_samples.append(f"第{p_idx}段字号 {spt}pt（要求 {target_size}pt）: {run.text[:16]}")

            if not is_heading:
                if body_cfg.get("allow_bold") is False and run.bold:
                    effect_samples.append(f"第{p_idx}段禁止加粗: {run.text[:16]}")
                if body_cfg.get("allow_italic") is False and run.italic:
                    effect_samples.append(f"第{p_idx}段禁止斜体: {run.text[:16]}")
                if body_cfg.get("allow_underline") is False and run.underline:
                    effect_samples.append(f"第{p_idx}段禁止下划线: {run.text[:16]}")
            if body_cfg.get("allow_color") is False and run.font.color and run.font.color.rgb:
                rgb = str(run.font.color.rgb).upper()
                if rgb not in ("000000", "AUTO"):
                    color_samples.append(f"第{p_idx}段存在彩色字体 #{rgb}: {run.text[:16]}")
            if body_cfg.get("allow_highlight") is False and run.font.highlight_color is not None:
                color_samples.append(f"第{p_idx}段存在高亮底纹: {run.text[:16]}")

    if empty_paras:
        add_issue("struct.empty_paragraph", "low", "structure", "存在多余空段", len(empty_paras), empty_paras, True)
    if toc_hits:
        add_issue("struct.toc", "high", "structure", "该地区暗标禁止设置目录", len(toc_hits), toc_hits, False)
    if font_samples:
        add_issue("font.chinese", "medium", "font", f"正文字体不符合要求（应为 {target_font}）", len(font_samples), font_samples, True)
    if size_samples:
        add_issue("font.size", "medium", "font", f"正文字号不符合要求（应为 {target_size}pt）", len(size_samples), size_samples, True)
    if line_samples:
        add_issue("para.line_spacing", "medium", "paragraph", f"正文固定行距不符合要求（应为 {target_line_pt}磅）", len(line_samples), line_samples, True)
    if effect_samples:
        add_issue("font.effects", "high", "font", "暗标正文禁止加粗/斜体/下划线标记", len(effect_samples), effect_samples, True)
    if color_samples:
        add_issue("font.color", "high", "font", "暗标禁止彩色文字或高亮底纹", len(color_samples), color_samples, True)
    if identity_hits:
        add_issue("privacy.keyword", "high", "privacy", "疑似包含投标人身份信息/联系方式（暗标极高危废标风险）", len(identity_hits), identity_hits, False)

    summary = {"high": 0, "medium": 0, "low": 0, "total_issues": len(issues)}
    for it in issues:
        summary[it["severity"]] += it["count"]

    return {"summary": summary, "issues": issues}


def fix_docx(docx_path: pathlib.Path, out_path: pathlib.Path, profile: dict) -> pathlib.Path:
    doc = docx.Document(str(docx_path))
    page_cfg = profile.get("page", {})
    body_cfg = profile.get("body", {})
    struct_cfg = profile.get("structure", {})
    privacy_cfg = profile.get("privacy", {})

    # 1. 清除核心文档属性中的身份线索
    if privacy_cfg.get("check_doc_props"):
        cp = doc.core_properties
        cp.author = ""
        cp.last_modified_by = ""
        cp.comments = ""
        cp.keywords = ""
        cp.subject = ""
        cp.category = ""

    # 2. 修复页面与页眉页脚
    for sec in doc.sections:
        if page_cfg.get("paper") == "A4":
            sec.page_width = Cm(21.0)
            sec.page_height = Cm(29.7)
        if page_cfg.get("margin_top_cm") is not None:
            sec.top_margin = Cm(float(page_cfg["margin_top_cm"]))
        if page_cfg.get("margin_bottom_cm") is not None:
            sec.bottom_margin = Cm(float(page_cfg["margin_bottom_cm"]))
        if page_cfg.get("margin_left_cm") is not None:
            sec.left_margin = Cm(float(page_cfg["margin_left_cm"]))
        if page_cfg.get("margin_right_cm") is not None:
            sec.right_margin = Cm(float(page_cfg["margin_right_cm"]))
        if page_cfg.get("forbid_columns"):
            cols = sec._sectPr.find(qn("w:cols"))
            if cols is not None:
                cols.set(qn("w:num"), "1")

        if struct_cfg.get("forbid_header_content"):
            sec.header.is_linked_to_previous = False
            for hp in sec.header.paragraphs:
                hp.text = ""
        if struct_cfg.get("forbid_footer_content") or struct_cfg.get("page_number") == "forbid":
            sec.footer.is_linked_to_previous = False
            for fp in sec.footer.paragraphs:
                fp.text = ""

    # 3. 修复段落和文字格式
    target_font = body_cfg.get("font_cn") or "宋体"
    target_size = float(body_cfg.get("size_pt") or 14)
    target_line_pt = body_cfg.get("line_pt")

    for p in list(doc.paragraphs):
        txt = p.text.strip()
        if struct_cfg.get("check_empty_paragraphs") and not txt:
            has_drawing = bool(p._p.findall(f".//{{{NS_W}}}drawing"))
            has_br = bool(p._p.findall(f".//{{{NS_W}}}br"))
            if not has_drawing and not has_br:
                p._element.getparent().remove(p._element)
                continue

        style_name = p.style.name if p.style else ""
        is_heading = style_name.startswith("Heading")

        if not is_heading:
            pf = p.paragraph_format
            if body_cfg.get("line_rule") == "exact" and target_line_pt:
                pf.line_spacing = Pt(float(target_line_pt))
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            if body_cfg.get("first_line_indent_chars") is not None:
                pf.first_line_indent = Pt(target_size * float(body_cfg["first_line_indent_chars"]))

        for run in p.runs:
            if body_cfg.get("allow_color") is False:
                run.font.color.rgb = RGBColor(0, 0, 0)
            if body_cfg.get("allow_highlight") is False:
                run.font.highlight_color = None
            if not is_heading:
                if body_cfg.get("allow_bold") is False:
                    run.bold = False
                if body_cfg.get("allow_italic") is False:
                    run.italic = False
                if body_cfg.get("allow_underline") is False:
                    run.underline = False
                run.font.name = target_font
                rPr = run._r.get_or_add_rPr()
                rFonts = rPr.find(qn("w:rFonts"))
                if rFonts is None:
                    rFonts = OxmlElement("w:rFonts")
                    rPr.insert(0, rFonts)
                rFonts.set(qn("w:eastAsia"), target_font)
                rFonts.set(qn("w:ascii"), target_font)
                rFonts.set(qn("w:hAnsi"), target_font)
                run.font.size = Pt(target_size)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))

    # 4. 从 ZIP 内清理批注与 Company 扩展属性（如开启）
    if struct_cfg.get("check_comments") or privacy_cfg.get("check_doc_props"):
        tmp_path = out_path.with_suffix(".tmp.docx")
        with zipfile.ZipFile(out_path, "r") as zin, zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if struct_cfg.get("check_comments") and item.filename == "word/comments.xml":
                    continue
                data = zin.read(item.filename)
                if privacy_cfg.get("check_doc_props") and item.filename == "docProps/app.xml":
                    text = data.decode("utf-8", errors="ignore")
                    text = re.sub(r"<Company>.*?</Company>", "<Company></Company>", text, flags=re.S)
                    data = text.encode("utf-8")
                zout.writestr(item, data)
        tmp_path.replace(out_path)

    return out_path


def main():
    parser = argparse.ArgumentParser(description="Word 暗标合规检查与一键自动修复引擎")
    parser.add_argument("docx_file", help="待检查的 .docx 文件路径")
    parser.add_argument("--preset", "-p", help="地区/行业预设 ID（如 guizhou_2026 / water_conservancy / general）")
    parser.add_argument("--config", "-c", help="由 regional_matcher.py 生成的 JSON 规范文件")
    parser.add_argument("--keyword", "-k", action="append", default=[], help="额外需要查杀的敏感词（如我方公司名、项目经理姓名）")
    parser.add_argument("--fix", action="store_true", help="自动生成合规修复副本 (<原名>_已修复.docx)")
    parser.add_argument("--out", "-o", help="指定修复输出文件路径")
    args = parser.parse_args()

    src = pathlib.Path(args.docx_file)
    if not src.exists():
        print(f"文件不存在: {src}", file=sys.stderr)
        sys.exit(1)

    profile_name, profile = load_profile(args.preset, args.config)
    before = audit_docx(src, profile, args.keyword)

    result = {
        "file": str(src),
        "profile_name": profile_name,
        "before_audit": before,
    }

    if args.fix:
        fixed_path = pathlib.Path(args.out) if args.out else src.with_name(f"{src.stem}_已修复.docx")
        fix_docx(src, fixed_path, profile)
        after = audit_docx(fixed_path, profile, args.keyword)
        result["fixed_file"] = str(fixed_path)
        result["after_audit"] = after

    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
