#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
多标书防雷同/串标查重引擎 (bid-writer-pro v2.0)
基于 OpenBidKit_Yibiao (AGPL-3.0) 的查重思路，纯本地计算多份标书（DOCX/MD/TXT）间的文本重合率与高度相似段落。
"""

import argparse
import json
import pathlib
import re
import sys

import docx


def load_paragraphs(path: pathlib.Path) -> list[str]:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        doc = docx.Document(str(path))
        paras = [p.text.strip() for p in doc.paragraphs if len(p.text.strip()) >= 15]
        return paras
    lines = path.read_text(encoding="utf-8", errors="ignore").splitlines()
    return [ln.strip() for ln in lines if len(ln.strip()) >= 15]


def normalize(text: str) -> str:
    return re.sub(r"[\s\W_]+", "", text)


def char_ngrams(text: str, n: int = 8) -> set[str]:
    norm = normalize(text)
    if len(norm) < n:
        return {norm} if norm else set()
    return {norm[i : i + n] for i in range(len(norm) - n + 1)}


def jaccard(a: set[str], b: set[str]) -> float:
    if not a or not b:
        return 0.0
    inter = len(a & b)
    union = len(a | b)
    return inter / union if union else 0.0


def compare_pair(path_a: pathlib.Path, path_b: pathlib.Path, threshold: float = 0.55) -> dict:
    paras_a = load_paragraphs(path_a)
    paras_b = load_paragraphs(path_b)
    full_a = char_ngrams("".join(paras_a), n=8)
    full_b = char_ngrams("".join(paras_b), n=8)
    doc_sim = round(jaccard(full_a, full_b) * 100, 2)

    b_grams = [(i + 1, p, char_ngrams(p, n=6)) for i, p in enumerate(paras_b)]
    hits = []
    for idx_a, pa in enumerate(paras_a, start=1):
        ga = char_ngrams(pa, n=6)
        if len(ga) < 5:
            continue
        best_sim = 0.0
        best_b = None
        for idx_b, pb, gb in b_grams:
            sim = jaccard(ga, gb)
            if sim > best_sim:
                best_sim = sim
                best_b = (idx_b, pb)
        if best_sim >= threshold and best_b is not None:
            hits.append({
                "similarity_pct": round(best_sim * 100, 1),
                "para_a": f"#{idx_a}: {pa[:80]}",
                "para_b": f"#{best_b[0]}: {best_b[1][:80]}",
            })

    hits.sort(key=lambda x: x["similarity_pct"], reverse=True)
    return {
        "file_a": str(path_a.name),
        "file_b": str(path_b.name),
        "overall_similarity_pct": doc_sim,
        "high_similarity_paragraphs": hits[:20],
        "risk_level": "high" if doc_sim >= 35 or len(hits) >= 5 else ("medium" if doc_sim >= 15 or hits else "low"),
    }


def main():
    parser = argparse.ArgumentParser(description="多标书查重引擎")
    parser.add_argument("files", nargs="+", help="至少 2 份标书文件路径 (.docx / .md / .txt)")
    parser.add_argument("--threshold", type=float, default=0.55, help="段落相似度阈值 (默认 0.55)")
    args = parser.parse_args()

    paths = [pathlib.Path(f) for f in args.files]
    for p in paths:
        if not p.exists():
            print(f"文件不存在: {p}", file=sys.stderr)
            sys.exit(1)
    if len(paths) < 2:
        print("至少需要提供 2 份文件进行比对", file=sys.stderr)
        sys.exit(1)

    pairs = []
    for i in range(len(paths)):
        for j in range(i + 1, len(paths)):
            pairs.append(compare_pair(paths[i], paths[j], args.threshold))

    print(json.dumps({"pairs": pairs}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
