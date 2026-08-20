#!/usr/bin/env python3
"""清理 Vision OCR 结果 → 干净的 Markdown（正文 + 章节标题）v3"""
import json, re, sys

COPYRIGHT_WORDS = ["ISBN", "图书在版编目", "著者：", "定价：", "承印者", "服务热线",
                   "版权所有", "CIP", "印张", "开本："]
PAGE_END_PUNCT = "。！？…”」』】》“”"
DECOR_LINES = {"—", "-", "–", "·", "■", "●", "◆", "一"}

def load(p):
    return json.load(open(p))

def is_skip_page(pdata):
    items = [it for it in pdata["items"] if it["t"].strip()]
    if len(items) < 5:
        return True
    text = "".join(it["t"] for it in items)
    if any(w in text for w in COPYRIGHT_WORDS):
        return True
    if "目录" in text and len(items) > 8:
        return True
    narrow = sum(1 for it in items if it["w"] < 0.35)
    if len(items) >= 10 and narrow / len(items) > 0.85:
        return True
    return False

def is_punct_line(t):
    return bool(re.fullmatch(r"[\s\-—–・·―—–−~_～〜．。•·\u2010-\u2015·‥…]+", t))

def join_title(parts):
    parts = [p for p in parts if p.strip()]
    parts = [p for p in parts if p not in DECOR_LINES and not is_punct_line(p)]
    out = ""
    for p in parts:
        if not out:
            out = p
            continue
        # 相邻字符去重（OCR 行首尾重复，如 俱+俱乐部 → 俱乐部）
        if out and p and out[-1] == p[0]:
            p = p[1:]
            if not p:
                continue
        if re.fullmatch(r"[\d０-９]+", p) or re.fullmatch(r"[\d０-９]+", out[-1]):
            out += " " + p
        else:
            out += p
    if out.startswith("序") and len(out) > 1 and out[1] not in "：:":
        out = "序：" + out[1:]
    if out.startswith("引") and len(out) > 1 and out[1] not in "：:":
        out = "引子：" + out[1:]
    out = out.replace("嬴家", "赢家").replace("乐部", "俱乐部").strip()
    return out

def detect_header_chapter(items, last_chapter):
    """页眉区（y>0.88）左侧章号+章名 → 页眉型章首标题"""
    header = [it for it in items if it["y"] > 0.88]
    left = sorted([it for it in header if it["x"] < 0.5], key=lambda it: -it["y"])
    chap = None
    for it in left:
        t = it["t"].strip()
        if re.fullmatch(r"\d{1,2}", t):
            chap = t
            break
    if chap is None or chap == last_chapter:
        return None, chap
    names = []
    for it in left:
        t = it["t"].strip()
        if re.fullmatch(r"\d{1,4}", t) or not t or is_punct_line(t):
            continue
        names.append(t)
    if not names:
        return None, chap
    return join_title([chap] + names), chap

def clean_page(pdata, last_chapter):
    items = [it for it in pdata["items"] if it["t"].strip()]
    # 页眉型章首（先于页眉删除提取）
    header_title, chap = detect_header_chapter(items, last_chapter)
    # 1. 删页眉与页脚数字
    body = []
    for it in items:
        if it["y"] > 0.88:
            continue
        if it["y"] < 0.06 and re.fullmatch(r"\d{1,4}", it["t"].strip()):
            continue
        body.append(it)
    if not body:
        return [], header_title, chap
    # 2. 独立标题检测：页面顶部连续窄行组
    top = sorted(body, key=lambda it: -it["y"])
    group = []
    prev_y = None
    for it in top:
        if it["w"] >= 0.35:
            break
        if prev_y is not None and (prev_y - it["y"]) > 0.15:
            break
        group.append(it)
        prev_y = it["y"]
    title = header_title
    title_lines = []
    if group:
        real = [it for it in group if not is_punct_line(it["t"].strip())]
        real_text = [it for it in real if not re.fullmatch(r"\d{1,2}", it["t"].strip())]
        first = real[0] if real else None
        ok = False
        if len(real_text) >= 2:
            ok = first is not None and first["h"] >= 0.024 and group[0]["y"] > 0.45
        elif len(real_text) == 1 and first is not None:
            t = real_text[0]["t"].strip()
            ok = (not t or t[-1] not in PAGE_END_PUNCT) and first["h"] >= 0.026 and group[0]["y"] > 0.45
        if ok and header_title is None:
            title_lines = group
            title = join_title([it["t"].strip() for it in group])
            m = re.match(r"^(\d{1,2})", title)
            if m:
                chap = m.group(1)
    body = [it for it in body if it not in title_lines]
    # 3. 双栏检测与重组
    body.sort(key=lambda it: -it["y"])
    if body:
        centers = [it["x"] + it["w"] / 2 for it in body]
        p_left = sum(1 for c in centers if c < 0.42) / len(centers)
        p_right = sum(1 for c in centers if c > 0.6) / len(centers)
        if p_left > 0.3 and p_right > 0.3:
            left = [it for it in body if it["x"] + it["w"] / 2 < 0.5]
            right = [it for it in body if it["x"] + it["w"] / 2 >= 0.5]
            left.sort(key=lambda it: (-it["y"], it["x"]))
            right.sort(key=lambda it: (-it["y"], it["x"]))
            body = left + right
    # 4. 行内清理（保留 OCR 原样，仅去空白）
    lines = []
    for it in body:
        t = re.sub(r"\s+", "", it["t"])
        if t:
            lines.append(t)
    return lines, title, chap

def end_of_para(line):
    if not line:
        return False
    last = line[-1]
    if last in "。！？…":
        return True
    if last in "”」』》":
        # 右引号：只有前一字是句读才算段末
        if len(line) >= 2 and line[-2] in "。！？…":
            return True
        return False
    if last in "“「『《":
        return False
    if line.endswith("——") or line.endswith("--"):
        return True
    if last in "·":
        return True
    return False

def build_md(pages):
    out = []
    buf = ""
    last_chapter = None
    for pdata in pages:
        if is_skip_page(pdata):
            continue
        lines, title, chap = clean_page(pdata, last_chapter)
        if chap:
            last_chapter = chap
        if title:
            if buf.strip():
                out.append(buf.strip())
                buf = ""
            out.append(f"# {title}")
        for line in lines:
            if not line:
                continue
            if end_of_para(line):
                buf += line
                out.append(buf.strip())
                buf = ""
            else:
                buf += line
    if buf.strip():
        out.append(buf.strip())
    return "\n\n".join(out)

if __name__ == "__main__":
    src, dst = sys.argv[1], sys.argv[2]
    pages = load(src)
    md = build_md(pages)
    with open(dst, "w", encoding="utf-8") as f:
        f.write(md + "\n")
    titles = [l for l in md.split("\n") if l.startswith("# ")]
    print(f"输出: {dst}")
    print(f"总字符: {len(md)}  标题数: {len(titles)}")
    for t in titles:
        print("  " + t)
