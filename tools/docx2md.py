#!/usr/bin/env python3
"""把整理好的《崩坏：星穹铁道》剧情 docx 转成 Markdown。

这些 docx 使用了一套固定的语义样式：
    Meta    —— 章节副标题 / 版本元数据
    Line    —— 英文台词（说话人加粗，形如 "March 7th: ..."）
    Zh      —— 对应中文台词
    Choice  —— 玩家可选分支（▸ 开头）
    Scene   —— 舞台提示 / 动作描述（括号开头）
    Note    —— 任务概要
    heading 1/2/3 —— 章 / 任务 / 目标

用法：
    python3 tools/docx2md.py          # 转换工作区里所有 docx
    python3 tools/docx2md.py 某文件.docx
"""

import html
import re
import sys
import zipfile
from pathlib import Path

# ---------------------------------------------------------------- 解析

# styleId -> 语义名。styleId 由 Word 生成，不稳定，所以优先用 styles.xml 里的 name。
KNOWN = {
    "Meta": "meta", "Line": "line", "Zh": "zh", "Choice": "choice",
    "Scene": "scene", "Note": "note",
    "heading 1": "h1", "heading 2": "h2", "heading 3": "h3",
    "Heading 1": "h1", "Heading 2": "h2", "Heading 3": "h3",
}
# 找不到 styles.xml 时的兜底（本批文件实测的映射）
FALLBACK = {
    "style4097": "meta", "style4098": "line", "style4099": "zh",
    "style4100": "choice", "style4101": "scene", "style4105": "note",
    "style1": "h1", "style2": "h2", "style3": "h3",
}

TAG = re.compile(r"<[^>]+>")
PARA = re.compile(r"<w:p[ >].*?</w:p>", re.S)
RUN = re.compile(r"<w:r[ >](.*?)</w:r>", re.S)
W_T = re.compile(r"<w:t[^>]*>(.*?)</w:t>", re.S)


def style_map(z):
    """styleId -> 语义名"""
    try:
        sx = z.read("word/styles.xml").decode("utf-8")
    except KeyError:
        return dict(FALLBACK)
    m = {}
    for sm in re.finditer(r'<w:style [^>]*w:styleId="([^"]+)".*?</w:style>', sx, re.S):
        sid, body = sm.group(1), sm.group(0)
        nm = re.search(r'<w:name w:val="([^"]+)"', body)
        if nm and nm.group(1) in KNOWN:
            m[sid] = KNOWN[nm.group(1)]
    return m or dict(FALLBACK)


def unescape(s):
    # 有些文件是双重转义的（&amp;#58; -> &#58;），所以解到不再变化为止
    for _ in range(3):
        n = html.unescape(s)
        if n == s:
            break
        s = n
    return s


def inline(para_xml):
    """把一个段落里的 run 拼成带 **粗体** 的行内 Markdown。"""
    out = []
    for rm in RUN.finditer(para_xml):
        run = rm.group(1)
        txt = "".join(W_T.findall(run))
        if not txt:
            continue
        txt = unescape(TAG.sub("", txt))
        bold = "<w:b/>" in run or re.search(r"<w:b [^>]*/>", run) or "<w:b " in run
        # 图片 run 里常带 alt 文本，丢掉
        if "<w:drawing" in run and not txt.strip():
            continue
        out.append((txt, bool(bold)))
    if not out:
        return ""

    # 合并相邻同样式的 run
    merged = []
    for txt, bold in out:
        if merged and merged[-1][1] == bold:
            merged[-1] = (merged[-1][0] + txt, bold)
        else:
            merged.append((txt, bold))

    # 粗体在这批文件里主要用于说话人名（"March 7th: " / "三月七："）。
    # 但表演提示也会被加粗（"*sigh*"），那种情况必须保持普通文本，
    # 否则会拼出 "**sigh*" 这种坏掉的 markdown。
    # 判据：只有以冒号结尾的粗体才算说话人。
    s = ""
    prev_bold = False
    for txt, bold in merged:
        if bold and txt.strip().endswith((":", "：", ":*")):
            core = txt.strip()
            if not core:
                continue
            if s and txt[:1].isspace() and not s.endswith(" "):
                s += " "
            s += f"**{core}**"
            prev_bold = True
        else:
            body = txt.lstrip(" ") if prev_bold else txt
            if prev_bold and body and ord(body[0]) < 128 and not s.endswith(" "):
                s += " "
            s += body
            prev_bold = False
    return s.strip()


def is_image_only(para_xml):
    return "<w:drawing" in para_xml and not TAG.sub("", "".join(W_T.findall(para_xml))).strip()


def convert(path):
    z = zipfile.ZipFile(path)
    doc = z.read("word/document.xml").decode("utf-8")
    smap = style_map(z)

    blocks = []          # (kind, text)
    for pm in PARA.finditer(doc):
        p = pm.group(0)
        st = re.search(r'<w:pStyle w:val="([^"]+)"', p)
        kind = smap.get(st.group(1)) if st else None
        if kind is None and st:
            kind = "plain"      # style0 之类
        text = inline(p)
        if not text:
            continue
        blocks.append((kind or "plain", text))

    # ------------------------------------------------------------ 渲染
    # 封面图注本来在最前面；挪到一级标题之后，别让 md 文件以图注开头。
    if (len(blocks) >= 2 and blocks[1][0] == "h1"
            and re.search(r"版本主视觉|图源：", blocks[0][1])):
        blocks[0], blocks[1] = blocks[1], blocks[0]

    out = []
    prev = None
    for kind, text in blocks:
        if kind == "h1":
            out.append(f"# {text}")
        elif kind == "h2":
            out.append(f"## {text}")
        elif kind == "h3":
            out.append(f"### {text}")
        elif kind == "meta":
            out.append(f"> {text}")
        elif kind == "note":
            # 有的 Note 正文本身就带 "*sigh*" 这类星号，再套一层斜体会把 markdown 弄坏
            out.append(f"> {text if '*' in text else f'*{text}*'}")
        elif kind == "zh":
            out.append(f"> {text}")
        elif kind == "choice":
            out.append(f"- {text}")
        elif kind == "scene":
            out.append(text if "*" in text else f"*{text}*")
        elif kind == "line":
            out.append(text)
        elif re.search(r"版本主视觉|图源：", text):
            # 封面图的图注：图片本身留在 docx 里没抽出来，所以标成斜体旁注
            out.append(f"*{text}*")
        else:                                   # plain / 无样式
            out.append(re.sub(r"^([#>+-]|\d+[.)])\s", r"\\\1 ", text))

    # 块间空行：列表项连续时不留空行
    body = []
    for i, line in enumerate(out):
        body.append(line)
        nxt = out[i + 1] if i + 1 < len(out) else None
        if nxt is None:
            continue
        if line.startswith("- ") and nxt.startswith("- "):
            continue
        body.append("")
    text = "\n".join(body).rstrip() + "\n"

    # 标题之后、以及正文之间多余的空白压缩
    text = re.sub(r"\n{3,}", "\n\n", text)

    # 反过来：Meta 之后紧跟另一个 Meta 时保持同块
    return text


def main(argv):
    root = Path(__file__).resolve().parent.parent
    targets = [Path(a).resolve() for a in argv[1:]] or sorted(root.rglob("*.docx"))
    for src in targets:
        if "~$" in src.name:
            continue
        md = convert(src)
        # "续闻" 那批文件名带 hsr.docx01_ 这样的排序前缀（合并工具的产物），
        # 输出 md 时去掉，保留原始的 "<版本号>_<标题>" 形式。
        stem = re.sub(r"^hsr\.docx\d+_", "", src.stem)
        dst = src.parent / f"{stem}.md"
        dst.write_text(md, encoding="utf-8")
        print(f"{src.relative_to(root)} -> {dst.relative_to(root)}  "
              f"({len(src.read_bytes())//1024}KB -> {len(md.encode())//1024}KB)")


if __name__ == "__main__":
    main(sys.argv)
