import streamlit as st, mammoth, re, os, uuid, requests, base64
from bs4 import BeautifulSoup

HR_STYLE_MAP = "p[style-name='Horizontal Line'] => hr"

def convert_docx_to_html(docx_file):
    image_dir = "/tmp/mammoth_images"
    os.makedirs(image_dir, exist_ok=True)
    image_map = {}

    def save_image(i):
        ext = i.content_type.split("/")[-1]
        name = f"img_{uuid.uuid4().hex[:8]}.{ext}"
        path = os.path.join(image_dir, name)
        with open(path, "wb") as f:
            with i.open() as img_stream:
                data = img_stream.read()
                f.write(data)
        with open(path, "rb") as f:
            encoded = base64.b64encode(f.read()).decode("utf-8")
        uri = f"data:image/{ext};base64,{encoded}"
        image_map[uri] = path
        return {"src": uri}

    result = mammoth.convert_to_html(docx_file, convert_image=mammoth.images.inline(save_image), style_map=HR_STYLE_MAP)
    html = re.sub(r"<p[^>]*>[-]{3,}\s*</p>", "<hr />", result.value, flags=re.I)

    soup = BeautifulSoup(html, "html.parser")
    for tbl in soup.find_all("table"):
        tbl["style"] = "border-collapse:collapse;width:100%;"
        for r_idx, row in enumerate(tbl.find_all("tr")):
            for cell in row.find_all(["th", "td"]):
                cell["style"] = (
                    "border:1px solid black;padding:4px;text-align:center;vertical-align:middle;"
                )
                if r_idx == 0:
                    cell.name = "th"
                    cell["style"] += "background-color:#f2f2f2;font-weight:bold;"
    return str(soup), image_map

def table_to_md(tbl):
    rows = [[c.get_text(" ", strip=True) for c in r.find_all(["th", "td"])] for r in tbl.find_all("tr")]
    if not rows: return ""
    widths = [max(len(row[i]) if i < len(row) else 0 for row in rows) for i in range(len(rows[0]))]
    fmt = lambda r: "| " + " | ".join((r[i] if i < len(r) else "").ljust(widths[i]) for i in range(len(widths))) + " |"
    return "```\n" + "\n".join(fmt(r) for r in rows) + "\n```"

def parse_html_blocks(html):
    soup = BeautifulSoup(html, "html.parser")
    return [tag for tag in soup.find_all(recursive=False) if tag.name in ["h1", "h2", "h3", "p", "ul", "ol", "hr", "table", "img"]]

def convert_tag_text_with_links(tag):
    parts = []
    for content in tag.contents:
        if isinstance(content, str):
            parts.append(content)
        elif content.name == "a" and content.get("href"):
            href = content["href"]
            text = content.get_text(" ", strip=True)
            parts.append(f"[{text}]({href})")
        else:
            parts.append(content.get_text(" ", strip=True))
    return ''.join(parts).strip()

def blocks_to_md(blocks, link=None, use_title=True, image_map=None):
    out = []

    def walk_list(tag, depth):
        items = []
        for li in tag.find_all("li", recursive=False):
            inner_text = convert_tag_text_with_links(li.find(text=True, recursive=False) or "")
            bullet = "  " * depth + "- " + inner_text.strip()
            items.append((bullet, None))
            for child in li.find_all(["ul", "ol"], recursive=False):
                items.extend(walk_list(child, depth + 1))
        return items

    for tag in blocks:
        nm = tag.name
        if nm in ["h1", "h2", "h3"]:
            lvl = {"h1": "#", "h2": "##", "h3": "###"}[nm]
            text = convert_tag_text_with_links(tag)
            out.append((f"{lvl} **{text}**", None))
        elif nm == "p":
            text = convert_tag_text_with_links(tag)
            if text:
                out.append((text, None))
        elif nm in ["ul", "ol"]:
            out.extend(walk_list(tag, 0))
        elif nm in ["hr", "br"]:
            out.append(("------------------------", None))
        elif nm == "table":
            out.append((table_to_md(tag), None))
        elif nm == "img":
            src = tag.get("src", "")
            img_path = image_map.get(src)
            if not img_path and src.startswith("data:image/") and image_map:
                img_path = list(image_map.values())[-1]
            out.append(("[[IMAGE]]", img_path))

    if use_title and out:
        txt, img = out[0]
        if img is None:
            core = txt.lstrip("# ").strip()
            out[0] = (f"# [**{core}**]({link})" if link else f"# **{core}**", None)
    return out

# ===== UI Section =====
st.set_page_config("📄 Word → Discord 변환기", layout="wide")
st.title("📄 Word → Discord 변환기")

docx = st.file_uploader("📎 .docx 업로드", type=["docx"])
link = st.text_input("🔗 제목 링크 (선택)")
use_title = st.checkbox("첫 줄을 제목으로 처리", value=True)
webhook = st.text_input("📬 Discord Webhook URL")

if docx:
    html, image_map = convert_docx_to_html(docx)
    blocks = parse_html_blocks(html)
    md_blocks = blocks_to_md(blocks, link, use_title, image_map=image_map)
    md_preview = "\n\n".join(t for t, i in md_blocks if t)
    html_preview = str(BeautifulSoup(html, "html.parser").prettify())

    tab_html, tab_md = st.tabs(["HTML", "Discord Markdown"])

    with tab_html:
        st.text_area("HTML", html_preview, height=400)

    with tab_md:
        st.text_area("Markdown", md_preview, height=400)
