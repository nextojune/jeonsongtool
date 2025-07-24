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

def blocks_to_md(blocks, link=None, use_title=True, image_map=None):
    out = []
    for tag in blocks:
        nm = tag.name
        if nm in ["h1", "h2", "h3"]:
            lvl = {"h1": "#", "h2": "##", "h3": "###"}[nm]
            out.append((f"{lvl} **{tag.get_text(' ', strip=True)}**", None))
        elif nm == "p":
            txt = tag.get_text(" ", strip=True)
            if txt: out.append((txt, None))
        elif nm in ["ul", "ol"]:
            for li in tag.find_all("li", recursive=True):
                d = len([p for p in li.parents if p.name in ["ul", "ol"]]) - 1
                bullet = "  " * d + "- " + li.get_text(" ", strip=True)
                out.append((bullet, None))
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

def group_md_blocks_for_sending(md_blocks, limit=1900):
    groups = []
    buf = ""
    for txt, img in md_blocks:
        if img or txt.startswith("```"):
            if buf.strip():
                groups.append((buf.strip(), None))
                buf = ""
            groups.append((txt, img))
        elif len(buf) + len(txt) + 2 > limit:
            groups.append((buf.strip(), None))
            buf = txt
        else:
            buf += ("\n" if buf else "") + txt
    if buf.strip():
        groups.append((buf.strip(), None))
    return groups

def send_discord(webhook, md_groups):
    for i, (txt, img_path) in enumerate(md_groups):
        payload = {}; files = {}

        if img_path:
            fname = os.path.basename(img_path)
            try:
                files["file"] = (fname, open(img_path, "rb"))
            except Exception as e:
                st.error(f"❌ 이미지 파일 열기 실패: {img_path}\n{e}")
                continue
            payload["content"] = txt if txt.strip() else "📎 이미지 첨부"
        elif txt.strip():
            payload["content"] = txt
        else:
            continue

        res = requests.post(webhook, data=payload, files=files if files else None)
        if res.status_code not in (200, 204):
            st.error(f"❌ 전송 실패 {i+1}/{len(md_groups)} - HTTP {res.status_code}: {res.text[:150]}")
            return
    st.success("✅ 순차 전송 완료")

def send_discord_embed(webhook, md_groups):
    for i, (txt, img_path) in enumerate(md_groups):
        payload = {
            "embeds": [{
                "title": "📄 Word → Discord 변환기",
                "description": txt if len(txt) < 2048 else txt[:2045] + "...",
                "color": 0x00BFFF
            }]
        }
        files = {}
        if img_path:
            fname = os.path.basename(img_path)
            try:
                files["file"] = (fname, open(img_path, "rb"))
            except Exception as e:
                st.error(f"❌ 이미지 파일 열기 실패: {img_path}\n{e}")
                continue

        res = requests.post(webhook, json=payload, files=files if files else None)
        if res.status_code not in (200, 204):
            st.error(f"❌ Embed 전송 실패 {i+1}/{len(md_groups)} - HTTP {res.status_code}: {res.text[:150]}")
            return
    st.success("✅ Embed 전송 완료")

def copy_btn(label, uid, text):
    import html as _h
    b64 = base64.b64encode(text.encode()).decode()
    btn_id = f"btn_{uid}_{uuid.uuid4().hex[:6]}"
    js = f"""
    <script>
    function copy_{btn_id}() {{
        navigator.clipboard.writeText(atob(\"{b64}\"));
    }}
    </script>
    <style>
    button.copy-btn {{
        position: absolute; top: 4px; right: 8px;
        background: rgba(0,0,0,0.3); color: white; border: none;
        padding: 2px 6px; font-size: 12px; border-radius: 4px;
        cursor: pointer; z-index: 999;
    }}
    </style>
    <button onclick="copy_{btn_id}()" class="copy-btn">{_h.escape(label)}</button>
    """
    st.markdown(js, unsafe_allow_html=True)

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
    md_groups = group_md_blocks_for_sending(md_blocks)
    md_preview = "\n\n".join(t for t, i in md_groups if t)
    html_preview = str(BeautifulSoup(html, "html.parser").prettify())

    tab_html, tab_md = st.tabs(["HTML", "Discord Markdown"])

    with tab_html:
        copy_btn("📋 복사", "html", html_preview)
        st.text_area("HTML", html_preview, height=400)

    with tab_md:
        copy_btn("📋 복사", "md", md_preview)
        st.text_area("Markdown", md_preview, height=400)

        col1, col2 = st.columns(2)
        with col1:
            if st.button("📤 순차 전송 (텍스트/이미지)"):
                send_discord(webhook, md_groups)
        with col2:
            if st.button("🖼️ Embed 전송"):
                send_discord_embed(webhook, md_groups)
