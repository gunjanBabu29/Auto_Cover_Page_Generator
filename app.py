
import io
import os
import math
import zipfile
import tempfile
from pathlib import Path

import streamlit as st
from PIL import Image, ImageDraw, ImageFont, ImageFilter
from docx import Document
from docx.shared import Inches

st.set_page_config(
    page_title="Proposal Cover Page Generator",
    page_icon="📄",
    layout="wide",
)

NAVY = "#0B2B68"
GREEN = "#63B52A"
ORANGE = "#F0A500"
LIGHT_GREEN = "#EAF4DE"
MID_GREY = "#D8D8D8"
TEXT = "#15366E"

FONT_CANDIDATES = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf",
    "/usr/share/fonts/truetype/liberation2/LiberationSerif-Bold.ttf",
]
FONT_REGULAR = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf",
    "/usr/share/fonts/truetype/liberation2/LiberationSerif-Regular.ttf",
]

def get_font(size, bold=True):
    paths = FONT_CANDIDATES if bold else FONT_REGULAR
    for p in paths:
        if os.path.exists(p):
            return ImageFont.truetype(p, size=size)
    return ImageFont.load_default()

def fit_font(text, max_width, start_size=60, min_size=18, bold=True):
    size = start_size
    while size >= min_size:
        f = get_font(size, bold)
        box = f.getbbox(text)
        if box[2] - box[0] <= max_width:
            return f
        size -= 2
    return get_font(min_size, bold)

def wrap_text(draw, text, font, max_width):
    words = text.split()
    lines = []
    current = ""
    for word in words:
        trial = word if not current else current + " " + word
        if draw.textbbox((0, 0), trial, font=font)[2] <= max_width:
            current = trial
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines

def rounded_rectangle(draw, xy, radius, fill, outline=None, width=1):
    draw.rounded_rectangle(xy, radius=radius, fill=fill, outline=outline, width=width)

def fit_crop(img, size):
    img = img.convert("RGB")
    target_w, target_h = size
    ratio = max(target_w / img.width, target_h / img.height)
    nw, nh = int(img.width * ratio), int(img.height * ratio)
    img = img.resize((nw, nh), Image.Resampling.LANCZOS)
    left = (nw - target_w) // 2
    top = (nh - target_h) // 2
    return img.crop((left, top, left + target_w, top + target_h))

def fit_contain(img, size, bg=(255,255,255)):
    img = img.convert("RGBA")
    target_w, target_h = size
    ratio = min(target_w / img.width, target_h / img.height)
    nw, nh = max(1, int(img.width * ratio)), max(1, int(img.height * ratio))
    img = img.resize((nw, nh), Image.Resampling.LANCZOS)
    canvas = Image.new("RGBA", size, bg + (255,))
    canvas.alpha_composite(img, ((target_w - nw)//2, (target_h - nh)//2))
    return canvas

def draw_image_collage(canvas, images, x, y, w, h):
    if not images:
        draw = ImageDraw.Draw(canvas)
        rounded_rectangle(draw, (x,y,x+w,y+h), 35, "#F2F2F2", NAVY, 6)
        draw.text((x+30,y+h//2-20), "UPLOAD PROJECT IMAGES", font=get_font(34), fill=NAVY)
        return

    n = len(images)
    gap = 12
    if n == 1:
        tiles = [(x,y,w,h)]
    elif n == 2:
        tiles = [(x,y,w//2-gap,h), (x+w//2+gap//2,y,w//2-gap//2,h)]
    elif n == 3:
        left_w = int(w*0.47)
        right_w = w-left_w-gap
        tiles = [(x,y,left_w,h), (x+left_w+gap,y,right_w,(h-gap)//2),
                 (x+left_w+gap,y+(h+gap)//2,right_w,(h-gap)//2)]
    else:
        left_w = int(w*0.46)
        right_x = x+left_w+gap
        right_w = w-left_w-gap
        tiles = [(x,y,left_w,h),
                 (right_x,y,right_w,(h-gap)//2),
                 (right_x,y+(h+gap)//2,right_w,(h-gap)//2)]

    draw = ImageDraw.Draw(canvas)
    for i, box in enumerate(tiles):
        img = images[i % len(images)]
        tw, th = box[2]-box[0], box[3]-box[1]
        tile = fit_crop(img, (tw,th))
        mask = Image.new("L", (tw,th), 0)
        ImageDraw.Draw(mask).rounded_rectangle((0,0,tw-1,th-1), radius=28, fill=255)
        canvas.paste(tile, (box[0],box[1]), mask)

def build_cover(project_type, project_name, reference, date_text,
                client_name, client_logo, funding_name, funding_logo,
                project_images, companies):
    W, H = 2048, 2816
    canvas = Image.new("RGB", (W,H), "white")
    draw = ImageDraw.Draw(canvas)

    # Background
    for yy in range(H):
        # subtle white-to-blue geometric feel
        t = yy / H
        if yy < 1200:
            c = int(255 - 14*t)
            draw.line((0,yy,W,yy), fill=(c,c+1,255))
    # geometric background polygons
    draw.polygon([(1550,0),(W,0),(W,900),(1760,690)], fill="#EDF4FA")
    draw.polygon([(1780,0),(W,0),(W,760)], fill="#E7F0F8")
    draw.polygon([(0,1180),(350,880),(650,1150),(300,1500)], fill="#F4F8FB")

    # Logos
    if client_logo:
        logo = fit_contain(client_logo, (230,180))
        canvas.paste(logo, (100,120), logo)
    if funding_logo:
        logo = fit_contain(funding_logo, (300,190))
        canvas.paste(logo, (W-400,105), logo)

    # Main heading
    heading = project_type.upper()
    hf = fit_font(76, 880, 76, 44, True)
    hb = draw.textbbox((0,0), heading, font=hf)
    hw = hb[2]-hb[0]
    hx = (W-hw)//2
    hy = 180
    # shadow / orange accent
    draw.polygon([(hx-80,hy+20),(hx+hw+80,hy+20),(hx+hw+105,hy+130),(hx-55,hy+130)], fill="#F6D78E")
    draw.polygon([(hx-70,hy),(hx+hw+70,hy),(hx+hw+95,hy+100),(hx-45,hy+100)], fill="white", outline=ORANGE)
    draw.text((hx,hy+13), heading, font=hf, fill=NAVY)

    # Project information box
    bx, by, bw, bh = 30, 490, W-60, 430
    rounded_rectangle(draw, (bx,by,bx+bw,by+bh), 58, LIGHT_GREEN, NAVY, 8)

    pf = fit_font(project_name, bw-90, 54, 28, True)
    lines = wrap_text(draw, project_name, pf, bw-90)
    line_h = pf.getbbox("Ag")[3] - pf.getbbox("Ag")[1] + 8
    ty = by+55
    # Alternate a few words in red for the same visual character.
    for line in lines[:6]:
        draw.text((bx+45,ty), line, font=pf, fill=NAVY)
        ty += line_h

    draw.line((bx+50, by+bh-118, bx+bw-50, by+bh-118), fill="#222222", width=4)
    rf = fit_font(f"Reference No: {reference}", 1200, 46, 28, True)
    rb = draw.textbbox((0,0), f"Reference No: {reference}", font=rf)
    draw.text(((W-(rb[2]-rb[0]))//2, by+bh-105), f"Reference No: {reference}", font=rf, fill="#41622D")

    # Date capsule
    df = fit_font(date_text, 360, 42, 28, True)
    db = draw.textbbox((0,0), date_text, font=df)
    dw = db[2]-db[0]
    rounded_rectangle(draw, ((W-dw)//2-55, 870, (W+dw)//2+55, 965), 15, NAVY, "white", 3)
    draw.text(((W-dw)//2, 885), date_text, font=df, fill="white")

    # Project images
    draw.rectangle((0, 1190, W, 1225), fill="#2C7EC7")
    draw_image_collage(canvas, project_images, 35, 1235, W-70, 930)

    # Submitted by panel
    sy = 2235
    rounded_rectangle(draw, (15,sy,W-15,H-25), 38, "white", NAVY, 6)
    draw.polygon([(W//2-270,sy+55),(W//2+270,sy+55),(W//2+220,sy+115),(W//2-220,sy+115)], fill=NAVY)
    sf = get_font(46, True)
    s = "SUBMITTED BY"
    sb = draw.textbbox((0,0),s,font=sf)
    draw.text(((W-(sb[2]-sb[0]))//2, sy+64), s, font=sf, fill="white")

    n = max(1, min(len(companies), 5))
    col_w = (W-90)//n
    for i, company in enumerate(companies[:5]):
        cx = 45+i*col_w
        if i:
            draw.line((cx, sy+155, cx, H-70), fill="#A9A9A9", width=2)
        logo = company.get("logo")
        if logo:
            l = fit_contain(logo, (col_w-60, 190))
            canvas.paste(l, (cx+30, sy+165), l)
        name = company.get("name","").strip()
        role = company.get("role","")
        nf = fit_font(name, col_w-55, 32, 18, True)
        nlines = wrap_text(draw, name, nf, col_w-55)
        ny = sy+365
        for line in nlines[:2]:
            nb = draw.textbbox((0,0), line, font=nf)
            draw.text((cx+(col_w-(nb[2]-nb[0]))//2, ny), line, font=nf, fill=NAVY)
            ny += 38
        rf = get_font(30, True)
        rs = f"[ {role.upper()} ]"
        rb = draw.textbbox((0,0),rs,font=rf)
        draw.text((cx+(col_w-(rb[2]-rb[0]))//2, ny+8), rs, font=rf, fill=GREEN)

    return canvas

def image_from_upload(upload):
    if upload is None:
        return None
    return Image.open(upload).convert("RGBA")

def make_docx(img):
    out = io.BytesIO()
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
        img.save(tmp.name, format="PNG", dpi=(300,300))
        tmp_path = tmp.name
    try:
        doc = Document()
        sec = doc.sections[0]
        sec.page_width = Inches(8.5)
        sec.page_height = Inches(11.69)
        sec.top_margin = Inches(0)
        sec.bottom_margin = Inches(0)
        sec.left_margin = Inches(0)
        sec.right_margin = Inches(0)
        p = doc.paragraphs[0]
        p.paragraph_format.space_before = 0
        p.paragraph_format.space_after = 0
        p.add_run().add_picture(tmp_path, width=Inches(8.5))
        doc.save(out)
    finally:
        os.unlink(tmp_path)
    out.seek(0)
    return out

st.title("📄 Proposal Cover Page Generator")
st.caption("Generate a reusable, project-specific cover page from one form and export it as an editable Word document containing the generated artwork.")

with st.sidebar:
    st.header("1. Project Details")
    project_type = st.selectbox("Document Type", ["Expression of Interest", "Request for Proposal", "Technical Proposal", "Financial Proposal", "Proposal"])
    project_name = st.text_area("Project Name", height=120, value="Consulting Services for Maritime Diagnosis and Modelling, Study and Alternatives, Executive Project and Environmental and Social Studies, for the Construction of the Coastal Protection Infrastructure, Rehabilitation and Requalification of the Marginal of the Municipality of the City of Vilankulo, Inhambane Province")
    reference = st.text_input("Reference No.", "MZ-MEF-DNT-556716-CS-QCBS")
    date_text = st.text_input("Date", "Aug, 2026")
    client_name = st.text_input("Client / Government", "")
    funding_name = st.text_input("Funding / Bank", "")

    st.header("2. Logos")
    client_logo_file = st.file_uploader("Client / Government Logo", type=["png","jpg","jpeg"], key="client")
    funding_logo_file = st.file_uploader("Funding / Bank Logo", type=["png","jpg","jpeg"], key="funding")

    st.header("3. Project Images")
    image_files = st.file_uploader("Upload 1–4 Project Images", type=["png","jpg","jpeg"], accept_multiple_files=True)
    image_files = image_files[:4]

    st.header("4. Submitted By")
    company_count = st.number_input("Number of Organizations", min_value=1, max_value=5, value=3, step=1)
    companies = []
    roles = ["LEAD", "JV", "SUB-CON", "ASSOCIATE", "PARTNER"]
    for i in range(int(company_count)):
        st.markdown(f"**Organization {i+1}**")
        name = st.text_input("Company Name", key=f"name_{i}")
        role = st.selectbox("Role", roles, key=f"role_{i}")
        logo_file = st.file_uploader("Company Logo", type=["png","jpg","jpeg"], key=f"logo_{i}")
        companies.append({"name": name, "role": role, "logo": image_from_upload(logo_file) if logo_file else None})

    if st.button("Load Example Cover", use_container_width=True):
        st.session_state["example"] = True
        st.rerun()

    generate = st.button("🚀 Generate Cover", type="primary", use_container_width=True)

if st.session_state.get("example"):
    project_name = "Consulting Services for Maritime Diagnosis and Modelling, Study and Alternatives, Executive Project and Environmental and Social Studies, for the Construction of the Coastal Protection Infrastructure, Rehabilitation and Requalification of the Marginal of the Municipality of the City of Vilankulo, Inhambane Province"
    reference = "MZ-MEF-DNT-556716-CS-QCBS"
    date_text = "Aug, 2026"
    asset_dir = Path(__file__).parent / "sample_assets"
    client_logo_file = asset_dir / "world_bank.png"
    funding_logo_file = asset_dir / "mozambique.png"
    image_files = [asset_dir / x for x in ["project1.jpg","project2.jpg","project3.jpg"]]
    companies = [
        {"name":"CONSULTING ENGINEERS GROUP LIMITED","role":"LEAD","logo":Image.open(asset_dir/"ceg.png").convert("RGBA")},
        {"name":"PROES CONSULTORES S.A., SPAIN","role":"JV","logo":Image.open(asset_dir/"proes.png").convert("RGBA")},
        {"name":"COSE LDA","role":"SUB-CON","logo":Image.open(asset_dir/"cose.png").convert("RGBA")},
    ]

client_logo = image_from_upload(client_logo_file) if client_logo_file else None
funding_logo = image_from_upload(funding_logo_file) if funding_logo_file else None
project_images = [image_from_upload(x) if not isinstance(x, Path) else Image.open(x).convert("RGBA") for x in image_files]

if not companies:
    companies = [{"name":"","role":"LEAD","logo":None}]

if generate or st.session_state.get("example"):
    cover = build_cover(
        project_type, project_name, reference, date_text,
        client_name, client_logo, funding_name, funding_logo,
        project_images, companies
    )
else:
    cover = build_cover(
        project_type, project_name, reference, date_text,
        client_name, client_logo, funding_name, funding_logo,
        project_images, companies
    )

left, right = st.columns([1.15, 0.85])
with left:
    st.subheader("Live Preview")
    st.image(cover, use_container_width=True)
with right:
    st.subheader("Output")
    docx_bytes = make_docx(cover)
    st.download_button(
        "⬇️ Download Word Cover Page",
        data=docx_bytes.getvalue(),
        file_name="Proposal_Cover_Page.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )
    png = io.BytesIO()
    cover.save(png, format="PNG", dpi=(300,300))
    st.download_button(
        "🖼️ Download PNG Cover",
        data=png.getvalue(),
        file_name="Proposal_Cover_Page.png",
        mime="image/png",
        use_container_width=True,
    )
    st.info("The Word file contains the generated cover as a high-resolution page image so the layout stays fixed across computers. A future version can place individual text boxes and images as native editable Word objects.")
