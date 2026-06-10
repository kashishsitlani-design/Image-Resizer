"""
Image Resizer Pro — Merged
Requirements: streamlit, Pillow, rembg, onnxruntime, pandas, requests, openpyxl
"""

import io
import os
import shutil
import zipfile
import requests
import pandas as pd
import streamlit as st
from PIL import Image, ImageChops

# rembg optional
try:
    from rembg import remove as rembg_remove, new_session
    REMBG_AVAILABLE = True
except Exception:
    REMBG_AVAILABLE = False

st.set_page_config(page_title="Image Resizer Pro", page_icon="🖼️", layout="centered")
st.title("🖼️ Image Resizer Pro")
st.caption("Bulk resize · Remove backgrounds · Marketplace presets · Excel links · Download as ZIP")

# ─────────────────────────────────────────────────────────────────
# MARKETPLACE PRESETS
# ─────────────────────────────────────────────────────────────────
PRESETS = {
    "— Custom —":       (800,  800,  None, None, None, None, "—"),
    "Allegro PL":       (2200, 2200, 500,  500,  2560, 2560, "1:1"),
    "Allegro One PL":   (2200, 2200, 1000, 1000, 2560, 2560, "1:1"),
    "Best Buy US":      (2000, 2000, 2000, 2000, None, None, "1:1"),
    "Best Buy CA":      (2000, 2000, 2000, 2000, None, None, "1:1"),
    "Bol NL":           (2400, 2400, 500,  500,  6000, 6000, "1:1"),
    "Bol BE":           (2400, 2400, 500,  500,  6000, 6000, "1:1"),
    "eBay US":          (1600, 1600, 500,  500,  9000, 9000, "1:1"),
    "eBay DE":          (1600, 1600, 500,  500,  9000, 9000, "1:1"),
    "eBay UK":          (1600, 1600, 500,  500,  9000, 9000, "1:1"),
    "Kohl's US":        (1000, 1000, 1000, 1000, None, None, "1:1"),
    "Lowes US":         (1000, 1000, 1000, 1000, None, None, "1:1"),
    "Macy's US":        (1000, 1000, 1000, 1000, None, None, "1:1"),
    "MediaMarkt DE":    (1200, 1200, 1000, 1000, None, None, "1:1"),
    "Mercado Libre US": (1600, 1600, 500,  500,  2500, 2500, "1:1"),
    "Nordstrom US":     (2600, 4000, 1300, 2000, None, None, "2:3"),
    "Octopia FR":       (1000, 1000, 500,  500,  2500, 2500, "1:1"),
    "OTTO DE":          (960,  480,  None, None, None, None, "2:1"),
    "Target US":        (2400, 2400, 1200, 1200, 5000, 5000, "1:1"),
    "Tesco UK":         (2400, 2400, 1000, 1000, None, None, "1:1"),
    "TikTok US":        (1000, 1000, 600,  600,  3000, 3000, "1:1"),
    "TikTok UK":        (1000, 1000, 600,  600,  3000, 3000, "1:1"),
    "Walmart US":       (2200, 2200, 1500, 1500, 5000, 5000, "1:1"),
    "Walmart CA":       (2200, 2200, 1500, 1500, 5000, 5000, "1:1"),
    "Zalando DE":       (2000, 2000, 800,  1200, 5000, 5000, "2:3"),
    "Shopify":          (2048, 2048, 800,  800,  4472, 4472, "1:1"),
}

EXT_MAP = {"PNG": "png", "JPEG": "jpg", "WEBP": "webp", "GIF": "gif", "BMP": "bmp", "TIFF": "tif"}

# ─────────────────────────────────────────────────────────────────
# SECTION 0 — UPLOAD (direct images OR Excel with links)
# ─────────────────────────────────────────────────────────────────
st.divider()
st.subheader("📁 Section 0 — Upload Images")

uploaded = st.file_uploader(
    "Upload images directly",
    type=["png", "jpg", "jpeg", "webp", "gif", "bmp", "tiff"],
    accept_multiple_files=True,
)

st.divider()
st.subheader("📊 Or Upload Excel File with Image Links")
excel_file = st.file_uploader("Excel file (.xlsx) with image URLs", type=["xlsx"])

col1, col2 = st.columns(2)
with col1:
    file_col1 = st.text_input("FileName column 1", value="FileName1")
    file_col2 = st.text_input("FileName column 2 (optional)", value="")
with col2:
    link_col1 = st.text_input("ImageLink column 1", value="ImageLink1")
    link_col2 = st.text_input("ImageLink column 2 (optional)", value="")

has_input = uploaded or excel_file
if not has_input:
    st.info("⬆️ Upload images or an Excel file to get started.")
    st.stop()

if uploaded:
    st.success(f"✅ {len(uploaded)} image{'s' if len(uploaded) > 1 else ''} ready")

# ─────────────────────────────────────────────────────────────────
# SECTION 1 — BACKGROUND
# ─────────────────────────────────────────────────────────────────
st.divider()
st.subheader("🪄 Section 1 — Background")

bg_choice = st.radio(
    "Background options",
    [
        "Keep original background",
        "Remove background — keep transparent",
        "Remove background — fill with white",
    ],
    index=0,
)

remove_bg  = bg_choice in ("Remove background — keep transparent",
                            "Remove background — fill with white")
fill_white = bg_choice == "Remove background — fill with white"

if remove_bg:
    if not REMBG_AVAILABLE:
        st.error("❌ rembg not installed. Add `rembg` and `onnxruntime` to requirements.txt")
    else:
        st.warning("⏳ Background removal uses AI. First run takes 20–40 seconds. Do not click Stop.")
    if fill_white:
        st.caption("✅ After removal, transparent areas filled with white.")
    else:
        st.caption("✅ After removal, transparency kept. Use PNG or WebP to preserve it.")

crop_white_opt = False
auto_fill      = False
fill_percent   = 90
margin_cm      = 0.0

# Padding sub-options — only show when relevant
if not remove_bg:
    crop_white_opt = st.checkbox("Crop white border before resizing", value=False,
                                  help="Removes any existing white padding around the product before resizing.")
    auto_fill = st.checkbox("Auto-fill canvas (fit product to % of canvas)", value=False)
    if auto_fill:
        fill_percent = st.slider("Product fill target (%)", 85, 95, 90, 1)
    margin_cm = st.number_input("Margin (cm)", min_value=0.0, max_value=10.0, value=0.0, step=0.1)

# ─────────────────────────────────────────────────────────────────
# SECTION 2 — MARKETPLACE DIMENSIONS
# ─────────────────────────────────────────────────────────────────
st.divider()
st.subheader("📐 Section 2 — Dimensions")

marketplace = st.selectbox("Choose marketplace (or Custom)", list(PRESETS.keys()))
rec_w, rec_h, min_w, min_h, max_w, max_h, ratio = PRESETS[marketplace]

if marketplace != "— Custom —":
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Recommended", f"{rec_w}×{rec_h}")
    c2.metric("Minimum",     f"{min_w}×{min_h}" if min_w else "—")
    c3.metric("Maximum",     f"{max_w}×{max_h}" if max_w else "—")
    c4.metric("Ratio",       ratio)

with st.expander("📋 View all marketplace dimension requirements"):
    st.markdown("""
| Marketplace | Recommended | Minimum | Maximum | Ratio |
|---|---|---|---|---|
| Allegro PL | 2200×2200 | 500×500 | 2560×2560 | 1:1 |
| Allegro One PL | 2200×2200 | 1000×1000 | 2560×2560 | 1:1 |
| Best Buy US/CA | 2000×2000 | 2000×2000 | — | 1:1 |
| Bol NL/BE | 2400×2400 | 500×500 | 6000×6000 | 1:1 |
| eBay US/DE/UK | 1600×1600 | 500×500 | 9000×9000 | 1:1 |
| Kohl's US | 1000×1000 | 1000×1000 | — | 1:1 |
| Lowes US | 1000×1000 | 1000×1000 | — | 1:1 |
| Macy's US | 1000×1000 | 1000×1000 | — | 1:1 |
| MediaMarkt DE | 1200×1200 | 1000×1000 | — | 1:1 |
| Mercado Libre US | 1600×1600 | 500×500 | 2500×2500 | 1:1 |
| Nordstrom US | 2600×4000 | 1300×2000 | — | 2:3 |
| Octopia FR | 1000×1000 | 500×500 | 2500×2500 | 1:1 |
| OTTO DE | 960×480 | — | — | 2:1 |
| Target US | 2400×2400 | 1200×1200 | 5000×5000 | 1:1 |
| Tesco UK | 2400×2400 | 1000×1000 | — | 1:1 |
| TikTok US/UK | 1000×1000 | 600×600 | 3000×3000 | 1:1 |
| Walmart US/CA | 2200×2200 | 1500×1500 | 5000×5000 | 1:1 |
| Zalando DE | 2000×2000 | 800×1200 | 5000×5000 | 2:3 |
| Shopify | 2048×2048 | 800×800 | 4472×4472 | 1:1 |
""")

resize_mode = st.radio(
    "Resize mode",
    ["By Width only", "By Height only", "Exact W × H"],
    horizontal=True,
    help=(
        "**By Width only** — height auto-calculated per each image ratio.\n\n"
        "**By Height only** — width auto-calculated per each image ratio.\n\n"
        "**Exact W × H** — image fits inside the box keeping its own ratio. "
        "Remaining space filled only if white bg option is chosen above."
    ),
)

col_w, col_h = st.columns(2)
target_w = int(rec_w)
target_h = int(rec_h)

if resize_mode in ("By Width only", "Exact W × H"):
    target_w = col_w.number_input("Width (px)", min_value=1, value=target_w, step=1)
if resize_mode in ("By Height only", "Exact W × H"):
    target_h = col_h.number_input("Height (px)", min_value=1, value=target_h, step=1)

output_dpi = st.number_input("DPI", min_value=50, max_value=1200, value=300, step=1)

# ─────────────────────────────────────────────────────────────────
# SECTION 3 — FORMAT & QUALITY
# ─────────────────────────────────────────────────────────────────
st.divider()
st.subheader("🎨 Section 3 — Output Format & Quality")

format_options = [
    "Keep original format",
    "PNG  — lossless, preserves transparency",
    "JPEG — smaller file, no transparency",
    "WebP — modern, preserves transparency",
]
default_format = (
    "JPEG — smaller file, no transparency"
    if marketplace != "— Custom —"
    else "Keep original format"
)
chosen_format = st.selectbox("Output format", format_options,
                              index=format_options.index(default_format))

FORMAT_MAP = {
    "Keep original format":                    None,
    "PNG  — lossless, preserves transparency": "PNG",
    "JPEG — smaller file, no transparency":    "JPEG",
    "WebP — modern, preserves transparency":   "WEBP",
}
out_fmt = FORMAT_MAP[chosen_format]

quality = 92
if out_fmt in ("JPEG", None):
    quality = st.slider("JPEG quality", 10, 100, 92, 1,
                        help="Higher = sharper, larger file. 85–95 recommended.")
elif out_fmt == "WEBP":
    quality = st.slider("WebP quality", 10, 100, 90, 1)

if out_fmt == "JPEG" and remove_bg and not fill_white:
    st.warning("⚠️ JPEG cannot store transparency. Switch to PNG/WebP, or enable white background fill above.")

# ─────────────────────────────────────────────────────────────────
# SECTION 4 — EXTRAS
# ─────────────────────────────────────────────────────────────────
st.divider()
st.subheader("⚙️ Section 4 — Extra Options")

amazon_check     = st.checkbox("Amazon compliance checker", value=False,
                                help="Checks image size, square ratio, product fill %, and white background.")
preview_enabled  = st.checkbox("Preview first 5 processed images", value=True)

# ─────────────────────────────────────────────────────────────────
# HELPER FUNCTIONS
# ─────────────────────────────────────────────────────────────────
def cm_to_px(cm, dpi):
    return int((cm / 2.54) * dpi)

def get_new_size(orig_w, orig_h):
    if resize_mode == "By Width only":
        return target_w, max(1, round(orig_h * target_w / orig_w))
    elif resize_mode == "By Height only":
        return max(1, round(orig_w * target_h / orig_h)), target_h
    else:
        scale = min(target_w / orig_w, target_h / orig_h)
        return max(1, round(orig_w * scale)), max(1, round(orig_h * scale))

def do_crop_white(img, tolerance=10):
    rgb = img.convert("RGB")
    bg  = Image.new("RGB", rgb.size, (255, 255, 255))
    diff = ImageChops.difference(rgb, bg).convert("L")
    mask = diff.point(lambda p: 255 if p > tolerance else 0)
    bbox = mask.getbbox()
    return img.crop(bbox) if bbox else img

def amazon_compliance(img):
    w, h = img.size
    rgb  = img.convert("RGB")
    bg   = Image.new("RGB", rgb.size, (255, 255, 255))
    diff = ImageChops.difference(rgb, bg).convert("L")
    mask = diff.point(lambda p: 255 if p > 12 else 0)
    bbox = mask.getbbox()
    fill = 0
    if bbox:
        fill = round(max((bbox[2]-bbox[0])/w, (bbox[3]-bbox[1])/h) * 100, 1)
    issues = []
    if w != h:             issues.append("Not square")
    if w < 1000:           issues.append("Below 1000px")
    if fill < 85:          issues.append(f"Product fill {fill}% < 85%")
    status = "✅ Pass" if not issues else "⚠️ Review"
    return status, fill, "; ".join(issues) if issues else "Looks good"

def process_image_bytes(raw_bytes, filename):
    orig_pil = Image.open(io.BytesIO(raw_bytes))
    orig_fmt = orig_pil.format or "PNG"

    # Step 1 — background removal
    if remove_bg and REMBG_AVAILABLE:
        session = new_session("u2netp")
        raw_bytes = rembg_remove(raw_bytes, session=session)
        img = Image.open(io.BytesIO(raw_bytes)).convert("RGBA")
        # Crop to content bounding box after removal (tight crop, no padding)
        bbox = img.getbbox()
        if bbox:
            img = img.crop(bbox)
    else:
        # Keep image exactly as-is — do NOT convert mode here
        img = orig_pil.copy()

    has_alpha = img.mode in ("RGBA", "LA", "PA") or (
        img.mode == "P" and "transparency" in img.info
    )

    # Step 2 — crop white border if requested (only when not removing bg)
    if crop_white_opt and not remove_bg:
        img = do_crop_white(img)

    # Step 3 — fill white background ONLY if user chose it
    if fill_white and has_alpha:
        if img.mode == "P":
            img = img.convert("RGBA")
        bg = Image.new("RGB", img.size, (255, 255, 255))
        bg.paste(img, mask=img.split()[-1] if img.mode in ("RGBA", "LA") else None)
        img = bg
        has_alpha = False

    # Step 4 — resize
    new_w, new_h = get_new_size(img.width, img.height)
    img = img.resize((new_w, new_h), Image.LANCZOS)

    # Step 5 — if Exact W×H, place on canvas preserving original background
    save_fmt = out_fmt or orig_fmt
    if resize_mode == "Exact W × H" and (new_w != target_w or new_h != target_h):
        offset_x = (target_w - new_w) // 2
        offset_y = (target_h - new_h) // 2
        if fill_white:
            # User explicitly chose white background
            canvas = Image.new("RGB", (target_w, target_h), (255, 255, 255))
            if img.mode == "RGBA":
                canvas.paste(img, (offset_x, offset_y), mask=img.split()[3])
            else:
                canvas.paste(img.convert("RGB"), (offset_x, offset_y))
        else:
            # Keep original background — transparent canvas so no colour is added
            canvas = Image.new("RGBA", (target_w, target_h), (0, 0, 0, 0))
            if img.mode != "RGBA":
                img = img.convert("RGBA")
            canvas.paste(img, (offset_x, offset_y))
        img = canvas

    # Step 6 — auto fill canvas scaling
    if auto_fill and not remove_bg:
        margin_px = cm_to_px(margin_cm, output_dpi)
        canvas_w  = target_w - 2 * margin_px
        canvas_h  = target_h - 2 * margin_px
        scale     = min(canvas_w / img.width, canvas_h / img.height) * (fill_percent / 100)
        new_fw    = max(1, round(img.width  * scale))
        new_fh    = max(1, round(img.height * scale))
        img       = img.resize((new_fw, new_fh), Image.LANCZOS)
        if fill_white:
            final = Image.new("RGB", (target_w, target_h), (255, 255, 255))
        else:
            # Transparent canvas — no bg added
            final = Image.new("RGBA", (target_w, target_h), (0, 0, 0, 0))
        ox = (target_w - new_fw) // 2
        oy = (target_h - new_fh) // 2
        if img.mode == "RGBA":
            final.paste(img, (ox, oy), mask=img.split()[3])
        else:
            if img.mode != "RGBA":
                img = img.convert("RGBA")
            final.paste(img, (ox, oy))
        img = final

    # Step 7 — convert for save format
    # JPEG has no alpha channel — drop it without adding any background colour
    if save_fmt == "JPEG":
        if img.mode == "RGBA" and not fill_white:
            # Drop alpha, keep pixel colours exactly — no white fill
            r, g, b, a = img.split()
            img = Image.merge("RGB", (r, g, b))
        else:
            img = img.convert("RGB")
    elif save_fmt in ("PNG", "WEBP"):
        if img.mode not in ("RGBA", "RGB", "L", "LA"):
            img = img.convert("RGBA" if has_alpha else "RGB")

    # Step 8 — save
    buf = io.BytesIO()
    save_kw = {}
    if save_fmt == "JPEG":
        save_kw = {"quality": quality, "optimize": True, "progressive": True,
                   "dpi": (output_dpi, output_dpi)}
    elif save_fmt == "WEBP":
        save_kw = {"quality": quality}
    elif save_fmt == "PNG":
        save_kw = {"dpi": (output_dpi, output_dpi)}
    img.save(buf, format=save_fmt, **save_kw)
    buf.seek(0)

    # Output filename
    stem     = filename.rsplit(".", 1)[0] if "." in filename else filename
    new_ext  = EXT_MAP.get(save_fmt, save_fmt.lower())
    orig_ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    out_name = f"{stem}.{new_ext}" if new_ext != orig_ext else filename

    return buf.read(), out_name, img.width, img.height, img

# ─────────────────────────────────────────────────────────────────
# PROCESS BUTTON
# ─────────────────────────────────────────────────────────────────
st.divider()

if st.button("🚀 Process & Resize All Images", type="primary", use_container_width=True):

    results          = []   # (name, data, w, h)
    errors           = []
    compliance_rows  = []
    preview_imgs     = []

    # Collect all sources
    sources = []  # list of (raw_bytes, filename)

    # Direct uploads
    if uploaded:
        for f in uploaded:
            f.seek(0)
            sources.append((f.read(), f.name))

    # Excel links
    if excel_file:
        try:
            df = pd.read_excel(excel_file)
            col_pairs = []
            if file_col1 and link_col1:
                col_pairs.append((file_col1, link_col1))
            if file_col2 and link_col2:
                col_pairs.append((file_col2, link_col2))

            for _, row in df.iterrows():
                for fc, lc in col_pairs:
                    fname = row.get(fc)
                    link  = row.get(lc)
                    if pd.notna(fname) and pd.notna(link):
                        try:
                            resp = requests.get(str(link), timeout=15)
                            resp.raise_for_status()
                            sources.append((resp.content, str(fname)))
                        except Exception as e:
                            errors.append(f"**{fname}** (Excel link) — {e}")
        except Exception as e:
            st.error(f"❌ Could not read Excel: {e}")

    total = len(sources)
    if total == 0:
        st.warning("No images to process.")
        st.stop()

    bar = st.progress(0, text="Starting…")

    for i, (raw_bytes, filename) in enumerate(sources):
        try:
            data, name, w, h, pil_img = process_image_bytes(raw_bytes, filename)
            results.append((name, data, w, h))

            if amazon_check:
                status, fill, notes = amazon_compliance(pil_img)
                compliance_rows.append({
                    "File": name, "Width": w, "Height": h,
                    "Fill %": fill, "Status": status, "Notes": notes
                })

            if len(preview_imgs) < 5:
                preview_imgs.append((filename, raw_bytes, name, data))

        except Exception as e:
            errors.append(f"**{filename}** — {e}")

        bar.progress((i + 1) / total, text=f"Processing {i+1}/{total}: {filename}")

    bar.empty()

    # Errors
    if errors:
        with st.expander(f"⚠️ {len(errors)} error(s)"):
            for err in errors:
                st.warning(err)

    if results:
        st.success(f"✅ {len(results)} image{'s' if len(results) > 1 else ''} processed!")

        # Preview
        if preview_enabled and preview_imgs:
            st.markdown("**Preview (first 5 images):**")
            for orig_name, orig_bytes, out_name, out_data in preview_imgs:
                pc1, pc2 = st.columns(2)
                pc1.image(Image.open(io.BytesIO(orig_bytes)),
                          caption=f"Original: {orig_name}", use_container_width=True)
                pc2.image(Image.open(io.BytesIO(out_data)),
                          caption=f"Processed: {out_name}", use_container_width=True)

        # Amazon compliance table
        if amazon_check and compliance_rows:
            st.subheader("Amazon Compliance Summary")
            comp_df = pd.DataFrame(compliance_rows)
            st.dataframe(comp_df, use_container_width=True)
            passes  = len(comp_df[comp_df["Status"] == "✅ Pass"])
            reviews = len(comp_df[comp_df["Status"] == "⚠️ Review"])
            st.info(f"✅ Pass: {passes}  |  ⚠️ Review: {reviews}")

        # ZIP download
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_STORED) as zf:
            for name, data, _, _ in results:
                zf.writestr(name, data)
            if amazon_check and compliance_rows:
                csv = pd.DataFrame(compliance_rows).to_csv(index=False).encode()
                zf.writestr("amazon_compliance_report.csv", csv)
        zip_buf.seek(0)

        st.download_button(
            label=f"⬇️ Download all {len(results)} images as ZIP",
            data=zip_buf,
            file_name="processed_images.zip",
            mime="application/zip",
            use_container_width=True,
            type="primary",
        )

st.divider()
st.caption("Images are processed only in your session — nothing is stored or sent anywhere.")
st.markdown(
    "<div style='text-align:center;color:#888;font-size:0.8rem;padding-top:8px;'>"
    "Created by <strong>Kashish Sitlani</strong></div>",
    unsafe_allow_html=True,
)
