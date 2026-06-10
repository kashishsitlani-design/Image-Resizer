"""
Image Resizer Pro — Built from scratch
Requirements: streamlit, Pillow, rembg
"""

import io
import zipfile
import streamlit as st
from PIL import Image

st.set_page_config(page_title="Image Resizer Pro", page_icon="🖼️", layout="centered")

st.title("🖼️ Image Resizer Pro")
st.caption("Bulk resize · Remove backgrounds · Marketplace presets · Download as ZIP")

# ─────────────────────────────────────────────────────────────────
# MARKETPLACE PRESETS
# (rec_w, rec_h, min_w, min_h, max_w, max_h, ratio)
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

EXT_MAP = {
    "PNG": "png", "JPEG": "jpg", "WEBP": "webp",
    "GIF": "gif", "BMP": "bmp", "TIFF": "tif"
}

# ─────────────────────────────────────────────────────────────────
# SECTION 0 — UPLOAD
# ─────────────────────────────────────────────────────────────────
st.divider()
st.subheader("📁 Upload Images")

uploaded = st.file_uploader(
    "Upload one or more images",
    type=["png", "jpg", "jpeg", "webp", "gif", "bmp", "tiff"],
    accept_multiple_files=True,
)

if not uploaded:
    st.info("⬆️ Upload images above to get started.")
    st.stop()

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

remove_bg = bg_choice in (
    "Remove background — keep transparent",
    "Remove background — fill with white",
)
fill_white = bg_choice == "Remove background — fill with white"

if remove_bg:
    st.info(
        "🤖 Background removal uses AI (`rembg`). "
        "The first run downloads the model (~170 MB) — this is automatic."
    )
    if fill_white:
        st.caption("✅ After removal, transparent areas will be filled with white.")
    else:
        st.caption("✅ After removal, transparency is kept. Use PNG or WebP to preserve it.")

# ─────────────────────────────────────────────────────────────────
# SECTION 2 — MARKETPLACE DIMENSIONS
# ─────────────────────────────────────────────────────────────────
st.divider()
st.subheader("📐 Section 2 — Dimensions")

marketplace = st.selectbox("Choose marketplace (or Custom)", list(PRESETS.keys()))
rec_w, rec_h, min_w, min_h, max_w, max_h, ratio = PRESETS[marketplace]

# Show info cards for chosen marketplace
if marketplace != "— Custom —":
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Recommended", f"{rec_w}×{rec_h}")
    c2.metric("Minimum",     f"{min_w}×{min_h}" if min_w else "—")
    c3.metric("Maximum",     f"{max_w}×{max_h}" if max_w else "—")
    c4.metric("Ratio",       ratio)

# Full reference table
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

# Resize mode
resize_mode = st.radio(
    "Resize mode",
    ["By Width only", "By Height only", "Exact W × H"],
    horizontal=True,
    help=(
        "**By Width only** — height auto-calculated per each image ratio.\n\n"
        "**By Height only** — width auto-calculated per each image ratio.\n\n"
        "**Exact W × H** — forces exact size (may stretch if ratio differs)."
    ),
)

col_w, col_h = st.columns(2)
target_w = rec_w
target_h = rec_h

if resize_mode in ("By Width only", "Exact W × H"):
    target_w = col_w.number_input("Width (px)", min_value=1, value=int(rec_w), step=1)
if resize_mode in ("By Height only", "Exact W × H"):
    target_h = col_h.number_input("Height (px)", min_value=1, value=int(rec_h), step=1)

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

chosen_format = st.selectbox(
    "Output format",
    format_options,
    index=format_options.index(default_format),
)

FORMAT_MAP = {
    "Keep original format":                    None,
    "PNG  — lossless, preserves transparency": "PNG",
    "JPEG — smaller file, no transparency":    "JPEG",
    "WebP — modern, preserves transparency":   "WEBP",
}
out_fmt = FORMAT_MAP[chosen_format]

quality = 92
if out_fmt == "JPEG" or out_fmt is None:
    quality = st.slider("JPEG quality", min_value=10, max_value=100, value=92, step=1,
                        help="Higher = sharper image, larger file. 85–95 is a good range.")
elif out_fmt == "WEBP":
    quality = st.slider("WebP quality", min_value=10, max_value=100, value=90, step=1)

if out_fmt == "JPEG" and remove_bg and not fill_white:
    st.warning("⚠️ JPEG cannot store transparency. Switch to PNG/WebP, or enable white background fill above.")

# ─────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────
def get_new_size(orig_w, orig_h):
    if resize_mode == "By Width only":
        return target_w, max(1, round(orig_h * target_w / orig_w))
    elif resize_mode == "By Height only":
        return max(1, round(orig_w * target_h / orig_h)), target_h
    else:
        return target_w, target_h


def process_one(file):
    file.seek(0)
    raw_bytes = file.read()

    # Read original format before any processing
    orig_pil  = Image.open(io.BytesIO(raw_bytes))
    orig_fmt  = orig_pil.format or "PNG"

    # Step 1 — background removal
    if remove_bg:
        from rembg import remove as rembg_remove
        raw_bytes = rembg_remove(raw_bytes)
        img = Image.open(io.BytesIO(raw_bytes)).convert("RGBA")
    else:
        img = orig_pil.copy()

    has_alpha = img.mode in ("RGBA", "LA", "PA") or (
        img.mode == "P" and "transparency" in img.info
    )

    # Step 2 — fill white background if requested
    if fill_white and has_alpha:
        if img.mode == "P":
            img = img.convert("RGBA")
        bg = Image.new("RGB", img.size, (255, 255, 255))
        bg.paste(img, mask=img.split()[-1] if img.mode in ("RGBA", "LA") else None)
        img = bg
        has_alpha = False

    # Step 3 — resize
    new_w, new_h = get_new_size(img.width, img.height)
    img = img.resize((new_w, new_h), Image.LANCZOS)

    # Step 4 — convert mode for chosen format
    save_fmt = out_fmt or orig_fmt
    if save_fmt == "JPEG":
        img = img.convert("RGB")   # drop alpha — no white fill unless user chose it
    elif save_fmt in ("PNG", "WEBP"):
        if img.mode not in ("RGBA", "RGB", "L", "LA"):
            img = img.convert("RGBA" if has_alpha else "RGB")

    # Step 5 — save to buffer
    buf = io.BytesIO()
    save_kwargs = {}
    if save_fmt == "JPEG":
        save_kwargs = {"quality": quality, "optimize": True, "progressive": True}
    elif save_fmt == "WEBP":
        save_kwargs = {"quality": quality}
    img.save(buf, format=save_fmt, **save_kwargs)
    buf.seek(0)

    # Build output filename — keep original name, only swap extension if format changed
    stem     = file.name.rsplit(".", 1)[0] if "." in file.name else file.name
    new_ext  = EXT_MAP.get(save_fmt, save_fmt.lower())
    orig_ext = file.name.rsplit(".", 1)[-1].lower() if "." in file.name else ""
    out_name = f"{stem}.{new_ext}" if new_ext != orig_ext else file.name

    return buf.read(), out_name, new_w, new_h


# ─────────────────────────────────────────────────────────────────
# PROCESS BUTTON
# ─────────────────────────────────────────────────────────────────
st.divider()

if st.button("🚀 Process & Resize All Images", type="primary", use_container_width=True):

    results = []
    errors  = []
    total   = len(uploaded)
    bar     = st.progress(0, text="Starting…")

    for i, f in enumerate(uploaded):
        try:
            data, name, w, h = process_one(f)
            results.append((name, data, w, h))
        except Exception as e:
            errors.append(f"**{f.name}** — {e}")
        bar.progress((i + 1) / total, text=f"Processing {i + 1}/{total}: {f.name}")

    bar.empty()

    # Show errors
    if errors:
        with st.expander(f"⚠️ {len(errors)} file(s) had errors"):
            for err in errors:
                st.warning(err)

    if results:
        st.success(f"✅ {len(results)} image{'s' if len(results) > 1 else ''} processed successfully!")

        # Before / after preview of first image
        st.markdown("**Preview (first image):**")
        uploaded[0].seek(0)
        orig_preview    = Image.open(uploaded[0])
        result_preview  = Image.open(io.BytesIO(results[0][1]))
        pc1, pc2 = st.columns(2)
        pc1.image(orig_preview,   caption=f"Original — {orig_preview.width}×{orig_preview.height}px",      use_container_width=True)
        pc2.image(result_preview, caption=f"Processed — {results[0][2]}×{results[0][3]}px", use_container_width=True)

        # Build ZIP
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_STORED) as zf:
            for name, data, _, _ in results:
                zf.writestr(name, data)
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
