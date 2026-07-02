"""
Creative Editing for Atlas


Requirements:
streamlit
Pillow
pandas
requests
openpyxl
rembg
onnxruntime
"""

import gc
import io
import re
import time
import zipfile
from collections import Counter
from typing import Dict, Iterable, List, Optional, Tuple
from urllib.parse import urlparse

import pandas as pd
import requests
import streamlit as st
from PIL import Image, ImageChops, ImageOps, UnidentifiedImageError

try:
    from rembg import remove as rembg_remove, new_session
    REMBG_AVAILABLE = True
except Exception:
    REMBG_AVAILABLE = False

st.set_page_config(page_title="Creative Editing for Atlas", page_icon="🎨", layout="centered")

NAVY = "#0B2E59"
LIGHT_NAVY = "#123E73"
BORDER = "#D9E2EC"
SOFT_BG = "#F5F8FC"

st.markdown(
    f"""
    <style>
    .stApp {{ background: white; }}
    h1, h2, h3 {{ color: {NAVY}; }}
    .stButton > button, .stDownloadButton > button {{
        background-color: {NAVY} !important;
        color: white !important;
        border: 1px solid {NAVY} !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
    }}
    .stButton > button:hover, .stDownloadButton > button:hover {{
        background-color: {LIGHT_NAVY} !important;
        border-color: {LIGHT_NAVY} !important;
        color: white !important;
    }}
    div[data-testid="stMetric"] {{
        background: {SOFT_BG};
        border: 1px solid {BORDER};
        border-radius: 10px;
        padding: 12px;
    }}
    div[data-testid="stFileUploader"] section {{
        border-color: {BORDER};
        background: {SOFT_BG};
    }}
    .small-note {{ font-size: 0.86rem; color: #4B5563; }}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Creative Editing for Atlas")
st.caption("Resize images as-is • No crop • No stretch • Optional background tools • Excel links • ZIP download")

# Recommended marketplace dimensions from the uploaded PXM Marketplace Image Naming Convention file.
# Format: dropdown label -> (recommended_width, recommended_height, minimum, maximum, aspect_ratio)
MARKETPLACE_PRESETS = {
    "Custom": (800, 800, "—", "—", "—"),
    "Allegro PL — 2200 x 2200": (2200, 2200, "500 x 500", "2560 x 2560", "1:1"),
    "Allegro One PL — 2200 x 2200": (2200, 2200, "1000 x 1000", "2560 x 2560", "1:1"),
    "Best Buy US — 2000 x 2000": (2000, 2000, "2000 x 2000", "—", "1:1"),
    "Best Buy CA — 2000 x 2000": (2000, 2000, "2000 x 2000", "—", "1:1"),
    "Bol NL — 2400 x 2400": (2400, 2400, "500 x 500", "6000 x 6000", "1:1"),
    "Bol BE — 2400 x 2400": (2400, 2400, "500 x 500", "6000 x 6000", "1:1"),
    "eBay US — 1600 x 1600": (1600, 1600, "500 x 500", "9000 x 9000", "1:1"),
    "eBay DE — 1600 x 1600": (1600, 1600, "500 x 500", "9000 x 9000", "1:1"),
    "eBay UK — 1600 x 1600": (1600, 1600, "500 x 500", "9000 x 9000", "1:1"),
    "Kohl's US — 1000 x 1000": (1000, 1000, "1000 x 1000", "—", "1:1"),
    "Lowes US — 1000 x 1000": (1000, 1000, "1000 x 1000", "—", "1:1"),
    "Macy's US — 1000 x 1000": (1000, 1000, "1000 x 1000", "—", "1:1"),
    "MediaMarkt DE — 1200 x 1200": (1200, 1200, "1000 x 1000", "—", "1:1"),
    "Mercado Libre US — 1600 x 1600": (1600, 1600, "500 x 500", "2500 x 2500", "1:1"),
    "Nordstrom US — 2600 x 4000": (2600, 4000, "1300 x 2000", "—", "2:3"),
    "Octopia FR — 500 x 500": (500, 500, "1000 x 1000", "2500 x 2500", "1:1"),
    "OTTO DE — 960 x 480": (960, 480, "—", "—", "2:1"),
    "Target US — 2400 x 2400": (2400, 2400, "1200 x 1200", "5000 x 5000", "1:1"),
    "Tesco UK — 2400 x 2400": (2400, 2400, "1000 x 1000", "—", "1:1"),
    "Tik Tok US — 1000 x 1000": (1000, 1000, "600 x 600", "3000 x 3000", "1:1"),
    "Tik Tok UK — 1000 x 1000": (1000, 1000, "600 x 600", "3000 x 3000", "1:1"),
    "Walmart US — 2200 x 2200": (2200, 2200, "1500 x 1500", "5000 x 5000", "1:1"),
    "Walmart CA — 2200 x 2200": (2200, 2200, "1500 x 1500", "5000 x 5000", "1:1"),
    "Zalando DE — 2000 x 2000": (2000, 2000, "800 x 1200", "5000 x 5000", "2:3"),
}

PRESETS = {name: (info[0], info[1]) for name, info in MARKETPLACE_PRESETS.items()}

EXT_MAP = {"PNG": "png", "JPEG": "jpg", "WEBP": "webp", "GIF": "gif", "BMP": "bmp", "TIFF": "tif"}
URL_KEYWORDS = ("image", "img", "link", "url", "photo", "picture", "media", "front", "back", "side", "lifestyle")
FILE_KEYWORDS = ("filename", "file name", "file", "sku", "productid", "product id", "item", "item code", "code", "name", "title")


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def is_url(value: str) -> bool:
    parsed = urlparse(value.strip()) if value else None
    return bool(parsed and parsed.scheme in ("http", "https") and parsed.netloc)


def safe_filename(name: str, fallback: str = "image") -> str:
    name = clean_text(name) or fallback
    name = re.sub(r"[\\/:*?\"<>|]+", "_", name)
    name = re.sub(r"\s+", "_", name).strip("._ ")
    return name or fallback


def unique_name(base_name: str, counter: Counter) -> str:
    base_name = safe_filename(base_name, "image")
    if "." in base_name:
        stem, ext = base_name.rsplit(".", 1)
        ext = "." + ext
    else:
        stem, ext = base_name, ""
    key = base_name.lower()
    counter[key] += 1
    return base_name if counter[key] == 1 else f"{stem}_{counter[key]}{ext}"


def detect_url_columns(df: pd.DataFrame) -> List[str]:
    columns = []
    for col in df.columns:
        col_name = str(col).strip().lower()
        name_hint = any(k in col_name for k in URL_KEYWORDS)
        sample = df[col].dropna().astype(str).head(40).tolist()
        url_count = sum(1 for v in sample if is_url(v))
        if url_count > 0 and (name_hint or url_count >= max(1, len(sample) // 3)):
            columns.append(col)
    return columns


def detect_filename_column(df: pd.DataFrame, url_columns: List[str]) -> Optional[str]:
    candidates = [c for c in df.columns if c not in url_columns]
    for col in candidates:
        col_name = str(col).strip().lower()
        if any(k == col_name or k in col_name for k in FILE_KEYWORDS):
            return col
    return candidates[0] if candidates else None


def build_excel_sources(df: pd.DataFrame) -> Tuple[List[Dict], List[str], Optional[str]]:
    url_columns = detect_url_columns(df)
    filename_column = detect_filename_column(df, url_columns)
    sources: List[Dict] = []
    seen_urls = set()
    for row_index, row in df.iterrows():
        base_name = safe_filename(row.get(filename_column), f"row_{row_index + 1}") if filename_column else f"row_{row_index + 1}"
        for col in url_columns:
            url = clean_text(row.get(col))
            if not is_url(url) or url in seen_urls:
                continue
            seen_urls.add(url)
            col_part = safe_filename(str(col), "image")
            sources.append({"type": "url", "url": url, "filename": f"{base_name}_{col_part}.jpg"})
    return sources, [str(c) for c in url_columns], str(filename_column) if filename_column else None


def download_with_retry(session: requests.Session, url: str, retries: int = 3, timeout: int = 25) -> bytes:
    headers = {"User-Agent": "Mozilla/5.0 CreativeEditingAtlas/1.0"}
    last_error = None
    for attempt in range(1, retries + 1):
        try:
            response = session.get(url, timeout=timeout, headers=headers)
            response.raise_for_status()
            if not response.content:
                raise ValueError("Empty image response")
            return response.content
        except Exception as exc:
            last_error = exc
            time.sleep(0.35 * attempt)
    raise RuntimeError(str(last_error))


def has_alpha(img: Image.Image) -> bool:
    return img.mode in ("RGBA", "LA", "PA") or (img.mode == "P" and "transparency" in img.info)


def corner_background_color(img: Image.Image) -> Tuple[int, int, int]:
    rgb = img.convert("RGB")
    w, h = rgb.size
    points = [(0, 0), (w - 1, 0), (0, h - 1), (w - 1, h - 1)]
    colors = [rgb.getpixel(p) for p in points]
    return tuple(sum(c[i] for c in colors) // 4 for i in range(3))


def remove_padding(img: Image.Image, tolerance: int = 6) -> Image.Image:
    """Safely remove only obvious outer padding.

    This never resizes, stretches, or intentionally cuts the product. It only trims:
    - fully transparent empty space, or
    - almost-white empty border around the image.

    The tolerance is intentionally low so product edges are not removed by mistake.
    """
    if has_alpha(img):
        rgba = img.convert("RGBA")
        alpha_bbox = rgba.split()[3].getbbox()
        return rgba.crop(alpha_bbox) if alpha_bbox else rgba

    rgb = img.convert("RGB")
    # Only trim if the outer border is clearly white/off-white. This avoids cutting
    # images that have grey/coloured backgrounds or shadows at the edges.
    border_pixels = []
    w, h = rgb.size
    step_x = max(1, w // 30)
    step_y = max(1, h // 30)
    for x in range(0, w, step_x):
        border_pixels.append(rgb.getpixel((x, 0)))
        border_pixels.append(rgb.getpixel((x, h - 1)))
    for y in range(0, h, step_y):
        border_pixels.append(rgb.getpixel((0, y)))
        border_pixels.append(rgb.getpixel((w - 1, y)))

    white_like = sum(1 for r, g, b in border_pixels if r >= 245 and g >= 245 and b >= 245)
    if white_like / max(1, len(border_pixels)) < 0.85:
        return img

    white_bg = Image.new("RGB", rgb.size, (255, 255, 255))
    diff = ImageChops.difference(rgb, white_bg).convert("L")
    mask = diff.point(lambda px: 255 if px > tolerance else 0)
    bbox = mask.getbbox()
    return img.crop(bbox) if bbox else img


def add_white_background(img: Image.Image) -> Image.Image:
    """Add white only behind transparent pixels. It does not add borders/shadows."""
    if has_alpha(img):
        rgba = img.convert("RGBA")
        bg = Image.new("RGB", rgba.size, (255, 255, 255))
        bg.paste(rgba, mask=rgba.split()[3])
        return bg
    return img.convert("RGB")


def flatten_for_jpeg(img: Image.Image, bg_color: Tuple[int, int, int]) -> Image.Image:
    if has_alpha(img):
        rgba = img.convert("RGBA")
        bg = Image.new("RGB", rgba.size, bg_color)
        bg.paste(rgba, mask=rgba.split()[3])
        return bg
    return img.convert("RGB")


def fit_inside_canvas(img: Image.Image, target_w: int, target_h: int, bg_mode: str, output_format: Optional[str]) -> Image.Image:
    """Resize the complete image into the requested size without crop or stretch.

    Important behaviour:
    - The full original image is always visible.
    - No grey/auto-colour border is added.
    - If Add white background is selected, the canvas is white.
    - Otherwise, PNG/WebP can keep transparent canvas; JPEG/BMP/TIFF use white because
      they cannot reliably store transparency.
    """
    scale = min(target_w / img.width, target_h / img.height)
    new_w = max(1, round(img.width * scale))
    new_h = max(1, round(img.height * scale))
    resized = img.resize((new_w, new_h), Image.LANCZOS)

    fmt = output_format
    transparent_canvas_allowed = fmt in (None, "PNG", "WEBP") and bg_mode != "Add white background"

    if transparent_canvas_allowed:
        canvas = Image.new("RGBA", (target_w, target_h), (0, 0, 0, 0))
    else:
        canvas = Image.new("RGB", (target_w, target_h), (255, 255, 255))

    x = (target_w - new_w) // 2
    y = (target_h - new_h) // 2
    if has_alpha(resized):
        rgba = resized.convert("RGBA")
        canvas.paste(rgba, (x, y), rgba.split()[3])
    else:
        canvas.paste(resized.convert("RGB" if canvas.mode == "RGB" else "RGBA"), (x, y))
    return canvas


def process_image(
    raw_bytes: bytes,
    filename: str,
    resize_mode: str,
    target_w: int,
    target_h: int,
    output_format: Optional[str],
    quality: int,
    output_dpi: int,
    bg_mode: str,
    padding_mode: str,
) -> Tuple[bytes, str, int, int, Optional[Image.Image]]:
    with Image.open(io.BytesIO(raw_bytes)) as opened:
        ImageOps.exif_transpose(opened, in_place=True)
        original_format = opened.format or "PNG"
        original_bg = corner_background_color(opened)
        img = opened.copy()

    if padding_mode == "Remove padding":
        img = remove_padding(img)

    if bg_mode == "Remove background":
        if not REMBG_AVAILABLE:
            raise RuntimeError("Background removal is selected but rembg/onnxruntime is not installed.")
        removed = rembg_remove(img.convert("RGBA"), session=new_session("u2netp"))
        img = removed if isinstance(removed, Image.Image) else Image.open(io.BytesIO(removed)).convert("RGBA")
    elif bg_mode == "Add white background":
        img = add_white_background(img)

    if resize_mode == "By Width only":
        new_w = target_w
        new_h = max(1, round(img.height * target_w / img.width))
        img = img.resize((new_w, new_h), Image.LANCZOS)
    elif resize_mode == "By Height only":
        new_h = target_h
        new_w = max(1, round(img.width * target_h / img.height))
        img = img.resize((new_w, new_h), Image.LANCZOS)
    else:
        img = fit_inside_canvas(img, target_w, target_h, bg_mode, output_format)

    save_fmt = output_format or original_format
    if save_fmt not in ("PNG", "JPEG", "WEBP", "GIF", "BMP", "TIFF"):
        save_fmt = "PNG"

    if save_fmt == "JPEG":
        img = flatten_for_jpeg(img, (255, 255, 255))
    elif save_fmt in ("PNG", "WEBP") and img.mode not in ("RGBA", "RGB", "L", "LA"):
        img = img.convert("RGBA" if has_alpha(img) else "RGB")

    buffer = io.BytesIO()
    save_kwargs = {}
    if save_fmt == "JPEG":
        save_kwargs = {"quality": quality, "subsampling": 0, "optimize": True, "progressive": True, "dpi": (output_dpi, output_dpi)}
    elif save_fmt == "WEBP":
        save_kwargs = {"quality": quality, "method": 6}
    elif save_fmt == "PNG":
        save_kwargs = {"dpi": (output_dpi, output_dpi), "compress_level": 3}

    img.save(buffer, format=save_fmt, **save_kwargs)
    buffer.seek(0)
    output_data = buffer.read()

    stem = safe_filename(filename.rsplit(".", 1)[0] if "." in filename else filename, "image")
    extension = EXT_MAP.get(save_fmt, save_fmt.lower())
    output_name = f"{stem}.{extension}"
    preview_img = img.copy() if img.width * img.height <= 16_000_000 else None
    width, height = img.size
    img.close()
    buffer.close()
    return output_data, output_name, width, height, preview_img


st.divider()
st.subheader("Upload Images")
uploaded_images = st.file_uploader("Upload image files", type=["png", "jpg", "jpeg", "webp", "gif", "bmp", "tiff"], accept_multiple_files=True)
excel_file = st.file_uploader("Or upload an Excel file with image links", type=["xlsx"])
st.markdown("<div class='small-note'>Excel image columns are detected automatically, including Image1, Image2, Image10, ImageLink, URL, Photo, Front, Back, Side, etc.</div>", unsafe_allow_html=True)

if not uploaded_images and excel_file is None:
    st.info("Upload images or an Excel file to start.")
    st.stop()

excel_sources: List[Dict] = []
if excel_file is not None:
    try:
        df_excel = pd.read_excel(excel_file)
        excel_sources, detected_url_columns, detected_filename_column = build_excel_sources(df_excel)
        if detected_url_columns:
            st.success(f"Found {len(excel_sources)} image link(s) from {len(detected_url_columns)} Excel column(s).")
            with st.expander("Detected Excel columns"):
                st.write("Image columns:", ", ".join(detected_url_columns))
                st.write("Filename column:", detected_filename_column or "Auto-generated")
        else:
            st.warning("No image URL columns were detected in the Excel file.")
    except Exception as exc:
        st.error(f"Could not read Excel file: {exc}")

if uploaded_images:
    st.success(f"{len(uploaded_images)} uploaded image(s) ready.")

st.divider()
st.subheader("Background")
bg_mode = st.radio(
    "Choose one",
    ["Keep original", "Remove background", "Add white background"],
    horizontal=True,
    index=0,
)
if bg_mode == "Remove background" and not REMBG_AVAILABLE:
    st.warning("Background removal needs rembg and onnxruntime in requirements.txt.")

st.divider()
st.subheader("Padding")
padding_mode = st.radio("Padding option", ["Keep padding", "Remove padding"], horizontal=True, index=0)
st.caption("Keep padding preserves the image exactly. Remove padding only trims obvious empty white/transparent space.")

st.divider()
st.subheader("Dimensions")
preset = st.selectbox("Choose marketplace recommended size", list(PRESETS.keys()))
default_w, default_h = PRESETS[preset]
if preset in MARKETPLACE_PRESETS and preset != "Custom":
    rec_w, rec_h, min_dim, max_dim, aspect_ratio = MARKETPLACE_PRESETS[preset]
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Recommended", f"{rec_w} x {rec_h}")
    m2.metric("Minimum", min_dim)
    m3.metric("Maximum", max_dim)
    m4.metric("Ratio", aspect_ratio)
st.caption("Safe resize is used: the app will not crop, stretch, add shadows, or auto-add grey borders.")
resize_mode = st.radio(
    "Resize mode",
    ["By Width only", "By Height only", "Exact W x H"],
    horizontal=True,
    help="By Width/Height keeps the image ratio. Exact W x H fits the full image inside the size without cropping or stretching.",
)
col_w, col_h = st.columns(2)
target_w = default_w
target_h = default_h
if resize_mode in ("By Width only", "Exact W x H"):
    target_w = int(col_w.number_input("Width (px)", min_value=1, value=int(default_w), step=1))
if resize_mode in ("By Height only", "Exact W x H"):
    target_h = int(col_h.number_input("Height (px)", min_value=1, value=int(default_h), step=1))
output_dpi = int(st.number_input("DPI", min_value=50, max_value=1200, value=300, step=1))

st.divider()
st.subheader("Output Format")
format_options = ["Keep original format", "PNG", "JPEG", "WEBP"]
chosen_format = st.selectbox("Output format", format_options, index=0)
output_format = None if chosen_format == "Keep original format" else chosen_format
quality = 98
if output_format == "JPEG":
    quality = st.slider("JPEG quality", 80, 100, 98, 1)
elif output_format == "WEBP":
    quality = st.slider("WebP quality", 80, 100, 98, 1)
else:
    st.caption("Keeping original format avoids unnecessary conversion or compression.")

st.divider()
st.subheader("Settings")
preview_enabled = st.checkbox("Preview first 5 processed images", value=True)
retry_count = 3

direct_count = len(uploaded_images) if uploaded_images else 0
url_count = len(excel_sources)
total_count = direct_count + url_count
st.info(f"Ready to process {total_count} image(s).")

if st.button("Process & Download ZIP", type="primary", use_container_width=True):
    if total_count == 0:
        st.warning("No images found to process.")
        st.stop()

    start_time = time.time()
    progress = st.progress(0, text="Starting...")
    status_box = st.empty()
    preview_box = st.container()

    zip_buffer = io.BytesIO()
    errors = []
    processed_rows = []
    preview_items = []
    name_counter: Counter = Counter()
    session = requests.Session()

    def iter_sources() -> Iterable[Dict]:
        if uploaded_images:
            for file in uploaded_images:
                yield {"type": "upload", "file": file, "filename": file.name}
        for source in excel_sources:
            yield source

    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for index, source in enumerate(iter_sources(), start=1):
            raw_bytes = None
            filename = source.get("filename", f"image_{index}.jpg")
            try:
                if source["type"] == "upload":
                    status_box.info(f"Reading image {index} of {total_count}: {filename}")
                    source["file"].seek(0)
                    raw_bytes = source["file"].read()
                else:
                    status_box.info(f"Downloading image {index} of {total_count}: {filename}")
                    raw_bytes = download_with_retry(session, source["url"], retries=retry_count)

                status_box.info(f"Processing image {index} of {total_count}: {filename}")
                output_data, out_name, w, h, preview_img = process_image(
                    raw_bytes=raw_bytes,
                    filename=filename,
                    resize_mode=resize_mode,
                    target_w=target_w,
                    target_h=target_h,
                    output_format=output_format,
                    quality=quality,
                    output_dpi=output_dpi,
                    bg_mode=bg_mode,
                    padding_mode=padding_mode,
                )
                out_name = unique_name(out_name, name_counter)
                zf.writestr(out_name, output_data)
                processed_rows.append({"File": out_name, "Width": w, "Height": h, "Status": "Success"})

                if preview_enabled and preview_img is not None and len(preview_items) < 5:
                    preview_items.append((out_name, preview_img))

            except UnidentifiedImageError:
                errors.append({"File": filename, "Error": "File could not be opened as an image"})
            except Exception as exc:
                errors.append({"File": filename, "Error": str(exc)})

            progress.progress(index / total_count, text=f"Processed {index} of {total_count}")
            raw_bytes = None
            if index % 25 == 0:
                gc.collect()

        if errors:
            error_csv = pd.DataFrame(errors).to_csv(index=False).encode("utf-8")
            zf.writestr("error_report.csv", error_csv)

    zip_buffer.seek(0)
    elapsed = time.time() - start_time
    progress.empty()
    status_box.empty()

    success_count = len(processed_rows)
    fail_count = len(errors)
    st.success(f"Done. {success_count} image(s) processed, {fail_count} failed.")

    c1, c2, c3 = st.columns(3)
    c1.metric("Successful", success_count)
    c2.metric("Failed", fail_count)
    c3.metric("Time", f"{elapsed:.1f}s")

    if preview_enabled and preview_items:
        with preview_box:
            st.markdown("Preview")
            for out_name, preview_img in preview_items:
                st.image(preview_img, caption=out_name, use_container_width=True)
                preview_img.close()

    if errors:
        with st.expander("View errors"):
            st.dataframe(pd.DataFrame(errors), use_container_width=True)
            st.caption("The ZIP includes error_report.csv.")

    st.download_button(
        label=f"Download ZIP ({success_count} images)",
        data=zip_buffer,
        file_name="creative_editing_for_atlas.zip",
        mime="application/zip",
        use_container_width=True,
        type="primary",
    )

st.divider()
st.caption("Default behavior keeps the original image look. No shadow is added. Cropping happens only if Remove padding is selected.")
