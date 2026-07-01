"""
Image Resizer Pro - Updated Streamlit App
Requirements: streamlit, Pillow, rembg, onnxruntime, pandas, requests, openpyxl
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


# -----------------------------------------------------------------------------
# Page setup and theme
# -----------------------------------------------------------------------------
st.set_page_config(page_title="Image Resizer Pro", page_icon="image", layout="centered")

NAVY = "#0B2E59"
LIGHT_NAVY = "#123E73"
BORDER = "#D9E2EC"
BG = "#FFFFFF"
SOFT_BG = "#F5F8FC"

st.markdown(
    f"""
    <style>
    .stApp {{
        background: {BG};
    }}
    h1, h2, h3, .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 {{
        color: {NAVY};
    }}
    div[data-testid="stCaptionContainer"] {{
        color: {LIGHT_NAVY};
    }}
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
    .small-note {{
        font-size: 0.86rem;
        color: #4B5563;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Image Resizer Pro")
st.caption("Bulk resize images from uploads or Excel links. Simple, fast, and marketplace-ready.")


# -----------------------------------------------------------------------------
# Presets - marketplace neutral, no Amazon-related code
# -----------------------------------------------------------------------------
PRESETS = {
    "Custom": (800, 800, None, None, None, None, "-"),
    "Allegro PL": (2200, 2200, 500, 500, 2560, 2560, "1:1"),
    "Allegro One PL": (2200, 2200, 1000, 1000, 2560, 2560, "1:1"),
    "Best Buy US": (2000, 2000, 2000, 2000, None, None, "1:1"),
    "Best Buy CA": (2000, 2000, 2000, 2000, None, None, "1:1"),
    "Bol NL": (2400, 2400, 500, 500, 6000, 6000, "1:1"),
    "Bol BE": (2400, 2400, 500, 500, 6000, 6000, "1:1"),
    "eBay US": (1600, 1600, 500, 500, 9000, 9000, "1:1"),
    "eBay DE": (1600, 1600, 500, 500, 9000, 9000, "1:1"),
    "eBay UK": (1600, 1600, 500, 500, 9000, 9000, "1:1"),
    "Kohl's US": (1000, 1000, 1000, 1000, None, None, "1:1"),
    "Lowes US": (1000, 1000, 1000, 1000, None, None, "1:1"),
    "Macy's US": (1000, 1000, 1000, 1000, None, None, "1:1"),
    "MediaMarkt DE": (1200, 1200, 1000, 1000, None, None, "1:1"),
    "Mercado Libre US": (1600, 1600, 500, 500, 2500, 2500, "1:1"),
    "Nordstrom US": (2600, 4000, 1300, 2000, None, None, "2:3"),
    "Octopia FR": (1000, 1000, 500, 500, 2500, 2500, "1:1"),
    "OTTO DE": (960, 480, None, None, None, None, "2:1"),
    "Target US": (2400, 2400, 1200, 1200, 5000, 5000, "1:1"),
    "Tesco UK": (2400, 2400, 1000, 1000, None, None, "1:1"),
    "TikTok US": (1000, 1000, 600, 600, 3000, 3000, "1:1"),
    "TikTok UK": (1000, 1000, 600, 600, 3000, 3000, "1:1"),
    "Walmart US": (2200, 2200, 1500, 1500, 5000, 5000, "1:1"),
    "Walmart CA": (2200, 2200, 1500, 1500, 5000, 5000, "1:1"),
    "Zalando DE": (2000, 2000, 800, 1200, 5000, 5000, "2:3"),
    "Shopify": (2048, 2048, 800, 800, 4472, 4472, "1:1"),
}

EXT_MAP = {"PNG": "png", "JPEG": "jpg", "WEBP": "webp", "GIF": "gif", "BMP": "bmp", "TIFF": "tif"}
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp", ".tif", ".tiff"}
URL_KEYWORDS = ("image", "img", "link", "url", "photo", "picture", "media", "front", "back", "side", "lifestyle")
FILE_KEYWORDS = ("filename", "file name", "file", "sku", "productid", "product id", "item", "item code", "code", "name", "title")


# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------
def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def is_probable_url(value: str) -> bool:
    if not value:
        return False
    value = value.strip()
    parsed = urlparse(value)
    return parsed.scheme in ("http", "https") and bool(parsed.netloc)


def is_probable_image_url(value: str) -> bool:
    if not is_probable_url(value):
        return False
    parsed = urlparse(value)
    path = parsed.path.lower()
    if any(path.endswith(ext) for ext in IMAGE_EXTENSIONS):
        return True
    # Many CDN links do not expose extensions, so accept valid URLs too.
    return True


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
    if counter[key] == 1:
        return base_name
    return f"{stem}_{counter[key]}{ext}"


def detect_url_columns(df: pd.DataFrame) -> List[str]:
    columns = []
    for col in df.columns:
        col_name = str(col).strip().lower()
        name_hint = any(k in col_name for k in URL_KEYWORDS)
        sample = df[col].dropna().astype(str).head(25).tolist()
        url_count = sum(1 for v in sample if is_probable_image_url(v))
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
        base_name = ""
        if filename_column:
            base_name = safe_filename(row.get(filename_column), f"row_{row_index + 1}")
        if not base_name:
            base_name = f"row_{row_index + 1}"

        for col in url_columns:
            url = clean_text(row.get(col))
            if not is_probable_image_url(url):
                continue
            if url in seen_urls:
                continue
            seen_urls.add(url)
            col_part = safe_filename(str(col), "image")
            sources.append({
                "type": "url",
                "url": url,
                "filename": f"{base_name}_{col_part}.jpg",
                "row": row_index + 1,
                "column": str(col),
            })

    return sources, [str(c) for c in url_columns], str(filename_column) if filename_column else None


def download_with_retry(session: requests.Session, url: str, timeout: int = 20, retries: int = 3) -> bytes:
    last_error = None
    headers = {"User-Agent": "Mozilla/5.0 ImageResizerPro/1.0"}
    for attempt in range(1, retries + 1):
        try:
            response = session.get(url, timeout=timeout, headers=headers, stream=True)
            response.raise_for_status()
            content_type = response.headers.get("content-type", "").lower()
            data = response.content
            if not data:
                raise ValueError("Empty response")
            if "text/html" in content_type and len(data) < 200000:
                raise ValueError("URL returned HTML, not an image")
            return data
        except Exception as exc:
            last_error = exc
            time.sleep(0.35 * attempt)
    raise RuntimeError(str(last_error))


def average_corner_color(img: Image.Image) -> Tuple[int, int, int]:
    rgb = img.convert("RGB")
    w, h = rgb.size
    points = [(0, 0), (w - 1, 0), (0, h - 1), (w - 1, h - 1)]
    colors = [rgb.getpixel(p) for p in points]
    return tuple(sum(c[i] for c in colors) // 4 for i in range(3))


def crop_white_border(img: Image.Image, tolerance: int = 10) -> Image.Image:
    rgb = img.convert("RGB")
    bg = Image.new("RGB", rgb.size, (255, 255, 255))
    diff = ImageChops.difference(rgb, bg).convert("L")
    mask = diff.point(lambda p: 255 if p > tolerance else 0)
    bbox = mask.getbbox()
    return img.crop(bbox) if bbox else img


def image_has_alpha(img: Image.Image) -> bool:
    return img.mode in ("RGBA", "LA", "PA") or (img.mode == "P" and "transparency" in img.info)


def flatten_to_background(img: Image.Image, color: Tuple[int, int, int]) -> Image.Image:
    if img.mode == "P":
        img = img.convert("RGBA")
    if image_has_alpha(img):
        rgba = img.convert("RGBA")
        bg = Image.new("RGB", rgba.size, color)
        bg.paste(rgba, mask=rgba.split()[3])
        return bg
    return img.convert("RGB")


def resize_fit(img: Image.Image, box_w: int, box_h: int) -> Image.Image:
    fitted = img.copy()
    fitted.thumbnail((box_w, box_h), Image.LANCZOS)
    return fitted


def paste_center(canvas: Image.Image, img: Image.Image) -> Image.Image:
    x = (canvas.width - img.width) // 2
    y = (canvas.height - img.height) // 2
    if image_has_alpha(img):
        rgba = img.convert("RGBA")
        canvas.paste(rgba, (x, y), rgba.split()[3])
    else:
        canvas.paste(img.convert(canvas.mode), (x, y))
    return canvas


def should_use_smart_canvas(orig_w: int, orig_h: int, resize_mode: str) -> bool:
    if resize_mode == "Exact W x H":
        return True
    ratio = max(orig_w / orig_h, orig_h / orig_w)
    return ratio >= 3.0


def get_target_size(orig_w: int, orig_h: int, resize_mode: str, target_w: int, target_h: int) -> Tuple[int, int]:
    if resize_mode == "By Width only":
        return target_w, max(1, round(orig_h * target_w / orig_w))
    if resize_mode == "By Height only":
        return max(1, round(orig_w * target_h / orig_h)), target_h
    scale = min(target_w / orig_w, target_h / orig_h)
    return max(1, round(orig_w * scale)), max(1, round(orig_h * scale))


def process_image_bytes(
    raw_bytes: bytes,
    filename: str,
    *,
    resize_mode: str,
    target_w: int,
    target_h: int,
    bg_choice: str,
    crop_white: bool,
    output_format: Optional[str],
    quality: int,
    output_dpi: int,
    rembg_session=None,
) -> Tuple[bytes, str, int, int, Optional[Image.Image]]:
    with Image.open(io.BytesIO(raw_bytes)) as opened:
        ImageOps.exif_transpose(opened, in_place=True)
        orig_pil = opened.copy()
        orig_fmt = opened.format or "PNG"

    remove_bg = bg_choice in ("Remove background - transparent", "Remove background - white")
    fill_white = bg_choice == "Remove background - white"

    if remove_bg and REMBG_AVAILABLE:
        raw_bytes = rembg_remove(raw_bytes, session=rembg_session)
        img = Image.open(io.BytesIO(raw_bytes)).convert("RGBA")
        bbox = img.getbbox()
        if bbox:
            img = img.crop(bbox)
    else:
        img = orig_pil.copy()

    if crop_white and not remove_bg:
        img = crop_white_border(img)

    bg_color = (255, 255, 255) if fill_white else average_corner_color(orig_pil)

    # If user selected white background, flatten before resize.
    if fill_white:
        img = flatten_to_background(img, (255, 255, 255))

    orig_w, orig_h = img.size
    save_fmt = output_format or orig_fmt
    if save_fmt not in ("PNG", "JPEG", "WEBP", "GIF", "BMP", "TIFF"):
        save_fmt = "PNG"

    smart_canvas = should_use_smart_canvas(orig_w, orig_h, resize_mode)

    if smart_canvas:
        fitted = resize_fit(img, target_w, target_h)
        if save_fmt == "JPEG" or fill_white:
            canvas = Image.new("RGB", (target_w, target_h), bg_color)
        else:
            canvas = Image.new("RGBA", (target_w, target_h), (bg_color[0], bg_color[1], bg_color[2], 0))
            if not image_has_alpha(img) and not remove_bg:
                canvas = Image.new("RGB", (target_w, target_h), bg_color)
        img = paste_center(canvas, fitted)
    else:
        new_w, new_h = get_target_size(orig_w, orig_h, resize_mode, target_w, target_h)
        img = img.resize((new_w, new_h), Image.LANCZOS)

    if save_fmt == "JPEG":
        img = flatten_to_background(img, bg_color)
    elif save_fmt in ("PNG", "WEBP"):
        if img.mode not in ("RGBA", "RGB", "L", "LA"):
            img = img.convert("RGBA" if image_has_alpha(img) else "RGB")

    buf = io.BytesIO()
    save_kwargs = {}
    if save_fmt == "JPEG":
        save_kwargs = {"quality": quality, "optimize": True, "progressive": True, "dpi": (output_dpi, output_dpi)}
    elif save_fmt == "WEBP":
        save_kwargs = {"quality": quality}
    elif save_fmt == "PNG":
        save_kwargs = {"dpi": (output_dpi, output_dpi)}

    img.save(buf, format=save_fmt, **save_kwargs)
    buf.seek(0)
    output_data = buf.read()

    stem = safe_filename(filename.rsplit(".", 1)[0] if "." in filename else filename, "image")
    new_ext = EXT_MAP.get(save_fmt, save_fmt.lower())
    out_name = f"{stem}.{new_ext}"

    preview_img = img.copy() if img.width * img.height <= 16_000_000 else None
    width, height = img.size

    img.close()
    orig_pil.close()
    buf.close()
    return output_data, out_name, width, height, preview_img


# -----------------------------------------------------------------------------
# UI
# -----------------------------------------------------------------------------
st.divider()
st.subheader("Upload Images")

uploaded_images = st.file_uploader(
    "Upload image files",
    type=["png", "jpg", "jpeg", "webp", "gif", "bmp", "tiff"],
    accept_multiple_files=True,
)

excel_file = st.file_uploader("Or upload an Excel file with image links", type=["xlsx"])

st.markdown(
    "<div class='small-note'>Excel columns are detected automatically. Columns like Image1, ImageLink1, Image2, Image10, URL, Photo, Front, Back, Side, etc. will be scanned.</div>",
    unsafe_allow_html=True,
)

has_input = bool(uploaded_images) or excel_file is not None
if not has_input:
    st.info("Upload images or an Excel file to start.")
    st.stop()

excel_sources: List[Dict] = []
detected_url_columns: List[str] = []
detected_filename_column: Optional[str] = None

if excel_file is not None:
    try:
        df_excel = pd.read_excel(excel_file)
        excel_sources, detected_url_columns, detected_filename_column = build_excel_sources(df_excel)
        if detected_url_columns:
            st.success(f"Found {len(excel_sources)} image link(s) from {len(detected_url_columns)} column(s).")
            with st.expander("Detected Excel columns"):
                st.write("Image columns:", ", ".join(detected_url_columns))
                st.write("Filename column:", detected_filename_column or "Auto-generated")
        else:
            st.warning("No image URL columns were detected in the Excel file.")
    except Exception as exc:
        st.error(f"Could not read Excel file: {exc}")
        excel_sources = []

if uploaded_images:
    st.success(f"{len(uploaded_images)} uploaded image(s) ready.")

st.divider()
st.subheader("Background")

bg_choice = st.radio(
    "Choose background handling",
    ["Keep original background", "Remove background - transparent", "Remove background - white"],
    index=0,
)

if bg_choice.startswith("Remove background") and not REMBG_AVAILABLE:
    st.error("Background removal is not available. Add rembg and onnxruntime to requirements.txt.")
elif bg_choice.startswith("Remove background"):
    st.caption("Background removal may be slower on the first run.")

crop_white = False
if bg_choice == "Keep original background":
    crop_white = st.checkbox("Crop existing white border before resizing", value=False)

st.divider()
st.subheader("Dimensions")

marketplace = st.selectbox("Choose preset", list(PRESETS.keys()))
rec_w, rec_h, min_w, min_h, max_w, max_h, ratio = PRESETS[marketplace]

if marketplace != "Custom":
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Recommended", f"{rec_w} x {rec_h}")
    c2.metric("Minimum", f"{min_w} x {min_h}" if min_w else "-")
    c3.metric("Maximum", f"{max_w} x {max_h}" if max_w else "-")
    c4.metric("Ratio", ratio)

resize_mode = st.radio(
    "Resize mode",
    ["By Width only", "By Height only", "Exact W x H"],
    horizontal=True,
    help="Exact W x H fits the image inside the canvas without stretching.",
)

col_w, col_h = st.columns(2)
target_w = int(rec_w)
target_h = int(rec_h)

if resize_mode in ("By Width only", "Exact W x H"):
    target_w = int(col_w.number_input("Width (px)", min_value=1, value=target_w, step=1))
if resize_mode in ("By Height only", "Exact W x H"):
    target_h = int(col_h.number_input("Height (px)", min_value=1, value=target_h, step=1))

if resize_mode == "By Width only":
    target_h = int(rec_h)
elif resize_mode == "By Height only":
    target_w = int(rec_w)

output_dpi = int(st.number_input("DPI", min_value=50, max_value=1200, value=300, step=1))

st.divider()
st.subheader("Output Format")

format_options = ["Keep original format", "PNG", "JPEG", "WEBP"]
default_format = "JPEG" if marketplace != "Custom" else "Keep original format"
chosen_format = st.selectbox("Output format", format_options, index=format_options.index(default_format))
output_format = None if chosen_format == "Keep original format" else chosen_format

quality = 92
if output_format in ("JPEG", None):
    quality = st.slider("JPEG quality", 10, 100, 92, 1)
elif output_format == "WEBP":
    quality = st.slider("WebP quality", 10, 100, 90, 1)

if output_format == "JPEG" and bg_choice == "Remove background - transparent":
    st.warning("JPEG cannot keep transparency. The app will use the detected background color when saving JPEG.")

st.divider()
st.subheader("Settings")

preview_enabled = st.checkbox("Preview first 5 processed images", value=True)
retry_count = 3

# Prepare task counts without loading everything into memory.
direct_count = len(uploaded_images) if uploaded_images else 0
url_count = len(excel_sources)
total_count = direct_count + url_count

st.info(f"Ready to process {total_count} image(s).")


# -----------------------------------------------------------------------------
# Processing
# -----------------------------------------------------------------------------
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

    rembg_session = None
    if bg_choice.startswith("Remove background") and REMBG_AVAILABLE:
        with st.spinner("Loading background removal model..."):
            rembg_session = new_session("u2netp")

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
                    status_box.info(f"Reading uploaded image {index} of {total_count}: {filename}")
                    source["file"].seek(0)
                    raw_bytes = source["file"].read()
                else:
                    status_box.info(f"Downloading image {index} of {total_count}: {filename}")
                    raw_bytes = download_with_retry(session, source["url"], retries=retry_count)

                status_box.info(f"Processing image {index} of {total_count}: {filename}")
                output_data, out_name, w, h, preview_img = process_image_bytes(
                    raw_bytes,
                    filename,
                    resize_mode=resize_mode,
                    target_w=target_w,
                    target_h=target_h,
                    bg_choice=bg_choice,
                    crop_white=crop_white,
                    output_format=output_format,
                    quality=quality,
                    output_dpi=output_dpi,
                    rembg_session=rembg_session,
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
        file_name="processed_images.zip",
        mime="application/zip",
        use_container_width=True,
        type="primary",
    )

st.divider()
st.caption("Images are processed only in your session. Created by Kashish Sitlani.")
