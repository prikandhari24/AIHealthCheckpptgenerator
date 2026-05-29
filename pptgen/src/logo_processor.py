from __future__ import annotations

from pathlib import Path
import re

from PIL import Image, ImageOps

from src.config import UPLOAD_DIR


def slugify(value: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "_", value.strip().lower())
    return slug.strip("_") or "logo"


def save_uploaded_file(uploaded_file, output_path: Path) -> Path:
    with open(output_path, "wb") as file_handle:
        file_handle.write(uploaded_file.getbuffer())
    return output_path


def resize_logo(
    input_path: Path,
    output_path: Path,
    size: tuple[int, int] = (400, 180),
    grayscale: bool = False,
) -> Path:
    image = Image.open(input_path).convert("RGBA")

    if grayscale:
        gray = ImageOps.grayscale(image.convert("RGB"))
        image = Image.merge("RGBA", (gray, gray, gray, image.getchannel("A")))

    bbox = image.getbbox()
    if bbox:
        image = image.crop(bbox)

    image.thumbnail(size, Image.LANCZOS)
    image.save(output_path)
    return output_path


def _raw_logo_path(prefix: str, uploaded_file) -> Path:
    suffix = Path(uploaded_file.name).suffix or ".png"
    return UPLOAD_DIR / f"{prefix}{suffix}"


def _process_uploaded_logo(
    uploaded_file,
    raw_prefix: str,
    processed_filename: str,
    grayscale: bool = False,
) -> str:
    raw_path = _raw_logo_path(raw_prefix, uploaded_file)
    processed_path = UPLOAD_DIR / processed_filename
    save_uploaded_file(uploaded_file, raw_path)
    resize_logo(raw_path, processed_path, grayscale=grayscale)
    return str(processed_path)


def process_logos(front_slide_logo, client_logo, peer_logos_by_company: dict[str, object]):
    UPLOAD_DIR.mkdir(exist_ok=True)

    client_logo_path = _process_uploaded_logo(
        client_logo,
        "client_logo_raw",
        "client_logo_processed.png",
        grayscale=False,
    )
    title_logo_path = (
        _process_uploaded_logo(
            front_slide_logo,
            "front_slide_logo_raw",
            "front_slide_logo_processed.png",
            grayscale=False,
        )
        if front_slide_logo is not None
        else client_logo_path
    )

    processed_peers: dict[str, str] = {}
    for company, logo_file in peer_logos_by_company.items():
        company_slug = slugify(company)
        processed_peers[company] = _process_uploaded_logo(
            logo_file,
            f"{company_slug}_peer_raw",
            f"{company_slug}_peer_processed.png",
            grayscale=True,
        )

    return title_logo_path, client_logo_path, processed_peers
