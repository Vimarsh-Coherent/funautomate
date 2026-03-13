"""Export PowerPoint slides as JPG images.

Supports multiple backends:
1. Windows + PowerPoint installed: uses comtypes COM automation
2. Linux + LibreOffice: headless conversion to PDF then to images
3. Fallback: returns empty list with warning
"""

import logging
import os
import shutil
import subprocess
from pathlib import Path

logger = logging.getLogger(__name__)

SLIDE_NAMES = [
    "Regional_Insights",
    "Impact_Analysis",
    "Key_Takeaways",
    "Segmental_Insights",
    "Market_KeyPlayer",
]


def _export_via_comtypes(pptx_path: Path, output_dir: Path) -> list[Path]:
    """Export slides as JPG using PowerPoint COM automation (Windows only)."""
    import comtypes.client

    pptx_abs = str(pptx_path.resolve())
    output_abs = str(output_dir.resolve())

    ppt_app = None
    presentation = None
    images = []

    try:
        ppt_app = comtypes.client.CreateObject("PowerPoint.Application")
        ppt_app.Visible = 1  # PowerPoint requires visible window

        presentation = ppt_app.Presentations.Open(pptx_abs, WithWindow=False)

        for i, name in enumerate(SLIDE_NAMES):
            if i < presentation.Slides.Count:
                img_path = Path(output_abs) / f"{name}.jpg"
                presentation.Slides(i + 1).Export(str(img_path), "JPG")
                if img_path.exists():
                    images.append(img_path)
                    logger.info(f"Exported slide {i + 1} -> {name}.jpg")
                else:
                    logger.warning(f"Export reported success but {name}.jpg not found")
    except Exception as e:
        logger.error(f"COM export failed: {e}")
    finally:
        try:
            if presentation:
                presentation.Close()
        except Exception:
            pass
        try:
            if ppt_app:
                ppt_app.Quit()
        except Exception:
            pass

    return images


def _export_via_libreoffice(pptx_path: Path, output_dir: Path) -> list[Path]:
    """Export slides via LibreOffice headless -> PDF -> images."""
    try:
        from PIL import Image
    except ImportError:
        logger.error("Pillow not installed, cannot convert PDF to images")
        return []

    # Step 1: Convert PPTX to PDF via LibreOffice
    pdf_path = output_dir / f"{pptx_path.stem}.pdf"
    result = subprocess.run(
        ["libreoffice", "--headless", "--convert-to", "pdf",
         "--outdir", str(output_dir), str(pptx_path)],
        capture_output=True, text=True, timeout=120,
    )

    if result.returncode != 0 or not pdf_path.exists():
        logger.error(f"LibreOffice conversion failed: {result.stderr}")
        return []

    # Step 2: Convert PDF pages to images
    images = []
    try:
        from pdf2image import convert_from_path
        pil_images = convert_from_path(str(pdf_path), dpi=200)
        for i, (pil_img, name) in enumerate(zip(pil_images, SLIDE_NAMES)):
            img_path = output_dir / f"{name}.jpg"
            pil_img.save(str(img_path), "JPEG", quality=90)
            images.append(img_path)
            logger.info(f"Exported page {i + 1} -> {name}.jpg")
    except ImportError:
        logger.warning("pdf2image not available, trying Pillow PDF reader")
        try:
            img = Image.open(str(pdf_path))
            for i in range(min(img.n_frames, len(SLIDE_NAMES))):
                img.seek(i)
                img_path = output_dir / f"{SLIDE_NAMES[i]}.jpg"
                img.save(str(img_path), "JPEG", quality=90)
                images.append(img_path)
        except Exception as e:
            logger.error(f"PDF to image conversion failed: {e}")
    finally:
        # Clean up PDF
        try:
            pdf_path.unlink()
        except Exception:
            pass

    return images


def export_slides_to_jpg(pptx_path: Path, output_dir: Path) -> list[Path]:
    """Export PPTX slides as JPG images using best available backend.

    Args:
        pptx_path: Path to the PPTX file
        output_dir: Directory to save JPG images

    Returns:
        List of paths to generated JPG files
    """
    pptx_path = Path(pptx_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    if not pptx_path.exists():
        logger.error(f"PPTX file not found: {pptx_path}")
        return []

    # Try Windows COM first
    if os.name == "nt":
        try:
            images = _export_via_comtypes(pptx_path, output_dir)
            if images:
                return images
            logger.warning("COM export returned no images, trying fallback")
        except ImportError:
            logger.info("comtypes not available")
        except Exception as e:
            logger.warning(f"COM export failed: {e}")

    # Try LibreOffice
    if shutil.which("libreoffice") or shutil.which("soffice"):
        try:
            images = _export_via_libreoffice(pptx_path, output_dir)
            if images:
                return images
        except Exception as e:
            logger.warning(f"LibreOffice export failed: {e}")

    logger.warning(
        "No image export backend available. "
        "Install PowerPoint (Windows) or LibreOffice (Linux) for slide-to-JPG export."
    )
    return []
