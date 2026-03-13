"""
Excel VBA macro runner.
Opens .xlsm files in Excel via VBScript, runs macros to generate chart images,
and collects the output files.
"""
import os
import shutil
import subprocess
import time
from pathlib import Path
from typing import Callable

EXPECTED_IMAGES = [
    "Key_Takeaways.jpg",
    "Impact_Analysis.jpg",
    "Segmental_Insights.jpg",
    "Regional_Insights.jpg",
    "Market_KeyPlayer.jpg",
]

VBS_PATH = Path(__file__).parent / "run_macro.vbs"


def is_macro_available() -> bool:
    """Check if we can run Excel macros (Windows only)."""
    if os.name != "nt":
        return False
    # Check if cscript exists (built into Windows)
    try:
        result = subprocess.run(
            ["cscript", "//Nologo", "//?"],
            capture_output=True, timeout=5,
        )
        return True
    except Exception:
        return False


def _snapshot_jpg_files(directory: Path) -> set:
    """Get set of .jpg files in a directory."""
    if not directory.exists():
        return set()
    return {f.name for f in directory.glob("*.jpg")}


def _find_generated_images(search_dirs: list[Path], pre_existing: set) -> list[Path]:
    """Search for expected images in multiple directories."""
    found = []
    found_names = set()

    for search_dir in search_dirs:
        if not search_dir.exists():
            continue
        for expected_name in EXPECTED_IMAGES:
            if expected_name in found_names:
                continue
            candidate = search_dir / expected_name
            if candidate.exists():
                found.append(candidate)
                found_names.add(expected_name)

    # Also check for any NEW jpg files (in case names differ)
    if len(found) < len(EXPECTED_IMAGES):
        for search_dir in search_dirs:
            if not search_dir.exists():
                continue
            for f in search_dir.glob("*.jpg"):
                if f.name not in found_names and f.name not in pre_existing:
                    found.append(f)
                    found_names.add(f.name)

    return found


def run_macro_and_collect_images(
    xlsm_path: str | Path,
    on_progress: Callable | None = None,
    timeout_seconds: int = 120,
) -> list[Path]:
    """
    Open .xlsm in Excel via VBScript, run CreateImageAndTable macro,
    collect generated images.
    Returns list of Path objects to the generated image files.
    """
    xlsm_path = Path(xlsm_path).resolve()
    xlsm_dir = xlsm_path.parent

    def progress(step, status, detail=""):
        if on_progress:
            on_progress(step, status, detail)

    if not is_macro_available():
        progress("macro", "warning", "Excel macro execution not available (requires Windows)")
        return []

    if not VBS_PATH.exists():
        progress("macro", "error", f"VBScript not found: {VBS_PATH}")
        return []

    # Snapshot existing jpg files before macro runs
    pre_existing = _snapshot_jpg_files(xlsm_dir)
    progress("macro", "running", f"Opening Excel and running macro: {xlsm_path.name}")

    try:
        # Run macro via VBScript
        result = subprocess.run(
            [
                "cscript", "//Nologo",
                str(VBS_PATH),
                str(xlsm_path),
                "CreateImageAndTable",
            ],
            capture_output=True,
            text=True,
            timeout=timeout_seconds,
            cwd=str(xlsm_dir),
        )

        output = result.stdout.strip()
        error = result.stderr.strip()

        if result.returncode == 0 and "SUCCESS" in output:
            progress("macro", "done", "Macro executed successfully")
        elif result.returncode == 3:
            progress("macro", "warning", f"Macro error (may have partially completed): {output}")
        else:
            progress("macro", "error", f"VBScript failed (code {result.returncode}): {output} {error}")

    except subprocess.TimeoutExpired:
        progress("macro", "error", f"Macro timed out after {timeout_seconds}s — killing Excel")
        try:
            subprocess.run(["taskkill", "/F", "/IM", "EXCEL.EXE"], capture_output=True)
        except Exception:
            pass
        return []
    except Exception as e:
        progress("macro", "error", f"Failed to run VBScript: {e}")
        return []

    # Give file system a moment
    time.sleep(2)

    # Search for generated images
    progress("macro", "running", "Searching for generated images...")

    search_dirs = [
        xlsm_dir,
        Path(os.environ.get("TEMP", "")),
        Path(os.environ.get("USERPROFILE", "")) / "Desktop",
        Path(os.environ.get("USERPROFILE", "")) / "Documents",
    ]

    images = _find_generated_images(search_dirs, pre_existing)

    if images:
        progress("macro", "done", f"Found {len(images)} images: {', '.join(p.name for p in images)}")
    else:
        progress("macro", "warning", "No images found after macro execution")

    missing = set(EXPECTED_IMAGES) - {p.name for p in images}
    if missing:
        progress("macro", "warning", f"Missing images: {', '.join(sorted(missing))}")

    return images


def copy_images_to_output(image_paths: list[Path], output_dir: Path) -> list[Path]:
    """Copy images to the final output directory. Returns list of destination paths."""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    copied = []
    for img in image_paths:
        dest = output_dir / img.name
        shutil.copy2(str(img), str(dest))
        copied.append(dest)
    return copied
