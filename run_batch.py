"""
Standalone batch automation runner.
Called by Streamlit as a subprocess. Reads job config from a JSON file,
runs automation for each Excel file, writes progress + results to output JSON.
"""
import json
import logging
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path

from core.excel_parser import parse_excel, parse_excel_full
from core.automator import run_automation_sync
from core.pptx_generator import generate_pptx
from core.image_exporter import export_slides_to_jpg
from core.doc_generator import generate_combined_doc, generate_toc_doc

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Template search order for PPTX
TEMPLATE_SEARCH_PATHS = [
    "Template/Collateral Designs.pptx",
    "temp_ref/Global Skin Packaging Market/Output.pptx",
]


def find_pptx_template() -> Path | None:
    """Find the PowerPoint template file."""
    base = Path(__file__).parent
    for rel_path in TEMPLATE_SEARCH_PATHS:
        p = base / rel_path
        if p.exists():
            return p
    return None


def main():
    if len(sys.argv) < 2:
        print("Usage: python run_batch.py <job_config.json>")
        sys.exit(1)

    job_path = Path(sys.argv[1])
    job = json.loads(job_path.read_text(encoding="utf-8"))

    progress_path = job_path.parent / "progress.json"
    results_path = job_path.parent / "results.json"

    files = job["files"]  # list of {"name": str, "path": str}
    app_url = job["app_url"]
    email = job["email"]
    password = job["password"]
    dd1 = job["collateral_dd1"]
    dd2 = job["collateral_dd2"]
    headless = job.get("headless", False)
    timeout_ms = job.get("timeout_ms", 180000)
    output_dir = job.get("output_dir", "outputs")

    all_results = []
    log_entries = []

    def write_progress(current_index, total, log_entries_list):
        progress_data = {
            "current_index": current_index,
            "total": total,
            "running": True,
            "log": log_entries_list[-100:],  # last 100 entries
        }
        progress_path.write_text(json.dumps(progress_data, ensure_ascii=False), encoding="utf-8")

    def write_results(results_list, running=True):
        results_data = {
            "running": running,
            "results": results_list,
        }
        results_path.write_text(json.dumps(results_data, ensure_ascii=False), encoding="utf-8")

    for i, file_info in enumerate(files):
        file_name = file_info["name"]
        file_path = file_info["path"]

        log_entries.append({
            "time": datetime.now().strftime("%H:%M:%S"),
            "step": "batch",
            "status": "running",
            "detail": f"--- File {i+1}/{len(files)}: {file_name} ---",
        })
        write_progress(i + 1, len(files), log_entries)

        # Parse the Excel
        try:
            data = parse_excel(file_path)
        except Exception as e:
            log_entries.append({
                "time": datetime.now().strftime("%H:%M:%S"),
                "step": "parse",
                "status": "error",
                "detail": f"Failed to parse {file_name}: {e}",
            })
            all_results.append({
                "name": file_name,
                "market": file_name,
                "success": False,
                "error": f"Parse error: {e}",
                "downloaded_file": None,
                "image_files": [],
                "generated_files": [],
            })
            write_progress(i + 1, len(files), log_entries)
            write_results(all_results)
            continue

        def progress_callback(step_id, status, detail=""):
            log_entries.append({
                "time": datetime.now().strftime("%H:%M:%S"),
                "step": step_id,
                "status": status,
                "detail": detail,
            })
            write_progress(i + 1, len(files), log_entries)

        # ---- NEW: Generate PPTX, images, and documents ----
        generated_files = []
        image_files = []
        try:
            full_data = parse_excel_full(file_path)
            market_name_safe = data.market_name.replace("/", "-").replace("\\", "-")
            ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            gen_output_dir = Path(output_dir) / market_name_safe / ts
            gen_output_dir.mkdir(parents=True, exist_ok=True)

            template = find_pptx_template()
            if template:
                # Generate modified PPTX
                progress_callback("generate", "running", "Generating PowerPoint...")
                pptx_path = gen_output_dir / "Output.pptx"
                generate_pptx(template, full_data, pptx_path)
                generated_files.append(str(pptx_path))
                progress_callback("generate", "running", "PowerPoint generated")

                # Export slides as JPG
                progress_callback("generate", "running", "Exporting slide images...")
                images = export_slides_to_jpg(pptx_path, gen_output_dir)
                image_files = [str(p) for p in images]
                progress_callback("generate", "running", f"Exported {len(images)} images")
            else:
                progress_callback("generate", "warning", "No PPTX template found, skipping image generation")

            # Generate Combined document
            progress_callback("generate", "running", "Generating Combined document...")
            combined_path = gen_output_dir / f"Combined - {market_name_safe}.docx"
            generate_combined_doc(full_data, gen_output_dir, combined_path)
            generated_files.append(str(combined_path))

            # Generate TOC document
            progress_callback("generate", "running", "Generating TOC document...")
            toc_path = gen_output_dir / f"TOC - {market_name_safe}.docx"
            generate_toc_doc(full_data, toc_path)
            generated_files.append(str(toc_path))

            progress_callback("generate", "done", f"Generated {len(generated_files)} files + {len(image_files)} images")

        except Exception as e:
            logger.error(f"Generation step failed for {file_name}: {e}")
            progress_callback("generate", "error", f"Generation failed: {e}")

        # ---- Web automation (existing) ----
        try:
            result = run_automation_sync(
                market_data=data,
                app_url=app_url,
                email=email,
                password=password,
                collateral_dd1=dd1,
                collateral_dd2=dd2,
                headless=headless,
                timeout_ms=timeout_ms,
                output_dir=output_dir,
                on_progress=progress_callback,
            )

            # Copy web-downloaded .doc to generation output dir
            if result.success and result.downloaded_file and gen_output_dir.exists():
                try:
                    shutil.copy2(result.downloaded_file, gen_output_dir / Path(result.downloaded_file).name)
                except Exception:
                    pass

            all_results.append({
                "name": file_name,
                "market": data.market_name,
                "success": result.success,
                "error": result.error,
                "downloaded_file": result.downloaded_file,
                "image_files": image_files,
                "generated_files": generated_files,
                "output_dir": str(gen_output_dir),
            })

            if result.success:
                log_entries.append({
                    "time": datetime.now().strftime("%H:%M:%S"),
                    "step": "batch",
                    "status": "done",
                    "detail": f"COMPLETED: {data.market_name}",
                })
            else:
                log_entries.append({
                    "time": datetime.now().strftime("%H:%M:%S"),
                    "step": "batch",
                    "status": "error",
                    "detail": f"FAILED: {result.error}",
                })

        except Exception as e:
            all_results.append({
                "name": file_name,
                "market": data.market_name,
                "success": False,
                "error": str(e),
                "downloaded_file": None,
                "image_files": image_files,
                "generated_files": generated_files,
                "output_dir": str(gen_output_dir) if 'gen_output_dir' in dir() else None,
            })
            log_entries.append({
                "time": datetime.now().strftime("%H:%M:%S"),
                "step": "batch",
                "status": "error",
                "detail": f"FATAL: {e}",
            })

        write_progress(i + 1, len(files), log_entries)
        write_results(all_results)

    # Final: mark as done
    log_entries.append({
        "time": datetime.now().strftime("%H:%M:%S"),
        "step": "batch",
        "status": "done",
        "detail": f"=== Batch complete: {len(files)} files processed ===",
    })
    write_progress(len(files), len(files), log_entries)
    write_results(all_results, running=False)
    print(f"Done. {len(all_results)} files processed.")


if __name__ == "__main__":
    main()
