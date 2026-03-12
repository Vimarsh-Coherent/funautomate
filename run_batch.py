"""
Standalone batch automation runner.
Called by Streamlit as a subprocess. Reads job config from a JSON file,
runs automation for each Excel file, writes progress + results to output JSON.
"""
import json
import sys
import time
from datetime import datetime
from pathlib import Path

from core.excel_parser import parse_excel
from core.automator import run_automation_sync


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

            all_results.append({
                "name": file_name,
                "market": data.market_name,
                "success": result.success,
                "error": result.error,
                "downloaded_file": result.downloaded_file,
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
