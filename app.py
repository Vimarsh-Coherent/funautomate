import base64
import json
import os
import subprocess
import sys
import time
import zipfile
from datetime import datetime
from io import BytesIO
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components
import yaml

from core.excel_parser import parse_excel, load_mapping
from core.models import MarketData

# ──────────────────────────────────────────────
# Ensure Playwright browsers are installed
# ──────────────────────────────────────────────
@st.cache_resource
def install_playwright_browser():
    """Install Chromium browser for Playwright (runs once per app lifecycle)."""
    subprocess.run(
        [sys.executable, "-m", "playwright", "install", "chromium"],
        check=True,
        capture_output=True,
    )
    return True

try:
    install_playwright_browser()
except Exception:
    pass  # Already installed or will fail at automation time with a clear error


def auto_download(file_bytes: bytes, file_name: str):
    """Trigger an automatic browser download using JavaScript."""
    b64 = base64.b64encode(file_bytes).decode()
    js = f"""
    <script>
    (function() {{
        var link = document.createElement('a');
        link.href = 'data:application/octet-stream;base64,{b64}';
        link.download = '{file_name}';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }})();
    </script>
    """
    components.html(js, height=0)


# Detect if running on Streamlit Cloud (Linux) vs local (Windows)
IS_CLOUD = os.name != "nt"

# -- Page Config --
st.set_page_config(
    page_title="MarketRytrAI Automation",
    page_icon="🤖",
    layout="wide",
)

BASE_DIR = Path(__file__).parent
CONFIG_PATH = BASE_DIR / "config" / "field_mapping.yaml"
OUTPUT_DIR = BASE_DIR / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)
JOBS_DIR = BASE_DIR / "jobs"
JOBS_DIR.mkdir(exist_ok=True)


def load_config() -> dict:
    return load_mapping(CONFIG_PATH)


def save_config(config: dict):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        yaml.dump(config, f, default_flow_style=False, allow_unicode=True)


# ──────────────────────────────────────────────
# Session State Initialization
# ──────────────────────────────────────────────
if "parsed_files" not in st.session_state:
    st.session_state.parsed_files = []
if "job_id" not in st.session_state:
    st.session_state.job_id = None


def get_job_dir(job_id: str) -> Path:
    return JOBS_DIR / job_id


def is_job_running(job_id: str | None) -> bool:
    if not job_id:
        return False
    results_path = get_job_dir(job_id) / "results.json"
    if not results_path.exists():
        progress_path = get_job_dir(job_id) / "progress.json"
        return progress_path.exists()
    try:
        data = json.loads(results_path.read_text(encoding="utf-8"))
        return data.get("running", False)
    except Exception:
        return False


def read_progress(job_id: str) -> dict | None:
    if not job_id:
        return None
    progress_path = get_job_dir(job_id) / "progress.json"
    if not progress_path.exists():
        return None
    try:
        return json.loads(progress_path.read_text(encoding="utf-8"))
    except Exception:
        return None


def read_results(job_id: str) -> dict | None:
    if not job_id:
        return None
    results_path = get_job_dir(job_id) / "results.json"
    if not results_path.exists():
        return None
    try:
        return json.loads(results_path.read_text(encoding="utf-8"))
    except Exception:
        return None


# ──────────────────────────────────────────────
# Header
# ──────────────────────────────────────────────
st.title("MarketRytrAI Automation Dashboard")
st.caption("Autonomous Excel-to-Web form filler with report download — upload multiple files")

# ──────────────────────────────────────────────
# Sidebar: Credentials
# ──────────────────────────────────────────────
with st.sidebar:
    st.header("MarketRytrAI Login")
    config = load_config()
    app_config = config.get("app", {})
    saved_email = app_config.get("email", "")
    saved_password = app_config.get("password", "")

    if saved_email and saved_password:
        st.success(f"Logged in as: **{saved_email}**")
        if st.button("Change Credentials"):
            st.session_state["show_cred_form"] = True
    else:
        st.warning("Credentials not configured.")
        st.session_state["show_cred_form"] = True

    if st.session_state.get("show_cred_form", not (saved_email and saved_password)):
        with st.form("cred_form"):
            email_input = st.text_input("Email", value=saved_email)
            password_input = st.text_input("Password", value=saved_password, type="password")
            if st.form_submit_button("Save Credentials"):
                config["app"]["email"] = email_input
                config["app"]["password"] = password_input
                save_config(config)
                st.session_state["show_cred_form"] = False
                st.success("Credentials saved!")
                st.rerun()

    st.divider()
    st.markdown(f"**App URL:** {app_config.get('url', 'N/A')}")
    st.markdown(f"**DD1:** {app_config.get('collateral_dropdown1', 'Report Summary')}")
    st.markdown(f"**DD2:** {app_config.get('collateral_dropdown2', 'CMI')}")
    if IS_CLOUD:
        st.caption("Running on Streamlit Cloud (headless mode forced)")


tab1, tab2, tab3, tab4 = st.tabs([
    "Upload & Preview",
    "Configuration",
    "Run Automation",
    "View Outputs",
])


# ══════════════════════════════════════════════
# TAB 1: Upload & Preview (Multiple Files)
# ══════════════════════════════════════════════
with tab1:
    st.header("Upload Excel Files")
    st.info("Upload **any number** of Excel files. Each will be parsed and queued for automation.")

    uploaded_files = st.file_uploader(
        "Choose .xlsm or .xlsx files",
        type=["xlsm", "xlsx"],
        accept_multiple_files=True,
        help="Select one or more market research Excel files",
    )

    if uploaded_files:
        temp_path = BASE_DIR / "data"
        temp_path.mkdir(exist_ok=True)

        existing_names = {f["name"] for f in st.session_state.parsed_files}
        new_files_added = 0

        for uploaded in uploaded_files:
            if uploaded.name not in existing_names:
                file_path = temp_path / uploaded.name
                with open(file_path, "wb") as f:
                    f.write(uploaded.getbuffer())

                entry = {"name": uploaded.name, "path": str(file_path), "data": None, "error": None}
                try:
                    data = parse_excel(file_path)
                    entry["data"] = data
                except Exception as e:
                    entry["error"] = str(e)

                st.session_state.parsed_files.append(entry)
                new_files_added += 1

        if new_files_added > 0:
            st.success(f"Added {new_files_added} new file(s). Total: {len(st.session_state.parsed_files)}")

    if st.session_state.parsed_files:
        st.divider()
        st.subheader(f"Uploaded Files ({len(st.session_state.parsed_files)})")

        for i, entry in enumerate(st.session_state.parsed_files):
            data = entry["data"]
            err = entry["error"]

            if err:
                st.error(f"**{i+1}. {entry['name']}** — Parse error: {err}")
                continue

            with st.expander(f"{i+1}. {data.market_name} ({entry['name']})", expanded=(i == 0)):
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown(f"**Market Name:** {data.market_name}")
                    st.markdown(f"**Market Size:** {data.market_size_text[:80]}...")
                    st.markdown(f"**Disease Name:** {data.disease_name}")
                    st.markdown(f"**Segments:** {len(data.segments)}")
                    for seg in data.segments:
                        st.markdown(f"  - {seg.name} ({len(seg.sub_segments)} sub-segs, dominating: {seg.dominating})")
                with col2:
                    st.markdown(f"**Driver 1:** {(data.driver_1 or 'N/A')[:80]}")
                    st.markdown(f"**Driver 2:** {(data.driver_2 or 'N/A')[:80]}")
                    st.markdown(f"**Restrain:** {(data.restrain or 'N/A')[:80]}")
                    st.markdown(f"**Opportunity:** {(data.opportunity or 'N/A')[:80]}")
                    st.markdown(f"**Dominating Region:** {data.dominating_region or 'N/A'}")
                    st.markdown(f"**Fastest Growing Region:** {data.fastest_growing_region or 'N/A'}")

                    missing = []
                    if not data.market_name: missing.append("Market Name")
                    if not data.segments: missing.append("Segments")
                    if not data.driver_1: missing.append("Driver 1")
                    if not data.dominating_region: missing.append("Dominating Region")
                    if missing:
                        st.warning(f"Missing: {', '.join(missing)}")

        if st.button("Clear All Files"):
            st.session_state.parsed_files = []
            st.session_state.job_id = None
            st.rerun()
    else:
        st.info("No files uploaded yet.")


# ══════════════════════════════════════════════
# TAB 2: Configuration
# ══════════════════════════════════════════════
with tab2:
    st.header("Configuration")
    config = load_config()

    st.subheader("App Settings")
    app_config = config.get("app", {})
    app_url = st.text_input("Application URL", value=app_config.get("url", "https://marketrytrai.com/application"))

    st.subheader("Login Credentials")
    login_email = st.text_input("Email", value=app_config.get("email", ""), key="cfg_email")
    login_password = st.text_input("Password", value=app_config.get("password", ""), type="password", key="cfg_password")

    st.subheader("Collateral Dropdowns")
    dd1_val = st.text_input("Collateral Dropdown 1", value=app_config.get("collateral_dropdown1", "Report Summary"))
    dd2_val = st.text_input("Collateral Dropdown 2", value=app_config.get("collateral_dropdown2", "CMI"))

    st.subheader("Field Mapping (Excel Cell Positions)")
    st.caption("Change only if your Excel layout differs.")

    fields = config.get("fields", {})
    col1, col2 = st.columns(2)
    with col1:
        market_name_cell = st.text_input("Market Name Cell", value=fields.get("market_name", "D2"))
        market_size_row1 = st.number_input("Market Size Row 1", value=fields.get("market_size_row1", 7), min_value=1)
        market_size_row2 = st.number_input("Market Size Row 2", value=fields.get("market_size_row2", 8), min_value=1)
        cagr_row = st.number_input("CAGR Row", value=fields.get("cagr_row", 9), min_value=1)
    with col2:
        disease_name = st.text_input("Disease Name (default)", value=fields.get("disease_name", "NA"))
        seg_start_col = st.text_input("Segments Start Column", value=config.get("segments", {}).get("start_col", "H"))
        seg_header_row = st.number_input("Segments Header Row", value=config.get("segments", {}).get("header_row", 2), min_value=1)

    st.subheader("Market Size Text Template")
    template = st.text_input(
        "Template",
        value=config.get("market_size_template", "Market size in {year1} is {size1} and for {year2} is {size2} and CAGR is {cagr}%"),
        help="Available variables: {year1}, {size1}, {year2}, {size2}, {cagr}",
    )

    if st.button("Save Configuration"):
        config["app"] = {
            "url": app_url,
            "email": login_email,
            "password": login_password,
            "collateral_dropdown1": dd1_val,
            "collateral_dropdown2": dd2_val,
        }
        config["fields"]["market_name"] = market_name_cell
        config["fields"]["market_size_row1"] = int(market_size_row1)
        config["fields"]["market_size_row2"] = int(market_size_row2)
        config["fields"]["cagr_row"] = int(cagr_row)
        config["fields"]["disease_name"] = disease_name
        config["segments"]["start_col"] = seg_start_col
        config["segments"]["header_row"] = int(seg_header_row)
        config["market_size_template"] = template
        save_config(config)
        st.success("Configuration saved!")


# ══════════════════════════════════════════════
# TAB 3: Run Automation (Batch via subprocess)
# ══════════════════════════════════════════════
with tab3:
    st.header("Run Automation")

    config = load_config()
    app_config = config.get("app", {})

    if not app_config.get("email") or not app_config.get("password"):
        st.error("Credentials not configured! Enter your email and password in the sidebar or Configuration tab.")
        st.stop()

    valid_files = [f for f in st.session_state.parsed_files if f["data"] is not None]
    failed_files = [f for f in st.session_state.parsed_files if f["error"] is not None]

    job_id = st.session_state.job_id
    job_running = is_job_running(job_id)

    if not valid_files and not job_running:
        st.warning("No Excel files uploaded yet. Go to **Upload & Preview** tab first.")
    else:
        if valid_files and not job_running:
            st.info(f"**{len(valid_files)} file(s)** ready for automation" +
                    (f" ({len(failed_files)} failed parsing)" if failed_files else ""))

            for i, f in enumerate(valid_files):
                st.markdown(f"⏳ **{i+1}.** {f['data'].market_name} (`{f['name']}`)")

        st.divider()

        col1, col2, col3 = st.columns(3)
        with col1:
            # Force headless on cloud
            if IS_CLOUD:
                headless = True
                st.checkbox("Run Headless (forced on Cloud)", value=True, disabled=True)
            else:
                headless = st.checkbox("Run Headless (no browser window)", value=False)
        with col2:
            timeout_sec = st.slider("Timeout per file (seconds)", min_value=30, max_value=600, value=180)
        with col3:
            st.markdown(f"**URL:** {app_config.get('url', 'N/A')}")

        # START BUTTON
        if st.button("Start Batch Automation", type="primary", disabled=job_running):
            job_id = datetime.now().strftime("%Y%m%d_%H%M%S")
            job_dir = get_job_dir(job_id)
            job_dir.mkdir(parents=True, exist_ok=True)

            job_config = {
                "files": [{"name": f["name"], "path": f["path"]} for f in valid_files],
                "app_url": app_config.get("url", "https://marketrytrai.com/application"),
                "email": app_config.get("email", ""),
                "password": app_config.get("password", ""),
                "collateral_dd1": app_config.get("collateral_dropdown1", "Report Summary"),
                "collateral_dd2": app_config.get("collateral_dropdown2", "CMI"),
                "headless": headless if IS_CLOUD else headless,
                "timeout_ms": timeout_sec * 1000,
                "output_dir": str(OUTPUT_DIR),
            }

            job_file = job_dir / "job.json"
            job_file.write_text(json.dumps(job_config, ensure_ascii=False), encoding="utf-8")

            # Launch subprocess using same Python executable
            log_file = job_dir / "subprocess.log"
            log_fh = open(log_file, "w")
            popen_kwargs = {
                "cwd": str(BASE_DIR),
                "stdout": log_fh,
                "stderr": log_fh,
            }
            if os.name == "nt":
                popen_kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW

            subprocess.Popen(
                [sys.executable, str(BASE_DIR / "run_batch.py"), str(job_file)],
                **popen_kwargs,
            )

            st.session_state.job_id = job_id
            st.success(f"Batch automation started! Job ID: {job_id}")
            time.sleep(2)
            st.rerun()

        # SHOW PROGRESS if a job is active
        if job_id:
            progress = read_progress(job_id)
            results = read_results(job_id)
            running = is_job_running(job_id)

            if progress:
                current = progress.get("current_index", 0)
                total = progress.get("total", 1)

                if running:
                    st.progress(current / total, text=f"Processing file {current} of {total}")

                st.subheader("Progress Log")
                log_entries = progress.get("log", [])
                recent = log_entries[-50:]
                for entry in recent:
                    if entry["status"] == "done":
                        icon = "✅"
                    elif entry["status"] == "running":
                        icon = "🔄"
                    elif entry["status"] == "filled":
                        icon = "📝"
                    elif entry["status"] == "warning":
                        icon = "⚠️"
                    else:
                        icon = "❌"
                    st.markdown(f"`{entry['time']}` {icon} **{entry['step']}** — {entry['detail']}")

                if len(log_entries) > 50:
                    st.caption(f"Showing last 50 of {len(log_entries)} log entries")

            if running:
                time.sleep(3)
                st.rerun()

            # RESULTS when done
            if results and not results.get("running", True):
                st.divider()
                st.subheader("Batch Results Summary")

                result_list = results.get("results", [])
                success_count = sum(1 for r in result_list if r["success"])
                fail_count = len(result_list) - success_count

                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total Files", len(result_list))
                with col2:
                    st.metric("Successful", success_count)
                with col3:
                    st.metric("Failed", fail_count)

                # Display results and collect output dirs
                for r in result_list:
                    if r["success"]:
                        out_dir = r.get("output_dir", "")
                        n_images = len(r.get("image_files", []))
                        n_gen = len(r.get("generated_files", []))
                        st.success(
                            f"**{r['market']}** (`{r['name']}`) — "
                            f"Report downloaded, {n_images} images, {n_gen} documents generated"
                        )
                    else:
                        # Even if web automation failed, show generated files
                        n_gen = len(r.get("generated_files", []))
                        n_images = len(r.get("image_files", []))
                        if n_gen > 0 or n_images > 0:
                            st.warning(
                                f"**{r['market']}** (`{r['name']}`) — "
                                f"Web automation failed ({r.get('error', 'Unknown')}), "
                                f"but {n_images} images + {n_gen} documents were generated"
                            )
                        else:
                            st.error(f"**{r['market']}** (`{r['name']}`) — Error: {r.get('error', 'Unknown')}")

                # AUTO-DOWNLOAD: Create ZIP per market with all outputs
                auto_dl_key = f"auto_downloaded_{job_id}"
                if not st.session_state.get(auto_dl_key, False):
                    # Check if there are any files to download
                    has_files = any(
                        r.get("output_dir") and os.path.isdir(r.get("output_dir", ""))
                        for r in result_list
                    )
                    if has_files:
                        st.session_state[auto_dl_key] = True

                        for r in result_list:
                            out_dir = r.get("output_dir", "")
                            if not out_dir or not os.path.isdir(out_dir):
                                continue

                            out_path = Path(out_dir)
                            output_files = list(out_path.iterdir())
                            if not output_files:
                                continue

                            market_safe = r.get("market", "output").replace(" ", "_")

                            if len(output_files) == 1:
                                # Single file — download directly
                                f = output_files[0]
                                with open(f, "rb") as fh:
                                    auto_download(fh.read(), f.name)
                            else:
                                # Multiple files — create ZIP
                                zip_buffer = BytesIO()
                                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                                    for f in output_files:
                                        if f.is_file():
                                            zf.write(str(f), f.name)
                                zip_buffer.seek(0)
                                auto_download(zip_buffer.read(), f"{market_safe}.zip")

                            time.sleep(0.5)

                        st.info("All output files auto-downloaded as ZIP packages!")

                # Manual download buttons
                st.divider()
                st.subheader("Download Individual Files")
                for r in result_list:
                    out_dir = r.get("output_dir", "")
                    if out_dir and os.path.isdir(out_dir):
                        with st.expander(f"Files for: {r.get('market', r['name'])}"):
                            for f in sorted(Path(out_dir).iterdir()):
                                if f.is_file():
                                    col1, col2 = st.columns([3, 1])
                                    with col1:
                                        st.markdown(f"**{f.name}** ({f.stat().st_size / 1024:.1f} KB)")
                                    with col2:
                                        if f.suffix in [".jpg", ".png"]:
                                            st.image(str(f), width=200)
                                        with open(f, "rb") as fh:
                                            st.download_button(
                                                "Download",
                                                data=fh.read(),
                                                file_name=f.name,
                                                key=f"dl_{r['name']}_{f.name}",
                                            )

                if st.button("Clear Results & Start New Batch"):
                    st.session_state.job_id = None
                    st.rerun()


# ══════════════════════════════════════════════
# TAB 4: View Outputs
# ══════════════════════════════════════════════
with tab4:
    st.header("Past Automation Outputs")

    if not OUTPUT_DIR.exists() or not any(OUTPUT_DIR.iterdir()):
        st.info("No outputs yet. Run an automation first.")
    else:
        markets = sorted([d.name for d in OUTPUT_DIR.iterdir() if d.is_dir()])

        if markets:
            selected_market = st.selectbox("Select Market", markets)
            market_dir = OUTPUT_DIR / selected_market

            runs = sorted([d.name for d in market_dir.iterdir() if d.is_dir()], reverse=True)

            if runs:
                selected_run = st.selectbox("Select Run", runs)
                run_dir = market_dir / selected_run

                st.subheader(f"Files from {selected_run}")
                files = list(run_dir.iterdir())

                for file in files:
                    col1, col2 = st.columns([3, 1])
                    with col1:
                        st.markdown(f"**{file.name}** ({file.stat().st_size / 1024:.1f} KB)")
                    with col2:
                        if file.suffix in [".png", ".jpg"]:
                            if st.button(f"View {file.name}", key=f"view_{file.name}"):
                                st.image(str(file))
                        else:
                            with open(file, "rb") as f:
                                st.download_button(
                                    "Download",
                                    data=f.read(),
                                    file_name=file.name,
                                    key=f"dl_{file.name}",
                                )
            else:
                st.info("No runs found for this market.")
        else:
            st.info("No market output folders found.")
