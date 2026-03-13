import asyncio
from datetime import datetime
from pathlib import Path
from typing import Callable

from playwright.async_api import async_playwright, Page, Download

from core.models import MarketData


class AutomationResult:
    def __init__(self):
        self.steps: list[dict] = []
        self.downloaded_file: str | None = None
        self.error: str | None = None
        self.success: bool = False

    def add_step(self, step_id: str, status: str, detail: str = ""):
        self.steps.append({"step": step_id, "status": status, "detail": detail})


# ─── Bubble.io selector map (discovered from live inspection) ───
# These are standard <select> and <textarea> elements inside Bubble.
# Ordered by vertical position on page (y coordinate).
SELECTORS = {
    # Login page
    "login_email": 'input[type="email"]',
    "login_password": 'input[type="password"]',
    "login_button": 'button:has-text("Sign In")',

    # Top dropdown (Option 1/2/3)
    "option_dropdown": 'select.bubble-element.Dropdown >> nth=1',

    # Market Name textarea (only textarea with placeholder "Type here...")
    "market_name_input": 'textarea[placeholder="Type here..."]',

    # Updated Collateral dropdowns (Report Summary / CMI)
    # Select[5] y=327 — Report Summary/Description/Press Release
    "collateral_dd1": 'select.bubble-element.Dropdown >> nth=5',
    # Select[6] y=351 — CMI/CoherentMI
    "collateral_dd2": 'select.bubble-element.Dropdown >> nth=6',
    # Updated Collateral output button
    "collateral_output_btn": 'button.baTaJpaZ0',

    # Segment Writeup dropdown (Select[2])
    "segment_dropdown": 'select.bubble-element.Dropdown >> nth=2',
    "segment_output_btn": 'button.baTaJpt0',

    # Holistic Insights dropdown (Select[4])
    "holistic_dropdown": 'select.bubble-element.Dropdown >> nth=4',
    "holistic_output_btn": 'button.baTaJqaD0',

    # Additional Insights dropdown (Select[3])
    "additional_dropdown": 'select.bubble-element.Dropdown >> nth=3',
    "additional_output_btn": 'button.baTaJqaJ0',

    # DRO's Outline Output button
    "dro_output_btn": 'button.baTaOaNy',

    # Company Profile output button
    "company_output_btn": 'button.baTaJpc0',

    # Know Market Classification / Players / Events output buttons
    "classification_output_btn": 'button.baTaOaNaC',
    "players_output_btn": 'button.baTaOaNaI',
    "events_output_btn": 'button.baTaSuaC',
}


async def _login(page: Page, email: str, password: str, progress: Callable):
    """Handle login to MarketRytrAI."""
    progress("login", "running", f"Logging in as {email}")

    await page.goto("https://marketrytrai.com/login?land=login", wait_until="networkidle", timeout=120000)
    await page.wait_for_timeout(2000)

    await page.locator(SELECTORS["login_email"]).fill(email)
    await page.locator(SELECTORS["login_password"]).fill(password)
    await page.locator(SELECTORS["login_button"]).click()

    # Wait for redirect after login
    await page.wait_for_timeout(8000)
    progress("login", "done", "Logged in successfully")


async def _navigate_to_form(page: Page, app_url: str, progress: Callable):
    """Navigate to the app and click 'Data Driven Insights' to show the form."""
    progress("navigate", "running", f"Opening {app_url}")
    # Bubble.io apps load slowly — use domcontentloaded first, then wait for idle
    try:
        await page.goto(app_url, wait_until="networkidle", timeout=120000)
    except Exception:
        # Retry with looser wait condition
        progress("navigate", "running", f"Retrying navigation to {app_url}...")
        await page.goto(app_url, wait_until="domcontentloaded", timeout=120000)
        await page.wait_for_timeout(10000)
    await page.wait_for_timeout(3000)

    # Click "Data Driven Insights" to switch to the form view
    ddi = page.get_by_text("Data Driven Insights")
    if await ddi.count() > 0:
        await ddi.first.click()
        progress("navigate", "running", "Switched to Data Driven Insights")
        await page.wait_for_timeout(5000)

    progress("navigate", "done", "Form page loaded")


async def _get_modal_inputs_sorted(page: Page) -> list:
    """Get all visible inputs/textareas sorted by Y then X, skipping 'Type here...' (market name behind modal)."""
    result = await page.evaluate("""() => {
        const inputs = document.querySelectorAll('input, textarea');
        const visible = [];
        for (let idx = 0; idx < inputs.length; idx++) {
            const el = inputs[idx];
            if (el.offsetParent !== null) {
                const rect = el.getBoundingClientRect();
                visible.push({
                    idx: idx,
                    placeholder: el.placeholder || '',
                    y: rect.y,
                    x: rect.x
                });
            }
        }
        return visible.sort((a, b) => a.y - b.y || a.x - b.x);
    }""")

    # Build locator list, skipping the market name textarea
    inputs = []
    all_elements = page.locator("input, textarea")
    for item in result:
        if item["placeholder"] == "Type here...":
            continue
        inputs.append(all_elements.nth(item["idx"]))
    return inputs


async def _fill_modal(page: Page, data: MarketData, progress: Callable):
    """Fill modal: Market Size → Disease Name → select segment count → fill segments/drivers/regions → Submit."""
    progress("fill_modal", "running", "Waiting for modal to appear")
    await page.wait_for_timeout(3000)

    # Phase 1: Fill Market Size and Disease Name (first 2 inputs visible in modal)
    inputs = await _get_modal_inputs_sorted(page)
    progress("fill_modal", "running", f"Phase 1: Found {len(inputs)} inputs")

    if len(inputs) >= 1:
        await inputs[0].click()
        await inputs[0].fill(data.market_size_text)
        progress("fill_modal", "filled", f"Market Size & CAGR: {data.market_size_text[:60]}")

    if len(inputs) >= 2:
        await inputs[1].click()
        await inputs[1].fill(data.disease_name)
        progress("fill_modal", "filled", f"Disease Name: {data.disease_name}")

    # Phase 2: Select number of segments from dropdown (options: 1, 2, 3)
    num_segments = min(len(data.segments), 3)
    progress("fill_modal", "running", f"Selecting {num_segments} segments")

    segment_dd = page.locator("select:visible").filter(has_text="Select range")
    if await segment_dd.count() > 0:
        await segment_dd.first.select_option(str(num_segments))
    else:
        # Fallback: find by option values
        all_selects = page.locator("select:visible")
        count = await all_selects.count()
        for i in range(count):
            text = await all_selects.nth(i).text_content() or ""
            if "range" in text.lower() or "1-3" in text:
                await all_selects.nth(i).select_option(str(num_segments))
                break

    progress("fill_modal", "filled", f"Segment count: {num_segments}")
    await page.wait_for_timeout(5000)  # Wait for segment rows to appear

    # Phase 3: Re-collect inputs after segment rows appeared
    inputs = await _get_modal_inputs_sorted(page)
    progress("fill_modal", "running", f"Phase 2: Found {len(inputs)} inputs after segment selection")

    # Modal field order (verified from live inspection):
    # [0] Market Size & CAGR    (already filled)
    # [1] Disease Name           (already filled)
    # [2] Segment 1 Name
    # [3] Sub-Segment 1
    # [4] Dominating Segment 1
    # [5] Segment 2 Name
    # [6] Sub-Segment 2
    # [7] Dominating Segment 2
    # [8] Segment 3 Name
    # [9] Sub-Segment 3
    # [10] Dominating Segment 3
    # [11] Driver 1 Outline
    # [12] Restrain Outline       (same row as Driver 1)
    # [13] Driver 2 Outline
    # [14] Opportunity Outline    (same row as Driver 2)
    # [15] Dominating Region
    # [16] Fastest Growing Region

    field_map = []  # list of (index, name, value)

    # Segments
    for i in range(num_segments):
        seg = data.segments[i]
        base = 2 + (i * 3)
        field_map.append((base, f"Segment {i+1}", seg.name))
        field_map.append((base + 1, f"Sub-Segments {i+1}", ", ".join(seg.sub_segments)))
        field_map.append((base + 2, f"Dominating {i+1}", seg.dominating))

    # Drivers, Restrain, Opportunity (after segments)
    driver_base = 2 + (num_segments * 3)
    field_map.append((driver_base, "Driver 1", data.driver_1))
    field_map.append((driver_base + 1, "Restrain", data.restrain))
    field_map.append((driver_base + 2, "Driver 2", data.driver_2))
    field_map.append((driver_base + 3, "Opportunity", data.opportunity))

    # Regions
    region_base = driver_base + 4
    field_map.append((region_base, "Dominating Region", data.dominating_region))
    field_map.append((region_base + 1, "Fastest Growing Region", data.fastest_growing_region))

    # Fill each field
    for idx, name, value in field_map:
        if idx >= len(inputs):
            progress("fill_modal", "warning", f"Input index {idx} out of range ({len(inputs)} inputs), skipping {name}")
            continue
        if not value:
            continue
        try:
            await inputs[idx].click()
            await inputs[idx].fill(value)
            progress("fill_modal", "filled", f"{name}: {value[:60]}")
        except Exception as e:
            progress("fill_modal", "error", f"{name}: {e}")

    progress("fill_modal", "done", "All modal fields filled")


async def run_automation(
    market_data: MarketData,
    app_url: str,
    email: str,
    password: str,
    collateral_dd1: str = "Report Summary",
    collateral_dd2: str = "CMI",
    headless: bool = False,
    timeout_ms: int = 120000,
    output_dir: str = "outputs",
    on_progress: Callable | None = None,
) -> AutomationResult:
    """
    Main automation: login → navigate → fill form → submit → download report.
    """
    result = AutomationResult()

    def progress(step_id: str, status: str, detail: str = ""):
        result.add_step(step_id, status, detail)
        if on_progress:
            on_progress(step_id, status, detail)

    # Prepare output directory
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    safe_name = market_data.market_name.replace("/", "_").replace("\\", "_")
    download_dir = Path(output_dir) / safe_name / timestamp
    download_dir.mkdir(parents=True, exist_ok=True)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=headless)
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()
        page.set_default_timeout(timeout_ms)

        try:
            # Step 1: Login
            await _login(page, email, password, progress)

            # Step 2: Navigate to form
            await _navigate_to_form(page, app_url, progress)

            # Step 3: Enter Market Name
            progress("market_name", "running", market_data.market_name)
            market_input = page.locator(SELECTORS["market_name_input"]).first
            await market_input.click()
            await market_input.fill(market_data.market_name)
            await page.wait_for_timeout(500)
            progress("market_name", "done", market_data.market_name)

            # Step 4: Select Updated Collateral Dropdown 1 (Report Summary)
            progress("collateral_dd1", "running", collateral_dd1)
            dd1 = page.locator(SELECTORS["collateral_dd1"])
            await dd1.select_option(label=collateral_dd1)
            await page.wait_for_timeout(500)
            progress("collateral_dd1", "done", collateral_dd1)

            # Step 5: Select Updated Collateral Dropdown 2 (CMI)
            progress("collateral_dd2", "running", collateral_dd2)
            dd2 = page.locator(SELECTORS["collateral_dd2"])
            await dd2.select_option(label=collateral_dd2)
            await page.wait_for_timeout(500)
            progress("collateral_dd2", "done", collateral_dd2)

            # Step 6: Click "output" button for Updated Collateral
            progress("collateral_output", "running", "Clicking output button")
            await page.locator(SELECTORS["collateral_output_btn"]).click()
            await page.wait_for_timeout(5000)
            progress("collateral_output", "done", "Output button clicked, waiting for modal")

            # Step 7: Fill modal form
            await _fill_modal(page, market_data, progress)

            # Take screenshot before submit
            screenshot_path = str(download_dir / "before_submit.png")
            await page.screenshot(path=screenshot_path, full_page=True)
            progress("screenshot", "done", f"Saved screenshot: {screenshot_path}")

            # Step 8: Click Submit
            progress("submit", "running", "Clicking Submit")
            submit_btn = page.locator("button:has-text('Submit')").first
            await submit_btn.click()
            progress("submit", "done", "Submitted, waiting for popup")

            # Step 9: Wait for the "Continue" popup to appear
            # After submit, a popup shows "We've got the market scoop! Hit 'Continue' to generate Report Summary."
            progress("continue_wait", "running", "Waiting for Continue popup")
            continue_btn = page.locator("button.baTaOpaC, button:has-text('Continue')")

            # Poll for Continue button (appears after a few seconds)
            for attempt in range(60):  # up to 5 minutes
                await page.wait_for_timeout(5000)
                if await continue_btn.count() > 0 and await continue_btn.first.is_visible():
                    break
                progress("continue_wait", "running", f"Waiting... ({(attempt + 1) * 5}s)")

            # Click Continue with force=True to bypass any overlay image
            await continue_btn.first.click(force=True)
            progress("continue_wait", "done", "Continue clicked, generating report")

            # Step 10: Wait for report generation
            # A loading animation plays ("Sit back & relax, CMI is doing the heavy lifting.")
            # Then the download icon (button.baTaQaVaI with fa-download) appears in the top-right
            progress("generating", "running", "Report is being generated, please wait...")

            # Wait for the download icon to appear (fa-download SVG icon)
            download_icon = page.locator("button.baTaQaVaI")

            for attempt in range(60):  # up to 5 minutes
                await page.wait_for_timeout(5000)
                if await download_icon.count() > 0 and await download_icon.first.is_visible():
                    progress("generating", "done", f"Report ready after {(attempt + 1) * 5}s")
                    break
                progress("generating", "running", f"Generating... ({(attempt + 1) * 5}s)")

            await page.wait_for_timeout(2000)

            # Step 11: Click download icon
            progress("download", "running", "Clicking download icon")
            try:
                async with page.expect_download(timeout=60000) as download_info:
                    await download_icon.first.click(force=True)

                download: Download = await download_info.value
                file_name = download.suggested_filename or "report.docx"
                save_path = str(download_dir / file_name)
                await download.save_as(save_path)
                result.downloaded_file = save_path
                progress("download", "done", f"Report saved to {save_path}")

            except Exception:
                # Bubble.io might open download in new page or use JS-based download
                progress("download", "running", "Trying alternative download method...")
                try:
                    # Check if a new page opened
                    pages = page.context.pages
                    if len(pages) > 1:
                        new_page = pages[-1]
                        await new_page.wait_for_timeout(3000)
                        progress("download", "done", f"Report opened in new tab: {new_page.url}")
                    else:
                        # Click again and wait
                        await download_icon.first.click(force=True)
                        await page.wait_for_timeout(10000)
                        await page.screenshot(path=str(download_dir / "after_download.png"))
                        progress("download", "warning", "Download clicked. Check outputs folder.")
                except Exception as e2:
                    progress("download", "error", f"Download failed: {e2}")

            result.success = True

        except Exception as e:
            try:
                await page.screenshot(
                    path=str(download_dir / "error.png"), full_page=True
                )
            except Exception:
                pass
            result.error = str(e)
            progress("error", "failed", str(e))

        finally:
            await browser.close()

    return result


def run_automation_sync(
    market_data: MarketData,
    app_url: str,
    email: str = "",
    password: str = "",
    collateral_dd1: str = "Report Summary",
    collateral_dd2: str = "CMI",
    headless: bool = False,
    timeout_ms: int = 120000,
    output_dir: str = "outputs",
    on_progress: Callable | None = None,
) -> AutomationResult:
    """Synchronous wrapper for run_automation (for use from Streamlit)."""
    return asyncio.run(run_automation(
        market_data=market_data,
        app_url=app_url,
        email=email,
        password=password,
        collateral_dd1=collateral_dd1,
        collateral_dd2=collateral_dd2,
        headless=headless,
        timeout_ms=timeout_ms,
        output_dir=output_dir,
        on_progress=on_progress,
    ))
