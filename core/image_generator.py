"""Generate slide images directly using matplotlib + Pillow (no LibreOffice needed).

Creates JPG images matching the 5 PPTX slide layouts:
1. Regional Insights
2. Impact Analysis
3. Key Takeaways
4. Segmental Insights (with doughnut chart)
5. Market Key Players
"""

import logging
from pathlib import Path

import matplotlib
matplotlib.use("Agg")  # Non-interactive backend for server
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch
import numpy as np
from PIL import Image, ImageDraw, ImageFont

from core.models import ExtendedMarketData

logger = logging.getLogger(__name__)

# Brand colors
DARK_BLUE = "#173461"
GREEN = "#75C892"
LIGHT_BLUE = "#4A90D9"
WHITE = "#FFFFFF"
LIGHT_GRAY = "#F0F2F5"
DARK_GRAY = "#333333"
MID_GRAY = "#666666"
ORANGE = "#E8804C"
RED_ACCENT = "#D94F4F"

SLIDE_W, SLIDE_H = 1280, 720  # 16:9 aspect ratio
DPI = 150


def _fig_to_image(fig) -> Image.Image:
    """Convert matplotlib figure to PIL Image."""
    fig.canvas.draw()
    buf = fig.canvas.buffer_rgba()
    img = Image.frombuffer("RGBA", fig.canvas.get_width_height(), buf)
    plt.close(fig)
    return img.convert("RGB")


def _create_header_bar(draw, width, title: str, subtitle: str = ""):
    """Draw a dark blue header bar at the top."""
    draw.rectangle([0, 0, width, 80], fill=DARK_BLUE)
    try:
        font_title = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 24)
        font_sub = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 14)
    except (IOError, OSError):
        font_title = ImageFont.load_default()
        font_sub = ImageFont.load_default()
    draw.text((40, 20), title, fill=WHITE, font=font_title)
    if subtitle:
        draw.text((40, 52), subtitle, fill="#B0C4DE", font=font_sub)


def _get_font(size, bold=False):
    """Try to load a system font, fallback to default."""
    try:
        if bold:
            return ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", size)
        return ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", size)
    except (IOError, OSError):
        return ImageFont.load_default()


def generate_regional_insights(data: ExtendedMarketData, output_path: Path) -> Path:
    """Generate Regional Insights slide image."""
    img = Image.new("RGB", (SLIDE_W, SLIDE_H), WHITE)
    draw = ImageDraw.Draw(img)

    _create_header_bar(draw, SLIDE_W, "Regional Insights", f"{data.market_name}")

    font_large = _get_font(36, bold=True)
    font_medium = _get_font(18, bold=True)
    font_small = _get_font(14)
    font_value = _get_font(48, bold=True)

    # Dominating region highlight
    dom_region = data.dominating_region or "N/A"
    fast_region = data.fastest_growing_region or "N/A"

    # Market share percentage - big display
    pct_str = f"{data.market_share_pct:.1f}%"
    draw.text((80, 120), pct_str, fill=DARK_BLUE, font=font_value)
    draw.text((80, 180), f"{dom_region}", fill=DARK_BLUE, font=font_medium)
    draw.text((80, 210), f"Estimated Market Revenue Share, {data.base_year + 1}", fill=MID_GRAY, font=font_small)

    # Total market size
    draw.text((80, 260), f"Total Market Size: {data.market_size_value}", fill=DARK_GRAY, font=font_medium)

    # Region boxes
    regions_data = {r.name: r.status for r in data.regions}
    all_regions = ["North America", "Europe", "Asia Pacific", "Latin America", "Middle East", "Africa"]

    y_start = 320
    col1_x, col2_x = 80, 660
    box_w, box_h = 520, 55

    for i, region in enumerate(all_regions):
        x = col1_x if i < 3 else col2_x
        y = y_start + (i % 3) * 70
        status = regions_data.get(region, "")

        if status == "Dominating":
            bg_color = DARK_BLUE
            text_color = WHITE
            badge = "DOMINATING"
        elif status == "Fastest Growing":
            bg_color = GREEN
            text_color = WHITE
            badge = "FASTEST GROWING"
        else:
            bg_color = LIGHT_GRAY
            text_color = DARK_GRAY
            badge = ""

        draw.rounded_rectangle([x, y, x + box_w, y + box_h], radius=8, fill=bg_color)
        draw.text((x + 15, y + 15), region, fill=text_color, font=font_medium)
        if badge:
            badge_font = _get_font(10, bold=True)
            draw.text((x + box_w - 180, y + 20), badge, fill=text_color, font=badge_font)

    # Footer
    draw.rectangle([0, SLIDE_H - 40, SLIDE_W, SLIDE_H], fill=DARK_BLUE)
    footer_font = _get_font(11)
    draw.text((40, SLIDE_H - 30), f"Source: Coherent Market Insights | {data.base_year + 1}", fill="#B0C4DE", font=footer_font)

    img.save(str(output_path), "JPEG", quality=92)
    return output_path


def generate_impact_analysis(data: ExtendedMarketData, output_path: Path) -> Path:
    """Generate Impact Analysis slide image."""
    fig, ax = plt.subplots(figsize=(SLIDE_W / DPI, SLIDE_H / DPI), dpi=DPI)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 6)
    ax.axis("off")
    fig.patch.set_facecolor("white")

    # Title bar
    title_rect = FancyBboxPatch((0, 5.2), 10, 0.8, boxstyle="square,pad=0",
                                 facecolor=DARK_BLUE, edgecolor="none")
    ax.add_patch(title_rect)
    ax.text(0.3, 5.55, f"Impact Analysis of Key Factors | {data.market_name}",
            fontsize=9, color="white", fontweight="bold", va="center")

    # Categories
    categories = [
        ("Drivers", [(data.driver_1, data.driver_1_indicator),
                     (data.driver_2, data.driver_2_indicator)], DARK_BLUE),
        ("Restraints", [(data.restraint_1, data.restraint_1_indicator),
                        (data.restraint_2, data.restraint_2_indicator)], ORANGE),
        ("Opportunities", [(data.opportunity_1, data.opportunity_1_indicator),
                           (data.opportunity_2, data.opportunity_2_indicator)], GREEN),
    ]

    y_pos = 4.5
    for cat_name, items, color in categories:
        ax.text(0.2, y_pos, cat_name, fontsize=8, fontweight="bold", color=color, va="center")
        for j, (text, indicator) in enumerate(items):
            bar_y = y_pos - 0.4 - j * 0.55
            # Factor name
            display_text = (text[:45] + "...") if len(text) > 45 else text
            ax.text(0.3, bar_y + 0.12, display_text, fontsize=6, color=DARK_GRAY, va="center")
            # Impact bar background
            bar_rect = FancyBboxPatch((3.5, bar_y - 0.08), 6, 0.25, boxstyle="round,pad=0.02",
                                      facecolor=LIGHT_GRAY, edgecolor="#DDD")
            ax.add_patch(bar_rect)
            # Impact bar fill
            bar_width = max(0.3, (indicator / 10) * 6)
            fill_rect = FancyBboxPatch((3.5, bar_y - 0.08), bar_width, 0.25,
                                        boxstyle="round,pad=0.02",
                                        facecolor=color, edgecolor="none", alpha=0.85)
            ax.add_patch(fill_rect)
            ax.text(3.5 + bar_width + 0.15, bar_y + 0.05, f"{indicator}/10",
                    fontsize=6, color=color, fontweight="bold", va="center")

        y_pos -= 1.6

    # Footer
    footer_rect = FancyBboxPatch((0, -0.05), 10, 0.35, boxstyle="square,pad=0",
                                  facecolor=DARK_BLUE, edgecolor="none")
    ax.add_patch(footer_rect)
    ax.text(0.3, 0.12, f"Source: Coherent Market Insights | {data.base_year + 1}",
            fontsize=6, color="#B0C4DE")

    plt.tight_layout(pad=0)
    img = _fig_to_image(fig)
    img.save(str(output_path), "JPEG", quality=92)
    return output_path


def generate_key_takeaways(data: ExtendedMarketData, output_path: Path) -> Path:
    """Generate Key Takeaways slide image."""
    img = Image.new("RGB", (SLIDE_W, SLIDE_H), WHITE)
    draw = ImageDraw.Draw(img)

    _create_header_bar(draw, SLIDE_W, "Key Takeaways from Lead Analyst", data.market_name)

    font_body = _get_font(13)
    font_bullet = _get_font(13, bold=True)

    y = 110
    for i, takeaway in enumerate(data.takeaways or []):
        # Bullet point
        draw.ellipse([80, y + 4, 90, y + 14], fill=DARK_BLUE)
        # Wrap text
        words = takeaway.split()
        lines = []
        current_line = ""
        for word in words:
            test = f"{current_line} {word}".strip()
            if len(test) > 95:
                lines.append(current_line)
                current_line = word
            else:
                current_line = test
        if current_line:
            lines.append(current_line)

        for j, line in enumerate(lines):
            draw.text((105, y), line, fill=DARK_GRAY, font=font_body)
            y += 22
        y += 15

        if y > SLIDE_H - 80:
            break

    # Footer
    draw.rectangle([0, SLIDE_H - 40, SLIDE_W, SLIDE_H], fill=DARK_BLUE)
    footer_font = _get_font(11)
    draw.text((40, SLIDE_H - 30), f"Source: Coherent Market Insights | {data.base_year + 1}", fill="#B0C4DE", font=footer_font)

    img.save(str(output_path), "JPEG", quality=92)
    return output_path


def generate_segmental_insights(data: ExtendedMarketData, output_path: Path) -> Path:
    """Generate Segmental Insights slide image with doughnut chart."""
    chart_items = data.segment_chart_data
    if not chart_items:
        # Create placeholder image
        img = Image.new("RGB", (SLIDE_W, SLIDE_H), WHITE)
        draw = ImageDraw.Draw(img)
        _create_header_bar(draw, SLIDE_W, "Segmental Insights", "No segment data available")
        img.save(str(output_path), "JPEG", quality=92)
        return output_path

    fig = plt.figure(figsize=(SLIDE_W / DPI, SLIDE_H / DPI), dpi=DPI)
    fig.patch.set_facecolor("white")

    # Title area
    ax_title = fig.add_axes([0, 0.88, 1, 0.12])
    ax_title.axis("off")
    ax_title.set_facecolor(DARK_BLUE)
    rect = FancyBboxPatch((0, 0), 1, 1, boxstyle="square,pad=0",
                           facecolor=DARK_BLUE, edgecolor="none",
                           transform=ax_title.transAxes)
    ax_title.add_patch(rect)
    segment_type = data.segments[0].name if data.segments else "By Type"
    ax_title.text(0.03, 0.55, f"Segmental Insights | {data.market_name}",
                  fontsize=10, color="white", fontweight="bold",
                  va="center", transform=ax_title.transAxes)
    ax_title.text(0.03, 0.15, f"{segment_type}, {data.base_year + 1}",
                  fontsize=7, color="#B0C4DE", va="center", transform=ax_title.transAxes)

    # Doughnut chart
    ax_chart = fig.add_axes([0.05, 0.1, 0.5, 0.75])

    labels = [item.label for item in chart_items]
    values = [item.value for item in chart_items]

    colors_list = [DARK_BLUE, GREEN, LIGHT_BLUE, ORANGE, RED_ACCENT,
                   "#8B5CF6", "#EC4899", "#06B6D4", "#84CC16", "#F59E0B"]
    pie_colors = [colors_list[i % len(colors_list)] for i in range(len(values))]

    wedges, texts, autotexts = ax_chart.pie(
        values, labels=None, autopct="%1.1f%%",
        colors=pie_colors, pctdistance=0.78,
        wedgeprops=dict(width=0.4, edgecolor="white", linewidth=2),
        textprops=dict(fontsize=7, color="white", fontweight="bold"),
    )
    ax_chart.set_aspect("equal")

    # Center text
    largest = max(chart_items, key=lambda x: x.value)
    ax_chart.text(0, 0.05, f"{largest.value * 100:.1f}%", ha="center", va="center",
                  fontsize=20, fontweight="bold", color=DARK_BLUE)
    ax_chart.text(0, -0.12, largest.label[:20], ha="center", va="center",
                  fontsize=7, color=MID_GRAY)

    # Legend on the right side
    ax_legend = fig.add_axes([0.58, 0.15, 0.38, 0.7])
    ax_legend.axis("off")

    ax_legend.text(0, 0.97, "Market Breakdown", fontsize=9, fontweight="bold",
                   color=DARK_BLUE, va="top", transform=ax_legend.transAxes)

    for i, (label, value) in enumerate(zip(labels, values)):
        y_pos = 0.90 - i * 0.1
        if y_pos < 0:
            break
        color = pie_colors[i]
        ax_legend.add_patch(FancyBboxPatch((0, y_pos - 0.03), 0.06, 0.06,
                                            boxstyle="round,pad=0.01",
                                            facecolor=color, edgecolor="none",
                                            transform=ax_legend.transAxes))
        display_label = (label[:30] + "...") if len(label) > 30 else label
        ax_legend.text(0.1, y_pos, f"{display_label}  ({value * 100:.1f}%)",
                       fontsize=7, color=DARK_GRAY, va="center",
                       transform=ax_legend.transAxes)

    # Market size info
    ax_legend.text(0, 0.05, f"Total Market Size: {data.market_size_value}",
                   fontsize=7, color=MID_GRAY, va="center", transform=ax_legend.transAxes)

    # Footer
    ax_footer = fig.add_axes([0, 0, 1, 0.06])
    ax_footer.axis("off")
    rect = FancyBboxPatch((0, 0), 1, 1, boxstyle="square,pad=0",
                           facecolor=DARK_BLUE, edgecolor="none",
                           transform=ax_footer.transAxes)
    ax_footer.add_patch(rect)
    ax_footer.text(0.03, 0.4, f"Source: Coherent Market Insights | {data.base_year + 1}",
                   fontsize=6, color="#B0C4DE", transform=ax_footer.transAxes)

    img = _fig_to_image(fig)
    img.save(str(output_path), "JPEG", quality=92)
    return output_path


def generate_market_key_players(data: ExtendedMarketData, output_path: Path) -> Path:
    """Generate Market Key Players slide image."""
    img = Image.new("RGB", (SLIDE_W, SLIDE_H), WHITE)
    draw = ImageDraw.Draw(img)

    _create_header_bar(draw, SLIDE_W, "Market Key Players", data.market_name)

    font_company = _get_font(14)
    font_medium = _get_font(16, bold=True)
    font_small = _get_font(12)
    font_label = _get_font(11, bold=True)

    # Concentration meter
    draw.text((80, 100), "Market Concentration", fill=DARK_BLUE, font=font_medium)

    meter_x, meter_y = 80, 135
    meter_w, meter_h = 300, 25

    # Background bar
    draw.rounded_rectangle([meter_x, meter_y, meter_x + meter_w, meter_y + meter_h],
                           radius=12, fill=LIGHT_GRAY)

    # Gradient fill based on concentration
    conc = data.concentration_value
    fill_w = max(20, int((conc / 10) * meter_w))
    if conc <= 3:
        fill_color = GREEN
        label = "Low"
    elif conc <= 6:
        fill_color = ORANGE
        label = "Medium"
    else:
        fill_color = RED_ACCENT
        label = "High"

    draw.rounded_rectangle([meter_x, meter_y, meter_x + fill_w, meter_y + meter_h],
                           radius=12, fill=fill_color)
    draw.text((meter_x + meter_w + 15, meter_y + 3), f"{label} ({conc}/10)",
              fill=fill_color, font=font_label)

    # Scale labels
    draw.text((meter_x, meter_y + 30), "Low", fill=MID_GRAY, font=_get_font(9))
    draw.text((meter_x + meter_w - 25, meter_y + 30), "High", fill=MID_GRAY, font=_get_font(9))

    # Company grid
    draw.text((80, 200), "Key Market Players", fill=DARK_BLUE, font=font_medium)

    companies = data.companies or []
    cols = 3
    col_w = 370
    row_h = 50
    start_y = 240

    for i, company in enumerate(companies[:15]):  # Max 15 companies
        col = i % cols
        row = i // cols
        x = 80 + col * col_w
        y = start_y + row * row_h

        # Company card
        draw.rounded_rectangle([x, y, x + col_w - 20, y + row_h - 8],
                               radius=6, fill=LIGHT_GRAY, outline="#DDD")
        # Number badge
        badge_size = 24
        draw.ellipse([x + 8, y + 10, x + 8 + badge_size, y + 10 + badge_size],
                     fill=DARK_BLUE)
        num_font = _get_font(10, bold=True)
        draw.text((x + 14, y + 13), str(i + 1), fill=WHITE, font=num_font)
        # Company name
        display_name = (company[:35] + "...") if len(company) > 35 else company
        draw.text((x + 40, y + 15), display_name, fill=DARK_GRAY, font=font_company)

    # Footer
    draw.rectangle([0, SLIDE_H - 40, SLIDE_W, SLIDE_H], fill=DARK_BLUE)
    footer_font = _get_font(11)
    draw.text((40, SLIDE_H - 30), f"Source: Coherent Market Insights | {data.base_year + 1}",
              fill="#B0C4DE", font=footer_font)

    img.save(str(output_path), "JPEG", quality=92)
    return output_path


def generate_all_slide_images(data: ExtendedMarketData, output_dir: Path) -> list[Path]:
    """Generate all 5 slide images and return list of paths.

    Args:
        data: ExtendedMarketData with all fields populated
        output_dir: Directory to save JPG images

    Returns:
        List of paths to generated JPG files
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    generators = [
        ("Regional_Insights", generate_regional_insights),
        ("Impact_Analysis", generate_impact_analysis),
        ("Key_Takeaways", generate_key_takeaways),
        ("Segmental_Insights", generate_segmental_insights),
        ("Market_KeyPlayer", generate_market_key_players),
    ]

    images = []
    for name, gen_func in generators:
        img_path = output_dir / f"{name}.jpg"
        try:
            gen_func(data, img_path)
            images.append(img_path)
            logger.info(f"Generated {name}.jpg")
        except Exception as e:
            logger.error(f"Failed to generate {name}.jpg: {e}")

    return images
