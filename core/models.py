from pydantic import BaseModel


class SegmentData(BaseModel):
    name: str                    # e.g., "By Product Type"
    sub_segments: list[str]      # e.g., ["Acrylic Paints", "Watercolors", ...]
    dominating: str              # First/top sub-segment, e.g., "Acrylic Paints"


class RegionData(BaseModel):
    name: str                    # e.g., "North America"
    status: str = ""             # "Dominating", "Fastest Growing", or "-"/empty


class ChartItem(BaseModel):
    label: str                   # Segment name (cleaned, no ">" prefix)
    value: float                 # Market share as decimal (e.g., 0.455)


class MarketData(BaseModel):
    market_name: str             # e.g., "Global Art Materials Market"
    market_size_text: str        # Pre-formatted: "Market size in 2026 is USD 27.67 Bn..."
    disease_name: str = "NA"     # "NA" for non-healthcare markets
    segments: list[SegmentData]  # Dynamic list of segments (1, 2, 3, or more)
    driver_1: str = ""
    driver_2: str = ""
    restrain: str = ""
    opportunity: str = ""
    dominating_region: str = ""
    fastest_growing_region: str = ""


class ExtendedMarketData(MarketData):
    """Extended data model with all fields needed for PPTX/DOCX generation."""
    # Impact analysis indicators (0-10 scale)
    driver_1_indicator: int = 0
    driver_2_indicator: int = 0
    restraint_1: str = ""
    restraint_2: str = ""
    restraint_1_indicator: int = 0
    restraint_2_indicator: int = 0
    opportunity_1: str = ""
    opportunity_2: str = ""
    opportunity_1_indicator: int = 0
    opportunity_2_indicator: int = 0

    # Market metrics
    market_share_pct: float = 0.0     # D22 - decimal (e.g., 0.421)
    market_size_value: str = ""        # D7 - e.g., "USD 12.56 Bn"
    forecast_size_value: str = ""      # D8 - e.g., "USD 17.21 Bn"
    base_year: int = 2025              # D4
    historical_start: int = 2020       # D5
    forecast_end: int = 2033           # D6
    cagr_value: float = 0.0           # D9 - decimal
    currency_type: str = "USD"         # D10

    # Concentration and key players
    concentration_value: int = 0       # D31 - 0-10 scale
    companies: list[str] = []          # G3:G_last
    takeaways: list[str] = []          # D40:D44 (non-empty)

    # Regions with status
    regions: list[RegionData] = []     # C23:D28

    # Chart data (filtered parent-level segments only)
    segment_chart_data: list[ChartItem] = []

    # Geographic breakdown from Geographies sheet
    geographic_data: dict[str, list[str]] = {}  # {"North America": ["U.S.", "Canada"], ...}
