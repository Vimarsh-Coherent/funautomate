from pydantic import BaseModel


class SegmentData(BaseModel):
    name: str                    # e.g., "By Product Type"
    sub_segments: list[str]      # e.g., ["Acrylic Paints", "Watercolors", ...]
    dominating: str              # First/top sub-segment, e.g., "Acrylic Paints"


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
