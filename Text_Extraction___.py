from concurrent.futures import ThreadPoolExecutor, as_completed
import copy
import difflib
import hashlib
import io
import json
import mimetypes
import os
import re
import struct
import threading
import time
from datetime import datetime
from pathlib import Path
from statistics import mean
from typing import Any, List, Literal, Optional, Tuple

from google import genai
from google.genai import types
from PIL import Image
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from pydantic import BaseModel, Field


# ---------------------------------------------------------------------------
# Project and Folder Configuration
# ---------------------------------------------------------------------------
PROJECT_ID = "gc-proj-aiml-dev-01fd"
LOCATION = "us-east4"
MODEL_ID = "gemini-2.5-flash"

TRENDENCE_FOLDER = Path("input_images") / "Trendence"
SPARTNASH_FOLDER = Path("input_images") / "spartnash"
OUTPUT_FOLDER = Path("comparison_results")
REPORTS_FOLDER = OUTPUT_FOLDER / "reports"
SKILL_PROMPT_PATH = Path(__file__).with_name("dashboard_image_comparator.prompt.md")
VALIDATOR_PROMPT_PATH = SKILL_PROMPT_PATH
MANIFEST_JSON = Path(os.getenv("MANIFEST_JSON", "input_images/manifest.json"))
OUTPUT_WORKBOOK = REPORTS_FOLDER / "dashboard_comparison_complete.xlsx"
OUTPUT_JSONL = OUTPUT_FOLDER / "dashboard_comparison_complete.jsonl"
OUTPUT_JSON = OUTPUT_FOLDER / "dashboard_comparison_complete.json"
TIMING_LOG = OUTPUT_FOLDER / "timing.log"
RAW_JSON_FOLDER = OUTPUT_FOLDER / "raw_json"
EXTRACTION_JSON_FOLDER = OUTPUT_FOLDER / "extraction_json"
TRENDENCE_EXTRACTION_FOLDER = EXTRACTION_JSON_FOLDER / "trendence"
SPARTNASH_EXTRACTION_FOLDER = EXTRACTION_JSON_FOLDER / "spartnash"
TRENDENCE_EXTRACTION_JSON = OUTPUT_FOLDER / "trendence_extractions.json"
SPARTNASH_EXTRACTION_JSON = OUTPUT_FOLDER / "spartnash_extractions.json"

FUZZY_MATCH_CUTOFF = 0.55

SECTION_ORDER = {
    "metadata": 0,
    "slicer": 1,
    "kpi": 2,
    "chart": 3,
    "treemap": 4,
    "table": 5,
    "layout": 6,
}

WEEKDAY_ORDER = {
    "sunday": 0,
    "monday": 1,
    "tuesday": 2,
    "wednesday": 3,
    "thursday": 4,
    "friday": 5,
    "saturday": 6,
}

os.environ["GOOGLE_CLOUD_PROJECT"] = PROJECT_ID
os.environ["GOOGLE_CLOUD_QUOTA_PROJECT"] = PROJECT_ID

client = genai.Client(
    vertexai=True,
    project=PROJECT_ID,
    location=LOCATION,
)
_prompt_token_count: Optional[int] = None
_timing_log_lock = threading.Lock()
_timing_run_id: Optional[str] = None


def start_timing_run() -> None:
    """Start a fresh timing report for one comparison run."""
    global _timing_run_id
    _timing_run_id = datetime.now().astimezone().isoformat(timespec="seconds")
    TIMING_LOG.parent.mkdir(parents=True, exist_ok=True)
    with _timing_log_lock:
        TIMING_LOG.write_text(
            f"run_id={_timing_run_id}\n[timing] run_started\n",
            encoding="utf-8",
        )
    print(f"[timing] run_started run_id={_timing_run_id}", flush=True)


def log_timing(message: str) -> None:
    """Write timing details to both the terminal and a persistent log file."""
    print(message, flush=True)
    try:
        TIMING_LOG.parent.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().astimezone().isoformat(timespec="seconds")
        with _timing_log_lock:
            with TIMING_LOG.open("a", encoding="utf-8") as log_file:
                run_prefix = f"run_id={_timing_run_id} " if _timing_run_id else ""
                log_file.write(f"{timestamp} {run_prefix}{message}\n")
    except OSError as exc:
        print(f"WARNING: could not write timing log: {exc}", flush=True)


# ---------------------------------------------------------------------------
# Pydantic Schemas for Structured Gemini Responses
# ---------------------------------------------------------------------------
class ExtractionItem(BaseModel):
    object_id: Optional[str] = Field(
        None, description="Stable identifier for this visible dashboard object"
    )
    selection: Optional[str] = Field(
        None,
        description="Routing group: selection_1 for charts, selection_2 for tables, selection_3 for KPI/UI/text",
    )
    bbox: Optional[List[float]] = Field(
        None,
        description="Normalized bounding box [ymin, xmin, ymax, xmax], each value from 0.0 to 1.0; null when not visibly determinable",
    )
    reading_order: Optional[int] = Field(
        None, description="Top-to-bottom, left-to-right visual reading order"
    )
    section: Literal["slicer", "data_label", "series_value", "chart"] = Field(
        description="Exactly one allowed semantic section: slicer, data_label, series_value, or chart"
    )
    item_name: str = Field(description="Visible label or concise descriptor of the item")
    visual_type: Optional[str] = Field(
        None,
        description=(
            "Optional descriptive metadata only; never use visual type classification to route or omit extraction"
        ),
    )
    value: Optional[str] = Field(None, description="Exact visible text or numerical value of the item")
    confidence: Optional[float] = Field(
        None, ge=0.0, le=1.0, description="Confidence score between 0.0 and 1.0"
    )
    background_color: str = Field(
        default="",
        description="Visible fill color when meaningful; empty string when unavailable or not meaningful",
    )
    kpi_title: Optional[str] = Field(
        None,
        description="Exact KPI card title text; populated only for kpi_card items",
    )
    kpi_value: Optional[str] = Field(
        None,
        description="Primary KPI display value as shown; populated only for kpi_card items",
    )
    series_name: Optional[str] = Field(
        None,
        description="Legend series name for a chart data row; null when the chart has no legend",
    )
    category: Optional[str] = Field(
        None,
        description="X-axis (or Y-axis) category label for a chart data row",
    )
    measure: Optional[str] = Field(
        None,
        description="Exact visible measure or unit associated with this bounded region",
    )
    percentage: Optional[float] = Field(
        None,
        description="Visible or derived percentage for a treemap tile",
    )
    scrollable: bool = Field(
        default=False,
        description="True when the visual shows a scrollbar or scroll affordance; otherwise False",
    )
    selected_options: List[str] = Field(
        default_factory=list,
        description="All currently selected options for a slicer/filter item; empty when none are selected",
    )
    box_x1: int = Field(default=0, description="Left edge in 0-1000 coordinate space; 0 for non-treemap")
    box_y1: int = Field(default=0, description="Top edge in 0-1000 coordinate space; 0 for non-treemap")
    box_x2: int = Field(default=0, description="Right edge in 0-1000 coordinate space; 0 for non-treemap")
    box_y2: int = Field(default=0, description="Bottom edge in 0-1000 coordinate space; 0 for non-treemap")


class DashboardExtraction(BaseModel):
    dashboard_title: Optional[str] = None
    page_header: Optional[str] = None
    active_page: Optional[str] = None
    refresh_date: Optional[str] = None
    items: List[ExtractionItem]
    notes: List[str] = Field(default_factory=list)


class DashboardSummaryFilter(BaseModel):
    filter_id: str = ""
    label: str = ""
    value: Optional[str] = None
    confidence: Optional[float] = Field(None, ge=0.0, le=1.0)


class DashboardSummaryVisual(BaseModel):
    visual_id: str = ""
    title: Optional[str] = None
    visual_type: Optional[str] = None
    bbox: Optional[List[float]] = Field(
        None,
        description="Normalized bounding box [ymin, xmin, ymax, xmax]",
    )
    reading_order: Optional[int] = None
    confidence: Optional[float] = Field(None, ge=0.0, le=1.0)


class DashboardSummary(BaseModel):
    dashboard_title: Optional[str] = None
    page_name: Optional[str] = None
    refresh_date: Optional[str] = None
    filters: List[DashboardSummaryFilter] = Field(default_factory=list)
    visuals: List[DashboardSummaryVisual] = Field(default_factory=list)


class FocusedVisualItem(BaseModel):
    section: Literal["slicer", "data_label", "series_value", "chart"] = Field(
        description="Exactly one allowed semantic section: slicer, data_label, series_value, or chart"
    )
    item_name: str = Field(description="Visible label or concise descriptor for this region or element")
    visual_type: Optional[str] = None
    value: Optional[str] = Field(None, description="Exact visible value, or category,value for a chart row")
    confidence: Optional[float] = Field(None, ge=0.0, le=1.0)
    series_name: Optional[str] = None
    category: Optional[str] = None
    measure: Optional[str] = None
    percentage: Optional[float] = None
    bbox: Optional[List[float]] = Field(
        None,
        description="Object bounding box [ymin, xmin, ymax, xmax] normalized to this crop; required when visible",
    )


class FocusedVisualExtraction(BaseModel):
    items: List[FocusedVisualItem] = Field(default_factory=list)


class VisualSeriesValue(BaseModel):
    name: Optional[str] = None
    value: Optional[str] = None
    percentage: Optional[float] = None
    confidence: Optional[float] = Field(None, ge=0.0, le=1.0)


class VisualCategory(BaseModel):
    name: Optional[str] = None
    series: List[VisualSeriesValue] = Field(default_factory=list)


class VisualElement(BaseModel):
    element_type: str = "element"
    label: Optional[str] = None
    value: Optional[str] = None
    category: Optional[str] = None
    series: Optional[str] = None
    percentage: Optional[float] = None
    bbox: Optional[List[float]] = None
    confidence: Optional[float] = Field(None, ge=0.0, le=1.0)


class VisualObject(BaseModel):
    object_id: Optional[str] = None
    object_type: str = "object"
    title: Optional[str] = None
    label: Optional[str] = None
    value: Optional[str] = None
    category: Optional[str] = None
    series: Optional[str] = None
    percentage: Optional[float] = None
    bbox: Optional[List[float]] = None
    color_token: Optional[str] = None
    legend_label: Optional[str] = None
    confidence: Optional[float] = Field(None, ge=0.0, le=1.0)


class VisualLegendBinding(BaseModel):
    label: Optional[str] = None
    color_token: Optional[str] = None
    confidence: Optional[float] = Field(None, ge=0.0, le=1.0)


class VisualDetailExtraction(BaseModel):
    visual_title: Optional[str] = None
    objects: List[VisualObject] = Field(
        description="Canonical bounded business objects; use this list for reportable extraction rows",
    )
    legend_bindings: List[VisualLegendBinding] = Field(default_factory=list)
    categories: List[VisualCategory] = Field(default_factory=list)
    elements: List[VisualElement] = Field(default_factory=list)
    notes: List[str] = Field(default_factory=list)


class TileBox(BaseModel):
    x1: int = Field(description="Left edge of tile, 0-1000 normalized")
    y1: int = Field(description="Top edge of tile, 0-1000 normalized")
    x2: int = Field(description="Right edge of tile, 0-1000 normalized")
    y2: int = Field(description="Bottom edge of tile, 0-1000 normalized")


class TreemapTileDetection(BaseModel):
    chart_title: Optional[str] = Field(
        None, description="Exact visible title/header text directly above the treemap chart"
    )
    measure: Optional[str] = Field(
        None, description="Exact visible measure represented by treemap area"
    )
    tiles: List[TileBox] = Field(
        default_factory=list,
        description="Bounding boxes for each distinct rectangular treemap tile, excluding navigation tabs and buttons",
    )


class TileReading(BaseModel):
    label: Optional[str] = Field(None, description="Exact visible label text inside this tile crop")
    value: Optional[str] = Field(None, description="Exact visible numerical value inside this tile crop, or null")
    percentage: Optional[float] = Field(
        None, description="Percentage visibly printed inside this tile, or null when absent"
    )
    confidence: Optional[float] = Field(None, ge=0.0, le=1.0)


class ValidationFinding(BaseModel):
    anomaly_id: Optional[str] = None
    element: Optional[str] = None
    category: Optional[str] = None
    status: Optional[str] = None
    reason: Optional[str] = None
    confidence: Optional[str] = None
    evidence: Optional[str] = None


class ReportValidation(BaseModel):
    status: str = "VALID_WITH_WARNINGS"
    agreement_score: Optional[float] = Field(None, ge=0.0, le=1.0)
    reported_similarity: Optional[float] = Field(None, ge=0.0, le=1.0)
    validated_similarity: Optional[float] = Field(None, ge=0.0, le=1.0)
    score_difference: Optional[float] = Field(None, ge=0.0, le=1.0)
    similarity_status: Optional[str] = None
    confirmed_anomalies: List[ValidationFinding] = Field(default_factory=list)
    disputed_anomalies: List[ValidationFinding] = Field(default_factory=list)
    false_positives: List[ValidationFinding] = Field(default_factory=list)
    false_negatives: List[ValidationFinding] = Field(default_factory=list)
    count_mismatches: List[str] = Field(default_factory=list)
    contradictions: List[str] = Field(default_factory=list)
    recommendations: List[str] = Field(default_factory=list)


def load_shared_skill_extraction_prompt() -> str:
    """Load the universal extraction contract from the canonical comparison prompt."""
    if not SKILL_PROMPT_PATH.is_file():
        return ""

    markdown = SKILL_PROMPT_PATH.read_text(encoding="utf-8")
    extraction_markers = (
        "## Universal Dashboard Visual Extraction System",
        "## Universal Visualization Extraction Contract",
    )
    extraction_positions = [markdown.find(marker) for marker in extraction_markers]
    extraction_positions = [position for position in extraction_positions if position >= 0]
    extraction_start = min(extraction_positions) if extraction_positions else -1
    workflow_start = markdown.find("## Core Workflow")
    if extraction_start < 0:
        return ""
    extraction_end = workflow_start if workflow_start > extraction_start else len(markdown)
    return markdown[extraction_start:extraction_end].strip()


def load_validator_prompt() -> str:
    """Load the canonical report-validator skill for pair-level orchestration."""
    if not VALIDATOR_PROMPT_PATH.is_file():
        return ""
    markdown = VALIDATOR_PROMPT_PATH.read_text(encoding="utf-8")
    frontmatter_end = markdown.find("---", 3)
    return markdown[frontmatter_end + 3 :].strip() if frontmatter_end >= 0 else markdown.strip()


_shared_skill_prompt = load_shared_skill_extraction_prompt()
if _shared_skill_prompt:
    EXTRACTION_PROMPT = _shared_skill_prompt


EXTRACTION_PROMPT_VERSION = "vision-first-v63-chart-series-binding-complete-categories"
EXTRACTION_MODELS = ("gemini-2.5-flash",)
# The full-dashboard pass can shift labels and values between small adjacent
# tiles, so treemaps receive one focused crop read after extraction.
ENABLE_TREEMAP_REFINEMENT = True
ENABLE_REPORT_VALIDATION = os.getenv("ENABLE_REPORT_VALIDATION", "0").lower() in {"1", "true", "yes"}
MAX_PARALLEL_EXTRACTIONS = max(1, int(os.getenv("MAX_PARALLEL_EXTRACTIONS", "2")))
MAX_PARALLEL_TREEMAP_TILES = max(1, int(os.getenv("MAX_PARALLEL_TREEMAP_TILES", "6")))
MAX_PARALLEL_VISUAL_DETAILS = max(1, int(os.getenv("MAX_PARALLEL_VISUAL_DETAILS", "2")))
_prompt_token_lock = threading.Lock()


# ---------------------------------------------------------------------------
# Helper Functions & Gemini API Calls
# ---------------------------------------------------------------------------
def image_mime_type(image_path: Path) -> str:
    mime_type, _ = mimetypes.guess_type(image_path.name)
    if mime_type not in {"image/png", "image/jpeg"}:
        raise ValueError(f"Unsupported image format for: {image_path.name}")
    return mime_type


def validate_comparison_report(
    result: dict,
    trendence_extraction: dict,
    spartnash_extraction: dict,
    trendence_image: Path,
    spartnash_image: Path,
) -> dict:
    """Validate one comparison pair against both images and structured extractions."""
    validator_prompt = load_validator_prompt()
    if not ENABLE_REPORT_VALIDATION or not validator_prompt:
        return {"status": "SKIPPED", "reason": "Report validation is disabled or unavailable."}

    evidence_payload = {
        "comparison_report": result,
        "baseline_extraction": trendence_extraction,
        "test_extraction": spartnash_extraction,
        "threshold": 0.95,
    }
    request_prompt = (
        f"{validator_prompt}\n\n"
        "### ORCHESTRATION INPUT\n"
        "Validate the comparison report below against the two supplied dashboard images and "
        "structured extractions. Return only JSON matching the ReportValidation schema. "
        "Images are the source of truth; do not invent findings.\n\n"
        f"{json.dumps(evidence_payload, ensure_ascii=False)}"
    )
    contents = [
        types.Part.from_bytes(data=trendence_image.read_bytes(), mime_type=image_mime_type(trendence_image)),
        types.Part.from_bytes(data=spartnash_image.read_bytes(), mime_type=image_mime_type(spartnash_image)),
        request_prompt,
    ]

    last_error: Optional[Exception] = None
    for model in ("gemini-2.5-flash",):
        try:
            response = client.models.generate_content(
                model=model,
                contents=contents,
                config=types.GenerateContentConfig(
                    temperature=0,
                    top_p=0,
                    top_k=1,
                    seed=42,
                    max_output_tokens=8192,
                    response_mime_type="application/json",
                    response_schema=ReportValidation,
                ),
            )
            if getattr(response, "parsed", None) is not None:
                validation = response.parsed
                return validation.model_dump() if hasattr(validation, "model_dump") else validation
            if response.text:
                return ReportValidation.model_validate_json(response.text).model_dump()
            last_error = ValueError("Validator returned an empty response")
        except Exception as exc:
            last_error = exc

    return {
        "status": "VALIDATION_UNAVAILABLE",
        "reason": f"{type(last_error).__name__}: {last_error}",
    }


def discover_numbered_images(folder: Path) -> dict[int, Path]:
    if not folder.is_dir():
        raise FileNotFoundError(f"Input folder does not exist: {folder}")

    images: dict[int, Path] = {}
    supported = {".png", ".jpg", ".jpeg"}
    for file_path in folder.iterdir():
        if file_path.is_file() and file_path.suffix.lower() in supported:
            if file_path.stem.isdigit():
                images[int(file_path.stem)] = file_path
    return images


def image_dimensions(image_path: Path) -> Tuple[int, int]:
    data = image_path.read_bytes()
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return struct.unpack(">II", data[16:24])
    if data.startswith(b"\xff\xd8"):
        index = 2
        while index + 9 < len(data):
            if data[index] != 0xFF:
                index += 1
                continue
            marker = data[index + 1]
            index += 2
            if marker in {0xD8, 0xD9}:
                continue
            segment_length = struct.unpack(">H", data[index:index + 2])[0]
            if marker in range(0xC0, 0xC4) or marker in range(0xC5, 0xC8) or marker in range(0xC9, 0xCC) or marker in range(0xCD, 0xD0):
                return struct.unpack(">HH", data[index + 3:index + 7])[::-1]
            index += segment_length
    raise ValueError(f"Invalid image format: {image_path.name}")


def image_content_hash(image_path: Path) -> str:
    return hashlib.sha256(image_path.read_bytes()).hexdigest()


def detect_treemap_tile_boxes(image_path: Path) -> Tuple[Optional[TreemapTileDetection], Any]:
    response = client.models.generate_content(
        model=MODEL_ID,
        contents=[
            types.Part.from_bytes(data=image_path.read_bytes(), mime_type=image_mime_type(image_path)),
            EXTRACTION_PROMPT,
        ],
        config=types.GenerateContentConfig(
            temperature=0,
            top_p=0,
            top_k=1,
            seed=42,
            response_mime_type="application/json",
            response_schema=TreemapTileDetection,
        ),
    )
    usage = getattr(response, "usage_metadata", None)
    if not response.text:
        return None, usage
    return TreemapTileDetection.model_validate_json(response.text), usage


def crop_tile_image(image_path: Path, box: TileBox, padding: int = 4) -> Optional[bytes]:
    with Image.open(image_path) as img:
        img = img.convert("RGB")
        width, height = img.size
        x1 = max(0, round(box.x1 / 1000 * width) - padding)
        y1 = max(0, round(box.y1 / 1000 * height) - padding)
        x2 = min(width, round(box.x2 / 1000 * width) + padding)
        y2 = min(height, round(box.y2 / 1000 * height) + padding)

        if x2 <= x1 or y2 <= y1:
            return None

        cropped = img.crop((x1, y1, x2, y2))
        scale = 3 if min(cropped.width, cropped.height) < 150 else 2
        cropped = cropped.resize((cropped.width * scale, cropped.height * scale), Image.LANCZOS)
        
        buffer = io.BytesIO()
        cropped.save(buffer, format="PNG")
        return buffer.getvalue()


def sample_tile_background_color(image_path: Path, box: TileBox) -> Optional[str]:
    """Sample color safely inside the tile interior to avoid borders and text."""
    with Image.open(image_path) as img:
        img = img.convert("RGB")
        width, height = img.size
        x1 = max(0, round(box.x1 / 1000 * width))
        y1 = max(0, round(box.y1 / 1000 * height))
        x2 = min(width, round(box.x2 / 1000 * width))
        y2 = min(height, round(box.y2 / 1000 * height))

        if x2 - x1 < 4 or y2 - y1 < 4:
            return None

        sample_x = x1 + max(2, int((x2 - x1) * 0.15))
        sample_y = y1 + max(2, int((y2 - y1) * 0.15))
        sample_x = min(sample_x, x2 - 1)
        sample_y = min(sample_y, y2 - 1)

        r, g, b = img.getpixel((sample_x, sample_y))
        return f"#{r:02X}{g:02X}{b:02X}"


def read_tile(crop_bytes: bytes) -> Tuple[Optional[TileReading], Any]:
    response = client.models.generate_content(
        model=MODEL_ID,
        contents=[
            types.Part.from_bytes(data=crop_bytes, mime_type="image/png"),
            EXTRACTION_PROMPT,
        ],
        config=types.GenerateContentConfig(
            temperature=0,
            top_p=0,
            top_k=1,
            seed=42,
            response_mime_type="application/json",
            response_schema=TileReading,
        ),
    )
    usage = getattr(response, "usage_metadata", None)
    if not response.text:
        return None, usage
    return TileReading.model_validate_json(response.text), usage


def normalize_treemap_items(items: list[dict]) -> list[dict]:
    normalized = []
    for item in items:
        item = dict(item)
        raw_value = item.get("value")
        if raw_value is None or str(raw_value).strip().casefold() in {"", "null", "none"}:
            item["value"] = "N/A"

        if normalize_label(item.get("section")) != "treemap":
            normalized.append(item)
            continue

        item_name = str(item.get("item_name") or "")
        separator = " - "
        if separator in item_name:
            chart_title, label = item_name.rsplit(separator, 1)
            prefix_match = re.match(r"^(\d[\d,]*)\s+(.+)$", label)
            if prefix_match:
                if not item.get("value"):
                    item["value"] = prefix_match.group(1)
                item["item_name"] = f"{chart_title}{separator}{prefix_match.group(2)}"
        normalized.append(item)
    return normalized


def refine_treemap_from_focus(image_number: int, image_path: Path, extraction: dict) -> int:
    """Re-read the treemap crop once so adjacent labels cannot borrow values."""
    if not ENABLE_TREEMAP_REFINEMENT:
        return 0

    original_items = [
        dict(item)
        for item in extraction.get("items", [])
        if normalize_label(item.get("section")) in {"chart", "treemap"}
    ]
    if not original_items:
        return 0

    def item_box(item: dict) -> tuple[int, int, int, int]:
        box = (
            int(item.get("box_x1") or 0),
            int(item.get("box_y1") or 0),
            int(item.get("box_x2") or 0),
            int(item.get("box_y2") or 0),
        )
        if box[2] > box[0] and box[3] > box[1]:
            return box
        bbox = item.get("bbox") or [0, 0, 0, 0]
        return int(bbox[1] * 1000), int(bbox[0] * 1000), int(bbox[3] * 1000), int(bbox[2] * 1000)

    boxes = [item_box(item) for item in original_items]
    min_x = min(box[0] for box in boxes)
    min_y = min(box[1] for box in boxes)
    max_x = max(box[2] for box in boxes)
    max_y = max(box[3] for box in boxes)
    if max_x <= min_x or max_y <= min_y:
        return 0

    with Image.open(image_path) as image:
        crop_box = (
            max(0, round((min_x - 300) * image.width / 1000)),
            max(0, round((min_y - 30) * image.height / 1000)),
            min(image.width, round((max_x + 30) * image.width / 1000)),
            min(image.height, round((max_y + 30) * image.height / 1000)),
        )
        if crop_box[2] <= crop_box[0] or crop_box[3] <= crop_box[1]:
            return 0
        crop_bytes = io.BytesIO()
        image.crop(crop_box).save(crop_bytes, format="PNG")

    focused_prompt = (
        f"{EXTRACTION_PROMPT}\n\n"
        "FOCUSED VISUAL READ: Inspect only the supplied crop. Return only the compact "
        "FocusedVisualExtraction JSON schema. For charts, emit one record per visible "
        "category and series with its exact value. For treemaps, emit one record per tile. "
        "Do not emit dashboard metadata, notes, insights, explanations, or duplicate "
        "unqualified chart summaries. Preserve visible zero values."
    )
    focused_started = time.perf_counter()
    response = client.models.generate_content(
        model=MODEL_ID,
        contents=[types.Part.from_bytes(data=crop_bytes.getvalue(), mime_type="image/png"), focused_prompt],
        config=types.GenerateContentConfig(
            temperature=0,
            top_p=0,
            top_k=1,
            response_mime_type="application/json",
            response_schema=FocusedVisualExtraction,
            max_output_tokens=8192,
        ),
    )
    log_timing(
        f"[timing] image={image_number} focused_treemap_seconds="
        f"{time.perf_counter() - focused_started:.2f}"
    )
    try:
        focused = (
            response.parsed.model_dump()
            if getattr(response, "parsed", None) is not None
            else json.loads(response.text)
        )
    except (json.JSONDecodeError, ValueError) as exc:
        retry_started = time.perf_counter()
        try:
            retry_response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=[types.Part.from_bytes(data=crop_bytes.getvalue(), mime_type="image/png"), focused_prompt],
                config=types.GenerateContentConfig(
                    temperature=0,
                    top_p=0,
                    top_k=1,
                    response_mime_type="application/json",
                    response_schema=FocusedVisualExtraction,
                    max_output_tokens=8192,
                ),
            )
            focused = (
                retry_response.parsed.model_dump()
                if getattr(retry_response, "parsed", None) is not None
                else json.loads(retry_response.text)
            )
            log_timing(
                f"[timing] image={image_number} focused_retry_seconds="
                f"{time.perf_counter() - retry_started:.2f} model=gemini-2.5-flash"
            )
        except (json.JSONDecodeError, ValueError) as retry_exc:
            log_timing(
                f"[timing] image={image_number} focused_treemap_parse_failed="
                f"{type(retry_exc).__name__} initial={type(exc).__name__}"
            )
            print(f"Image {image_number}: focused visual response could not be parsed: {retry_exc}")
            return 0
    focused_items = [
        dict(item)
        for item in focused.get("items", [])
        if normalize_label(item.get("section")) in {"chart", "treemap"}
    ]
    if not focused_items:
        return 0

    treemap_focus_items = [
        item for item in focused_items
        if normalize_label(item.get("section")) == "treemap"
    ]
    source_treemap = next(
        (item for item in original_items if normalize_label(item.get("section")) == "treemap"),
        None,
    )
    source_treemap_label = normalize_label(source_treemap.get("category")) if source_treemap else ""
    has_tile_rows = any(
        normalize_label(item.get("item_name")) != source_treemap_label
        or normalize_label(item.get("category")) != source_treemap_label
        for item in treemap_focus_items
    )
    if not has_tile_rows:
        if source_treemap:
            x1, y1, x2, y2 = item_box(source_treemap)
            with Image.open(image_path) as image:
                treemap_crop = io.BytesIO()
                image.crop(
                    (
                        max(0, round((x1 - 30) * image.width / 1000)),
                        max(0, round((y1 - 30) * image.height / 1000)),
                        min(image.width, round((x2 + 30) * image.width / 1000)),
                        min(image.height, round((y2 + 30) * image.height / 1000)),
                    )
                ).save(treemap_crop, format="PNG")
            treemap_started = time.perf_counter()
            try:
                treemap_response = client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=[
                        types.Part.from_bytes(data=treemap_crop.getvalue(), mime_type="image/png"),
                        f"{focused_prompt}\nTREEMAP ONLY: Return one record for every visible colored tile. "
                        "Never return one aggregate record for the whole treemap. For every tile, put "
                        "the tile's visible text label in item_name and category, and put only that tile's "
                        "exact numeric value in value. Do not put labels in value. Do not use the treemap "
                        "title or total as a tile label. Example: item_name='Benefits', category='Benefits', "
                        "value='174'.",
                    ],
                    config=types.GenerateContentConfig(
                        temperature=0,
                        top_p=0,
                        top_k=1,
                        response_mime_type="application/json",
                        response_schema=FocusedVisualExtraction,
                        max_output_tokens=8192,
                    ),
                )
                treemap_focus = (
                    treemap_response.parsed.model_dump()
                    if getattr(treemap_response, "parsed", None) is not None
                    else json.loads(treemap_response.text)
                )
                dedicated_treemap_items = [
                    dict(item)
                    for item in treemap_focus.get("items", [])
                    if normalize_label(item.get("section")) == "treemap"
                ]
                if any(
                    normalize_label(item.get("item_name")) != source_treemap_label
                    or normalize_label(item.get("category")) != source_treemap_label
                    for item in dedicated_treemap_items
                ):
                    focused_items = [
                        item for item in focused_items
                        if normalize_label(item.get("section")) != "treemap"
                    ] + dedicated_treemap_items
            except (json.JSONDecodeError, ValueError) as exc:
                log_timing(
                    f"[timing] image={image_number} treemap_focus_parse_failed={type(exc).__name__}"
                )
            log_timing(
                f"[timing] image={image_number} treemap_focus_seconds="
                f"{time.perf_counter() - treemap_started:.2f}"
            )

    focused_chart_items = [
        item for item in focused_items
        if normalize_label(item.get("section")) == "chart"
        and str(item.get("value") or "").strip().casefold() not in {"", "n/a", "null", "none"}
    ]
    original_chart_items = [
        item for item in original_items if normalize_label(item.get("section")) == "chart"
    ]
    refined_chart_items = []
    for focused_item in focused_chart_items:
        focused_name = str(focused_item.get("item_name") or "").strip()
        focused_title = focused_name.split(" - ", 1)[0]
        match = next(
            (
                item for item in original_chart_items
                if normalize_label(item.get("item_name")) == normalize_label(focused_title)
                or normalize_label(item.get("item_name")) == normalize_label(focused_name)
            ),
            None,
        )
        if match:
            focused_item["bbox"] = match.get("bbox")
            focused_item["visual_type"] = focused_item.get("visual_type") or match.get("visual_type")
        refined_chart_items.append(focused_item)

    def tile_label(item: dict) -> str:
        name = str(item.get("item_name") or item.get("category") or "").strip()
        return name.rsplit(" - ", 1)[-1].strip()

    refined_items = []
    for focused_item in focused_items:
        if normalize_label(focused_item.get("section")) != "treemap":
            continue
        label = tile_label(focused_item)
        source_treemap = next(
            (item for item in original_items if normalize_label(item.get("section")) == "treemap"),
            {},
        )
        tile = dict(focused_item)
        chart_title = str(source_treemap.get("item_name") or "Treemap").rsplit(" - ", 1)[0]
        tile["item_name"] = f"{chart_title} - {label}" if label else chart_title
        tile["category"] = label or tile.get("category")
        tile["bbox"] = source_treemap.get("bbox")
        tile["box_x1"] = source_treemap.get("box_x1", 0)
        tile["box_y1"] = source_treemap.get("box_y1", 0)
        tile["box_x2"] = source_treemap.get("box_x2", 0)
        tile["box_y2"] = source_treemap.get("box_y2", 0)
        refined_items.append(tile)

    if not refined_items and not refined_chart_items:
        return 0
    non_visual_items = [
        item for item in extraction.get("items", [])
        if normalize_label(item.get("section")) not in {"chart", "treemap"}
    ]
    original_treemap_items = [
        item for item in extraction.get("items", [])
        if normalize_label(item.get("section")) == "treemap"
    ]
    extraction["items"] = non_visual_items + refined_chart_items + (refined_items or original_treemap_items)
    return len(refined_chart_items) + len(refined_items)

def refine_treemap_items(image_number: int, image_path: Path, extraction: dict) -> int:
    """Detect and read treemap tiles independently of the dashboard pass."""
    if not ENABLE_TREEMAP_REFINEMENT:
        return 0

    detection, _ = detect_treemap_tile_boxes(image_path)
    if not detection or not detection.tiles:
        return 0

    unique_tiles = []
    seen_boxes = set()
    for tile in detection.tiles:
        coordinates = (tile.x1, tile.y1, tile.x2, tile.y2)
        if tile.x2 <= tile.x1 or tile.y2 <= tile.y1 or coordinates in seen_boxes:
            continue
        seen_boxes.add(coordinates)
        unique_tiles.append(tile)
    # Group nearby top edges into visual rows before reading color blocks left-to-right.
    # This prevents staggered rectangles beside a tall tile from being interleaved
    # solely because their individual y1 coordinates differ.
    heights = sorted(tile.y2 - tile.y1 for tile in unique_tiles)
    median_height = heights[len(heights) // 2] if heights else 0
    row_tolerance = max(10, min(60, round(median_height * 0.45)))
    rows = []
    for tile in sorted(unique_tiles, key=lambda candidate: (candidate.y1, candidate.x1)):
        matching_row = next(
            (row for row in rows if abs(tile.y1 - row["top"]) <= row_tolerance),
            None,
        )
        if matching_row is None:
            matching_row = {"top": tile.y1, "tiles": []}
            rows.append(matching_row)
        matching_row["tiles"].append(tile)
        matching_row["top"] = min(matching_row["top"], tile.y1)
    unique_tiles = [
        tile
        for row in sorted(rows, key=lambda row: row["top"])
        for tile in sorted(row["tiles"], key=lambda candidate: (candidate.x1, candidate.y1))
    ]

    def read_refined_tile(tile_index: int, tile: TileBox) -> Optional[dict]:
        crop_bytes = crop_tile_image(image_path, tile)
        if not crop_bytes:
            return None
        reading, _ = read_tile(crop_bytes)
        if not reading:
            return None

        label = reading.label.strip() if reading.label else None
        if label and re.fullmatch(r"[$€£]?\s*-?\d[\d,]*(?:\.\d+)?%?", label):
            label = None
        value = reading.value.strip() if reading.value else None
        if not label and not value:
            return None

        chart_title = (detection.chart_title or "").strip()
        item_name = f"{chart_title} - {label}" if chart_title and label else label or chart_title
        return {
            "object_id": f"image-{image_number}-treemap-tile-{tile_index:04d}",
            "selection": "selection_1",
            "bbox": [tile.y1 / 1000, tile.x1 / 1000, tile.y2 / 1000, tile.x2 / 1000],
            "reading_order": tile_index,
            "section": "Treemap",
            "item_name": item_name,
            "visual_type": "treemap",
            "value": value,
            "measure": detection.measure,
            "percentage": reading.percentage,
            "confidence": reading.confidence,
            "background_color": sample_tile_background_color(image_path, tile) or "",
            "category": label,
            "box_x1": tile.x1,
            "box_y1": tile.y1,
            "box_x2": tile.x2,
            "box_y2": tile.y2,
        }

    refined_items = []
    worker_count = min(MAX_PARALLEL_TREEMAP_TILES, len(unique_tiles))
    with ThreadPoolExecutor(max_workers=worker_count, thread_name_prefix="treemap-tile") as executor:
        futures = {
            executor.submit(read_refined_tile, tile_index, tile): tile_index
            for tile_index, tile in enumerate(unique_tiles, start=1)
        }
        for future in as_completed(futures):
            try:
                item = future.result()
            except Exception as exc:
                print(f"Treemap tile {futures[future]} failed: {type(exc).__name__}: {exc}")
                continue
            if item:
                refined_items.append(item)
    refined_items.sort(key=lambda item: item["reading_order"])

    if not refined_items:
        return 0

    visible_values = []
    for item in refined_items:
        numeric_value = _numeric_value(item.get("value"))
        if numeric_value is not None:
            visible_values.append(numeric_value)
    visible_total = sum(visible_values)
    if visible_total > 0:
        for item in refined_items:
            if item.get("percentage") is None:
                numeric_value = _numeric_value(item.get("value"))
                if numeric_value is not None:
                    item["percentage"] = round(numeric_value / visible_total * 100, 6)

    non_treemap_items = [
        item for item in extraction.get("items", [])
        if normalize_label(item.get("section")) != "treemap"
    ]
    extraction["items"] = non_treemap_items + refined_items
    return len(refined_items)


def _selection_for_item(item: dict) -> str:
    section = normalize_label(item.get("section"))
    if section in {"chart", "treemap"}:
        return "selection_1"
    if section in {"table", "matrix"}:
        return "selection_2"
    return "selection_3"


def build_structured_dashboard(extraction: dict, image_number: int) -> dict:
    """Add stable traceability fields and the requested dashboard tree."""
    normalized_items = []
    used_object_ids = set()
    for index, raw_item in enumerate(extraction.get("items", []), start=1):
        item = dict(raw_item)
        object_id = item.get("object_id") or f"image-{image_number}-object-{index:04d}"
        if object_id in used_object_ids:
            object_id = f"image-{image_number}-object-{index:04d}"
        used_object_ids.add(object_id)
        item["object_id"] = object_id
        item["selection"] = item.get("selection") or _selection_for_item(item)
        item["reading_order"] = index

        if not item.get("bbox"):
            coordinates = [item.get("box_y1"), item.get("box_x1"), item.get("box_y2"), item.get("box_x2")]
            if all(isinstance(value, (int, float)) for value in coordinates) and coordinates[2] > coordinates[0] and coordinates[3] > coordinates[1]:
                item["bbox"] = [round(max(0.0, min(1.0, value / 1000)), 6) for value in coordinates]
            else:
                item["bbox"] = None
        normalized_items.append(item)

    def in_selection(selection: str) -> list[dict]:
        return [item for item in normalized_items if item["selection"] == selection]

    metadata = {
        "dashboard_title": extraction.get("dashboard_title"),
        "page_header": extraction.get("page_header"),
        "active_page": extraction.get("active_page"),
        "refresh_date": extraction.get("refresh_date"),
    }
    visual_groups: dict[tuple[str, str], dict] = {}
    for item in normalized_items:
        visual_title = item.get("_visual_title") or item.get("item_name")
        visual_type = item.get("visual_type") or normalize_label(item.get("section")) or "unknown"
        visual_key = (normalize_label(visual_title), normalize_label(item.get("_visual_bbox") or item.get("bbox")))
        visual = visual_groups.setdefault(
            visual_key,
            {
                "id": item["object_id"],
                "type": visual_type,
                "title": visual_title,
                "position": {"bbox": item.get("_visual_bbox") or item.get("bbox"), "reading_order": item["reading_order"]},
                "categories": {},
                "elements": [],
                "confidence": item.get("confidence"),
                "rendering": {"blank": True, "broken": False, "overlapping": False},
            },
        )
        visual["rendering"]["blank"] = visual["rendering"]["blank"] and item.get("value") in (None, "") and item.get("kpi_value") in (None, "")
        visual["rendering"]["broken"] = visual["rendering"]["broken"] or bool(item.get("error"))
        if isinstance(item.get("confidence"), (int, float)):
            visual["confidence"] = item["confidence"] if visual["confidence"] is None else mean([visual["confidence"], item["confidence"]])

        category = item.get("category")
        series = item.get("series_name")
        if category and (series or normalize_label(item.get("section")) == "chart"):
            category_record = visual["categories"].setdefault(str(category), {"name": category, "series": []})
            category_record["series"].append({"name": series, "value": item.get("value"), "percentage": item.get("percentage")})
        else:
            visual["elements"].append({
                "element_type": normalize_label(item.get("section")) or "element",
                "label": item.get("item_name"),
                "value": item.get("value"),
                "percentage": item.get("percentage"),
                "bbox": item.get("bbox"),
            })

    visuals = []
    for visual in visual_groups.values():
        visual["data"] = {"categories": list(visual.pop("categories").values()), "elements": visual.pop("elements")}
        visuals.append(visual)
    return {
        "dashboard": {
            "title": metadata["dashboard_title"],
            "page": metadata["active_page"] or metadata["page_header"],
            "refresh_date": metadata["refresh_date"],
            "dashboard_type": "Power BI dashboard",
        },
        "filters": [
            {
                "name": item.get("item_name"),
                "type": item.get("visual_type") or "slicer",
                "value": item.get("selected_options") or item.get("value"),
                "selection_state": item.get("selection_state"),
                "confidence": item.get("confidence"),
            }
            for item in normalized_items
            if normalize_label(item.get("section")) == "slicer"
        ],
        "visuals": visuals,
        "metadata": metadata,
        "filters_legacy": [item for item in normalized_items if normalize_label(item.get("section")) == "slicer"],
        "kpis": [item for item in normalized_items if normalize_label(item.get("section")) == "kpi"],
        "charts": in_selection("selection_1"),
        "tables": in_selection("selection_2"),
        "ui_text": [item for item in in_selection("selection_3") if normalize_label(item.get("section")) not in {"slicer", "kpi"}],
        "layout": {
            "image_number": image_number,
            "width": extraction.get("_image_stats", {}).get("width"),
            "height": extraction.get("_image_stats", {}).get("height"),
            "regions": [
                {
                    "object_id": item["object_id"],
                    "object_type": item.get("visual_type") or normalize_label(item.get("section")) or "unknown",
                    "bbox": item.get("bbox"),
                    "confidence": item.get("confidence"),
                    "reading_order": item["reading_order"],
                }
                for item in normalized_items
            ],
        },
    }


def generate_flash_json(contents: list[Any], response_schema: Any, max_output_tokens: int) -> tuple[dict, Any]:
    request_contents = list(contents)
    for attempt in range(2):
        if attempt:
            request_contents = [
                *contents,
                "Return only valid JSON matching the supplied schema. Do not truncate strings, omit closing quotes, or add markdown.",
            ]
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=request_contents,
            config=types.GenerateContentConfig(
                temperature=0,
                top_p=0,
                top_k=1,
                seed=42,
                max_output_tokens=max_output_tokens,
                response_mime_type="application/json",
                response_schema=response_schema,
            ),
        )
        if getattr(response, "parsed", None) is not None:
            return response.parsed.model_dump(), getattr(response, "usage_metadata", None)
        if not response.text:
            if attempt == 0:
                continue
            raise ValueError("Gemini returned an empty JSON response.")
        try:
            return json.loads(response.text), getattr(response, "usage_metadata", None)
        except json.JSONDecodeError:
            if attempt == 1:
                raise


def extract_dashboard_summary(image_number: int, image_path: Path) -> tuple[dict, Any]:
    summary_prompt = """Analyze this complete dashboard screenshot as a collection of bounded visual containers.
Return only the compact DashboardSummary JSON schema.
Discover the dashboard title, page name, refresh date, visible filters, and one
inventory record per distinct bounded visual. A visual is a container such as a
KPI panel, chart, table, slicer, or composite region, not an isolated OCR label,
number, icon, axis mark, or legend entry.
For each visual return only visual_id, visible title when supported by the
container, optional descriptive type, normalized bbox [ymin,xmin,ymax,xmax],
reading order, and confidence. Do not extract chart points, table rows, KPI
values, pie slices, or treemap tile values at this stage. Do not include loose
OCR fragments, explanations, or duplicate containers."""
    return generate_flash_json(
        [
            types.Part.from_bytes(data=image_path.read_bytes(), mime_type=image_mime_type(image_path)),
            summary_prompt,
        ],
        DashboardSummary,
        8192,
    )


def extract_visual_detail(
    image: Image.Image,
    visual: dict,
    image_number: int,
    visual_index: int,
) -> tuple[list[dict], Any, list[str]]:
    bbox = visual.get("bbox") or [0, 0, 1, 1]
    crop_margin = 0.04
    crop_box = (
        max(0, round((bbox[1] - crop_margin) * image.width)),
        max(0, round((bbox[0] - crop_margin) * image.height)),
        min(image.width, round((bbox[3] + crop_margin) * image.width)),
        min(image.height, round((bbox[2] + crop_margin) * image.height)),
    )
    if crop_box[2] <= crop_box[0] or crop_box[3] <= crop_box[1]:
        return [], None, []
    crop_bytes = io.BytesIO()
    image.crop(crop_box).save(crop_bytes, format="PNG")
    detail_prompt = f"""Understand only visual {visual.get('visual_id') or f'VIS{visual_index:03d}'} from this crop.
Use vision-first schema extraction: understand the visual structure and business
objects before reading text. Do not run an OCR-style text dump. Return
VisualDetailExtraction JSON. First identify bounded business objects; do not
extract loose OCR fragments. Put every reportable object only in the `objects` list
and leave legacy `categories` and `elements` empty when `objects` is populated.
Each object must keep its own boundary, label, value, category/series relationship,
and confidence. Never associate text across card, tile, slice, bar, row, legend, or
other object boundaries.

For KPI regions, treat every bordered card as an independent object containing its
title and primary numeric value. Do not reverse labels and values; nearby icons are
decorative unless they contain text. For circular charts, identify each slice,
record its color token, displayed value and percentage, then bind it to a legend
label only when the slice color visibly corresponds to that legend color. Return
separate `legend_bindings` with one `{{label, color_token}}` entry per visible legend
swatch when needed. Use color association before proximity or list order. If the
binding is unclear, keep legend_label null and lower confidence.

For multi-series charts, return one object per category/series value. For treemaps,
return one object per tile. For tables and matrices, keep each row/cell label and
value in the same object. Never flatten a chart into a comma-separated value string.
For bar and column charts, include every visible axis category in reading order,
including categories with zero-height bars or a displayed value of zero. Bind each
value to its exact category and series label, and verify that every visible category
has been represented before returning. Never emit a category label as a data value.
Preserve exact visible text, including square brackets and parentheses such as
`[All]`, `(Blank)`, and `(>40 days)`, as well as zero values, N/A, blank,
and unreadable values. Do not strip delimiters or skip text because it
looks like cache/debug content.
Do not return logos, brand marks, watermarks, or decorative text as objects. Use
null for text or values that cannot be read; do not create a truncation or text-loss
record.
Use the same reading order and naming decisions on every run; never silently
drop important text or invent missing characters.
Do not infer values or return content outside this crop.
Visible title context: {visual.get('title') or 'unknown'}
"""
    detail_contents = [
        types.Part.from_bytes(data=crop_bytes.getvalue(), mime_type="image/png"),
        detail_prompt,
    ]
    try:
        detail, usage = generate_flash_json(detail_contents, VisualDetailExtraction, 8192)
    except (json.JSONDecodeError, ValueError) as detail_error:
        fallback_prompt = """Return compact valid JSON using the supplied FocusedVisualExtraction schema.
Extract every visible business label and value in this crop, preserving exact
brackets, parentheses, zeros, percentages, and N/A text. Keep one item per
    bounded object and include that object's crop-relative bbox whenever visible.
    Read the complete label from the supplied crop. Do not emit an isolated OCR
    fragment such as `Pri`, `eek`, `ate`, `Ag`, or `didates` when it is only a
    partial piece of a larger label. Do not create a second row for the same object.
    Do not include logos,
    watermarks, decorative text, explanations, markdown, or guessed characters.
    This is a fallback after a larger response failed, so prioritize complete valid
    JSON over commentary."""
        fallback, usage = generate_flash_json(
            [*detail_contents, fallback_prompt],
            FocusedVisualExtraction,
            4096,
        )

        def crop_bbox_to_image(item_bbox: object) -> Optional[list[float]]:
            if not isinstance(item_bbox, list) or len(item_bbox) != 4:
                return None
            try:
                crop_ymin, crop_xmin, crop_ymax, crop_xmax = [float(value) for value in item_bbox]
                crop_width = max(1, crop_box[2] - crop_box[0])
                crop_height = max(1, crop_box[3] - crop_box[1])
                return [
                    max(0.0, min(1.0, (crop_box[1] + crop_ymin * crop_height) / image.height)),
                    max(0.0, min(1.0, (crop_box[0] + crop_xmin * crop_width) / image.width)),
                    max(0.0, min(1.0, (crop_box[1] + crop_ymax * crop_height) / image.height)),
                    max(0.0, min(1.0, (crop_box[0] + crop_xmax * crop_width) / image.width)),
                ]
            except (TypeError, ValueError):
                return None

        detail = {
            "visual_title": visual.get("title") or visual.get("visual_id"),
            "objects": [
                {
                    "object_type": item.get("visual_type") or "object",
                    "label": item.get("item_name"),
                    "value": item.get("value"),
                    "category": item.get("category"),
                    "series": item.get("series_name"),
                    "percentage": item.get("percentage"),
                    "confidence": item.get("confidence"),
                    "bbox": crop_bbox_to_image(item.get("bbox")),
                }
                for item in fallback.get("items", [])
            ],
            "legend_bindings": [],
        }
        fallback_note = (
            f"Visual {visual.get('visual_id') or visual_index} used compact fallback extraction "
            f"after invalid detail JSON: {type(detail_error).__name__}"
        )
    else:
        fallback_note = None
    items = []
    visual_title = detail.get("visual_title") or visual.get("title") or visual.get("visual_id")
    text_notes: list[str] = []
    if fallback_note:
        text_notes.append(fallback_note)
    legend_by_color = {
        normalize_color_token(binding.get("color_token")): binding.get("label")
        for binding in detail.get("legend_bindings", [])
        if normalize_color_token(binding.get("color_token")) and binding.get("label")
    }

    def add_item(item: dict, element_index: int) -> None:
        item = dict(item)
        object_kind = normalize_label(item.get("object_type") or item.get("visual_type"))
        if object_kind in {"logo", "brand_logo", "watermark", "decorative_logo"}:
            return
        item["object_id"] = f"image-{image_number}-visual-{visual_index:04d}-element-{element_index:04d}"
        item["reading_order"] = item.get("reading_order") or element_index
        item["bbox"] = item.get("bbox") or visual.get("bbox")
        item["_visual_title"] = visual_title
        item["_visual_bbox"] = visual.get("bbox")
        item["item_name"] = item.get("item_name") or visual_title or visual.get("visual_id")
        items.append(item)

    element_index = 1
    canonical_objects = detail.get("objects", [])
    for business_object in canonical_objects:
        business_object = dict(business_object)
        object_type = normalize_label(business_object.get("object_type"))
        bound_legend_label = legend_by_color.get(
            normalize_color_token(business_object.get("color_token"))
        )
        object_label = (
            business_object.get("legend_label")
            or bound_legend_label
            or business_object.get("label")
            or business_object.get("category")
        )
        object_category = business_object.get("category") or object_label
        object_series = business_object.get("series")
        object_value = business_object.get("value")
        item_name = (
            f"{object_series}-{object_category}"
            if object_series and object_category
            else f"{visual_title}-{object_label}"
            if object_label
            else visual_title
        )
        add_item(
            {
                "section": "series_value" if object_series else "data_label",
                "item_name": item_name,
                "visual_type": visual.get("visual_type"),
                "object_type": object_type or "object",
                "category": object_category,
                "series_name": object_series,
                "value": object_value,
                "percentage": business_object.get("percentage"),
                "confidence": business_object.get("confidence"),
                "attribute": business_object.get("color_token"),
                "legend_label": business_object.get("legend_label") or bound_legend_label,
                "bbox": business_object.get("bbox"),
            },
            element_index,
        )
        element_index += 1

    if canonical_objects:
        return items, usage, text_notes

    for category in detail.get("categories", []):
        category_name = category.get("name")
        for series in category.get("series", []):
            series_name = series.get("name")
            item_name = (
                f"{series_name}-{category_name}"
                if series_name and category_name
                else f"{visual_title}-{category_name}"
                if category_name
                else series_name or visual_title
            )
            add_item(
                {
                    "section": "series_value" if series_name else "data_label",
                    "item_name": item_name,
                    "visual_type": visual.get("visual_type"),
                    "category": category_name,
                    "series_name": series_name,
                    "value": series.get("value"),
                    "percentage": series.get("percentage"),
                    "confidence": series.get("confidence"),
                },
                element_index,
            )
            element_index += 1

    for element in detail.get("elements", []):
        element = dict(element)
        element_type = normalize_label(element.get("element_type"))
        element_value = element.get("value")
        element_category = element.get("category")
        element_series = element.get("series")
        if element_value is not None and (element_category or element_series):
            element["section"] = "series_value" if element_series else "data_label"
            element["item_name"] = (
                f"{element_series}-{element_category}"
                if element_series and element_category
                else f"{visual_title}-{element_category}"
                if element_category
                else element_series or visual_title
            )
        elif element_value is not None:
            element["section"] = "data_label"
            element["item_name"] = visual_title
        else:
            # Keep value-less visual metadata out of the atomic value sections.
            element["section"] = "chart"
            element["item_name"] = (
                f"{visual_title}-{element.get('label')}"
                if element.get("label")
                else visual_title
            )
        element["series_name"] = element.get("series")
        add_item(element, element_index)
        element_index += 1
    return items, usage, text_notes


def extract_dashboard_image_legacy(image_number: int, image_path: Path) -> dict:
    global _prompt_token_count
    image_started = time.perf_counter()
    preparation_started = time.perf_counter()
    print(f"Processing image {image_number}: {image_path.name}")

    width, height = image_dimensions(image_path)
    if _prompt_token_count is None:
        with _prompt_token_lock:
            if _prompt_token_count is None:
                token_count_started = time.perf_counter()
                prompt_usage = client.models.count_tokens(model=MODEL_ID, contents=EXTRACTION_PROMPT)
                _prompt_token_count = prompt_usage.total_tokens
                log_timing(
                    f"[timing] prompt_token_count_seconds={time.perf_counter() - token_count_started:.2f} "
                    f"prompt_tokens={_prompt_token_count}"
                )

    request_contents = [
        types.Part.from_bytes(data=image_path.read_bytes(), mime_type=image_mime_type(image_path)),
        EXTRACTION_PROMPT,
    ]
    log_timing(
        f"[timing] image={image_number} preparation_seconds="
        f"{time.perf_counter() - preparation_started:.2f} width={width} height={height}"
    )
    request_config = types.GenerateContentConfig(
        temperature=0,
        top_p=0,
        top_k=1,
        seed=42,
        max_output_tokens=32768,
        response_mime_type="application/json",
        response_schema=DashboardExtraction,
    )
    response = None
    usage = None
    last_error = None
    llm_calls = 0
    for model_index, model in enumerate(EXTRACTION_MODELS):
        contents = request_contents
        config = request_config if model_index == 0 else types.GenerateContentConfig(
            temperature=0,
            top_p=0,
            top_k=1,
            response_mime_type="application/json",
            response_schema=DashboardExtraction,
            max_output_tokens=32768,
        )
        try:
            model_started = time.perf_counter()
            print(f"Processing image {image_number} using {model}")
            response = client.models.generate_content(
                model=model,
                contents=contents,
                config=config,
            )
            log_timing(
                f"[timing] image={image_number} model={model} "
                f"model_seconds={time.perf_counter() - model_started:.2f}"
            )
            llm_calls += 1
            candidates = getattr(response, "candidates", None) or []
            finish_reason = getattr(candidates[0], "finish_reason", None) if candidates else None
            if str(finish_reason).endswith("MAX_TOKENS"):
                last_error = ValueError(f"{model} returned MAX_TOKENS")
                print(f"Image {image_number}: {model} hit MAX_TOKENS; trying fallback")
                continue
            if getattr(response, "parsed", None) is not None:
                extraction = response.parsed.model_dump() if hasattr(response.parsed, "model_dump") else response.parsed
                print(f"Image {image_number}: success using {model}")
                usage = getattr(response, "usage_metadata", None)
                break
            if not response.text:
                last_error = ValueError("Gemini returned an empty extraction response.")
                continue
            extraction = json.loads(response.text)
            print(f"Image {image_number}: success using {model}")
            usage = getattr(response, "usage_metadata", None)
            break
        except (json.JSONDecodeError, ValueError) as exc:
            last_error = exc
            log_timing(
                f"[timing] image={image_number} model={model} "
                f"failed_seconds={time.perf_counter() - model_started:.2f}"
            )
            print(f"Image {image_number}: {model} failed: {exc}; trying fallback")
        except Exception as exc:
            last_error = exc
            log_timing(
                f"[timing] image={image_number} model={model} "
                f"error_seconds={time.perf_counter() - model_started:.2f}"
            )
            print(f"Image {image_number}: {model} API error: {exc}; trying fallback")
    else:
        raise ValueError(f"Gemini extraction failed for image {image_number}: {last_error}") from last_error

    refined_treemap_count = refine_treemap_from_focus(image_number, image_path, extraction)
    extraction["items"] = normalize_treemap_items(extraction.get("items", []))
    if refined_treemap_count:
        extraction.setdefault("notes", []).append(
            f"Treemap values read from {refined_treemap_count} isolated tiles."
        )

    p_tokens = getattr(usage, "prompt_token_count", 0) or 0
    c_tokens = getattr(usage, "candidates_token_count", 0) or 0
    t_tokens = getattr(usage, "total_token_count", 0) or 0

    extraction["_image_stats"] = {
        "width": width,
        "height": height,
        "pixels": width * height,
        "llm_calls": llm_calls,
        "prompt_tokens": p_tokens,
        "prompt_text_tokens": _prompt_token_count,
        "image_input_tokens": max(0, p_tokens - (_prompt_token_count or 0)),
        "output_tokens": c_tokens,
        "total_tokens": t_tokens,
        "image_hash": image_content_hash(image_path),
    }
    extraction["structured_dashboard"] = build_structured_dashboard(extraction, image_number)
    extraction["_image_stats"]["elapsed_seconds"] = round(time.perf_counter() - image_started, 3)
    log_timing(
        f"[timing] image={image_number} total_seconds={extraction['_image_stats']['elapsed_seconds']:.2f} "
        f"llm_calls={llm_calls}"
    )
    return extraction


def extract_dashboard_image(image_number: int, image_path: Path) -> dict:
    image_started = time.perf_counter()
    width, height = image_dimensions(image_path)
    print(f"Processing image {image_number}: {image_path.name}")
    summary_started = time.perf_counter()
    summary, summary_usage = extract_dashboard_summary(image_number, image_path)
    log_timing(
        f"[timing] image={image_number} summary_seconds={time.perf_counter() - summary_started:.2f} "
        f"visuals={len(summary.get('visuals', []))}"
    )

    items = []
    notes = ["Extraction used a compact dashboard inventory followed by per-visual detail reads."]
    total_prompt_tokens = getattr(summary_usage, "prompt_token_count", 0) or 0
    total_output_tokens = getattr(summary_usage, "candidates_token_count", 0) or 0
    total_tokens = getattr(summary_usage, "total_token_count", 0) or 0
    visuals = list(summary.get("visuals", []))

    def extract_visual(index_and_visual: tuple[int, dict]) -> tuple[int, dict, list[dict], Any, list[str], Optional[Exception], float]:
        visual_index, visual = index_and_visual
        detail_started = time.perf_counter()
        try:
            with Image.open(image_path) as image:
                visual_items, usage, text_notes = extract_visual_detail(
                    image, visual, image_number, visual_index
                )
            return visual_index, visual, visual_items, usage, text_notes, None, time.perf_counter() - detail_started
        except Exception as exc:
            return visual_index, visual, [], None, [], exc, time.perf_counter() - detail_started

    detail_results: list[tuple[int, dict, list[dict], Any, list[str], Optional[Exception], float]] = []
    if visuals:
        worker_count = min(MAX_PARALLEL_VISUAL_DETAILS, len(visuals))
        with ThreadPoolExecutor(max_workers=worker_count, thread_name_prefix="visual-detail") as executor:
            futures = [
                executor.submit(extract_visual, (visual_index, visual))
                for visual_index, visual in enumerate(visuals, start=1)
            ]
            detail_results = [future.result() for future in futures]

    for visual_index, visual, visual_items, usage, text_notes, error, elapsed in detail_results:
        if error is not None:
            notes.append(
                f"Visual {visual.get('visual_id') or visual_index} detail extraction failed: "
                f"{type(error).__name__}: {error}"
            )
        else:
            items.extend(visual_items)
            notes.extend(text_notes)
            total_prompt_tokens += getattr(usage, "prompt_token_count", 0) or 0
            total_output_tokens += getattr(usage, "candidates_token_count", 0) or 0
            total_tokens += getattr(usage, "total_token_count", 0) or 0
        log_timing(
            f"[timing] image={image_number} visual={visual.get('visual_id') or visual_index} "
            f"detail_seconds={elapsed:.2f}"
        )

    for filter_index, filter_item in enumerate(summary.get("filters", []), start=1):
        items.append(
            {
                "object_id": filter_item.get("filter_id") or f"image-{image_number}-filter-{filter_index:04d}",
                "selection": "selection_3",
                "section": "Slicer",
                "item_name": filter_item.get("label") or f"Filter {filter_index}",
                "visual_type": "filter",
                "value": filter_item.get("value"),
                "confidence": filter_item.get("confidence"),
                "selected_options": [filter_item["value"]] if filter_item.get("value") else [],
                "reading_order": len(items) + 1,
            }
        )

    extraction = {
        "dashboard_title": summary.get("dashboard_title"),
        "page_header": summary.get("page_name"),
        "active_page": summary.get("page_name"),
        "refresh_date": summary.get("refresh_date"),
        "items": normalize_treemap_items(items),
        "notes": notes,
        "_image_stats": {
            "width": width,
            "height": height,
            "pixels": width * height,
            "llm_calls": 1 + len(summary.get("visuals", [])),
            "prompt_tokens": total_prompt_tokens,
            "prompt_text_tokens": 0,
            "image_input_tokens": 0,
            "output_tokens": total_output_tokens,
            "total_tokens": total_tokens,
            "image_hash": image_content_hash(image_path),
        },
    }
    extraction["structured_dashboard"] = build_structured_dashboard(extraction, image_number)
    extraction["_image_stats"]["elapsed_seconds"] = round(time.perf_counter() - image_started, 3)
    log_timing(
        f"[timing] image={image_number} total_seconds={extraction['_image_stats']['elapsed_seconds']:.2f} "
        f"llm_calls={extraction['_image_stats']['llm_calls']}"
    )
    return extraction


def extract_all_images(folder: Path, cache_folder: Path, images: dict[int, Path]) -> dict[int, dict]:
    folder_started = time.perf_counter()
    cache_folder.mkdir(parents=True, exist_ok=True)
    version_path = cache_folder / ".prompt_version"
    cache_version_current = (
        version_path.is_file()
        and version_path.read_text(encoding="utf-8").strip() == EXTRACTION_PROMPT_VERSION
    )
    extractions: dict[int, dict] = {}

    pending: dict[int, tuple[Path, Path]] = {}
    for image_number, image_path in sorted(images.items()):
        cache_path = cache_folder / f"{image_number}.json"
        current_hash = image_content_hash(image_path)
        cached_extraction = None
        if cache_version_current and cache_path.is_file():
            cached_extraction = json.loads(cache_path.read_text(encoding="utf-8"))

        cached_hash = (cached_extraction or {}).get("_image_stats", {}).get("image_hash")
        cache_is_current = cached_extraction is not None and (
            cached_hash == current_hash or cached_hash is None
        )

        if cache_is_current:
            print(f"Using cached extraction for {folder.name}/{image_number}")
            extractions[image_number] = cached_extraction
            continue

        pending[image_number] = (image_path, cache_path)

    def extract_and_cache(image_number: int, paths: tuple[Path, Path]) -> tuple[int, dict]:
        image_path, cache_path = paths
        extraction = extract_dashboard_image(image_number, image_path)
        cache_path.write_text(json.dumps(extraction, indent=2, ensure_ascii=False), encoding="utf-8")
        return image_number, extraction

    if pending:
        worker_count = min(MAX_PARALLEL_EXTRACTIONS, len(pending))
        with ThreadPoolExecutor(max_workers=worker_count, thread_name_prefix="dashboard-extract") as executor:
            futures = {
                executor.submit(extract_and_cache, image_number, paths): image_number
                for image_number, paths in pending.items()
            }
            for future in as_completed(futures):
                image_number = futures[future]
                try:
                    _, extraction = future.result()
                except Exception as exc:
                    print(f"Image {image_number} failed: {type(exc).__name__}: {exc}")
                    extraction = {
                        "items": [],
                        "notes": [f"Extraction failed: {type(exc).__name__}: {exc}"],
                        "_image_stats": {
                            "width": image_dimensions(images[image_number])[0],
                            "height": image_dimensions(images[image_number])[1],
                            "pixels": image_dimensions(images[image_number])[0] * image_dimensions(images[image_number])[1],
                            "llm_calls": 0,
                            "prompt_tokens": 0,
                            "prompt_text_tokens": _prompt_token_count or 0,
                            "image_input_tokens": 0,
                            "output_tokens": 0,
                            "total_tokens": 0,
                            "image_hash": image_content_hash(images[image_number]),
                            "failed": True,
                        },
                    }
                extractions[image_number] = extraction

    version_path.write_text(EXTRACTION_PROMPT_VERSION, encoding="utf-8")
    log_timing(
        f"[timing] folder={folder.name} total_seconds={time.perf_counter() - folder_started:.2f} "
        f"cached={len(images) - len(pending)} pending={len(pending)}"
    )
    return {image_number: extractions[image_number] for image_number in sorted(extractions)}


def write_aggregate_extractions(
    trendence_extractions: dict[int, dict],
    spartnash_extractions: dict[int, dict],
) -> None:
    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
    for output_path, extractions in (
        (TRENDENCE_EXTRACTION_JSON, trendence_extractions),
        (SPARTNASH_EXTRACTION_JSON, spartnash_extractions),
    ):
        output_path.write_text(
            json.dumps(
                {str(k): v for k, v in sorted(extractions.items())},
                indent=2,
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )


def load_manifest(path: Path = MANIFEST_JSON) -> dict:
    """Load per-image view context (page name + applied filter) supplied by the main app.

    Expected shape: {"views": [{"image_number": 1, "page_name": "Overview",
    "filter_name": "Time Period", "selected": ["Last 4 Weeks"], "state": "..."}]}.
    Without a manifest, views fall back to numbered-pair behaviour with
    state = "no_filter".
    """
    if not path.is_file():
        print(f"No manifest found at {path}; using numbered-pair views with state=no_filter")
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        print(f"WARNING: could not read manifest {path}: {exc}")
        return {}
    entries = data.get("views", data) if isinstance(data, dict) else data
    context: dict = {}
    for entry in entries:
        if not isinstance(entry, dict):
            continue
        try:
            image_number = int(entry.get("image_number"))
        except (TypeError, ValueError):
            continue
        filter_name = entry.get("filter_name") or None
        selected = list(entry.get("selected") or [])
        page_name = entry.get("page_name") or entry.get("page") or None
        state = entry.get("state") or (
            f"{filter_name}: {', '.join(selected)}" if filter_name and selected else "no_filter"
        )
        context[image_number] = {
            "page_name": page_name,
            "filter_name": filter_name,
            "selected": selected,
            "state": state,
        }
    return context


def view_context_for(context: dict, image_number: int) -> dict:
    entry = (context or {}).get(image_number) or {}
    return {
        "page_name": entry.get("page_name"),
        "filter_name": entry.get("filter_name"),
        "selected": list(entry.get("selected") or []),
        "state": entry.get("state") or "no_filter",
    }


def filters_from_extraction(extraction: dict) -> list[dict]:
    """Extract one {filter_name, selected} entry per Slicer-section item."""
    filters = []
    for item in extraction.get("items", []):
        if normalize_label(item.get("section")) != "slicer":
            continue
        selected = item.get("selected_options")
        if not isinstance(selected, list):
            selected = [
                part.strip()
                for part in str(item.get("value") or "").split(",")
                if part.strip()
            ]
        filters.append({"filter_name": item.get("item_name"), "selected": list(selected)})
    return filters


# ---------------------------------------------------------------------------
# Comparison Logic
# ---------------------------------------------------------------------------
def normalize_label(text: object) -> str:
    return " ".join(str(text or "").strip().lower().split())


def normalize_color_token(value: object) -> str:
    """Normalize model color tokens for exact slice-to-legend joins."""
    return re.sub(r"[^a-z0-9#]", "", str(value or "").casefold())


def canonical_section(value: object) -> str:
    """Map legacy extraction labels to the four comparison sections."""
    section = normalize_label(value).replace("-", "_")
    if section in {"slicer", "filter", "filters"}:
        return "slicer"
    if section in {"series_value", "series", "multi_series", "data_point"}:
        return "series_value"
    if section in {"data_label", "kpi", "kpi_card", "value", "label", "label_value", "text"}:
        return "data_label"
    if section in {"chart", "treemap", "table", "matrix", "bar", "axis_label", "category_label", "title", "icon"}:
        return "chart"
    return "chart"


def normalize_value(text: object) -> str:
    return normalize_label(text).replace(",", "")


def comparison_category(item: dict) -> str:
    """Use the complete item label to resolve conflicting model categories."""
    item_name = str(item.get("item_name") or "").strip()
    if "-" in item_name:
        label_category = item_name.rsplit("-", 1)[1].strip()
        if label_category:
            return normalize_label(label_category)
    return normalize_label(item.get("category"))


def category_match_score(left: object, right: object) -> float:
    """Score likely OCR truncations without equating short or qualified labels."""
    left_label = normalize_label(left).rstrip(".… ")
    right_label = normalize_label(right).rstrip(".… ")
    if not left_label or not right_label:
        return 0.0
    if left_label == right_label:
        return 1.0

    similarity = difflib.SequenceMatcher(None, left_label, right_label).ratio()
    if similarity >= 0.78:
        return similarity

    shorter, longer = sorted((left_label, right_label), key=len)
    if (
        len(shorter) >= 6
        and longer.startswith(shorter)
        and len(shorter) / len(longer) >= 0.4
        and longer[len(shorter)].isalnum()
    ):
        return 0.78
    return 0.0


def standardized_item_name(left: dict, right: dict) -> str:
    """Use consistent spacing and the more complete spelling for chart categories."""
    left_name = str(left.get("item_name") or "").strip()
    right_name = str(right.get("item_name") or "").strip()
    if canonical_section(left.get("section")) not in {"data_label", "series_value", "chart"}:
        return left_name or right_name or "Unnamed item"
    left_category = str(left.get("category") or left_name.rsplit("-", 1)[-1]).strip()
    right_category = str(right.get("category") or right_name.rsplit("-", 1)[-1]).strip()
    if category_match_score(left_category, right_category) < 0.78:
        return left_name or right_name or "Unnamed item"

    category = max(
        (left_category, right_category),
        key=lambda label: len(label.rstrip(".… ")),
    )

    def item_prefix(name: str, item_category: str) -> str:
        if item_category and name.casefold().endswith(item_category.casefold()):
            return name[: -len(item_category)].rstrip(" -")
        return name.rsplit("-", 1)[0].strip()

    prefix = item_prefix(left_name, left_category) or item_prefix(right_name, right_category)
    return f"{prefix} - {category}" if prefix else category


def extract_comparable_value(text: object) -> str:
    raw = str(text or "")
    match = re.search(r"value\s*:\s*(.*)$", raw, flags=re.IGNORECASE)
    return normalize_value(match.group(1) if match else raw)


def comma_separated_values(text: object) -> tuple[str, ...] | None:
    raw = str(text or "")
    if "," not in raw:
        return None
    if _is_grouped_numeric_value(raw):
        return None
    values = [normalize_value(value) for value in re.split(r",\s+", raw)]
    if len(values) == 1:
        values = [normalize_value(value) for value in raw.split(",")]
    return tuple(values) if len(values) > 1 and all(values) else None


def _is_grouped_numeric_value(value: str) -> bool:
    return bool(
        re.fullmatch(
            r"[+-]?\$?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?(?:%|[kKmM])?",
            value.strip(),
        )
    )


def split_chart_values(value: object, expected_count: int) -> list[str]:
    """Split a multi-series value without treating thousands separators as delimiters."""
    raw = ("" if value is None else str(value)).strip()
    if not raw or expected_count < 2 or "," not in raw:
        return [raw] if raw else []
    if _is_grouped_numeric_value(raw):
        return [raw]

    def split_numeric_values(start: int, remaining: int) -> Optional[list[str]]:
        if remaining == 1:
            final_value = raw[start:].strip()
            return [final_value] if _is_grouped_numeric_value(final_value) else None
        for comma_index in range(start, len(raw)):
            if raw[comma_index] != ",":
                continue
            component = raw[start:comma_index].strip()
            if not _is_grouped_numeric_value(component):
                continue
            remainder = split_numeric_values(comma_index + 1, remaining - 1)
            if remainder is not None:
                return [component, *remainder]
        return None

    numeric_parts = split_numeric_values(0, expected_count)
    if numeric_parts is not None:
        return numeric_parts

    spaced_parts = [part.strip() for part in re.split(r",\s+", raw)]
    if len(spaced_parts) == expected_count and all(spaced_parts):
        return spaced_parts

    parts = [part.strip() for part in raw.split(",")]
    return parts if len(parts) == expected_count and all(parts) else [raw]


def _numeric_value(text: object) -> Optional[float]:
    """Parse a display value as a plain number allowing $, comma, % and k/m suffixes."""
    raw = display_value(text)
    if not isinstance(raw, str):
        return None
    cleaned = raw.strip().replace(",", "").replace("$", "").replace("%", "").strip()
    if not cleaned:
        return None
    multiplier = 1.0
    suffix = cleaned[-1:]
    if suffix in {"k", "K"}:
        multiplier = 1000.0
        cleaned = cleaned[:-1]
    elif suffix in {"m", "M"}:
        multiplier = 1000000.0
        cleaned = cleaned[:-1]
    if cleaned.casefold() in {"null", "none", "n/a"}:
        return None
    try:
        return float(cleaned) * multiplier
    except ValueError:
        return None


def _selected_options(item: dict) -> Optional[list[str]]:
    options = item.get("selected_options")
    if isinstance(options, list) and options:
        return sorted(str(option).strip().casefold() for option in options if str(option).strip())
    return None


def values_match_for_item(left: dict, right: dict) -> bool:
    left_series = left.get("series_values")
    right_series = right.get("series_values")
    if isinstance(left_series, dict) or isinstance(right_series, dict):
        if not isinstance(left_series, dict) or not isinstance(right_series, dict):
            return False
        normalized_left = {normalize_label(key): value for key, value in left_series.items()}
        normalized_right = {normalize_label(key): value for key, value in right_series.items()}
        if set(normalized_left) != set(normalized_right):
            return False
        return all(
            values_match_for_item(
                {"value": normalized_left[key]},
                {"value": normalized_right[key]},
            )
            for key in normalized_left
        )

    left_selected = _selected_options(left)
    right_selected = _selected_options(right)
    if left_selected is not None and right_selected is not None:
        return left_selected == right_selected

    left_number = _numeric_value(left.get("value"))
    right_number = _numeric_value(right.get("value"))
    if left_number is not None and right_number is not None:
        return abs(left_number - right_number) <= 1e-6 * max(1.0, abs(left_number), abs(right_number))

    left_values = comma_separated_values(left.get("value"))
    right_values = comma_separated_values(right.get("value"))
    if left_values is not None and right_values is not None:
        return left_values == right_values
    return extract_comparable_value(left.get("value")) == extract_comparable_value(right.get("value"))


def extraction_value_is_missing(item: dict) -> bool:
    value = item.get("value")
    return value is None or str(value).strip().casefold() in {"", "n/a", "na", "null", "none", "unreadable"}


def display_value(text: object) -> object:
    raw = "" if text is None else str(text)
    match = re.search(r"value\s*:\s*(.*)$", raw, flags=re.IGNORECASE)
    return match.group(1).strip() if match else text


def format_chart_component(value: object) -> str:
    value = display_value(value)
    if value is None:
        return "N/A"
    text = str(value).strip()
    return text if text and text.lower() not in {"null", "none"} else "N/A"


def format_chart_value(value: object) -> str:
    text = display_value(value)
    if not isinstance(text, str) or "," not in text:
        return format_chart_component(text)
    components = [format_chart_component(component) for component in text.split(",")]
    return ",".join(components)


def format_pie_value(value: object, percentage: object) -> object:
    """Show circular-chart values with the percentage printed beside the slice."""
    formatted_value = format_chart_component(value)
    if formatted_value == "N/A" or percentage is None:
        return formatted_value
    try:
        return f"{formatted_value} ({float(percentage):.2f}%)"
    except (TypeError, ValueError):
        return formatted_value


def format_chart_pair(first: object, second: object) -> str:
    components = [format_chart_component(first), format_chart_component(second)]
    while components and not components[-1]:
        components.pop()
    return ",".join(components)


def normalize_single_series_items(items: list[dict]) -> list[dict]:
    """Give repeated single-series chart rows a stable category-aware item_name when safe.

    The model sometimes returns the chart title repeatedly and places a category/value
    pair in `value`, e.g. "GRAND RAPIDS,300". Only rewrite the name when the final
    comma-separated component is clearly numeric and the prefix contains non-numeric
    text. This avoids treating ordinary numeric values such as "1,200" as categories.
    """
    normalized = []

    category_value_pattern = re.compile(
        r"^(?P<category>.+?),\s*(?P<number>-?\d[\d,]*(?:\.\d+)?)$"
    )

    for item in items:
        item = dict(item)
        section = canonical_section(item.get("section"))
        item["section"] = section
        if section in {"chart", "treemap"}:
            item_name = str(item.get("item_name") or "").strip()
            value = str(item.get("value") or "").strip()

            if item_name and " - " not in item_name and value:
                match = category_value_pattern.match(value)
                if match:
                    category = match.group("category").strip()
                    if category and not re.fullmatch(r"[-+]?\d[\d,]*(?:\.\d+)?", category):
                        item["item_name"] = f"{item_name} - {category}"

        normalized.append(item)

    return normalized


_LEGACY_SERIES_ROW_PATTERN = re.compile(r"^(.*?) - ([^-]+) \(([^()]+)\)$")


def sort_and_format_chart_items(items: list[dict]) -> list[dict]:
    """Normalize legacy chart records without regrouping atomic series values."""
    normalized: list[dict] = []
    for source_item in items:
        item = dict(source_item)
        section = canonical_section(item.get("section"))
        item["section"] = section
        if section != "chart":
            normalized.append(item)
            continue

        item_name = str(item.get("item_name") or "").strip()
        series_name = item.get("series_name")
        legacy_match = (
            _LEGACY_SERIES_ROW_PATTERN.match(item_name)
            if not item.get("category") and not series_name
            else None
        )
        if legacy_match:
            chart_title, legacy_series, category = legacy_match.groups()
            series_name = series_name or legacy_series
        elif " - " in item_name:
            chart_title, category = item_name.split(" - ", 1)
        else:
            chart_title = item_name
            category = item.get("category")

        chart_title = (chart_title or item_name or "Chart").strip()
        category = str(category or item.get("category") or "").strip() or None
        item["_chart_title"] = chart_title
        series_parts = [part.strip() for part in str(series_name or "").split(",") if part.strip()]
        value_parts = split_chart_values(item.get("value"), len(series_parts))
        if len(series_parts) > 1 and len(value_parts) == len(series_parts):
            for series_part, value_part in zip(series_parts, value_parts):
                atomic_item = dict(item)
                atomic_item["section"] = "series_value"
                atomic_item["series_name"] = series_part
                atomic_item["value"] = value_part
                atomic_item["category"] = category
                atomic_item["item_name"] = (
                    f"{series_part}-{category}" if category else series_part
                )
                normalized.append(atomic_item)
            continue

        title_series = []
        if not series_parts and " vs " in chart_title.casefold():
            title_parts = re.split(r"\s+vs\s+", chart_title, maxsplit=1, flags=re.IGNORECASE)
            if len(title_parts) == 2:
                title_series = [part.strip() for part in title_parts]
                value_parts = split_chart_values(item.get("value"), len(title_series))
        if title_series and len(value_parts) == len(title_series):
            for series_part, value_part in zip(title_series, value_parts):
                atomic_item = dict(item)
                atomic_item["section"] = "series_value"
                atomic_item["series_name"] = series_part.strip()
                atomic_item["value"] = value_part
                atomic_item["category"] = category
                atomic_item["item_name"] = (
                    f"{series_part.strip()}-{category}"
                    if category
                    else series_part.strip()
                )
                normalized.append(atomic_item)
            continue

        if not item.get("value") and not item.get("category") and not series_name:
            item["section"] = "chart"
            item["item_name"] = chart_title
            normalized.append(item)
            continue
        item["section"] = "series_value" if series_name else "data_label"
        item["item_name"] = (
            f"{series_name}-{category}"
            if series_name and category
            else f"{chart_title}-{category}"
            if category
            else series_name or chart_title
        )
        item["series_name"] = series_name
        item["category"] = category
        normalized.append(item)
    return normalized


def bundle_chart_items_for_comparison(items: list[dict]) -> list[dict]:
    """Use one comparable row for each non-circular visual/category."""
    bundled: dict[tuple[str, str, str], dict] = {}
    output: list[dict] = []
    pie_types = {"pie", "pie_chart", "donut", "donut_chart", "funnel", "slice"}

    for source_item in items:
        item = dict(source_item)
        section = canonical_section(item.get("section"))
        object_type = normalize_label(item.get("object_type") or item.get("visual_type"))
        category = str(item.get("category") or "").strip()
        title = str(
            item.get("_visual_title")
            or item.get("_chart_title")
            or str(item.get("item_name") or "").split(" - ", 1)[0]
        ).strip()
        can_bundle = section == "series_value" and bool(category) and object_type not in pie_types
        if not can_bundle:
            output.append(item)
            continue

        key = (normalize_label(title), normalize_label(category), normalize_label(object_type))
        group = bundled.get(key)
        if group is None:
            group = dict(item)
            group.update(
                {
                    "section": "data_label",
                    "item_name": f"{title}-{category}",
                    "_chart_title": title,
                    "_visual_title": item.get("_visual_title") or title,
                    "series_name": None,
                    "value": [],
                    "series_values": {},
                }
            )
            bundled[key] = group
            output.append(group)
        value = item.get("value")
        if value is not None and str(value).strip() != "":
            group["value"].append(str(value).strip())
            series_name = normalize_label(item.get("series_name"))
            if series_name:
                group["series_values"][series_name] = str(value).strip()

    for item in output:
        if isinstance(item.get("value"), list):
            item["value"] = ",".join(item["value"]) or None
            if not item["series_values"]:
                item.pop("series_values", None)
    return output


def deduplicate_extraction_items(items: list[dict]) -> list[dict]:
    """Drop semantic duplicates while retaining repeated objects from distinct visuals."""
    deduplicated: list[dict] = []
    seen: set[tuple] = set()
    for item in items:
        bbox = item.get("_visual_bbox") or item.get("bbox")
        spatial_key = tuple(round(float(value), 4) for value in bbox) if isinstance(bbox, list) and len(bbox) == 4 else None
        visual_key = normalize_label(item.get("_visual_title") or item.get("_chart_title"))
        key = (
            canonical_section(item.get("section")),
            visual_key,
            normalize_label(item.get("item_name")),
            normalize_label(item.get("category")),
            normalize_label(item.get("series_name")),
            normalize_value(item.get("value")),
            spatial_key if not visual_key else None,
        )
        if key in seen:
            continue
        seen.add(key)
        deduplicated.append(item)
    return deduplicated


def compare_extractions(trendence_data: dict, spartnash_data: dict) -> dict:
    trendence_items = deduplicate_extraction_items(
        bundle_chart_items_for_comparison(
            sort_and_format_chart_items(
            normalize_single_series_items(list(trendence_data.get("items", [])))
            )
        )
    )
    spartnash_items = deduplicate_extraction_items(
        bundle_chart_items_for_comparison(
            sort_and_format_chart_items(
            normalize_single_series_items(list(spartnash_data.get("items", [])))
            )
        )
    )

    for idx, item in enumerate(trendence_items):
        item["_orig_idx"] = idx
    for idx, item in enumerate(spartnash_items):
        item["_orig_idx"] = idx

    comparison_items: list[dict] = []

    def visual_signature(item: dict) -> tuple[str, str, str]:
        title = normalize_label(item.get("_visual_title") or item.get("_chart_title"))
        if not title:
            title = normalize_label(str(item.get("item_name") or "").split(" - ", 1)[0])
        visual_type = normalize_label(item.get("visual_type") or item.get("section"))
        # Child-object coordinates distinguish repeated cards/labels inside one visual.
        bbox = item.get("bbox") or item.get("_visual_bbox")
        if isinstance(bbox, list) and len(bbox) == 4:
            position = ",".join(str(round(float(value), 2)) for value in bbox)
        else:
            position = ""
        return title, visual_type, position

    def comparison_identity(item: dict) -> str:
        visual_title = visual_signature(item)[0]
        section = canonical_section(item.get("section"))
        category = comparison_category(item)
        series = normalize_label(item.get("series_name"))
        item_name = normalize_label(item.get("item_name"))
        object_type = normalize_label(item.get("object_type") or item.get("visual_type"))
        color = normalize_color_token(item.get("attribute") or item.get("color_token"))
        color_identity = color if object_type in {"slice", "pie", "pie_chart", "donut", "donut_chart"} else ""
        # Series is part of identity: two legend values may share one category.
        # The visual title keeps repeated labels such as "Opened" in separate cards.
        return "|".join((visual_title, section, category, series, item_name, color_identity))

    def key_of(item: dict) -> tuple[str, str]:
        return (canonical_section(item.get("section")), comparison_identity(item))

    remaining_spartnash = list(spartnash_items)

    def visual_exists_on_other_side(item: dict, other_items: list[dict]) -> bool:
        title = visual_signature(item)[0]
        return bool(
            title
            and any(visual_signature(candidate)[0] == title for candidate in other_items)
        )

    def series_signature(item: dict) -> tuple[str, ...]:
        series_values = item.get("series_values")
        if isinstance(series_values, dict) and series_values:
            return tuple(sorted(normalize_label(name) for name in series_values))
        series_name = normalize_label(item.get("series_name"))
        return (series_name,) if series_name else ()

    def pop_exact_match(item: dict) -> Optional[dict]:
        key = key_of(item)
        if key[0] == "treemap" and not key[1]:
            return None
        for index, candidate in enumerate(remaining_spartnash):
            if key_of(candidate) == key:
                return remaining_spartnash.pop(index)
        return None

    def pop_best_fuzzy_match(item: dict) -> Optional[dict]:
        section = canonical_section(item.get("section"))
        all_candidates = [
            c for c in remaining_spartnash if canonical_section(c.get("section")) == section
        ]
        target_category = comparison_category(item)
        target_series = normalize_label(item.get("series_name"))
        target_type = normalize_label(item.get("object_type") or item.get("visual_type"))
        target_color = normalize_color_token(item.get("attribute") or item.get("color_token"))
        target_is_circular = target_type in {"slice", "pie", "pie_chart", "donut", "donut_chart"}
        candidates = [
            candidate
            for candidate in all_candidates
            if (
                comparison_category(candidate) == target_category
                and normalize_label(candidate.get("series_name")) == target_series
            )
            or (
                target_is_circular
                and target_color
                and target_color
                == normalize_color_token(candidate.get("attribute") or candidate.get("color_token"))
            )
        ]
        layout_only = False
        if not candidates and target_category and not target_is_circular:
            target_title = visual_signature(item)[0]
            fuzzy_candidates = []
            for candidate in all_candidates:
                if (
                    visual_signature(candidate)[0] != target_title
                    or series_signature(candidate) != series_signature(item)
                ):
                    continue
                similarity = category_match_score(
                    target_category,
                    comparison_category(candidate),
                )
                if similarity >= 0.78:
                    fuzzy_candidates.append((similarity, candidate))

            fuzzy_candidates.sort(key=lambda entry: entry[0], reverse=True)
            if fuzzy_candidates and (
                len(fuzzy_candidates) == 1
                or fuzzy_candidates[0][0] - fuzzy_candidates[1][0] >= 0.08
            ):
                candidates = [fuzzy_candidates[0][1]]

        target_parent_bbox = item.get("_visual_bbox")
        target_bbox = item.get("bbox") or target_parent_bbox

        def bbox_distance(left: object, right: object) -> Optional[float]:
            if not isinstance(left, list) or not isinstance(right, list) or len(left) != 4 or len(right) != 4:
                return None
            try:
                return sum(abs(float(a) - float(b)) for a, b in zip(left, right))
            except (TypeError, ValueError):
                return None

        if not candidates and section in {"data_label", "series_value", "chart"}:
            target_title = visual_signature(item)[0]
            layout_candidates = []
            for candidate in all_candidates:
                candidate_title = visual_signature(candidate)[0]
                parent_distance = bbox_distance(target_parent_bbox, candidate.get("_visual_bbox"))
                object_distance = bbox_distance(target_bbox, candidate.get("bbox"))
                title_score = difflib.SequenceMatcher(None, target_title, candidate_title).ratio()
                if (
                    parent_distance is not None
                    and parent_distance <= 0.12
                    and (object_distance is None or object_distance <= 0.20)
                    and (title_score >= 0.35 or parent_distance <= 0.04)
                ):
                    layout_candidates.append(candidate)
            candidates = layout_candidates
            layout_only = bool(candidates)

        candidates = [
            candidate
            for candidate in candidates
            if not layout_only or canonical_section(candidate.get("section")) == section
        ]
        if not candidates:
            return None
        target_title, target_type, target_position = visual_signature(item)
        target_name = normalize_label(item.get("item_name"))

        def score(candidate: dict) -> float:
            candidate_title, candidate_type, candidate_position = visual_signature(candidate)
            title_score = difflib.SequenceMatcher(None, target_title, candidate_title).ratio()
            name_score = difflib.SequenceMatcher(
                None, target_name, normalize_label(candidate.get("item_name"))
            ).ratio()
            semantic_score = max(title_score, name_score)
            if not layout_only and semantic_score < 0.68:
                return -1.0
            type_score = 1.0 if target_type and target_type == candidate_type else 0.0
            position_score = 0.0
            if target_position and candidate_position:
                target_values = [float(value) for value in target_position.split(",")]
                candidate_values = [float(value) for value in candidate_position.split(",")]
                distance = sum(abs(left - right) for left, right in zip(target_values, candidate_values))
                position_score = max(0.0, 1.0 - distance / 1.5)
            return semantic_score * 0.65 + position_score * 0.25 + type_score * 0.1

        match = max(candidates, key=score)
        if score(match) < (0.55 if layout_only else FUZZY_MATCH_CUTOFF):
            return None
        remaining_spartnash.remove(match)
        return match

    def _carry(primary: dict, fallback: Optional[dict] = None) -> dict:
        fallback = fallback or {}

        def pick(*keys):
            for candidate in (primary, fallback):
                for key in keys:
                    value = candidate.get(key)
                    if value:
                        return value
            return None

        return {
            "visual_type": pick("visual_type"),
            "object_type": pick("object_type"),
            "attribute": pick("attribute"),
            "kpi_title": pick("kpi_title"),
            "kpi_value": pick("kpi_value"),
            "series_name": pick("series_name"),
            "category": pick("category"),
            "percentage": pick("percentage"),
            "attribute": pick("attribute", "color_token"),
            "_visual_title": pick("_visual_title"),
            "_chart_title": pick("_chart_title"),
            "_visual_bbox": pick("_visual_bbox"),
            "bbox": pick("bbox"),
            "scrollable": bool(pick("scrollable")),
            "selected_options": pick("selected_options"),
        }

    for trendence_item in trendence_items:
        match = pop_exact_match(trendence_item) or pop_best_fuzzy_match(trendence_item)
        trendence_value = trendence_item.get("value")
        confidences = [c for c in (trendence_item.get("confidence"),) if isinstance(c, (int, float))]
        orig_idx = trendence_item.get("_orig_idx", 9999)

        if match is None:
            visual_counterpart_exists = visual_exists_on_other_side(
                trendence_item,
                spartnash_items,
            )
            comparison_items.append(
                {
                    "section": trendence_item.get("section") or "Unspecified",
                    "item_name": trendence_item.get("item_name") or "Unnamed item",
                    "trendence_value": trendence_value,
                    "spartnash_value": None,
                    "trendence_series": trendence_item.get("series_values"),
                    "spartnash_series": None,
                    "trendence_percentage": trendence_item.get("percentage"),
                    "spartnash_percentage": None,
                    "status": "Uncertain" if visual_counterpart_exists else "Trendence Only",
                    "difference": (
                        "Could not align this item within a visual present on both dashboards"
                        if visual_counterpart_exists
                        else "Only present in Trendence"
                    ),
                    "confidence": mean(confidences) if confidences else None,
                    **_carry(trendence_item),
                    "_orig_idx": orig_idx,
                }
            )
            continue

        spartnash_value = match.get("value")
        confidences += [c for c in (match.get("confidence"),) if isinstance(c, (int, float))]
        values_match = values_match_for_item(trendence_item, match)
        data_section = normalize_label(trendence_item.get("section")) in {
            "chart", "treemap", "table", "matrix"
        }
        extraction_incomplete = (
            data_section
            and extraction_value_is_missing(trendence_item)
            and extraction_value_is_missing(match)
        )

        comparison_items.append(
            {
                "section": trendence_item.get("section") or match.get("section") or "Unspecified",
                "item_name": standardized_item_name(trendence_item, match),
                "trendence_value": trendence_value,
                "spartnash_value": spartnash_value,
                "trendence_series": trendence_item.get("series_values"),
                "spartnash_series": match.get("series_values"),
                "trendence_percentage": trendence_item.get("percentage"),
                "spartnash_percentage": match.get("percentage"),
                "status": "Uncertain" if extraction_incomplete else "Match" if values_match else "Different",
                "difference": (
                    "Insufficient extraction evidence"
                    if extraction_incomplete
                    else None if values_match
                    else "Values differ between Trendence and Spartnash"
                ),
                "confidence": mean(confidences) if confidences else None,
                **_carry(trendence_item, match),
                "_orig_idx": orig_idx,
            }
        )

    for spartnash_item in remaining_spartnash:
        confidence = spartnash_item.get("confidence")
        visual_counterpart_exists = visual_exists_on_other_side(
            spartnash_item,
            trendence_items,
        )
        comparison_items.append(
            {
                "section": spartnash_item.get("section") or "Unspecified",
                "item_name": spartnash_item.get("item_name") or "Unnamed item",
                "trendence_value": None,
                "spartnash_value": spartnash_item.get("value"),
                "trendence_series": None,
                "spartnash_series": spartnash_item.get("series_values"),
                "trendence_percentage": None,
                "spartnash_percentage": spartnash_item.get("percentage"),
                "status": "Uncertain" if visual_counterpart_exists else "Spartnash Only",
                "difference": (
                    "Could not align this item within a visual present on both dashboards"
                    if visual_counterpart_exists
                    else "Only present in Spartnash"
                ),
                "confidence": confidence if isinstance(confidence, (int, float)) else None,
                **_carry(spartnash_item),
                "_orig_idx": spartnash_item.get("_orig_idx", 9999) + 1000,
            }
        )

    def comparison_sort_key(item: dict) -> tuple:
        section = normalize_label(item.get("section"))
        section_idx = SECTION_ORDER.get(section, 99)
        item_name = normalize_label(item.get("item_name"))
        orig_idx = item.get("_orig_idx", 9999)

        bbox = item.get("_visual_bbox") or item.get("bbox")
        if isinstance(bbox, list) and len(bbox) == 4:
            try:
                left, top = float(bbox[0]), float(bbox[1])
            except (TypeError, ValueError):
                pass
            else:
                # Dashboard reading order is top-to-bottom, then left-to-right.
                return (0, round(top, 4), round(left, 4), orig_idx, item_name)

        weekday = next((day for day in WEEKDAY_ORDER if f"({day})" in item_name), None)
        if weekday is not None and " - " in item_name:
            chart_title = item_name.split(" - ", 1)[0]
            return (1, section_idx, 0, chart_title, WEEKDAY_ORDER[weekday], 0, orig_idx, item_name)

        return (1, section_idx, 1, "", -1, 0, orig_idx, item_name)

    comparison_items.sort(key=comparison_sort_key)

    for item in comparison_items:
        item.pop("_orig_idx", None)

    notes = list(
        dict.fromkeys(
            list(trendence_data.get("notes", []))
            + list(spartnash_data.get("notes", []))
        )
    )
    return {
        "trendence_dashboard_title": trendence_data.get("dashboard_title"),
        "spartnash_dashboard_title": spartnash_data.get("dashboard_title"),
        "trendence_page_header": trendence_data.get("page_header"),
        "spartnash_page_header": spartnash_data.get("page_header"),
        "trendence_active_page": trendence_data.get("active_page"),
        "spartnash_active_page": spartnash_data.get("active_page"),
        "trendence_refresh_date": trendence_data.get("refresh_date"),
        "spartnash_refresh_date": spartnash_data.get("refresh_date"),
        "comparison_items": comparison_items,
        "comparison_notes": notes,
    }


# ---------------------------------------------------------------------------
# Excel Generation
# ---------------------------------------------------------------------------
STATUS_NAMES = {
    "match": "Match",
    "different": "Different",
    "trendence only": "Trendence Only",
    "spartnash only": "Spartnash Only",
    "uncertain": "Uncertain",
}


def canonical_status(value: object) -> str:
    normalized = str(value or "").strip().lower().replace("_", " ").replace("-", " ")
    normalized = " ".join(normalized.split())
    return STATUS_NAMES.get(normalized, "Uncertain")


def calculate_metrics(items: list[dict]) -> dict:
    counts = {status: 0 for status in STATUS_NAMES.values()}
    confidences = []

    for item in items:
        status = canonical_status(item.get("status"))
        item["status"] = status
        counts[status] += 1

        confidence = item.get("confidence")
        if isinstance(confidence, (int, float)) and 0 <= confidence <= 1:
            confidences.append(float(confidence))

    evaluated = len(items) - counts["Uncertain"]
    match_percentage = counts["Match"] / evaluated if evaluated else None

    return {
        "total": len(items),
        "evaluated": evaluated,
        "counts": counts,
        "match_percentage": match_percentage,
        "average_confidence": mean(confidences) if confidences else None,
    }


HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
SUBHEADER_FILL = PatternFill("solid", fgColor="D9EAF7")
MATCH_FILL = PatternFill("solid", fgColor="E2F0D9")
DIFFERENT_FILL = PatternFill("solid", fgColor="FCE4D6")
ONLY_FILL = PatternFill("solid", fgColor="FFF2CC")
UNCERTAIN_FILL = PatternFill("solid", fgColor="E7E6E6")
TITLE_FILL = PatternFill("solid", fgColor="17365D")
ALT_ROW_FILL = PatternFill("solid", fgColor="F7FAFC")
SUMMARY_VALUE_FILL = PatternFill("solid", fgColor="FFFFFF")

WHITE_BOLD_FONT = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
BODY_FONT = Font(name="Segoe UI", size=10)
BOLD_FONT = Font(name="Segoe UI", size=10, bold=True)
TITLE_FONT = Font(name="Segoe UI", size=16, bold=True, color="FFFFFF")

THIN_BORDER = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)


def style_header_row(ws, row: int, headers: list[str]) -> None:
    for column, header in enumerate(headers, 1):
        cell = ws.cell(row=row, column=column, value=header)
        cell.font = WHITE_BOLD_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER
    ws.row_dimensions[row].height = 30


def auto_fit_columns(ws, maximum_width: int = 45) -> None:
    for column_cells in ws.columns:
        column_letter = get_column_letter(column_cells[0].column)
        content_width = max(len(str(cell.value or "")) for cell in column_cells) + 3
        ws.column_dimensions[column_letter].width = min(max(content_width, 12), maximum_width)


def combine_grouped_chart_rows(items: list[dict]) -> list[dict]:
    """Bundle multi-series chart values by visual and category for reporting."""
    grouped: dict[tuple[str, str, str], dict] = {}
    output: list[dict] = []
    pie_types = {"pie", "pie_chart", "donut", "donut_chart", "funnel", "slice"}

    for item in items:
        item = dict(item)
        section = canonical_section(item.get("section"))
        object_type = normalize_label(item.get("object_type") or item.get("visual_type"))
        category = str(item.get("category") or "").strip()
        series = str(item.get("series_name") or "").strip()
        can_bundle = section == "series_value" and bool(category) and object_type not in pie_types
        if not can_bundle:
            output.append(item)
            continue

        title = (
            item.get("_visual_title")
            or item.get("_chart_title")
            or str(item.get("item_name") or "").split("-", 1)[0]
            or "Chart"
        )
        group_key = (normalize_label(title), normalize_label(category), normalize_label(object_type))
        group = grouped.get(group_key)
        if group is None:
            group = {
                "section": "series_value",
                "item_name": f"{title} - {category}",
                "_series": {},
                "_series_order": [],
                "confidences": [],
                "visual_type": item.get("visual_type"),
                "object_type": item.get("object_type"),
                "category": category,
                "_visual_title": item.get("_visual_title") or title,
                "_visual_bbox": item.get("_visual_bbox"),
                "_output_index": len(output),
            }
            grouped[group_key] = group
            output.append(group)
        series_key = normalize_label(series) or f"series_{len(group['_series_order']) + 1}"
        if series_key not in group["_series"]:
            group["_series_order"].append(series_key)
            group["_series"][series_key] = {
                "name": series or None,
                "trendence": None,
                "spartnash": None,
                "status": item.get("status"),
            }
        series_record = group["_series"][series_key]
        if item.get("trendence_value") is not None:
            series_record["trendence"] = item.get("trendence_value")
        if item.get("spartnash_value") is not None:
            series_record["spartnash"] = item.get("spartnash_value")
        if item.get("status") not in {None, "Match"}:
            series_record["status"] = item.get("status")
        if isinstance(item.get("confidence"), (int, float)):
            group["confidences"].append(item["confidence"])

    for group in grouped.values():
        series_records = [group["_series"][key] for key in group["_series_order"]]
        trendence_values = [format_chart_component(record["trendence"]) for record in series_records]
        spartnash_values = [format_chart_component(record["spartnash"]) for record in series_records]

        def format_grouped_values(values: list[str]) -> str:
            present_values = [value for value in values if value != "N/A"]
            return "N/A" if not present_values else ",".join(present_values)

        statuses = {record["status"] for record in series_records if record["status"]}
        if statuses == {"Match"}:
            status = "Match"
            difference = None
        elif statuses == {"Trendence Only"}:
            status = "Trendence Only"
            difference = "Only present in Trendence"
        elif statuses == {"Spartnash Only"}:
            status = "Spartnash Only"
            difference = "Only present in Spartnash"
        elif "Uncertain" in statuses:
            status = "Uncertain"
            difference = "Insufficient extraction evidence"
        else:
            status = "Different"
            difference = "Values differ between Trendence and Spartnash"
        group.update(
            {
                "trendence_value": format_grouped_values(trendence_values),
                "spartnash_value": format_grouped_values(spartnash_values),
                "status": status,
                "difference": difference,
                "confidence": mean(group["confidences"]) if group["confidences"] else None,
                "series_name": ", ".join(
                    record["name"] for record in series_records if record["name"]
                ) or None,
            }
        )
        for key in ("_series", "_series_order", "confidences"):
            group.pop(key, None)
        output_index = group.pop("_output_index")
        output[output_index] = group

    return output


def write_pair_sheet(
    wb: openpyxl.Workbook,
    image_number: int,
    result: dict,
    trendence_image: Path,
    spartnash_image: Path,
) -> None:
    ws = wb.create_sheet(title=str(image_number))
    ws.sheet_view.showGridLines = True

    items = combine_grouped_chart_rows(result.get("comparison_items", []))
    metrics = calculate_metrics(items)
    counts = metrics["counts"]

    ws.merge_cells("A1:G1")
    title_cell = ws["A1"]
    title_cell.value = f"Dashboard Comparison - Pair {image_number}"
    title_cell.font = TITLE_FONT
    title_cell.fill = TITLE_FILL
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    summary_rows = [
        ("Trendence File", str(trendence_image)),
        ("Spartnash File", str(spartnash_image)),
        ("Trendence Title", result.get("trendence_dashboard_title") or "Not detected"),
        ("Spartnash Title", result.get("spartnash_dashboard_title") or "Not detected"),
        ("Total Items", metrics["total"]),
        ("Matches", counts["Match"]),
        ("Differences", counts["Different"]),
        ("Trendence Only", counts["Trendence Only"]),
        ("Spartnash Only", counts["Spartnash Only"]),
        ("Uncertain", counts["Uncertain"]),
        ("Match Accuracy", metrics["match_percentage"]),
        ("Average AI Confidence", metrics["average_confidence"]),
    ]

    for row_number, (label, value) in enumerate(summary_rows, 3):
        label_cell = ws.cell(row=row_number, column=1, value=label)
        value_cell = ws.cell(row=row_number, column=2, value=value)
        label_cell.font = BOLD_FONT
        label_cell.fill = SUBHEADER_FILL
        label_cell.border = THIN_BORDER
        value_cell.font = BODY_FONT
        value_cell.fill = SUMMARY_VALUE_FILL
        value_cell.border = THIN_BORDER
        if label == "Match Accuracy" and isinstance(value, float):
            value_cell.number_format = "0.0%"
        elif label == "Average AI Confidence" and isinstance(value, float):
            value_cell.number_format = "0.0%"

    detail_header_row = 18
    detail_headers = [
        "Section",
        "Item Name",
        "Trendence Value",
        "Spartnash Value",
        "Status",
        "Difference Summary",
        "AI Confidence",
    ]
    style_header_row(ws, detail_header_row, detail_headers)

    status_fills = {
        "Match": MATCH_FILL,
        "Different": DIFFERENT_FILL,
        "Trendence Only": ONLY_FILL,
        "Spartnash Only": ONLY_FILL,
        "Uncertain": UNCERTAIN_FILL,
    }

    for row_number, item in enumerate(items, detail_header_row + 1):
        status = canonical_status(item.get("status"))
        trendence_value = item.get("trendence_value")
        spartnash_value = item.get("spartnash_value")
        if normalize_label(item.get("section")) == "chart":
            trendence_value = format_chart_value(trendence_value)
            spartnash_value = format_chart_value(spartnash_value)
        object_type = normalize_label(item.get("object_type") or item.get("visual_type"))
        if object_type in {"pie", "pie_chart", "donut", "donut_chart", "slice"}:
            trendence_value = format_pie_value(
                trendence_value, item.get("trendence_percentage")
            )
            spartnash_value = format_pie_value(
                spartnash_value, item.get("spartnash_percentage")
            )
        values = [
            item.get("section") or "Unspecified",
            item.get("item_name") or "Unnamed item",
            display_value(trendence_value) if trendence_value is not None else "N/A",
            display_value(spartnash_value) if spartnash_value is not None else "N/A",
            status,
            item.get("difference") or ("No difference" if status == "Match" else "Not provided"),
            item.get("confidence"),
        ]
        for column, value in enumerate(values, 1):
            cell = ws.cell(row=row_number, column=column, value=value)
            cell.font = BODY_FONT
            if row_number % 2 == 1:
                cell.fill = ALT_ROW_FILL
            cell.border = THIN_BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)

        ws.cell(row=row_number, column=5).fill = status_fills[status]
        if item.get("confidence") is not None:
            ws.cell(row=row_number, column=7).number_format = "0.0%"

    last_row = max(detail_header_row, detail_header_row + len(items))
    ws.auto_filter.ref = f"A{detail_header_row}:G{last_row}"
    extraction_notes = [
        str(note)
        for note in result.get("comparison_notes", [])
        if str(note).strip()
    ]
    if extraction_notes:
        notes_header_row = last_row + 3
        ws.merge_cells(start_row=notes_header_row, start_column=1, end_row=notes_header_row, end_column=7)
        notes_header = ws.cell(row=notes_header_row, column=1, value="Extraction Notes")
        notes_header.font = WHITE_BOLD_FONT
        notes_header.fill = DIFFERENT_FILL
        notes_header.alignment = Alignment(vertical="center", wrap_text=True)
        notes_header.border = THIN_BORDER
        for note_index, note in enumerate(extraction_notes, notes_header_row + 1):
            ws.merge_cells(start_row=note_index, start_column=1, end_row=note_index, end_column=7)
            note_cell = ws.cell(row=note_index, column=1, value=note)
            note_cell.font = BODY_FONT
            note_cell.fill = DIFFERENT_FILL if "[HIGH]" in note else UNCERTAIN_FILL
            note_cell.alignment = Alignment(vertical="top", wrap_text=True)
            note_cell.border = THIN_BORDER
    auto_fit_columns(ws)


def write_error_sheet(wb: openpyxl.Workbook, image_number: int, message: str) -> None:
    ws = wb.create_sheet(title=str(image_number))
    ws["A1"] = f"Dashboard Comparison - Pair {image_number}"
    ws["A3"] = "Status"
    ws["B3"] = "Failed"
    ws["A4"] = "Reason"
    ws["B4"] = message


def write_jsonl_output(results: list[dict], output_path: Path = OUTPUT_JSONL) -> None:
    """Write one dashboard comparison result per JSONL line."""
    output_started = time.perf_counter()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with output_path.open("w", encoding="utf-8") as f:
        for result in results:
            f.write(
                json.dumps(
                    result,
                    ensure_ascii=False,
                    separators=(",", ":"),
                )
                + "\n"
            )

    print(f"JSONL output saved to:\n{output_path.resolve()}")
    log_timing(f"[timing] output=jsonl_write seconds={time.perf_counter() - output_started:.2f}")


def to_dashboard_json(
    image_number: int,
    result: dict,
    trendence_image: Path,
    spartnash_image: Path,
    *,
    page_name: Optional[str] = None,
    state: Optional[str] = None,
    filters: Optional[list] = None,
) -> dict:
    """Convert an internal comparison result to the public dashboard JSON shape."""
    items = combine_grouped_chart_rows(result.get("comparison_items", []))
    metrics = calculate_metrics(items)
    match_percentage = metrics["match_percentage"]
    percentage = round(match_percentage * 100) if match_percentage is not None else 0
    overall_status = "MATCH" if percentage == 100 else "MISMATCH" if percentage == 0 else "PARTIAL_MATCH"

    if filters is None:
        filters = []
        for item in items:
            if normalize_label(item.get("section")) != "slicer":
                continue
            selected = item.get("selected_options")
            if not isinstance(selected, list):
                selected = [
                    part.strip()
                    for part in str(
                        item.get("trendence_value") or item.get("spartnash_value") or ""
                    ).split(",")
                    if part.strip()
                ]
            filters.append({"filter_name": item.get("item_name"), "selected": selected})

    kpis = []
    charts = []
    tables = []
    mismatch_summary = []
    visuals = []
    for item in items:
        status = canonical_status(item.get("status"))
        section = normalize_label(item.get("section"))
        visual_type = item.get("visual_type")
        if not visual_type:
            visual_type = {
                "kpi": "kpi_card",
                "slicer": "slicer",
                "table": "table",
                "treemap": "treemap",
                "chart": None,
            }.get(section)

        match_status = "MATCH" if status == "Match" else "MISMATCH"

        if section == "kpi":
            kpis.append(
                {
                    "kpi_title": item.get("kpi_title") or item.get("item_name"),
                    "kpi_value": item.get("kpi_value")
                    or (item.get("trendence_value") or item.get("spartnash_value")),
                    "source_value": item.get("trendence_value"),
                    "target_value": item.get("spartnash_value"),
                    "status": match_status,
                }
            )
        elif section == "chart":
            charts.append(
                {
                    "chart_title": item.get("item_name"),
                    "chart_type": visual_type,
                    "scrollable": bool(item.get("scrollable", False)),
                    "source_value": item.get("trendence_value"),
                    "target_value": item.get("spartnash_value"),
                    "status": match_status,
                }
            )
        elif section in {"table", "matrix"}:
            tables.append(
                {
                    "table_title": item.get("item_name"),
                    "visual_type": visual_type,
                    "source_value": item.get("trendence_value"),
                    "target_value": item.get("spartnash_value"),
                    "status": match_status,
                }
            )

        if status not in {"Match", "Uncertain"}:
            mismatch_summary.append(
                {
                    "section": item.get("section") or "Unspecified",
                    "item_name": item.get("item_name"),
                    "source_value": item.get("trendence_value"),
                    "target_value": item.get("spartnash_value"),
                    "difference": item.get("difference") or "Values differ",
                }
            )

        visuals.append(
            {
                "visual_name": item.get("item_name") or "Unnamed visual",
                "visual_type": visual_type,
                "section": item.get("section") or "Unspecified",
                "status": match_status,
                "source": {"value": item.get("trendence_value"), "details": None},
                "target": {"value": item.get("spartnash_value"), "details": None},
                "differences": [] if status == "Match" else [item.get("difference") or "Values differ"],
                "confidence": item.get("confidence"),
                "scrollable": bool(item.get("scrollable", False)),
            }
        )

    return {
        "dashboard_name": result.get("trendence_dashboard_title") or result.get("spartnash_dashboard_title") or f"Dashboard {image_number}",
        "page_name": page_name,
        "state": state,
        "filters": filters,
        "metadata": {
            "dashboard_title": result.get("trendence_dashboard_title") or result.get("spartnash_dashboard_title"),
            "source_dashboard_title": result.get("trendence_dashboard_title"),
            "target_dashboard_title": result.get("spartnash_dashboard_title"),
            "page_header": result.get("trendence_page_header") or result.get("spartnash_page_header"),
            "active_page": result.get("trendence_active_page") or result.get("spartnash_active_page"),
            "refresh_date": result.get("trendence_refresh_date") or result.get("spartnash_refresh_date"),
        },
        "source": {
            "dashboard": result.get("trendence_dashboard_title") or "Trendence",
            "screenshot": trendence_image.name,
        },
        "target": {
            "dashboard": result.get("spartnash_dashboard_title") or "Spartnash",
            "screenshot": spartnash_image.name,
        },
        "overall": {"status": overall_status, "match_percentage": percentage},
        "analysis": {
            "overall_status": overall_status,
            "match_percentage": percentage,
            "slicers": filters,
            "kpis": kpis,
            "charts": charts,
            "tables": tables,
            "mismatch_summary": mismatch_summary,
        },
        "visuals": visuals,
    }


def write_json_output(results: list[dict], output_path: Path = OUTPUT_JSON) -> None:
    """Write the structured dashboard comparison JSON output."""
    output_started = time.perf_counter()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(results, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"JSON output saved to:\n{output_path.resolve()}")
    log_timing(f"[timing] output=json_write seconds={time.perf_counter() - output_started:.2f}")


def synchronize_identical_pair_extractions(
    trendence_images: dict[int, Path],
    spartnash_images: dict[int, Path],
    trendence_extractions: dict[int, dict],
    spartnash_extractions: dict[int, dict],
) -> list[int]:
    """Reuse the richest extraction when both sides contain identical image bytes."""

    def quality(extraction: dict) -> tuple[int, int, int, int, float]:
        items = extraction.get("items", [])
        usable_items = [
            item
            for item in items
            if not extraction_value_is_missing(item)
        ]
        linked_items = [
            item
            for item in usable_items
            if item.get("category") and item.get("series_name")
        ]
        confidences = [
            float(item["confidence"])
            for item in usable_items
            if isinstance(item.get("confidence"), (int, float))
        ]
        confidence = mean(confidences) if confidences else 0.0
        return (
            int(not extraction.get("_image_stats", {}).get("failed")),
            len(linked_items),
            len(usable_items),
            len(items),
            confidence,
        )

    synchronized = []
    for image_number in sorted(set(trendence_images) & set(spartnash_images)):
        if image_content_hash(trendence_images[image_number]) != image_content_hash(
            spartnash_images[image_number]
        ):
            continue

        trendence_extraction = trendence_extractions[image_number]
        spartnash_extraction = spartnash_extractions[image_number]
        shared_extraction = max(
            (trendence_extraction, spartnash_extraction),
            key=quality,
        )
        trendence_extractions[image_number] = copy.deepcopy(shared_extraction)
        spartnash_extractions[image_number] = copy.deepcopy(shared_extraction)
        synchronized.append(image_number)

    return synchronized


def process_dashboard_comparisons(
    trendence_folder: Path = TRENDENCE_FOLDER,
    spartnash_folder: Path = SPARTNASH_FOLDER,
    output_workbook: Path = OUTPUT_WORKBOOK,
) -> None:
    start_timing_run()
    process_started = time.perf_counter()
    trendence_images = discover_numbered_images(trendence_folder)
    spartnash_images = discover_numbered_images(spartnash_folder)
    image_numbers = sorted(set(trendence_images) | set(spartnash_images))

    if not image_numbers:
        raise ValueError("No numbered dashboard images found in input directories.")

    manifest = load_manifest()

    output_workbook.parent.mkdir(parents=True, exist_ok=True)
    RAW_JSON_FOLDER.mkdir(parents=True, exist_ok=True)

    folder_started = time.perf_counter()
    with ThreadPoolExecutor(max_workers=2, thread_name_prefix="folder-extract") as executor:
        trendence_future = executor.submit(
            extract_all_images,
            trendence_folder,
            TRENDENCE_EXTRACTION_FOLDER,
            trendence_images,
        )
        spartnash_future = executor.submit(
            extract_all_images,
            spartnash_folder,
            SPARTNASH_EXTRACTION_FOLDER,
            spartnash_images,
        )
        trendence_extractions = trendence_future.result()
        spartnash_extractions = spartnash_future.result()
    synchronized_pairs = synchronize_identical_pair_extractions(
        trendence_images,
        spartnash_images,
        trendence_extractions,
        spartnash_extractions,
    )
    for image_number in synchronized_pairs:
        shared_extraction = trendence_extractions[image_number]
        for cache_folder in (TRENDENCE_EXTRACTION_FOLDER, SPARTNASH_EXTRACTION_FOLDER):
            (cache_folder / f"{image_number}.json").write_text(
                json.dumps(shared_extraction, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )
        log_timing(
            f"[timing] pair={image_number} identical_images_shared_extraction"
        )
    folder_elapsed = time.perf_counter() - folder_started
    log_timing(f"[timing] stage=parallel_folder_extraction seconds={folder_elapsed:.2f}")
    write_aggregate_extractions(trendence_extractions, spartnash_extractions)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    jsonl_results: list[dict] = []
    json_results: list[dict] = []

    for image_number in image_numbers:
        trendence_image = trendence_images.get(image_number)
        spartnash_image = spartnash_images.get(image_number)

        if trendence_image is None or spartnash_image is None:
            write_error_sheet(wb, image_number, f"Missing pair image for pair {image_number}")
            continue

        trendence_extraction = trendence_extractions.get(image_number, {})
        spartnash_extraction = spartnash_extractions.get(image_number, {})
        failed_extraction = next(
            (
                extraction.get("notes", ["Extraction failed"])[0]
                for extraction in (trendence_extraction, spartnash_extraction)
                if extraction.get("_image_stats", {}).get("failed")
            ),
            None,
        )
        if failed_extraction:
            write_error_sheet(wb, image_number, failed_extraction)
            continue

        view = view_context_for(manifest, image_number)

        try:
            pair_started = time.perf_counter()
            compare_started = time.perf_counter()
            result = compare_extractions(
                trendence_extraction,
                spartnash_extraction,
            )
            log_timing(
                f"[timing] pair={image_number} compare_seconds="
                f"{time.perf_counter() - compare_started:.2f}"
            )
            validation_started = time.perf_counter()
            result["validation"] = validate_comparison_report(
                result,
                trendence_extraction,
                spartnash_extraction,
                trendence_image,
                spartnash_image,
            )
            log_timing(
                f"[timing] pair={image_number} validation_seconds="
                f"{time.perf_counter() - validation_started:.2f} enabled={ENABLE_REPORT_VALIDATION}"
            )
            raw_write_started = time.perf_counter()
            (RAW_JSON_FOLDER / f"{image_number}.json").write_text(
                json.dumps(result, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )
            log_timing(
                f"[timing] pair={image_number} raw_json_write_seconds="
                f"{time.perf_counter() - raw_write_started:.2f}"
            )
            filters = filters_from_extraction(trendence_extraction) or filters_from_extraction(
                spartnash_extraction
            )
            jsonl_results.append(
                {
                    "image_number": image_number,
                    "page_name": view["page_name"],
                    "state": view["state"],
                    "filter_name": view["filter_name"],
                    "selected": view["selected"],
                    "trendence_image": str(trendence_image),
                    "spartnash_image": str(spartnash_image),
                    **result,
                }
            )
            public_result = to_dashboard_json(
                image_number,
                result,
                trendence_image,
                spartnash_image,
                page_name=view["page_name"],
                state=view["state"],
                filters=filters,
            )
            public_result["validation"] = result["validation"]
            json_results.append(public_result)
            sheet_started = time.perf_counter()
            write_pair_sheet(wb, image_number, result, trendence_image, spartnash_image)
            log_timing(
                f"[timing] pair={image_number} workbook_sheet_seconds="
                f"{time.perf_counter() - sheet_started:.2f}"
            )
            log_timing(
                f"[timing] pair={image_number} compare_and_write_seconds="
                f"{time.perf_counter() - pair_started:.2f}"
            )
        except Exception as exc:
            write_error_sheet(wb, image_number, f"{type(exc).__name__}: {exc}")

    workbook_started = time.perf_counter()
    wb.save(output_workbook)
    log_timing(f"[timing] stage=workbook_save seconds={time.perf_counter() - workbook_started:.2f}")
    write_jsonl_output(jsonl_results)
    write_json_output(json_results)
    print(
        f"\nProcessing finished successfully."
        f"\nOutput workbook saved to:\n{output_workbook.resolve()}"
        f"\nJSONL output saved to:\n{OUTPUT_JSONL.resolve()}"
    )
    log_timing(f"[timing] stage=total seconds={time.perf_counter() - process_started:.2f}")

# /vijykfhwer,k
if __name__ == "__main__":
    process_dashboard_comparisons()