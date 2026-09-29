from concurrent.futures import ThreadPoolExecutor, as_completed
import difflib
import hashlib
import io
import json
import mimetypes
import os
import re
import struct
import threading
from pathlib import Path
from statistics import mean
from typing import Any, List, Optional, Tuple

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
OUTPUT_WORKBOOK = OUTPUT_FOLDER / "dashboard_comparison_complete.xlsx"
OUTPUT_JSONL = OUTPUT_FOLDER / "dashboard_comparison_complete.jsonl"
OUTPUT_JSON = OUTPUT_FOLDER / "dashboard_comparison_complete.json"
RAW_JSON_FOLDER = OUTPUT_FOLDER / "raw_json"
EXTRACTION_JSON_FOLDER = OUTPUT_FOLDER / "extraction_json"
TRENDENCE_EXTRACTION_FOLDER = EXTRACTION_JSON_FOLDER / "trendence"
SPARTNASH_EXTRACTION_FOLDER = EXTRACTION_JSON_FOLDER / "spartnash"
TRENDENCE_EXTRACTION_JSON = OUTPUT_FOLDER / "trendence_extractions.json"
SPARTNASH_EXTRACTION_JSON = OUTPUT_FOLDER / "spartnash_extractions.json"

FUZZY_MATCH_CUTOFF = 0.72

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


# ---------------------------------------------------------------------------
# Pydantic Schemas for Structured Gemini Responses
# ---------------------------------------------------------------------------
class ExtractionItem(BaseModel):
    section: str = Field(
        description="Standard section category: Metadata, Slicer, KPI, Chart, Treemap, Table, or Layout"
    )
    item_name: str = Field(description="Visible label or concise descriptor of the item")
    visual_type: Optional[str] = Field(
        None,
        description=(
            "Specific visual type when reliably identifiable, such as kpi_card, bar_chart, "
            "stacked_bar_chart, clustered_bar_chart, line_chart, pie_chart, donut_chart, "
            "funnel_chart, map, scatter_plot, treemap, table, matrix, slicer, or null if unknown"
        ),
    )
    value: Optional[str] = Field(None, description="Exact visible text or numerical value of the item")
    confidence: Optional[float] = Field(
        None, ge=0.0, le=1.0, description="Confidence score between 0.0 and 1.0"
    )
    background_color: str = Field(
        default="",
        description="Hex fill color code for Treemap tiles; empty string for non-treemap items",
    )
    truncated: bool = Field(
        default=False,
        description="True if the label ends with ellipses or is visibly clipped; otherwise False",
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


class TileBox(BaseModel):
    x1: int = Field(description="Left edge of tile, 0-1000 normalized")
    y1: int = Field(description="Top edge of tile, 0-1000 normalized")
    x2: int = Field(description="Right edge of tile, 0-1000 normalized")
    y2: int = Field(description="Bottom edge of tile, 0-1000 normalized")


class TreemapTileDetection(BaseModel):
    chart_title: Optional[str] = Field(
        None, description="Exact visible title/header text directly above the treemap chart"
    )
    tiles: List[TileBox] = Field(
        default_factory=list,
        description="Bounding boxes for each distinct rectangular treemap tile, excluding navigation tabs and buttons",
    )


class TileReading(BaseModel):
    label: Optional[str] = Field(None, description="Exact visible label text inside this tile crop")
    value: Optional[str] = Field(None, description="Exact visible numerical value inside this tile crop, or null")
    truncated: Optional[bool] = Field(None, description="True if text ends with '...' or is visually clipped")
    confidence: Optional[float] = Field(None, ge=0.0, le=1.0)


# ---------------------------------------------------------------------------
# Enhanced Universal Extraction Prompt
# ---------------------------------------------------------------------------
EXTRACTION_PROMPT = """
You are a precise dashboard transcription engine. Extract all visible elements from this dashboard image for automated structural comparison.

### MANDATORY SPATIAL READING RULE (CHARTS)
Reconstruct all chart data strictly from visual X-coordinate order. Never use OCR scan order, numeric value order, Y-coordinate order, or bar height order.

1. READ X-AXIS CATEGORIES: Identify all X-axis categories from LEFT to RIGHT (e.g., Sunday -> Monday -> Tuesday -> Wednesday -> Thursday -> Friday -> Saturday).
2. READ BARS WITHIN CATEGORIES: For each category, read bars strictly from LEFT to RIGHT:
    - LEFT BAR (Dark Green) = Cases Registered
    - RIGHT BAR (Light Green) = Cases Resolved
3. PAIRING RULE: Combine both values into a single deterministic pair per category: [Left Bar Value, Right Bar Value].

### MANDATORY TRANSCRIPTION RULES
1. STRICT VISUAL TRANSCRIPTION: Transcribe strictly what is visible. Do not infer, calculate, or guess missing information.
2. METADATA: Capture Dashboard Title, Page Header, Active Page/Tab Name, and Refresh Date/Time when visibly present. Store them in the corresponding top-level metadata fields as well as Metadata items when useful.
3. SLICERS: Section "Slicer". Capture ONLY currently selected options (filled radio buttons/highlights).
4. KPI CARDS: Section "KPI". Pair card title with primary value. Set visual_type = "kpi_card". Include prior period values or percentage changes in the value field if present.
5. TABLES: Section "Table". Capture column headers and row cell values. Set visual_type = "table" or "matrix" only when the visual is clearly identifiable.
6. VISUAL TYPE: For every KPI, chart, treemap, table, matrix, slicer, or other visual item, populate visual_type when it can be reliably identified from the screenshot. Never guess; use null when uncertain.

---

### SECTION CATEGORIZATION & EXTRACTION RULES

#### 1. METADATA (section: "Metadata")
- Extract Dashboard Title, Page Header, Active Page/Tab Name, and "Data Refresh Date" (e.g., "12/16/2025 7:05:19 AM").
- Populate the top-level fields dashboard_title, page_header, active_page, and refresh_date whenever the corresponding text is visibly available.
- Format: item_name = "<Metadata Category>", value = "<Exact Text Value>"

#### 2. SLICERS & FILTERS (section: "Slicer")
- Extract all top header slicers, side filter panels, and radio button options.
- Capture ONLY currently selected options (filled radio buttons, highlighted values, active dropdown options). Omit unselected options.
- Format: item_name = "<Slicer Name>", value = "<Selected Option Text>"

#### 3. KPI CARDS (section: "KPI")
- Extract primary stat callouts, scorecard metrics, and KPI containers.
- Pair the KPI title with its primary display value.
- Set visual_type = "kpi_card".
- If comparison values exist (e.g., prior period figures or percentage changes like "P10: 44 (+18.18%)"), append them to the value string.
- Format: item_name = "<KPI Title>", value = "<Primary Value> [<Comparison Text>]"

#### 4. TABLES & DATA GRIDS (section: "Table")
- Applies to full-screen tables or embedded data grids (e.g., Version logs, User lists, Skill metrics tables).
- Extract every visible row cell value associated with its column header.
- Set visual_type = "table" or "matrix" only when clearly identifiable.
- Format: item_name = "<Table Title or Column Header> - Row <Row Index or Identifier>", value = "<Cell Text>"

#### 5. CHARTS & DATA VISUALIZATIONS (section: "Chart")
- Set visual_type to the most specific reliably identifiable type, such as bar_chart, stacked_bar_chart, clustered_bar_chart, line_chart, pie_chart, donut_chart, funnel_chart, map, scatter_plot, or null.
A. DUAL-SERIES / GROUPED BAR CHARTS:
    - Process categories strictly from LEFT to RIGHT across the X-axis.
    - Match numeric labels above bars using their horizontal center X-coordinate.
    - item_name = "<Chart Title> - <Category Label>"
    - value = "<Left Bar Value>, <Right Bar Value>"

B. SINGLE SERIES / PIE / LINE CHARTS:
    - Single-Bar / Line Charts: item_name = "<Chart Title> - <Category>", value = "<Displayed Value>"
   - Extract category slice label, count value, and share percentage.
    - Pie Charts: item_name = "<Chart Title> - <Slice Label>", value = "<Value> (<Percentage>)"

#### 6. TREEMAPS (section: "Treemap")
- Set visual_type = "treemap".
- Extract every solid-colored tile in the treemap grid. Exclude navigation tabs/buttons around the visual.
- Format: item_name = "<Chart Title> - <Visible Tile Label>" (or "<Visible Tile Label>" if title is omitted).
- Format: value = "<Numeric Value inside Tile>" (or null if no number is visible).
- Populate 0-1000 bounding box coordinates (box_x1, box_y1, box_x2, box_y2). Set `truncated = true` for clipped text.
"""

TILE_DETECTION_PROMPT = """
Analyze the treemap chart within this dashboard image:
1. Identify the exact visible header/title text directly above the treemap plotting area.
2. Detect every individual rectangular colored tile in the treemap grid.
3. Exclude external elements such as page titles, KPI cards, side filters, legends, and bottom tabs/buttons.
4. Return normalized 0-1000 bounding coordinates (x1, y1, x2, y2) for each detected tile box.
"""

TILE_READING_PROMPT = """
This cropped image represents a single isolated treemap tile.
1. Read ONLY the label text and numerical value contained inside this specific tile crop.
2. Preserve capitalization, spacing, and truncation ellipses ('...') exactly as rendered.
3. Set truncated to true if the text ends with '...' or is visually cut off.
4. If no numerical value is present in this crop, set value to null. Do NOT invent or infer numbers.
"""

TREEMAP_REGION_PROMPT = """
Extract only the visible Treemap in this cropped chart panel. Return one item for
each colored tile, in reading order. Keep the exact visible tile label in item_name and the exact
number printed inside that tile in value. Never put a number in item_name.
Preserve truncation exactly (e.g., use labels such as "Talent ...", "Leav...", "Poin...").
Do not include panel titles, navigation tabs, legends, or content outside the Treemap grid.
"""

EXTRACTION_PROMPT_VERSION = "strict-left-to-right-chart-pairs-v29-metadata-visual-type"
EXTRACTION_MODELS = ("gemini-2.5-flash-lite", "gemini-2.5-flash")
ENABLE_TREEMAP_REFINEMENT = os.getenv("ENABLE_TREEMAP_REFINEMENT", "1").lower() in {"1", "true", "yes"}
MAX_PARALLEL_EXTRACTIONS = max(1, int(os.getenv("MAX_PARALLEL_EXTRACTIONS", "4")))
_prompt_token_lock = threading.Lock()


# ---------------------------------------------------------------------------
# Helper Functions & Gemini API Calls
# ---------------------------------------------------------------------------
def image_mime_type(image_path: Path) -> str:
    mime_type, _ = mimetypes.guess_type(image_path.name)
    if mime_type not in {"image/png", "image/jpeg"}:
        raise ValueError(f"Unsupported image format for: {image_path.name}")
    return mime_type


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
            TILE_DETECTION_PROMPT,
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
            TILE_READING_PROMPT,
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


def extract_treemap_region(image_path: Path) -> dict:
    crop_bytes = crop_tile_image(
        image_path,
        TileBox(x1=405, y1=575, x2=915, y2=910),
        padding=0,
    )
    if crop_bytes is None:
        return {"items": [], "llm_calls": 0, "prompt_tokens": 0, "output_tokens": 0, "total_tokens": 0}

    response = client.models.generate_content(
        model=MODEL_ID,
        contents=[
            types.Part.from_bytes(data=crop_bytes, mime_type="image/png"),
            TREEMAP_REGION_PROMPT,
        ],
        config=types.GenerateContentConfig(
            temperature=0,
            top_p=0,
            top_k=1,
            seed=42,
            response_mime_type="application/json",
            response_schema=DashboardExtraction,
        ),
    )
    usage = getattr(response, "usage_metadata", None)
    if not response.text:
        return {"items": [], "llm_calls": 1, "prompt_tokens": 0, "output_tokens": 0, "total_tokens": 0}

    parsed = DashboardExtraction.model_validate_json(response.text)
    items = []
    for item in parsed.items:
        item_data = item.model_dump()
        item_data["section"] = "Treemap"
        if not item_data["item_name"].startswith("Cases Registered"):
            item_data["item_name"] = f"Cases Registered by Disposition Area - {item_data['item_name']}"
        items.append(item_data)
    return {
        "items": items,
        "llm_calls": 1,
        "prompt_tokens": getattr(usage, "prompt_token_count", 0) or 0,
        "output_tokens": getattr(usage, "candidates_token_count", 0) or 0,
        "total_tokens": getattr(usage, "total_token_count", 0) or 0,
    }


def normalize_treemap_items(items: list[dict]) -> list[dict]:
    normalized = []
    for item in items:
        if normalize_label(item.get("section")) != "treemap":
            normalized.append(item)
            continue

        item = dict(item)
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


def extract_dashboard_image(image_number: int, image_path: Path) -> dict:
    global _prompt_token_count
    print(f"Processing image {image_number}: {image_path.name}")

    width, height = image_dimensions(image_path)
    if _prompt_token_count is None:
        with _prompt_token_lock:
            if _prompt_token_count is None:
                prompt_usage = client.models.count_tokens(model=MODEL_ID, contents=EXTRACTION_PROMPT)
                _prompt_token_count = prompt_usage.total_tokens

    request_contents = [
        types.Part.from_bytes(data=image_path.read_bytes(), mime_type=image_mime_type(image_path)),
        EXTRACTION_PROMPT,
    ]
    request_config = types.GenerateContentConfig(
        temperature=0,
        top_p=0,
        top_k=1,
        seed=42,
        max_output_tokens=16384,
        response_mime_type="application/json",
        response_schema=DashboardExtraction,
    )
    response = None
    usage = None
    last_error = None
    llm_calls = 0
    compact_prompt = (
        "Return ONLY the required dashboard facts as valid JSON. "
        "Omit notes, explanations, duplicate rows, and uncertain details. "
        "Keep item_name and value concise. Extract metadata, selected slicers, KPIs, "
        "visible chart values, treemap labels/values, and table cells."
    )
    for model_index, model in enumerate(EXTRACTION_MODELS):
        contents = request_contents if model_index == 0 else [request_contents[0], compact_prompt]
        config = request_config if model_index == 0 else types.GenerateContentConfig(
            temperature=0,
            top_p=0,
            top_k=1,
            response_mime_type="application/json",
            response_schema=DashboardExtraction,
            max_output_tokens=16384,
        )
        try:
            print(f"Processing image {image_number} using {model}")
            response = client.models.generate_content(
                model=model,
                contents=contents,
                config=config,
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
            print(f"Image {image_number}: {model} failed: {exc}; trying fallback")
        except Exception as exc:
            last_error = exc
            print(f"Image {image_number}: {model} API error: {exc}; trying fallback")
    else:
        raise ValueError(f"Gemini extraction failed for image {image_number}: {last_error}") from last_error

    extraction["items"] = normalize_treemap_items(extraction.get("items", []))
    has_treemap = any(normalize_label(item.get("section")) == "treemap" for item in extraction.get("items", []))

    tile_refinement = (
        extract_treemap_region(image_path)
        if ENABLE_TREEMAP_REFINEMENT and has_treemap
        else {"items": [], "llm_calls": 0, "prompt_tokens": 0, "output_tokens": 0, "total_tokens": 0}
    )
    if ENABLE_TREEMAP_REFINEMENT and tile_refinement["items"]:
        extraction["items"] = [
            item for item in extraction.get("items", [])
            if normalize_label(item.get("section")) != "treemap"
        ] + tile_refinement["items"]

    p_tokens = getattr(usage, "prompt_token_count", 0) or 0
    c_tokens = getattr(usage, "candidates_token_count", 0) or 0
    t_tokens = getattr(usage, "total_token_count", 0) or 0

    extraction["_image_stats"] = {
        "width": width,
        "height": height,
        "pixels": width * height,
        "llm_calls": llm_calls + tile_refinement["llm_calls"],
        "prompt_tokens": p_tokens,
        "prompt_text_tokens": _prompt_token_count,
        "image_input_tokens": max(0, p_tokens - (_prompt_token_count or 0)),
        "output_tokens": c_tokens + tile_refinement["output_tokens"],
        "total_tokens": t_tokens + tile_refinement["total_tokens"],
        "image_hash": image_content_hash(image_path),
    }
    return extraction


def extract_all_images(folder: Path, cache_folder: Path, images: dict[int, Path]) -> dict[int, dict]:
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


# ---------------------------------------------------------------------------
# Comparison Logic
# ---------------------------------------------------------------------------
def normalize_label(text: object) -> str:
    return " ".join(str(text or "").strip().lower().split())


def normalize_value(text: object) -> str:
    return normalize_label(text).replace(",", "")


def extract_comparable_value(text: object) -> str:
    raw = str(text or "")
    match = re.search(r"value\s*:\s*(.*)$", raw, flags=re.IGNORECASE)
    return normalize_value(match.group(1) if match else raw)


def comma_separated_values(text: object) -> tuple[str, ...] | None:
    raw = str(text or "")
    if "," not in raw:
        return None
    values = [normalize_value(value) for value in raw.split(",")]
    return tuple(sorted(values)) if len(values) > 1 and all(values) else None


def values_match_for_item(left: dict, right: dict) -> bool:
    left_values = comma_separated_values(left.get("value"))
    right_values = comma_separated_values(right.get("value"))
    if left_values is not None and right_values is not None:
        return left_values == right_values
    return extract_comparable_value(left.get("value")) == extract_comparable_value(right.get("value"))


def display_value(text: object) -> object:
    raw = str(text or "")
    match = re.search(r"value\s*:\s*(.*)$", raw, flags=re.IGNORECASE)
    return match.group(1).strip() if match else text


def format_chart_component(value: object) -> str:
    value = display_value(value)
    if value is None:
        return ""
    text = str(value).strip()
    if not text or text.lower() in {"null", "none"}:
        return ""
    try:
        if float(text) == 0:
            return ""
    except ValueError:
        pass
    return text


def format_chart_value(value: object) -> str:
    text = display_value(value)
    if not isinstance(text, str) or "," not in text:
        return format_chart_component(text)
    components = [format_chart_component(component) for component in text.split(",")]
    while components and not components[-1]:
        components.pop()
    return ",".join(components)


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
        section = normalize_label(item.get("section"))
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


def sort_and_format_chart_items(items: list[dict]) -> list[dict]:
    """Normalize grouped bars to one left-to-right pair per X-axis category."""
    grouped: dict[tuple[str, str], dict] = {}
    passthrough: list[dict] = []

    for item in items:
        if normalize_label(item.get("section")) != "chart":
            passthrough.append(item)
            continue

        item_name = str(item.get("item_name") or "")
        match = re.match(r"^(.*?) - (Cases Registered|Cases Resolved) \(([^()]+)\)$", item_name)
        if not match:
            if "," in str(item.get("value") or ""):
                item["value"] = format_chart_value(item.get("value"))
            passthrough.append(item)
            continue

        chart_title, series_name, category = match.groups()
        key = (chart_title, category)
        pair = grouped.setdefault(
            key,
            {
                "section": item.get("section") or "Chart",
                "item_name": f"{chart_title} - {category}",
                "series_values": {},
                "confidence": [],
                "original_index": len(grouped),
            },
        )
        pair["series_values"][series_name] = format_chart_component(item.get("value"))
        confidence = item.get("confidence")
        if isinstance(confidence, (int, float)):
            pair["confidence"].append(confidence)

    normalized_pairs = []
    for pair in grouped.values():
        registered = pair["series_values"].get("Cases Registered", "")
        resolved = pair["series_values"].get("Cases Resolved", "")
        normalized_pairs.append(
            {
                "section": pair["section"],
                "item_name": pair["item_name"],
                "value": format_chart_pair(registered, resolved),
                "confidence": mean(pair["confidence"]) if pair["confidence"] else None,
                "_dual_series_clustered": True,
                "_category_order": WEEKDAY_ORDER.get(normalize_label(pair["item_name"]).split("(")[-1].rstrip(")"), 99),
                "_original_index": pair["original_index"],
            }
        )

    normalized_pairs.sort(key=lambda item: (item["_category_order"], item["_original_index"]))
    for item in normalized_pairs:
        item.pop("_category_order", None)
        item.pop("_original_index", None)
    return passthrough + normalized_pairs


def compare_extractions(trendence_data: dict, spartnash_data: dict) -> dict:
    trendence_items = sort_and_format_chart_items(
        normalize_single_series_items(list(trendence_data.get("items", [])))
    )
    spartnash_items = sort_and_format_chart_items(
        normalize_single_series_items(list(spartnash_data.get("items", [])))
    )

    for idx, item in enumerate(trendence_items):
        item["_orig_idx"] = idx
    for idx, item in enumerate(spartnash_items):
        item["_orig_idx"] = idx

    comparison_items: list[dict] = []

    def key_of(item: dict) -> tuple[str, str]:
        return (normalize_label(item.get("section")), normalize_label(item.get("item_name")))

    remaining_spartnash = list(spartnash_items)

    def pop_exact_match(item: dict) -> Optional[dict]:
        key = key_of(item)
        for index, candidate in enumerate(remaining_spartnash):
            if key_of(candidate) == key:
                return remaining_spartnash.pop(index)
        return None

    def pop_best_fuzzy_match(item: dict) -> Optional[dict]:
        section = normalize_label(item.get("section"))
        candidates = [
            c for c in remaining_spartnash if normalize_label(c.get("section")) == section
        ]
        if not candidates:
            return None
        names = [normalize_label(c.get("item_name")) for c in candidates]
        best = difflib.get_close_matches(
            normalize_label(item.get("item_name")), names, n=1, cutoff=FUZZY_MATCH_CUTOFF
        )
        if not best:
            return None
        match = candidates[names.index(best[0])]
        remaining_spartnash.remove(match)
        return match

    for trendence_item in trendence_items:
        match = pop_exact_match(trendence_item) or pop_best_fuzzy_match(trendence_item)
        trendence_value = trendence_item.get("value")
        confidences = [c for c in (trendence_item.get("confidence"),) if isinstance(c, (int, float))]
        orig_idx = trendence_item.get("_orig_idx", 9999)

        if match is None:
            comparison_items.append(
                {
                    "section": trendence_item.get("section") or "Unspecified",
                    "item_name": trendence_item.get("item_name") or "Unnamed item",
                    "trendence_value": trendence_value,
                    "spartnash_value": None,
                    "status": "Trendence Only",
                    "difference": "Only present in Trendence",
                    "confidence": mean(confidences) if confidences else None,
                    "_orig_idx": orig_idx,
                }
            )
            continue

        spartnash_value = match.get("value")
        confidences += [c for c in (match.get("confidence"),) if isinstance(c, (int, float))]
        values_match = values_match_for_item(trendence_item, match)

        comparison_items.append(
            {
                "section": trendence_item.get("section") or match.get("section") or "Unspecified",
                "item_name": trendence_item.get("item_name") or match.get("item_name") or "Unnamed item",
                "trendence_value": trendence_value,
                "spartnash_value": spartnash_value,
                "status": "Match" if values_match else "Different",
                "difference": None if values_match else "Values differ between Trendence and Spartnash",
                "confidence": mean(confidences) if confidences else None,
                "_orig_idx": orig_idx,
            }
        )

    for spartnash_item in remaining_spartnash:
        confidence = spartnash_item.get("confidence")
        comparison_items.append(
            {
                "section": spartnash_item.get("section") or "Unspecified",
                "item_name": spartnash_item.get("item_name") or "Unnamed item",
                "trendence_value": None,
                "spartnash_value": spartnash_item.get("value"),
                "status": "Spartnash Only",
                "difference": "Only present in Spartnash",
                "confidence": confidence if isinstance(confidence, (int, float)) else None,
                "_orig_idx": spartnash_item.get("_orig_idx", 9999) + 1000,
            }
        )

    def comparison_sort_key(item: dict) -> tuple:
        section = normalize_label(item.get("section"))
        section_idx = SECTION_ORDER.get(section, 99)
        item_name = normalize_label(item.get("item_name"))
        orig_idx = item.get("_orig_idx", 9999)

        weekday = next((day for day in WEEKDAY_ORDER if f"({day})" in item_name), None)
        if weekday is not None and " - " in item_name:
            parts = item_name.split(" - ", 1)
            chart_title = parts[0]
            series_label = parts[1] if len(parts) > 1 else ""
            series_order = 0 if "cases registered" in series_label else 1
            return (section_idx, 0, chart_title, WEEKDAY_ORDER[weekday], series_order, orig_idx, item_name)

        return (section_idx, 1, "", -1, 0, orig_idx, item_name)

    comparison_items.sort(key=comparison_sort_key)

    for item in comparison_items:
        item.pop("_orig_idx", None)

    notes = list(trendence_data.get("notes", [])) + list(spartnash_data.get("notes", []))
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
    grouped: dict[tuple[str, str], dict] = {}
    output: list[dict] = []

    for item in items:
        match = re.match(
            r"^(.*?) - (Cases Registered|Cases Resolved) \(([^()]+)\)$",
            str(item.get("item_name") or ""),
        )
        if normalize_label(item.get("section")) != "chart" or not match:
            output.append(item)
            continue

        chart_title, series, category = match.groups()
        key = (chart_title, category)
        combined = grouped.setdefault(
            key,
            {
                "section": item.get("section") or "Chart",
                "item_name": f"{chart_title} ({category})",
                "trendence": {},
                "spartnash": {},
                "confidences": [],
            },
        )
        combined["trendence"][series] = format_chart_component(item.get("trendence_value"))
        combined["spartnash"][series] = format_chart_component(item.get("spartnash_value"))
        if isinstance(item.get("confidence"), (int, float)):
            combined["confidences"].append(item["confidence"])

    output.extend(grouped.values())
    for item in output:
        if "trendence" not in item:
            continue
        registered = "Cases Registered"
        resolved = "Cases Resolved"
        trendence_value = format_chart_pair(
            item["trendence"].get(registered, ""), item["trendence"].get(resolved, "")
        )
        spartnash_value = format_chart_pair(
            item["spartnash"].get(registered, ""), item["spartnash"].get(resolved, "")
        )
        values_match = values_match_for_item(
            {"value": trendence_value},
            {"value": spartnash_value},
        )
        item.update(
            {
                "trendence_value": trendence_value,
                "spartnash_value": spartnash_value,
            "status": "Match" if values_match else "Different",
            "difference": None if values_match else "Values differ between Trendence and Spartnash",
                "confidence": mean(item["confidences"]) if item["confidences"] else None,
            }
        )
        for key in ("trendence", "spartnash", "confidences"):
            item.pop(key, None)
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


def to_dashboard_json(
    image_number: int,
    result: dict,
    trendence_image: Path,
    spartnash_image: Path,
) -> dict:
    """Convert an internal comparison result to the public dashboard JSON shape."""
    items = combine_grouped_chart_rows(result.get("comparison_items", []))
    metrics = calculate_metrics(items)
    match_percentage = metrics["match_percentage"]
    percentage = round(match_percentage * 100) if match_percentage is not None else 0
    overall_status = "MATCH" if percentage == 100 else "MISMATCH" if percentage == 0 else "PARTIAL_MATCH"

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

        visuals.append(
            {
                "visual_name": item.get("item_name") or "Unnamed visual",
                "visual_type": visual_type,
                "section": item.get("section") or "Unspecified",
                "status": "MATCH" if status == "Match" else "MISMATCH",
                "source": {"value": item.get("trendence_value"), "details": None},
                "target": {"value": item.get("spartnash_value"), "details": None},
                "differences": [] if status == "Match" else [item.get("difference") or "Values differ"],
                "confidence": item.get("confidence"),
            }
        )

    return {
        "dashboard_name": result.get("trendence_dashboard_title") or result.get("spartnash_dashboard_title") or f"Dashboard {image_number}",
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
        "visuals": visuals,
    }


def write_json_output(results: list[dict], output_path: Path = OUTPUT_JSON) -> None:
    """Write the structured dashboard comparison JSON output."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(results, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"JSON output saved to:\n{output_path.resolve()}")


def process_dashboard_comparisons(
    trendence_folder: Path = TRENDENCE_FOLDER,
    spartnash_folder: Path = SPARTNASH_FOLDER,
    output_workbook: Path = OUTPUT_WORKBOOK,
) -> None:
    trendence_images = discover_numbered_images(trendence_folder)
    spartnash_images = discover_numbered_images(spartnash_folder)
    image_numbers = sorted(set(trendence_images) | set(spartnash_images))

    if not image_numbers:
        raise ValueError("No numbered dashboard images found in input directories.")

    output_workbook.parent.mkdir(parents=True, exist_ok=True)
    RAW_JSON_FOLDER.mkdir(parents=True, exist_ok=True)

    trendence_extractions = extract_all_images(trendence_folder, TRENDENCE_EXTRACTION_FOLDER, trendence_images)
    spartnash_extractions = extract_all_images(spartnash_folder, SPARTNASH_EXTRACTION_FOLDER, spartnash_images)
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

        try:
            result = compare_extractions(
                trendence_extraction,
                spartnash_extraction,
            )
            (RAW_JSON_FOLDER / f"{image_number}.json").write_text(
                json.dumps(result, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )
            jsonl_results.append(
                {
                    "image_number": image_number,
                    "trendence_image": str(trendence_image),
                    "spartnash_image": str(spartnash_image),
                    **result,
                }
            )
            json_results.append(to_dashboard_json(image_number, result, trendence_image, spartnash_image))
            write_pair_sheet(wb, image_number, result, trendence_image, spartnash_image)
        except Exception as exc:
            write_error_sheet(wb, image_number, f"{type(exc).__name__}: {exc}")

    wb.save(output_workbook)
    write_jsonl_output(jsonl_results)
    write_json_output(json_results)
    print(
        f"\nProcessing finished successfully."
        f"\nOutput workbook saved to:\n{output_workbook.resolve()}"
        f"\nJSONL output saved to:\n{OUTPUT_JSONL.resolve()}"
    )

# /vijykfhwer,k
if __name__ == "__main__":
    process_dashboard_comparisons()