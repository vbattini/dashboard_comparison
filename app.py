"""FastAPI app: upload Spartnash/Trendence dashboard images and run the
Gemini-based comparison, then download the resulting Excel workbook."""

from __future__ import annotations

import asyncio
import io
import json
import uuid
from pathlib import Path
from typing import Literal

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from PIL import Image

import Text_Extraction___ as compare_engine

BASE_DIR = Path(__file__).parent
FRONTEND_DIR = BASE_DIR / "frontend"

FOLDER_MAP = {
    "spartnash": compare_engine.SPARTNASH_FOLDER,
    "trendence": compare_engine.TRENDENCE_FOLDER,
}
SUPPORTED_EXTENSIONS = {".png", ".jpg", ".jpeg"}

app = FastAPI(title="Dashboard Comparison")

# Serve uploaded images so the UI can show thumbnails.
app.mount("/images", StaticFiles(directory=BASE_DIR / "input_images"), name="images")
app.mount("/static", StaticFiles(directory=FRONTEND_DIR / "static"), name="static")

# Guards against overlapping compare runs (Gemini calls are sequential anyway).
_compare_lock = asyncio.Lock()
_compare_jobs: dict[str, dict] = {}
_JOB_FOLDER = compare_engine.OUTPUT_FOLDER / "jobs"
INPUT_COST_PER_MILLION = 0.30
OUTPUT_COST_PER_MILLION = 2.50
USD_TO_INR = 95.50


def _lossless_optimize_png(content: bytes) -> bytes:
    """Reduce PNG file size while preserving every pixel and its dimensions."""
    with Image.open(io.BytesIO(content)) as image:
        optimized = io.BytesIO()
        image.save(optimized, format="PNG", optimize=True, compress_level=9)
        return optimized.getvalue()


def _folder_path(folder: str) -> Path:
    if folder not in FOLDER_MAP:
        raise HTTPException(status_code=400, detail="folder must be 'spartnash' or 'trendence'")
    path = FOLDER_MAP[folder]
    path.mkdir(parents=True, exist_ok=True)
    return path


def _job_path(job_id: str) -> Path:
    _JOB_FOLDER.mkdir(parents=True, exist_ok=True)
    return _JOB_FOLDER / f"{job_id}.json"


def _save_job(job_id: str, job: dict) -> None:
    _job_path(job_id).write_text(json.dumps(job), encoding="utf-8")


def _load_job(job_id: str) -> dict | None:
    job = _compare_jobs.get(job_id)
    if job is not None:
        return job
    path = _job_path(job_id)
    if not path.is_file():
        return None
    return json.loads(path.read_text(encoding="utf-8"))


def _has_running_job() -> bool:
    if any(job["status"] == "running" for job in _compare_jobs.values()):
        return True
    for path in _JOB_FOLDER.glob("*.json"):
        try:
            job = json.loads(path.read_text(encoding="utf-8"))
            if job.get("status") != "running":
                continue
            if path.stem in _compare_jobs:
                return True

            # A process restart or forced shutdown can leave a persisted job
            # marked running even though no worker owns it anymore.
            job.update(
                {
                    "status": "failed",
                    "error": "Comparison interrupted before the server restarted.",
                }
            )
            path.write_text(json.dumps(job), encoding="utf-8")
        except (OSError, json.JSONDecodeError):
            continue
    return False


@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    return HTMLResponse((FRONTEND_DIR / "index.html").read_text(encoding="utf-8"))


@app.get("/api/images/{folder}")
def list_images(folder: Literal["spartnash", "trendence"]):
    path = _folder_path(folder)
    images = compare_engine.discover_numbered_images(path)
    return [
        {"number": number, "filename": file_path.name, "url": f"/images/{path.name}/{file_path.name}"}
        for number, file_path in sorted(images.items())
    ]


@app.post("/api/upload/{folder}")
async def upload_images(folder: Literal["spartnash", "trendence"], files: list[UploadFile] = File(...)):
    path = _folder_path(folder)
    saved = []
    errors = []
    uploads = []
    for upload in files:
        name = Path(upload.filename or "").name
        stem, ext = Path(name).stem, Path(name).suffix.lower()
        if ext not in SUPPORTED_EXTENSIONS:
            errors.append(f"{name}: file must be PNG or JPEG")
            continue
        if len(files) == 1 and not stem.isdigit():
            name = f"1{ext}"
        content = await upload.read()
        if ext == ".png":
            content = _lossless_optimize_png(content)
        uploads.append((name, content))

    if not uploads:
        raise HTTPException(status_code=400, detail="; ".join(errors))

    # Each upload is a replacement for that folder, so old sample images do not
    # remain mixed with the user's new image set.
    for existing in path.iterdir():
        if existing.is_file() and existing.suffix.lower() in SUPPORTED_EXTENSIONS:
            existing.unlink()
    cache_folder = (
        compare_engine.SPARTNASH_EXTRACTION_FOLDER
        if folder == "spartnash"
        else compare_engine.TRENDENCE_EXTRACTION_FOLDER
    )
    for cached_file in cache_folder.glob("*.json"):
        cached_file.unlink()
    for name, content in uploads:
        (path / name).write_bytes(content)
        saved.append(name)

    return {"saved": saved, "errors": errors}


@app.delete("/api/images/{folder}/{filename}")
def delete_image(folder: Literal["spartnash", "trendence"], filename: str):
    path = _folder_path(folder)
    target = path / Path(filename).name
    if not target.is_file():
        raise HTTPException(status_code=404, detail="Image not found")
    target.unlink()
    return {"deleted": filename}


@app.post("/api/compare")
async def compare():
    if _compare_lock.locked() or _has_running_job():
        raise HTTPException(status_code=409, detail="A comparison is already running")

    try:
        trendence_images = compare_engine.discover_numbered_images(compare_engine.TRENDENCE_FOLDER)
        spartnash_images = compare_engine.discover_numbered_images(compare_engine.SPARTNASH_FOLDER)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    if not trendence_images or not spartnash_images:
        raise HTTPException(
            status_code=400,
            detail="Upload at least one Spartnash image and one Trendence image before comparing",
        )

    job_id = uuid.uuid4().hex
    job = {"status": "running"}
    _compare_jobs[job_id] = job
    _save_job(job_id, job)
    asyncio.create_task(_run_compare_job(job_id))
    return {
        "job_id": job_id,
        "status": "running",
        "trendence_images": len(trendence_images),
        "spartnash_images": len(spartnash_images),
        "common_image_numbers": len(set(trendence_images) & set(spartnash_images)),
    }


async def _run_compare_job(job_id: str) -> None:
    async with _compare_lock:
        try:
            await asyncio.to_thread(compare_engine.process_dashboard_comparisons)
            job = {"status": "completed", "result": _build_summary()}
            _compare_jobs[job_id] = job
            _save_job(job_id, job)
        except Exception as exc:
            job = {
                "status": "failed",
                "error": f"{type(exc).__name__}: {exc}",
            }
            _compare_jobs[job_id] = job
            _save_job(job_id, job)


@app.get("/api/compare/{job_id}")
def compare_status(job_id: str):
    job = _load_job(job_id)
    if job is None:
        if compare_engine.OUTPUT_WORKBOOK.is_file():
            return {"status": "completed", **_build_summary()}
        raise HTTPException(status_code=404, detail="Comparison job not found and no report is available yet")
    if job["status"] == "failed":
        raise HTTPException(status_code=500, detail=job["error"])
    if job["status"] == "completed":
        return {"status": "completed", **job["result"]}
    return {"status": "running", "job_id": job_id}


@app.post("/api/extract")
async def extract():
    if _compare_lock.locked():
        raise HTTPException(status_code=409, detail="An extraction or comparison is already running")

    async with _compare_lock:
        try:
            trendence_images = compare_engine.discover_numbered_images(compare_engine.TRENDENCE_FOLDER)
            spartnash_images = compare_engine.discover_numbered_images(compare_engine.SPARTNASH_FOLDER)
            if not trendence_images or not spartnash_images:
                raise HTTPException(
                    status_code=400,
                    detail="Upload at least one Spartnash image and one Trendence image before extracting",
                )
            trendence_extractions = await asyncio.to_thread(
                compare_engine.extract_all_images,
                compare_engine.TRENDENCE_FOLDER,
                compare_engine.TRENDENCE_EXTRACTION_FOLDER,
                trendence_images,
            )
            spartnash_extractions = await asyncio.to_thread(
                compare_engine.extract_all_images,
                compare_engine.SPARTNASH_FOLDER,
                compare_engine.SPARTNASH_EXTRACTION_FOLDER,
                spartnash_images,
            )
            compare_engine.write_aggregate_extractions(trendence_extractions, spartnash_extractions)
        except FileNotFoundError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

    return {
        "trendence_file": str(compare_engine.TRENDENCE_EXTRACTION_JSON),
        "spartnash_file": str(compare_engine.SPARTNASH_EXTRACTION_JSON),
        "trendence_images": len(trendence_extractions),
        "spartnash_images": len(spartnash_extractions),
    }


@app.get("/api/results")
def get_results():
    return _build_summary()


@app.get("/api/download-json")
def download_json():
    json_path = compare_engine.OUTPUT_JSON
    if not json_path.is_file():
        raise HTTPException(status_code=404, detail="No comparison JSON has been generated yet")
    return FileResponse(json_path, media_type="application/json", filename=json_path.name)


def _build_summary() -> dict:
    workbook_path = compare_engine.OUTPUT_WORKBOOK
    pairs = []
    current_trendence = compare_engine.discover_numbered_images(compare_engine.TRENDENCE_FOLDER)
    current_spartnash = compare_engine.discover_numbered_images(compare_engine.SPARTNASH_FOLDER)
    current_numbers = set(current_trendence) | set(current_spartnash)
    image_stats = []
    for folder_name, images, cache_folder in (
        ("Trendence", current_trendence, compare_engine.TRENDENCE_EXTRACTION_FOLDER),
        ("Spartnash", current_spartnash, compare_engine.SPARTNASH_EXTRACTION_FOLDER),
    ):
        for number, image_path in sorted(images.items()):
            width, height = compare_engine.image_dimensions(image_path)
            stats = {}
            cache_path = cache_folder / f"{number}.json"
            if cache_path.is_file():
                import json as json_module

                stats = json_module.loads(cache_path.read_text(encoding="utf-8")).get("_image_stats", {})
            prompt_tokens = stats.get("prompt_text_tokens", stats.get("prompt_tokens"))
            image_tokens = stats.get("image_input_tokens")
            if image_tokens is None and stats.get("total_tokens") is not None and stats.get("output_tokens") is not None:
                image_tokens = stats["total_tokens"] - stats["output_tokens"] - prompt_tokens
            total_input_tokens = (
                prompt_tokens + image_tokens
                if prompt_tokens is not None and image_tokens is not None
                else stats.get("prompt_tokens")
            )
            output_tokens = stats.get("output_tokens")
            total_tokens = stats.get("total_tokens")
            input_cost = total_input_tokens * INPUT_COST_PER_MILLION / 1_000_000 if total_input_tokens is not None else None
            output_cost = output_tokens * OUTPUT_COST_PER_MILLION / 1_000_000 if output_tokens is not None else None
            total_cost = input_cost + output_cost if input_cost is not None and output_cost is not None else None
            image_stats.append(
                {
                    "folder": folder_name,
                    "image": str(number),
                    "width": width,
                    "height": height,
                    "pixels": width * height,
                    "prompt_tokens": prompt_tokens,
                    "image_input_tokens": image_tokens,
                    "total_input_tokens": total_input_tokens,
                    "output_tokens": output_tokens,
                    "total_tokens": total_tokens,
                    "input_cost_usd": input_cost,
                    "output_cost_usd": output_cost,
                    "total_cost_usd": total_cost,
                    "total_cost_inr": total_cost * USD_TO_INR if total_cost is not None else None,
                    "llm_calls": stats.get("llm_calls", 0),
                }
            )
    if compare_engine.RAW_JSON_FOLDER.is_dir():
        for json_path in sorted(
            compare_engine.RAW_JSON_FOLDER.glob("*.json"),
            key=lambda p: int(p.stem) if p.stem.isdigit() else p.stem,
        ):
            if not json_path.stem.isdigit() or int(json_path.stem) not in current_numbers:
                continue
            import json as json_module

            result = json_module.loads(json_path.read_text(encoding="utf-8"))
            items = result.get("comparison_items", [])
            metrics = compare_engine.calculate_metrics(items)
            pairs.append(
                {
                    "pair": json_path.stem,
                    "trendence_title": result.get("trendence_dashboard_title"),
                    "spartnash_title": result.get("spartnash_dashboard_title"),
                    "total_items": metrics["total"],
                    "matches": metrics["counts"]["Match"],
                    "differences": metrics["counts"]["Different"],
                    "trendence_only": metrics["counts"]["Trendence Only"],
                    "spartnash_only": metrics["counts"]["Spartnash Only"],
                    "uncertain": metrics["counts"]["Uncertain"],
                    "match_percentage": metrics["match_percentage"],
                }
            )

    return {
        "pairs": pairs,
        "image_stats": image_stats,
        "total_llm_calls": sum(item["llm_calls"] for item in image_stats),
        "total_tokens": sum(item["total_tokens"] or 0 for item in image_stats),
        "total_prompt_tokens": sum(item["prompt_tokens"] or 0 for item in image_stats),
        "total_image_tokens": sum(item["image_input_tokens"] or 0 for item in image_stats),
        "total_input_tokens": sum(item["total_input_tokens"] or 0 for item in image_stats),
        "total_output_tokens": sum(item["output_tokens"] or 0 for item in image_stats),
        "total_input_cost_usd": sum(item["input_cost_usd"] or 0 for item in image_stats),
        "total_output_cost_usd": sum(item["output_cost_usd"] or 0 for item in image_stats),
        "total_cost_usd": sum(item["total_cost_usd"] or 0 for item in image_stats),
        "total_cost_inr": sum(item["total_cost_inr"] or 0 for item in image_stats),
        "workbook_available": workbook_path.is_file(),
        "download_url": "/api/download" if workbook_path.is_file() else None,
        "json_available": compare_engine.OUTPUT_JSON.is_file(),
        "json_download_url": "/api/download-json" if compare_engine.OUTPUT_JSON.is_file() else None,
    }


@app.get("/api/download")
def download_workbook():
    workbook_path = compare_engine.OUTPUT_WORKBOOK
    if not workbook_path.is_file():
        raise HTTPException(status_code=404, detail="No comparison workbook has been generated yet")
    return FileResponse(
        workbook_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=workbook_path.name,
    )
