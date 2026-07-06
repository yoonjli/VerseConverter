from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import Response, FileResponse
from fastapi.staticfiles import StaticFiles
from pathlib import Path

from parser import parse_verses, match_verses
from builder import build_pptx

app = FastAPI(title="Bilingual Scripture PPTX Generator")

# Serve the frontend
app.mount("/static", StaticFiles(directory="static"), name="static")


@app.get("/")
async def root():
    return FileResponse("static/index.html")


@app.post("/parse")
async def parse(
    korean: UploadFile = File(...),
    english: UploadFile = File(...),
):
    """Parse two scripture files and return matched verse pairs for preview."""
    try:
        ko_text = (await korean.read()).decode("utf-8")
        en_text = (await english.read()).decode("utf-8")
    except UnicodeDecodeError:
        raise HTTPException(status_code=400, detail="Files must be UTF-8 encoded text.")

    ko_verses = parse_verses(ko_text)
    en_verses = parse_verses(en_text)
    matched, only_ko, only_en = match_verses(ko_verses, en_verses)

    return {
        "matched": matched,
        "only_ko": only_ko,
        "only_en": only_en,
        "ko_count": len(ko_verses),
        "en_count": len(en_verses),
        "matched_count": len(matched),
    }


@app.post("/generate")
async def generate(
    korean: UploadFile = File(...),
    english: UploadFile = File(...),
):
    """Generate and return a PPTX file from two scripture text files."""
    try:
        ko_text = (await korean.read()).decode("utf-8")
        en_text = (await english.read()).decode("utf-8")
    except UnicodeDecodeError:
        raise HTTPException(status_code=400, detail="Files must be UTF-8 encoded text.")

    ko_verses = parse_verses(ko_text)
    en_verses = parse_verses(en_text)
    matched, _, _ = match_verses(ko_verses, en_verses)

    if not matched:
        raise HTTPException(status_code=400, detail="No matching verses found between the two files.")

    pptx_bytes = build_pptx(matched)

    # Derive a filename from the uploaded files
    stem = Path(korean.filename).stem if korean.filename else "scripture"
    filename = f"{stem}_bilingual.pptx"

    return Response(
        content=pptx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
