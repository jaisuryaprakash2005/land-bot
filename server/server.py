"""
Probridge Land Bot - local AI server.

Runs on your laptop. It:
  * turns voice notes into text           (Whisper, free, offline)
  * pulls land details out of the text    (Ollama + Llama, free, offline)
  * serves the chat screen used by the Excel add-in and the phone app

Start it with start.ps1 (or: python server/server.py).
"""

import json
import os
import re
import tempfile
import threading
import time
import uuid
from pathlib import Path

import requests
import uvicorn
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles

ROOT = Path(__file__).resolve().parent.parent
WEB_DIR = ROOT / "web"
CONFIG = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))

OLLAMA_URL = CONFIG.get("ollama_url", "http://localhost:11434").rstrip("/")
MODEL = CONFIG.get("ollama_model", "llama3.1:8b")
AI_KEYS = [c["key"] for c in CONFIG["columns"] if c.get("fill") == "ai"]

# ---------------------------------------------------------------------------
# AI instructions
# ---------------------------------------------------------------------------

PROMPT = """You read land offers that mediators send to a real estate company near
Chengalpattu / Chennai, Tamil Nadu. Messages may be in English, Tanglish, or translated Tamil.
Extract the details into the JSON schema.

RULES
- If a detail is NOT in the message, return "" (empty string). Never guess or invent.
- Always write the values in English (e.g. "Dry", not "Punjai"; "Burial ground", not "Sudukadu").
- exact_area: village / place name, plus distance from a known town if given.
  Example: "Alanjeri - 70km from Chengalpattu". Do not include greetings like "Sir" or "Good morning".
- source: "Mediator" unless the message says the owner is selling directly, then "Owner".
- mediator_name: the person sending/offering the land. Usually the name written right before
  or after the phone number, or after "this is" / "contact". Example: "Kumuda 8883325106" -> "Kumuda".
- number: phone number exactly as written.
- extent: keep the mediator's units (ac, cent, ground, sq ft) and any split,
  e.g. "6.5 ac (road side 2 ac dry + 4.5 ac wet)".
- cost_per_acre: exactly as the mediator said, with its unit, e.g. "75L per ac".
  If the price is a TOTAL for the whole land (not per acre), write "TOTAL: 4cr".
  If it is per cent or per sq ft, keep that unit, e.g. "6,50,000 per cent".
- land_category: Dry, Wet, Dry+Wet, Plot, Commercial, or as stated. "Patta" is a document, not a category.
- road_width: width of the ROAD itself, e.g. "40 feet bus road".
- frontage: length of the land touching the road, e.g. "140+ feet on GST road".
  If the message says "frontage", that number goes in frontage, NOT road_width.
- land_level: height of the land compared to the road, e.g. "5 feet below road", "road level".
- negative: burial ground, dump yard, water body (lake, pond, canal), govt sealing,
  high-tension line, temple land, etc. Write "Nil" ONLY if the message clearly says there are none.
- google_pin: latitude, longitude or a Google Maps link, copied exactly.

WORDS YOU MAY SEE
nanjai = wet land, punjai = dry land, manai = plot, L / lakh / lakhs / lacs / lox = lakh,
cr / crore = crore, ac / acre / acres, cent, ground (2400 sq ft), adi = feet,
patta / chitta = land documents, illa / illai = no / not there,
sudukadu / mayanam = burial ground, eri / kulam / kuttai = water body,
GST road / NH = national highway, ECR / OMR = main roads.

EXAMPLE MESSAGE
Kupan sir, Alanjeri 70km from chengalpattu, 6.5 acre road side 2ac dry 4.5ac wet,
75L per acre, 40 feet bus road, land 5 feet low from road, no burial no water body.
location 12.647456, 79.825731

EXAMPLE ANSWER
{"exact_area": "Alanjeri - 70km from Chengalpattu", "source": "Mediator",
 "mediator_name": "Kupan", "number": "", "extent": "6.5 ac (road side 2 ac dry + 4.5 ac wet)",
 "cost_per_acre": "75L per ac", "land_category": "Dry+Wet", "road_width": "40 feet bus road",
 "frontage": "", "land_level": "5 feet below road", "negative": "Nil",
 "google_pin": "12.647456, 79.825731"}"""

SCHEMA = {
    "type": "object",
    "properties": {k: {"type": "string"} for k in AI_KEYS},
    "required": AI_KEYS,
}

# ---------------------------------------------------------------------------
# Checks done in plain code (more reliable than the AI for numbers)
# ---------------------------------------------------------------------------

PHONE_RE = re.compile(r"(?<!\d)(?:\+?91[\s-]?)?([6-9]\d{4})[\s-]?(\d{5})(?!\d)")
COORD_RE = re.compile(r"(?<![\d.])(-?\d{1,2}\.\d{4,})\s*,\s*(-?\d{1,3}\.\d{4,})")
MAPS_RE = re.compile(
    r"https?://(?:maps\.app\.goo\.gl|goo\.gl/maps|(?:www\.)?google\.[a-z.]+/maps|maps\.google\.[a-z.]+)\S*",
    re.I,
)
# Fields whose words should appear in the message (to catch AI guesses)
GROUNDED_KEYS = {"exact_area", "mediator_name", "extent", "cost_per_acre",
                 "road_width", "frontage", "land_level"}


def find_phone(text: str) -> str:
    m = PHONE_RE.search(text or "")
    return m.group(1) + m.group(2) if m else ""


def extract_coords_from_url(url: str) -> str:
    """Extract latitude, longitude from common Google Maps URL formats."""
    # Format: @12.647456,79.825737 or 12.647456,79.825737
    m = re.search(r"@(-?\d{1,2}\.\d{4,})\s*,\s*(-?\d{1,3}\.\d{4,})", url)
    if m:
        return f"{m.group(1)}, {m.group(2)}"
    # Query parameters: ?q=12.647456,79.825737 or query= or ll=
    m = re.search(r"[?&](?:q|query|ll)=(-?\d{1,2}\.\d{4,})\s*,\s*(-?\d{1,3}\.\d{4,})", url)
    if m:
        return f"{m.group(1)}, {m.group(2)}"
    # Data parameters: !3d12.647456!4d79.825737
    m = re.search(r"!3d(-?\d{1,2}\.\d{4,})!4d(-?\d{1,3}\.\d{4,})", url)
    if m:
        return f"{m.group(1)}, {m.group(2)}"
    return ""


def resolve_maps_url(url: str) -> str:
    """Follow redirects on shortened Google Maps links (e.g. maps.app.goo.gl) to extract coordinates."""
    try:
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
        resp = requests.get(url, headers=headers, allow_redirects=True, timeout=5, stream=True)
        final_url = resp.url
        coords = extract_coords_from_url(final_url)
        if coords:
            return coords
    except Exception:
        pass
    return ""


def find_pin(text: str) -> str:
    # 1. Direct coordinates written in the message (e.g. 12.647456, 79.825737)
    m = COORD_RE.search(text or "")
    if m:
        return f"{m.group(1)}, {m.group(2)}"

    # 2. Google Maps link
    m = MAPS_RE.search(text or "")
    if m:
        url = m.group(0).rstrip(").,")
        coords = extract_coords_from_url(url)
        if not coords and ("goo.gl" in url or "maps.app" in url):
            coords = resolve_maps_url(url)
        return coords if coords else url
    return ""


OWNER_RE = re.compile(r"\bowner\b|\bdirect\s+(sale|owner|party)\b|\bno\s+(broker|commission|mediator)", re.I)
GREETING_RE = re.compile(r"^\s*((good\s+(morning|afternoon|evening))|hello|hi|vanakkam|dear|sir|madam|mam)\b[\s,.!-]*",
                         re.I)
NOT_NAMES = {"call", "contact", "number", "mobile", "phone", "ph", "cell", "sir", "madam", "mam",
             "whatsapp", "no", "num", "is", "my", "me", "at", "on", "to", "details"}
NAME_BEFORE = re.compile(r"([A-Z][a-zA-Z]{2,})(?:\s+(?:sir|madam|mam|anna|akka))?\s*[-:,]?\s*(?:\+?91[\s-]?)?[6-9]\d{4}[\s-]?\d{5}")
NAME_AFTER = re.compile(r"(?:this is|i am|contact|call|name)\s*[:-]?\s*([A-Z][a-zA-Z]{2,})", re.I)


def find_name(text: str) -> str:
    for rx in (NAME_BEFORE, NAME_AFTER):
        for m in rx.finditer(text or ""):
            name = m.group(1)
            if name.lower() not in NOT_NAMES:
                return name[0].upper() + name[1:]
    return ""


TAMIL_TO_ENGLISH = [(re.compile(rf"\b{t}\b", re.I), e) for t, e in [
    (r"punjai(\s+nilam)?", "Dry"), (r"nanjai(\s+nilam)?", "Wet"), (r"manai", "Plot"),
    (r"sudukadu|sudugadu|mayanam", "Burial ground"), (r"eri|kulam|kuttai", "Water body"),
]]

COMMON = {"feet", "foot", "road", "from", "near", "acre", "acres", "per", "the", "and",
          "land", "total", "cent", "cents", "with", "side", "bus", "main", "below", "above",
          "level", "lakh", "lakhs", "crore", "village", "town", "limit", "dry", "wet"}


def numbers(s: str) -> set:
    return set(re.findall(r"\d+(?:\.\d+)?", (s or "").replace(",", "")))


def words(s: str) -> set:
    return {w for w in re.findall(r"[a-z]+", (s or "").lower()) if len(w) >= 3 and w not in COMMON}


def looks_grounded(value: str, text: str) -> bool:
    """True if the value's numbers (or, without numbers, its key words) appear in the message."""
    nums = numbers(value)
    if nums:
        return bool(nums & numbers(text))
    w = words(value)
    return not w or bool(w & words(text))


def clean_up(fields: dict, text: str) -> tuple[dict, dict]:
    """Fix numbers with regex and flag values the AI may have guessed."""
    out = {k: (fields.get(k) or "").strip() for k in AI_KEYS}
    warnings = {}

    # Phone: trust the message text first
    phone = find_phone(text) or find_phone(out.get("number", ""))
    if "number" in out:
        out["number"] = phone
        if not phone:
            warnings["number"] = "No phone number found in the message."

    # Google pin: copy exactly from the message
    pin = find_pin(text)
    if "google_pin" in out:
        if pin:
            out["google_pin"] = pin
        elif out["google_pin"]:
            out["google_pin"] = ""  # AI invented a location that is not in the message

    # Source: "Owner" only when the message really says so
    if "source" in out:
        out["source"] = "Owner" if OWNER_RE.search(text or "") else CONFIG.get("default_source", "Mediator")

    # Mediator name: if the AI missed it, take the name written next to the phone number
    if "mediator_name" in out and not out["mediator_name"]:
        name = find_name(text)
        if name:
            out["mediator_name"] = name
            warnings["mediator_name"] = "Name taken from next to the phone number. Please check."

    # Exact area: remove greetings the AI sometimes keeps
    if out.get("exact_area"):
        out["exact_area"] = GREETING_RE.sub("", out["exact_area"]).strip(" ,.-")

    # Frontage: the AI sometimes puts the frontage figure into road width
    if out.get("road_width") and not out.get("frontage"):
        front_nums = set()
        for m in re.finditer(r"frontage", text or "", re.I):
            front_nums |= numbers(text[max(0, m.start() - 30): m.end() + 30])
        if numbers(out["road_width"]) & front_nums:
            out["frontage"], out["road_width"] = out["road_width"], ""

    # Patta is a document, not a land category
    if re.fullmatch(r"(patta|chitta)( land)?", out.get("land_category", ""), re.I):
        out["land_category"] = ""

    # Keep the sheet in English
    for k in ("land_category", "negative", "extent"):
        if out.get(k):
            for rx, eng in TAMIL_TO_ENGLISH:
                out[k] = rx.sub(eng, out[k])

    cost = out.get("cost_per_acre", "")
    if cost.upper().startswith("TOTAL"):
        warnings["cost_per_acre"] = "Looks like a TOTAL price, not per acre. Please check."
    elif re.search(r"cent|sq\.?\s*ft|sqft|ground", cost, re.I):
        warnings["cost_per_acre"] = "Price is not per acre (per cent / sq ft / ground). Please check."

    neg = out.get("negative", "")
    if neg.lower() in ("nil", "none", "no") and not re.search(
            r"\b(no|nil|none|illa|illai|without)\b", text or "", re.I):
        out["negative"] = ""
        warnings["negative"] = "Mediator did not clearly say 'no negatives'. Please confirm."

    for k in GROUNDED_KEYS & set(out):
        v = out[k]
        if v and not looks_grounded(v, text):
            warnings[k] = "Not found in the message - AI may have guessed. Please check."

    if not out.get("exact_area"):
        warnings["exact_area"] = "Location not found in the message."
    return out, warnings


# ---------------------------------------------------------------------------
# Ollama (AI) and Whisper (voice)
# ---------------------------------------------------------------------------

def ollama_extract(text: str) -> dict:
    r = requests.post(f"{OLLAMA_URL}/api/chat", timeout=600, json={
        "model": MODEL,
        "messages": [
            {"role": "system", "content": PROMPT},
            {"role": "user", "content": text},
        ],
        "format": SCHEMA,
        "stream": False,
        "keep_alive": "60m",
        "options": {"temperature": 0, "num_ctx": 4096, "num_predict": 700},
    })
    if r.status_code == 404:
        raise HTTPException(503, f"AI model '{MODEL}' is not installed. Run: ollama pull {MODEL}")
    r.raise_for_status()
    return json.loads(r.json()["message"]["content"])


_whisper = None
_whisper_lock = threading.Lock()


def get_whisper():
    global _whisper
    with _whisper_lock:
        if _whisper is None:
            from faster_whisper import WhisperModel  # loaded only when first needed
            _whisper = WhisperModel(CONFIG.get("whisper_model", "small"),
                                    device="cpu", compute_type="int8")
        return _whisper


def transcribe_file(path: str) -> dict:
    segments, info = get_whisper().transcribe(
        path, task=CONFIG.get("whisper_task", "translate"), vad_filter=True,
        initial_prompt=CONFIG.get("whisper_hint") or None)  # helps spell local places and terms
    text = " ".join(s.text.strip() for s in segments).strip()
    return {"text": text, "language": info.language}


def warm_up():
    """Load the AI model into memory at start, so the first message is faster."""
    try:
        requests.post(f"{OLLAMA_URL}/api/generate",
                      json={"model": MODEL, "prompt": "", "keep_alive": "60m"}, timeout=600)
        print(f"[ok] AI model {MODEL} loaded")
    except Exception as e:  # noqa: BLE001
        print(f"[warn] Could not load AI model yet: {e}")


# ---------------------------------------------------------------------------
# Web server
# ---------------------------------------------------------------------------

app = FastAPI(title="Probridge Land Bot")
_shared: dict[str, dict] = {}  # items shared from WhatsApp, kept for 1 hour


@app.middleware("http")
async def no_cache(request: Request, call_next):
    resp = await call_next(request)
    if request.url.path.startswith("/app"):
        resp.headers["Cache-Control"] = "no-cache"
    return resp


@app.get("/")
def home():
    return RedirectResponse("/app/")


@app.get("/api/health")
def health():
    info = {"server": True, "model": MODEL, "ollama": False, "model_installed": False}
    try:
        tags = requests.get(f"{OLLAMA_URL}/api/tags", timeout=5).json()
        info["ollama"] = True
        names = {m["name"] for m in tags.get("models", [])}
        info["model_installed"] = MODEL in names or f"{MODEL}:latest" in names
    except Exception:  # noqa: BLE001
        pass
    return info


@app.get("/api/config")
def public_config():
    return {
        "columns": CONFIG["columns"],
        "excel": CONFIG["excel"],
        "microsoft_login": CONFIG["microsoft_login"],
        "status_options": CONFIG["status_options"],
        "default_status": CONFIG["default_status"],
        "default_source": CONFIG.get("default_source", "Mediator"),
    }


@app.post("/api/extract")
def extract(body: dict):
    text = (body.get("text") or "").strip()
    if not text:
        raise HTTPException(400, "Message is empty.")
    if len(text) > 6000:
        raise HTTPException(400, "Message is too long. Please send one property at a time.")
    start = time.time()
    try:
        raw = ollama_extract(text)
    except requests.ConnectionError:
        raise HTTPException(503, "Ollama is not running. Open the Ollama app and try again.")
    fields, warnings = clean_up(raw, text)
    return {"fields": fields, "warnings": warnings, "seconds": round(time.time() - start, 1)}


async def _save_upload(file: UploadFile) -> str:
    suffix = Path(file.filename or "voice").suffix or ".audio"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as f:
        f.write(await file.read())
        return f.name


@app.post("/api/transcribe")
async def transcribe(file: UploadFile = File(...)):
    path = await _save_upload(file)
    try:
        return transcribe_file(path)
    except Exception as e:  # noqa: BLE001
        raise HTTPException(400, f"Could not read this audio file: {e}")
    finally:
        os.remove(path)


@app.post("/app/share")
async def share_target(title: str = Form(""), text: str = Form(""), url: str = Form(""),
                       audio: list[UploadFile] | None = File(None)):
    """Receives text / voice notes shared from WhatsApp (Android 'Share' menu)."""
    now = time.time()
    for k in [k for k, v in _shared.items() if now - v["t"] > 3600]:
        _shared.pop(k, None)
    item = {"t": now, "text": "\n".join(p for p in (title, text, url) if p), "audio": []}
    for f in audio or []:
        if f.filename:
            item["audio"].append(await _save_upload(f))
    sid = uuid.uuid4().hex
    _shared[sid] = item
    return RedirectResponse(f"/app/?shared={sid}", status_code=303)


@app.get("/api/shared/{sid}")
def get_shared(sid: str):
    item = _shared.pop(sid, None)
    if not item:
        raise HTTPException(404, "Shared item expired. Please share it again.")
    parts = [item["text"]] if item["text"] else []
    for path in item["audio"]:
        try:
            parts.append(transcribe_file(path)["text"])
        finally:
            os.remove(path)
    return {"text": "\n".join(parts).strip(), "had_audio": bool(item["audio"])}


@app.exception_handler(requests.HTTPError)
def ollama_error(_req, exc):
    return JSONResponse({"detail": f"AI error: {exc}"}, status_code=502)


app.mount("/app", StaticFiles(directory=WEB_DIR, html=True), name="app")


def main():
    cert_dir = Path.home() / ".office-addin-dev-certs"
    cert, key = cert_dir / "localhost.crt", cert_dir / "localhost.key"
    threading.Thread(target=warm_up, daemon=True).start()
    if cert.exists() and key.exists():
        print("\n  Excel add-in / laptop:  https://localhost:8443/app/\n")
        uvicorn.run(app, host="127.0.0.1", port=8443,
                    ssl_certfile=str(cert), ssl_keyfile=str(key))
    else:
        print("\n  [warn] HTTPS certificate not found - Excel add-in needs it (run setup.ps1).")
        print("  Test page only:  http://localhost:8000/app/\n")
        uvicorn.run(app, host="127.0.0.1", port=8000)


if __name__ == "__main__":
    main()
