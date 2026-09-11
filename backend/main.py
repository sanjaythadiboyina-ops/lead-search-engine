import os

import csv

import re

import base64

import hashlib

import hmac

import json

import secrets

from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo

from io import BytesIO

from pathlib import Path

from typing import Optional, List

from collections import defaultdict, deque

from urllib.parse import quote_plus, urlparse

from urllib import robotparser

import jwt

import psycopg

import requests

from dotenv import load_dotenv

from fastapi import FastAPI, HTTPException, Depends, Header, Query

from fastapi.middleware.cors import CORSMiddleware

from fastapi.responses import FileResponse, StreamingResponse

from openpyxl import Workbook

from pydantic import BaseModel, Field

# ============================================================

# Configuration

# ============================================================

BASE_DIR = Path(__file__).resolve().parent

PROJECT_DIR = BASE_DIR.parent

FRONTEND_FILE = PROJECT_DIR / "frontend" / "index.html"

load_dotenv(PROJECT_DIR / ".env")

load_dotenv()

DATABASE_URL = os.getenv("DATABASE_URL")

JWT_SECRET = os.getenv("JWT_SECRET") or secrets.token_urlsafe(32)

JWT_ALGORITHM = os.getenv("JWT_ALGORITHM", "HS256")

LATLNG_API_KEY = os.getenv("LATLNG_API_KEY")

SERPAPI_API_KEY = os.getenv("SERPAPI_API_KEY")

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

JWT_EXPIRE_HOURS = int(os.getenv("JWT_EXPIRE_HOURS", "168"))

LATLNG_GEOCODE_URL = "https://api.latlng.work/api"

LATLNG_PLACES_NEARBY_URL = "https://api.latlng.work/v1/places/nearby"

LATLNG_PLACES_SEARCH_URL = "https://api.latlng.work/v1/places/search"

GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.5-flash-lite")

SEARCH_RADIUS_METERS = min(5000, max(250, int(os.getenv("SEARCH_RADIUS_METERS", "5000"))))

MAX_LIMIT = min(100, max(1, int(os.getenv("MAX_SEARCH_RESULTS", "50"))))

CITY_WIDE_MAX_ZONES = min(9, max(1, int(os.getenv("CITY_WIDE_MAX_ZONES", "9"))))

CITY_WIDE_GRID_SPACING_METERS = min(4500, max(1000, int(os.getenv("CITY_WIDE_GRID_SPACING_METERS", "3500"))))

ENRICH_TIMEOUT = min(15, max(3, int(os.getenv("ENRICH_TIMEOUT_SECONDS", "8"))))

app = FastAPI(title="Lead Searching Engine", version="4.0")
_SEARCH_EVENTS = defaultdict(deque)
SEARCH_WINDOW_SECONDS = 300
SEARCH_MAX_REQUESTS = 20

app.add_middleware(

    CORSMiddleware,

    allow_origins=["*"],

    allow_credentials=True,

    allow_methods=["*"],

    allow_headers=["*"],

)

CATEGORY_MAP = {

    "Restaurant": "restaurant", "Cafe": "cafe", "Bakery": "bakery", "Bar": "bar",

    "Gym": "fitness_centre", "Salon": "beauty_salon", "Dentist": "dentist", "Clinic": "clinic",

    "Pharmacy": "pharmacy", "Hotel": "hotel", "Real Estate": "estate_agent", "Supermarket": "supermarket",

    "Jewellery": "jewellery", "Furniture": "furniture", "Car Repair": "car_repair",

    "Travel Agency": "travel_agency", "Photography": "photographer", "Event Management": "event_venue",

    "Interior Design": "interior_design", "Boutique": "clothes",

}

VALID_STATUSES = {"New", "Contacted", "Follow-up", "Negotiation", "Converted", "Lost"}

VALID_ROLES = {"Admin", "Agent"}


def is_admin(user: dict) -> bool:
    return str(user.get("role") or "Agent") == "Admin"


def require_admin(user: dict) -> None:
    if not is_admin(user):
        raise HTTPException(status_code=403, detail="Admin permission required.")


def lead_access_clause(user: dict, alias: str = ""):
    prefix = f"{alias}." if alias else ""
    if is_admin(user):
        return "TRUE", []
    return f"{prefix}assigned_to=%s", [user["id"]]


def lead_owner_check(cur, lead_id: int, user: dict, for_update: bool = False):
    scope, params = lead_access_clause(user)
    sql = f"SELECT id,assigned_to,user_id FROM leads WHERE id=%s AND {scope}"
    if for_update:
        sql += " FOR UPDATE"
    cur.execute(sql, [lead_id, *params])
    row = cur.fetchone()
    if not row:
        raise HTTPException(status_code=404, detail="Lead not found or you do not have access to it.")
    return row


def log_activity(cur, user_id: int, lead_id: Optional[int], action: str, details: Optional[dict] = None):
    cur.execute(
        "INSERT INTO activity_logs(user_id,lead_id,action,details) VALUES(%s,%s,%s,%s::jsonb)",
        (user_id, lead_id, action, json.dumps(details or {}, default=str)),
    )


def whatsapp_url(phone: Optional[str]) -> Optional[str]:
    if not phone:
        return None
    digits = re.sub(r"\D", "", str(phone))
    if digits.startswith("00"):
        digits = digits[2:]
    if len(digits) == 10:
        digits = "91" + digits
    elif digits.startswith("0") and len(digits) == 11:
        digits = "91" + digits[1:]
    if len(digits) < 10:
        return None
    return f"https://wa.me/{digits}"


def user_timezone_name() -> str:
    return os.getenv("APP_TIMEZONE", "Asia/Kolkata")


def local_today_utc_bounds() -> tuple[datetime, datetime, str]:
    try:
        tz = ZoneInfo(user_timezone_name())
    except Exception:
        tz = timezone.utc
    now_local = datetime.now(tz)
    start_local = datetime.combine(now_local.date(), datetime.min.time(), tzinfo=tz)
    end_local = start_local + timedelta(days=1)
    return start_local.astimezone(timezone.utc), end_local.astimezone(timezone.utc), now_local.date().isoformat()

# ============================================================

# Frontend

# ============================================================

@app.get("/")

def home():

    if not FRONTEND_FILE.exists():

        raise HTTPException(status_code=500, detail="frontend/index.html was not found.")

    return FileResponse(FRONTEND_FILE)

# ============================================================

# Database / safe migrations

# ============================================================

def db_conn():

    if not DATABASE_URL:

        raise HTTPException(status_code=500, detail="DATABASE_URL is not configured.")

    return psycopg.connect(DATABASE_URL)



def _add_columns(cur, table: str, migrations: dict):

    existing = {

        row[0] for row in cur.execute(

            """SELECT column_name FROM information_schema.columns

               WHERE table_schema='public' AND table_name=%s""", (table,)

        ).fetchall()

    }

    for column, sql in migrations.items():

        if column not in existing:

            try:

                cur.execute(sql)

            except Exception as exc:

                print(f"Migration warning {table}.{column}: {exc}")



def init_db():

    if not DATABASE_URL:

        print("WARNING: DATABASE_URL is not configured.")

        return

    with db_conn() as conn:

        with conn.cursor() as cur:

            cur.execute("""CREATE TABLE IF NOT EXISTS users (

                id SERIAL PRIMARY KEY, email TEXT UNIQUE NOT NULL, name TEXT,

                password_hash TEXT NOT NULL, salt TEXT NOT NULL,

                created_at TIMESTAMPTZ DEFAULT NOW(), role TEXT DEFAULT 'Agent')""")

            _add_columns(cur, "users", {

                "name": "ALTER TABLE users ADD COLUMN name TEXT",

                "password_hash": "ALTER TABLE users ADD COLUMN password_hash TEXT",

                "salt": "ALTER TABLE users ADD COLUMN salt TEXT",

                "created_at": "ALTER TABLE users ADD COLUMN created_at TIMESTAMPTZ DEFAULT NOW()",
                "role": "ALTER TABLE users ADD COLUMN role TEXT DEFAULT 'Agent'",

            })

            cur.execute("""CREATE TABLE IF NOT EXISTS leads (

                id SERIAL PRIMARY KEY,

                user_id INTEGER REFERENCES users(id) ON DELETE CASCADE,

                place_id TEXT,

                provider_place_id TEXT,

                name TEXT NOT NULL,

                address TEXT, city TEXT, region TEXT, country TEXT,

                latitude DOUBLE PRECISION, longitude DOUBLE PRECISION,

                phone TEXT, website TEXT, email TEXT,

                instagram TEXT, facebook TEXT, twitter TEXT,

                category TEXT, rating DOUBLE PRECISION, reviews INTEGER,

                maps_url TEXT, source TEXT DEFAULT 'LatLng Places',

                sources JSONB DEFAULT '[]'::jsonb,

                lead_score INTEGER, lead_type TEXT,

                ai_recommendation TEXT, ai_reason TEXT,

                status TEXT DEFAULT 'New', notes TEXT,
                assigned_to INTEGER REFERENCES users(id) ON DELETE SET NULL,
                last_contacted_date TIMESTAMPTZ, next_followup_date TIMESTAMPTZ,
                is_duplicate BOOLEAN DEFAULT FALSE, duplicate_of INTEGER REFERENCES leads(id) ON DELETE SET NULL,

                enrichment_source TEXT, enrichment_at TIMESTAMPTZ,

                created_at TIMESTAMPTZ DEFAULT NOW(), updated_at TIMESTAMPTZ DEFAULT NOW())""")

            _add_columns(cur, "leads", {

                "user_id": "ALTER TABLE leads ADD COLUMN user_id INTEGER REFERENCES users(id) ON DELETE CASCADE",

                "place_id": "ALTER TABLE leads ADD COLUMN place_id TEXT",

                "provider_place_id": "ALTER TABLE leads ADD COLUMN provider_place_id TEXT",

                "city": "ALTER TABLE leads ADD COLUMN city TEXT",

                "region": "ALTER TABLE leads ADD COLUMN region TEXT",

                "country": "ALTER TABLE leads ADD COLUMN country TEXT",

                "latitude": "ALTER TABLE leads ADD COLUMN latitude DOUBLE PRECISION",

                "longitude": "ALTER TABLE leads ADD COLUMN longitude DOUBLE PRECISION",

                "email": "ALTER TABLE leads ADD COLUMN email TEXT",

                "instagram": "ALTER TABLE leads ADD COLUMN instagram TEXT",

                "facebook": "ALTER TABLE leads ADD COLUMN facebook TEXT",

                "twitter": "ALTER TABLE leads ADD COLUMN twitter TEXT",

                "maps_url": "ALTER TABLE leads ADD COLUMN maps_url TEXT",

                "source": "ALTER TABLE leads ADD COLUMN source TEXT DEFAULT 'LatLng Places'",

                "sources": "ALTER TABLE leads ADD COLUMN sources JSONB DEFAULT '[]'::jsonb",

                "lead_score": "ALTER TABLE leads ADD COLUMN lead_score INTEGER",

                "lead_type": "ALTER TABLE leads ADD COLUMN lead_type TEXT",

                "ai_recommendation": "ALTER TABLE leads ADD COLUMN ai_recommendation TEXT",

                "ai_reason": "ALTER TABLE leads ADD COLUMN ai_reason TEXT",

                "status": "ALTER TABLE leads ADD COLUMN status TEXT DEFAULT 'New'",
                "assigned_to": "ALTER TABLE leads ADD COLUMN assigned_to INTEGER REFERENCES users(id) ON DELETE SET NULL",
                "last_contacted_date": "ALTER TABLE leads ADD COLUMN last_contacted_date TIMESTAMPTZ",
                "next_followup_date": "ALTER TABLE leads ADD COLUMN next_followup_date TIMESTAMPTZ",
                "is_duplicate": "ALTER TABLE leads ADD COLUMN is_duplicate BOOLEAN DEFAULT FALSE",
                "duplicate_of": "ALTER TABLE leads ADD COLUMN duplicate_of INTEGER REFERENCES leads(id) ON DELETE SET NULL",

                "notes": "ALTER TABLE leads ADD COLUMN notes TEXT",

                "enrichment_source": "ALTER TABLE leads ADD COLUMN enrichment_source TEXT",

                "enrichment_at": "ALTER TABLE leads ADD COLUMN enrichment_at TIMESTAMPTZ",
                "updated_at": "ALTER TABLE leads ADD COLUMN updated_at TIMESTAMPTZ DEFAULT NOW()",

            })

            # Preserve old data and old group memberships. Never drop group_leads on startup.

            cur.execute("""CREATE TABLE IF NOT EXISTS search_history (

                id SERIAL PRIMARY KEY, user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,

                city TEXT, category TEXT, keyword TEXT, city_wide BOOLEAN DEFAULT FALSE,

                result_count INTEGER, sources_used JSONB DEFAULT '[]'::jsonb, created_at TIMESTAMPTZ DEFAULT NOW(), role TEXT DEFAULT 'Agent')""")

            _add_columns(cur, "search_history", {

                "keyword": "ALTER TABLE search_history ADD COLUMN keyword TEXT",

                "city_wide": "ALTER TABLE search_history ADD COLUMN city_wide BOOLEAN DEFAULT FALSE",
                "sources_used": "ALTER TABLE search_history ADD COLUMN sources_used JSONB DEFAULT '[]'::jsonb",

            })

            cur.execute("""CREATE TABLE IF NOT EXISTS lead_groups (

                id SERIAL PRIMARY KEY, user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,

                name TEXT NOT NULL, created_at TIMESTAMPTZ DEFAULT NOW())""")

            cur.execute("""CREATE TABLE IF NOT EXISTS group_leads (

                id SERIAL PRIMARY KEY,

                group_id INTEGER NOT NULL REFERENCES lead_groups(id) ON DELETE CASCADE,

                lead_id INTEGER NOT NULL REFERENCES leads(id) ON DELETE CASCADE,

                created_at TIMESTAMPTZ DEFAULT NOW(), UNIQUE(group_id, lead_id))""")

            cur.execute("CREATE INDEX IF NOT EXISTS leads_user_created_idx ON leads(user_id, created_at DESC)")

            cur.execute("CREATE INDEX IF NOT EXISTS leads_user_city_idx ON leads(user_id, city)")

            cur.execute("CREATE INDEX IF NOT EXISTS leads_user_place_idx ON leads(user_id, place_id)")

            cur.execute("CREATE INDEX IF NOT EXISTS leads_assigned_status_idx ON leads(assigned_to,status,created_at DESC)")
            cur.execute("UPDATE leads SET status='New' WHERE status='Not Contacted'")
            cur.execute("UPDATE leads SET status='Negotiation' WHERE status='Interested'")
            cur.execute("""CREATE TABLE IF NOT EXISTS call_logs (
                id SERIAL PRIMARY KEY, lead_id INTEGER NOT NULL REFERENCES leads(id) ON DELETE CASCADE,
                called_by INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
                call_date TIMESTAMPTZ DEFAULT NOW(), call_outcome TEXT NOT NULL, notes TEXT,
                created_at TIMESTAMPTZ DEFAULT NOW())""")
            cur.execute("""CREATE TABLE IF NOT EXISTS activity_logs (
                id SERIAL PRIMARY KEY, user_id INTEGER REFERENCES users(id) ON DELETE SET NULL,
                lead_id INTEGER REFERENCES leads(id) ON DELETE CASCADE, action TEXT NOT NULL,
                details JSONB DEFAULT '{}'::jsonb, created_at TIMESTAMPTZ DEFAULT NOW())""")
            cur.execute("CREATE INDEX IF NOT EXISTS leads_followup_idx ON leads(next_followup_date)")
            cur.execute("CREATE INDEX IF NOT EXISTS activity_logs_lead_created_idx ON activity_logs(lead_id,created_at DESC)")
            cur.execute("CREATE INDEX IF NOT EXISTS call_logs_lead_date_idx ON call_logs(lead_id,call_date DESC)")
            # Ensure legacy leads remain visible to their original creator and guarantee one Admin.
            cur.execute("UPDATE leads SET assigned_to=user_id WHERE assigned_to IS NULL AND user_id IS NOT NULL")
            cur.execute("SELECT COUNT(*) FROM users WHERE role='Admin'")
            if (cur.fetchone() or [0])[0] == 0:
                cur.execute("SELECT id FROM users ORDER BY created_at ASC NULLS LAST,id ASC LIMIT 1")
                first_user = cur.fetchone()
                if first_user:
                    cur.execute("UPDATE users SET role='Admin' WHERE id=%s", (first_user[0],))

        conn.commit()



@app.on_event("startup")

def startup():

    try:

        init_db()

        print("Database initialization complete.")

    except Exception as exc:

        print(f"Database initialization error: {exc}")

# ============================================================

# Passwords / JWT

# ============================================================

PBKDF2_ITERATIONS = 260_000

def hash_password(password: str, salt: Optional[bytes] = None):

    salt = salt or secrets.token_bytes(16)

    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, PBKDF2_ITERATIONS)

    return base64.b64encode(digest).decode("ascii"), base64.b64encode(salt).decode("ascii")



def verify_password(password: str, password_hash: str, salt_b64: str) -> bool:

    try:

        salt = base64.b64decode(salt_b64)

        expected = base64.b64decode(password_hash)

        actual = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, PBKDF2_ITERATIONS)

        return hmac.compare_digest(actual, expected)

    except Exception:

        return False



def create_token(user_id: int, email: str):

    now = datetime.now(timezone.utc)

    return jwt.encode({"sub": str(user_id), "email": email, "iat": now, "exp": now + timedelta(hours=JWT_EXPIRE_HOURS)}, JWT_SECRET, algorithm=JWT_ALGORITHM)



def decode_token(token: str):

    try:

        return jwt.decode(token, JWT_SECRET, algorithms=[JWT_ALGORITHM])

    except jwt.ExpiredSignatureError:

        raise HTTPException(status_code=401, detail="Session expired. Please log in again.")

    except jwt.InvalidTokenError:

        raise HTTPException(status_code=401, detail="Invalid authentication token.")



def get_current_user(authorization: Optional[str] = Header(None)):

    if not authorization or not authorization.startswith("Bearer "):

        raise HTTPException(status_code=401, detail="Authentication required.")

    payload = decode_token(authorization.split(" ", 1)[1])

    try:

        user_id = int(payload["sub"])

    except (KeyError, ValueError, TypeError):

        raise HTTPException(status_code=401, detail="Invalid authentication token.")

    with db_conn() as conn:

        with conn.cursor() as cur:

            cur.execute("SELECT id, email, name, COALESCE(role, 'Agent') FROM users WHERE id=%s", (user_id,))

            user = cur.fetchone()

    if not user:

        raise HTTPException(status_code=401, detail="User not found.")

    return {"id": user[0], "email": user[1], "name": user[2], "role": user[3]}

# ============================================================

# Models

# ============================================================

class SignupRequest(BaseModel):

    email: str

    password: str

    name: Optional[str] = None

class LoginRequest(BaseModel):

    email: str

    password: str

class SearchRequest(BaseModel):

    city: str = Field(min_length=2, max_length=120)

    category: Optional[str] = None

    keyword: Optional[str] = None

    limit: int = Field(default=10, ge=1, le=100)

    city_wide: bool = False

    radius_meters: int = Field(default=SEARCH_RADIUS_METERS, ge=250, le=5000)

    enrich: bool = False

class StatusUpdateRequest(BaseModel):

    status: str

class NotesUpdateRequest(BaseModel):

    notes: Optional[str] = None

class LeadUpdateRequest(BaseModel):
    name: Optional[str] = Field(None, min_length=1, max_length=250)
    address: Optional[str] = Field(None, max_length=1000)
    city: Optional[str] = Field(None, max_length=120)
    region: Optional[str] = Field(None, max_length=120)
    country: Optional[str] = Field(None, max_length=120)
    phone: Optional[str] = Field(None, max_length=100)
    website: Optional[str] = Field(None, max_length=500)
    email: Optional[str] = Field(None, max_length=320)
    instagram: Optional[str] = Field(None, max_length=500)
    facebook: Optional[str] = Field(None, max_length=500)
    twitter: Optional[str] = Field(None, max_length=500)
    category: Optional[str] = Field(None, max_length=150)
    notes: Optional[str] = Field(None, max_length=5000)
    status: Optional[str] = None
    assigned_to: Optional[int] = None
    last_contacted_date: Optional[datetime] = None
    next_followup_date: Optional[datetime] = None

class AssignmentRequest(BaseModel):
    assigned_to: Optional[int] = None

class FollowupRequest(BaseModel):
    next_followup_date: Optional[datetime] = None

class CallLogRequest(BaseModel):
    call_date: Optional[datetime] = None
    call_outcome: str = Field(min_length=1, max_length=100)
    notes: Optional[str] = Field(None, max_length=5000)

class BulkLeadActionRequest(BaseModel):
    lead_ids: List[int] = Field(min_length=1, max_length=500)
    action: str
    status: Optional[str] = None
    assigned_to: Optional[int] = None

class GroupCreateRequest(BaseModel):

    name: str

class GroupBulkAddRequest(BaseModel):

    lead_ids: List[int]

# ============================================================

# Auth endpoints

# ============================================================

@app.post("/api/signup")

def signup(req: SignupRequest):

    email = req.email.strip().lower()

    name = (req.name or "").strip()

    if "@" not in email or "." not in email.split("@")[-1]:

        raise HTTPException(status_code=400, detail="Please enter a valid email address.")

    if len(req.password) < 6:

        raise HTTPException(status_code=400, detail="Password must be at least 6 characters.")

    password_hash, salt = hash_password(req.password)

    try:

        with db_conn() as conn:

            with conn.cursor() as cur:

                cur.execute("SELECT id FROM users WHERE email=%s", (email,))

                if cur.fetchone():

                    raise HTTPException(status_code=400, detail="An account with this email already exists.")

                cur.execute("SELECT COUNT(*) FROM users")
                role = "Admin" if cur.fetchone()[0] == 0 else "Agent"
                cur.execute("INSERT INTO users(email,name,password_hash,salt,role) VALUES(%s,%s,%s,%s,%s) RETURNING id", (email, name or None, password_hash, salt, role))

                user_id = cur.fetchone()[0]

            conn.commit()

    except HTTPException:

        raise

    except Exception as exc:

        print(f"Signup database error: {exc}")

        raise HTTPException(status_code=500, detail=f"Signup database error: {type(exc).__name__}.")

    return {"token": create_token(user_id, email), "user": {"id": user_id, "email": email, "name": name, "role": role}}

@app.post("/api/login")

def login(req: LoginRequest):

    email = req.email.strip().lower()

    try:

        with db_conn() as conn:

            with conn.cursor() as cur:

                cur.execute("SELECT id,email,name,password_hash,salt,COALESCE(role, 'Agent') FROM users WHERE email=%s", (email,))

                user = cur.fetchone()

    except Exception as exc:

        raise HTTPException(status_code=500, detail=f"Login database error: {type(exc).__name__}.")

    if not user or not verify_password(req.password, user[3], user[4]):

        raise HTTPException(status_code=401, detail="Invalid email or password.")

    return {"token": create_token(user[0], user[1]), "user": {"id": user[0], "email": user[1], "name": user[2], "role": user[5]}}

@app.get("/api/me")

def me(current_user: dict = Depends(get_current_user)):

    return {"user": current_user}

@app.get("/api/categories")

def categories():

    return {"categories": list(CATEGORY_MAP.keys())}

# ============================================================

# SerpApi Google Maps provider (primary local-business source)

# ============================================================

def serpapi_get(params: dict, timeout: int = 30):

    if not SERPAPI_API_KEY:

        raise ProviderError("SERPAPI_API_KEY is not configured.")

    query = dict(params)

    query.update({"api_key": SERPAPI_API_KEY, "engine": "google_maps"})

    try:

        response = requests.get("https://serpapi.com/search.json", params=query, timeout=timeout)

    except requests.RequestException:

        raise ProviderError("SerpApi is temporarily unavailable. Please try again later.")

    if response.status_code == 401:

        raise ProviderError("SerpApi API key is invalid or unauthorized.")

    if response.status_code == 429:

        raise ProviderError("SerpApi rate limit reached. Please try again later.")

    if response.status_code != 200:

        try:

            detail = response.json().get("error")

        except Exception:

            detail = None

        raise ProviderError(f"SerpApi returned HTTP {response.status_code}" + (f": {detail}" if detail else "."))

    try:

        data = response.json()

    except ValueError:

        raise ProviderError("SerpApi returned an invalid response.")

    if data.get("error"):

        raise ProviderError(str(data["error"]))

    return data



def search_serpapi_maps(lat: Optional[float], lon: Optional[float], query: str, limit: int):

    """Search Google Maps through SerpApi. Coordinates are optional; when

    available they improve geographic targeting and enable pagination."""

    target = min(max(int(limit), 1), MAX_LIMIT)

    results = []

    # SerpApi documents ll + start for paginated Maps searches. Without a

    # reliable geocoded center, use the first Maps page and rely on the city

    # included in q instead of requiring the old LatLng provider.

    max_pages = min(5, (target + 19) // 20) if lat is not None and lon is not None else 1

    for page in range(max_pages):

        params = {

            "q": query,

            "type": "search",

            "hl": "en",

            "gl": "in",

        }

        if lat is not None and lon is not None:

            params["ll"] = f"@{float(lat)},{float(lon)},14z"

            params["start"] = page * 20

        data = serpapi_get(params)

        page_results = data.get("local_results") or []

        if not isinstance(page_results, list):

            raise ProviderError("SerpApi returned an unexpected Google Maps response.")

        results.extend(page_results)

        if len(results) >= target or len(page_results) < 20:

            break

    return results[:target]



def extract_serpapi_lead(

    place: dict,

    requested_category: Optional[str],

    resolved_city: str,

    center_lat: float,

    center_lon: float,

    radius_meters: int = SEARCH_RADIUS_METERS,

):

    name = place.get("title")

    gps = place.get("gps_coordinates") or {}

    lat, lon = gps.get("latitude"), gps.get("longitude")

    if not name or lat is None or lon is None:

        return None

    try:

        lat, lon = float(lat), float(lon)

    except (ValueError, TypeError):

        return None

    # Google Maps results are not guaranteed to stay inside ll, so enforce

    # our own geographic boundary.

    if center_lat is not None and center_lon is not None:

        if distance_km(center_lat, center_lon, lat, lon) > radius_meters / 1000:

            return None

    provider_id = place.get("place_id") or place.get("data_cid") or place.get("data_id")

    place_id = str(provider_id or f"{name}|{lat:.6f}|{lon:.6f}")

    links = place.get("links") or {}

    website = place.get("website") or links.get("website")

    maps_url = (

        links.get("directions")

        or "https://www.google.com/maps/search/?api=1&query="

        + quote_plus(f"{name}, {resolved_city}")

    )

    return {

        "place_id": place_id,

        "provider_place_id": str(provider_id) if provider_id else None,

        "name": str(name),

        "address": place.get("address"),

        "city": resolved_city,

        "region": None,

        "country": place.get("country") or "India",

        "latitude": lat,

        "longitude": lon,

        "phone": place.get("phone"),

        "website": website,

        "email": place.get("email"),

        "instagram": place.get("instagram"),

        "facebook": place.get("facebook"),

        "twitter": place.get("twitter"),

        "category": requested_category or place.get("type"),

        "rating": place.get("rating"),

        "reviews": place.get("reviews"),

        "maps_url": maps_url,

        "source": "Google Maps via SerpApi",

        "sources": ["Google Maps via SerpApi"],

        "provider_category": place.get("type"),

        "distance_m": place.get("distance"),

        "match_score": None,

    }



# ============================================================

# LatLng provider

# ============================================================

class ProviderError(Exception):

    pass



def latlng_headers():

    if not LATLNG_API_KEY:

        raise ProviderError("LATLNG_API_KEY is not configured.")

    return {"X-Api-Key": LATLNG_API_KEY, "Accept": "application/json"}



def _latlng_get(url: str, params: dict, timeout: int = 20):

    try:

        response = requests.get(url, params=params, headers=latlng_headers(), timeout=timeout)

    except requests.RequestException:

        raise ProviderError("LatLng API is temporarily unavailable. Please try again later.")

    if response.status_code == 429:

        raise ProviderError("LatLng rate limit reached. Please try again later.")

    if response.status_code == 401:

        raise ProviderError("LatLng API key is invalid or not authorized.")

    if response.status_code != 200:

        raise ProviderError(f"LatLng API returned HTTP {response.status_code}.")

    try:

        return response.json()

    except ValueError:

        raise ProviderError("LatLng API returned an invalid response.")



def resolve_city(city: str):

    """Resolve the requested city without making LatLng mandatory.

    SerpApi is the primary provider. If its city lookup does not expose a

    reliable coordinate, the search can still proceed using the city name in

    the Google Maps query. LatLng is only an optional coordinate fallback.

    """

    raw = city.strip()

    city_name = raw.split(",", 1)[0].strip()

    if not city_name:

        raise ProviderError("City is required.")

    if SERPAPI_API_KEY:

        try:

            data = serpapi_get({

                "q": raw,

                "type": "search",

                "hl": "en",

                "gl": "in",

            })

            results = data.get("local_results") or []

            if isinstance(results, list):

                target = city_name.casefold()

                best = None

                for item in results:

                    title = str(item.get("title") or "").strip()

                    address = str(item.get("address") or "").casefold()

                    if title.casefold() == target or target in address:

                        best = item

                        break

                # Do not treat an arbitrary business as the city center.

                # If Maps does not return a city-level result, search by the

                # city name without a coordinate filter.

                gps = (best or {}).get("gps_coordinates") or {}

                lat = gps.get("latitude")

                lon = gps.get("longitude")

                try:

                    lat = float(lat) if lat is not None else None

                    lon = float(lon) if lon is not None else None

                except (ValueError, TypeError):

                    lat, lon = None, None

                return {

                    "lat": lat,

                    "lon": lon,

                    "resolved_city": city_name,

                    "state": "",

                    "country": "India",

                }

        except ProviderError:

            # If LatLng is configured, it remains a safe coordinate fallback.

            if not LATLNG_API_KEY:

                raise

    if LATLNG_API_KEY:

        data = _latlng_get(LATLNG_GEOCODE_URL, {"q": raw, "limit": 10, "lang": "en"}, 15)

        features = data.get("features", [])

        target = city_name.casefold()

        for feature in features:

            props = feature.get("properties") or {}

            coords = (feature.get("geometry") or {}).get("coordinates") or []

            if len(coords) < 2:

                continue

            candidates = {str(props.get(k) or "").strip().casefold() for k in ("city", "town", "municipality", "village", "name")}

            candidates.discard("")

            country = str(props.get("countrycode") or props.get("country_code") or "").strip().casefold()

            if target not in candidates and city_name.casefold() not in str(props.get("display_name") or "").casefold():

                continue

            if country and country != "in":

                continue

            try:

                lon, lat = float(coords[0]), float(coords[1])

            except (ValueError, TypeError):

                continue

            return {

                "lat": lat,

                "lon": lon,

                "resolved_city": str(props.get("city") or props.get("town") or props.get("municipality") or props.get("village") or props.get("name") or city_name),

                "state": props.get("state") or "",

                "country": props.get("country") or "India",

            }

    raise ProviderError(f"Could not locate '{city_name}'. Check your SerpApi API key and try the city name only.")



def distance_km(lat1, lon1, lat2, lon2):

    from math import radians, sin, cos, sqrt, atan2

    r = 6371.0088

    dlat, dlon = radians(lat2 - lat1), radians(lon2 - lon1)

    a = sin(dlat / 2) ** 2 + cos(radians(lat1)) * cos(radians(lat2)) * sin(dlon / 2) ** 2

    return 2 * r * atan2(sqrt(a), sqrt(1 - a))



def search_nearby(lat, lon, category, limit, radius=SEARCH_RADIUS_METERS):

    """Search businesses around a coordinate using LatLng Places Nearby."""

    latlng_category = CATEGORY_MAP.get(category)

    if not latlng_category:

        raise ProviderError(f"Category '{category}' is not supported.")

    params = {

        "lat": float(lat),

        "lon": float(lon),

        "radius": min(max(int(radius), 250), 5000),

        "category": latlng_category,

        "country": "IN",

        "limit": min(max(int(limit), 1), 50),

    }

    data = _latlng_get(LATLNG_PLACES_NEARBY_URL, params, 20)

    places = data.get("places", []) if isinstance(data, dict) else []

    if not isinstance(places, list):

        raise ProviderError("LatLng returned an unexpected Places response.")

    return places

def search_keyword(lat, lon, keyword, limit):

    q = keyword.strip()

    if len(q) < 2:

        raise ProviderError("Keyword must contain at least 2 characters.")

    data = _latlng_get(LATLNG_PLACES_SEARCH_URL, {"q": q, "lat": lat, "lon": lon, "limit": min(max(limit, 1), 50)}, 20)

    return data.get("places", [])



def city_grid(location: dict, enabled: bool, spacing_meters: Optional[int] = None):
    if not enabled or location.get("lat") is None or location.get("lon") is None:
        return [location]
    from math import cos, radians
    spacing = spacing_meters or CITY_WIDE_GRID_SPACING_METERS
    offsets = [(0,0),(0,1),(1,0),(0,-1),(-1,0),(1,1),(1,-1),(-1,1),(-1,-1)][:CITY_WIDE_MAX_ZONES]
    lat_delta = spacing / 111_320
    lon_delta = spacing / max(1, 111_320 * abs(cos(radians(location["lat"]))))
    return [{**location, "lat": location["lat"] + y * lat_delta, "lon": location["lon"] + x * lon_delta} for x, y in offsets]



def extract_lead(place: dict, requested_category: Optional[str], resolved_city: str, center_lat: float, center_lon: float, radius_meters: int = SEARCH_RADIUS_METERS):

    name, lat, lon = place.get("name"), place.get("lat"), place.get("lon")

    if not name or lat is None or lon is None:

        return None

    try:

        lat, lon = float(lat), float(lon)

    except (ValueError, TypeError):

        return None

    # Keep results geographically relevant to the requested city/grid zone.

    if distance_km(center_lat, center_lon, lat, lon) > radius_meters / 1000:

        return None

    provider_id = place.get("place_id") or place.get("id")

    place_id = str(provider_id or f"{name}|{lat:.6f}|{lon:.6f}")

    maps_url = "https://www.google.com/maps/search/?api=1&query=" + quote_plus(f"{name}, {resolved_city}")

    return {

        "place_id": place_id,

        "provider_place_id": str(provider_id) if provider_id else None,

        "name": str(name),

        "address": place.get("address") or place.get("formatted_address"),

        "city": resolved_city,

        "region": None,

        "country": "India",

        "latitude": lat,

        "longitude": lon,

        "phone": place.get("phone"),

        "website": place.get("website"),

        "email": place.get("email"),

        "instagram": place.get("instagram"),

        "facebook": place.get("facebook"),

        "twitter": place.get("twitter"),

        "category": requested_category or place.get("category"),

        "rating": place.get("rating"),

        "reviews": place.get("reviews"),

        "maps_url": maps_url,

        "source": "LatLng Places",

        "sources": ["LatLng Places"],

        "provider_category": place.get("category"),

        "distance_m": place.get("distance_m"),

        "match_score": place.get("match_score"),

    }

# ============================================================

# Dedupe / merge

# ============================================================

def norm(value):

    return re.sub(r"[^a-z0-9]", "", str(value or "").lower())



def phone_key(value):

    return re.sub(r"\D", "", str(value or ""))[-10:] if value else ""



def domain_key(value):

    if not value:

        return ""

    try:

        host = urlparse(value if "//" in value else "https://" + value).netloc.lower().split(":")[0]

        return host[4:] if host.startswith("www.") else host

    except Exception:

        return ""



def dedupe_leads(leads: list[dict], limit: int):

    groups = []

    indexes = {}

    for lead in leads:

        keys = [

            ("id", norm(lead.get("provider_place_id") or lead.get("place_id"))),

            ("phone", phone_key(lead.get("phone"))),

            ("domain", domain_key(lead.get("website"))),

            ("nameaddr", norm((lead.get("name") or "") + " " + (lead.get("address") or ""))),

        ]

        found = None

        for kind, key in keys:

            if key and (kind, key) in indexes:

                found = indexes[(kind, key)]

                break

        if found is None:

            found = len(groups)

            groups.append(dict(lead))

        else:

            target = groups[found]

            for key, value in lead.items():

                if key in {"sources", "source"}:

                    continue

                if not target.get(key) and value not in (None, ""):

                    target[key] = value

            target["sources"] = list(dict.fromkeys([*(target.get("sources") or []), *(lead.get("sources") or []), lead.get("source") or ""]))

            target["sources"] = [x for x in target["sources"] if x]

            target["source"] = ", ".join(target["sources"])

        for kind, key in keys:

            if key:

                indexes[(kind, key)] = found

        if len(groups) >= limit and len(groups) >= limit * 2:

            break

    return groups[:limit]

# ============================================================

# Public website enrichment (conservative)

# ============================================================

EMAIL_RE = re.compile(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", re.I)

SOCIAL_RE = {

    "instagram": re.compile(r"https?://(?:www\.)?instagram\.com/[^\"'\s<>]+", re.I),

    "facebook": re.compile(r"https?://(?:www\.)?facebook\.com/[^\"'\s<>]+", re.I),

    "twitter": re.compile(r"https?://(?:www\.)?(?:twitter\.com|x\.com)/[^\"'\s<>]+", re.I),

}



def enrich_public_website(url: Optional[str]):

    if not url or not str(url).startswith(("http://", "https://")):

        return {}

    try:

        rp = robotparser.RobotFileParser()

        base = f"{urlparse(url).scheme}://{urlparse(url).netloc}"

        rp.set_url(base.rstrip("/") + "/robots.txt")

        try:

            rp.read()

            if not rp.can_fetch("LeadSearchingEngine/1.0", url):

                return {}

        except Exception:

            # A failed robots fetch is not treated as permission to crawl; skip enrichment.

            return {}

        response = requests.get(url, headers={"User-Agent": "LeadSearchingEngine/1.0"}, timeout=ENRICH_TIMEOUT, allow_redirects=True)

        if response.status_code != 200 or "text/html" not in response.headers.get("content-type", "").lower():

            return {}

        text = response.text[:1_000_000]

        result = {"sources": ["Business Website Enrichment"], "enrichment_source": response.url, "enrichment_at": datetime.now(timezone.utc).isoformat()}

        email = EMAIL_RE.search(text)

        if email:

            result["email"] = email.group(0)

        for key, pattern in SOCIAL_RE.items():

            match = pattern.search(text)

            if match:

                result[key] = match.group(0).rstrip(".,);'\"")

        return result

    except Exception:

        return {}

# ============================================================

# AI analysis

# ============================================================

def fallback_analysis(lead: dict):

    phone, website = bool(lead.get("phone")), bool(lead.get("website"))

    social = bool(lead.get("instagram") or lead.get("facebook") or lead.get("twitter"))

    score = 45 + (10 if phone else 0) + (10 if website else 0) + (5 if social else 0)

    if lead.get("rating") is not None: score += 5

    if lead.get("reviews") is not None and lead.get("reviews") >= 20: score += 5

    score = min(score, 80)

    if not phone and not website and not social:

        lead_type, recommendation = "Needs More Information", "Verify contact details manually before outreach."

    elif phone and not website:

        lead_type, recommendation = "Medium Potential", "Consider contacting the business about its website or online presence."

    else:

        lead_type, recommendation = "Needs More Information", "Review the available business signals before prioritizing outreach."

    return {"lead_score": score, "lead_type": lead_type, "ai_recommendation": recommendation,

            "ai_reason": "Conservative assessment based only on fields returned by the authorized data source. Missing fields do not prove the business lacks them."}



def score_to_lead_priority(score: int) -> str:
    score = max(0, min(100, int(score)))
    if score >= 70:
        return "Hot Lead"
    if score >= 45:
        return "Warm Lead"
    return "Cold Lead"

def gemini_analysis(lead: dict):
    fallback = fallback_analysis(lead)
    if not GEMINI_API_KEY:
        fallback["lead_type"] = score_to_lead_priority(fallback["lead_score"])
        return fallback
    prompt = f'''Analyze this business lead for a digital marketing/web-services agency.\nBusiness: {lead.get("name")}\nCategory: {lead.get("category")}\nPhone: {"yes" if lead.get("phone") else "no"}\nWebsite: {"yes" if lead.get("website") else "no"}\nSocial: {"yes" if (lead.get("instagram") or lead.get("facebook") or lead.get("twitter")) else "no"}\nRating: {lead.get("rating", "not available")}\nReviews: {lead.get("reviews", "not available")}\nIMPORTANT: Missing fields only mean the source did not return that field. Do not invent facts. Return ONLY JSON with lead_score and reason.'''
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}"
    try:
        response = requests.post(url, json={"contents":[{"parts":[{"text":prompt}]}]}, timeout=12)
        if response.status_code != 200:
            fallback["lead_type"] = score_to_lead_priority(fallback["lead_score"]); return fallback
        text = response.json()["candidates"][0]["content"]["parts"][0]["text"].strip()
        text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text, flags=re.I)
        parsed=json.loads(text)
        score=max(0,min(100,int(parsed.get("lead_score",fallback["lead_score"]))))
        return {"lead_score":score,"lead_type":score_to_lead_priority(score),"ai_recommendation":fallback["ai_recommendation"],"ai_reason":str(parsed.get("reason") or fallback["ai_reason"])[:1000]}
    except Exception:
        fallback["lead_type"] = score_to_lead_priority(fallback["lead_score"]); return fallback
# Search

# ============================================================

def _check_search_rate(user_id: int):
    now = datetime.now(timezone.utc).timestamp()
    events = _SEARCH_EVENTS[user_id]
    while events and now - events[0] > SEARCH_WINDOW_SECONDS:
        events.popleft()
    if len(events) >= SEARCH_MAX_REQUESTS:
        raise HTTPException(status_code=429, detail="Search limit reached temporarily. Please wait a few minutes and try again.")
    events.append(now)

@app.post("/api/search")
def search(req: SearchRequest, current_user: dict = Depends(get_current_user)):
    _check_search_rate(current_user["id"])
    city = req.city.strip()
    category = req.category.strip() if req.category else None
    keyword = req.keyword.strip() if req.keyword else None
    if not city:
        raise HTTPException(status_code=400, detail="City is required.")
    if bool(category) == bool(keyword):
        raise HTTPException(status_code=400, detail="Provide either a category or a custom keyword, not both.")
    if category and category not in CATEGORY_MAP:
        raise HTTPException(status_code=400, detail=f"Category '{category}' is not supported.")
    if keyword and (len(keyword) < 2 or len(keyword) > 100 or any(ord(c) < 32 for c in keyword)):
        raise HTTPException(status_code=400, detail="Keyword must be 2-100 characters and contain no control characters.")

    try:
        location = resolve_city(city)
    except ProviderError as exc:
        raise HTTPException(status_code=502, detail=str(exc))

    limit = min(max(int(req.limit), 1), MAX_LIMIT)
    radius_meters = min(max(int(req.radius_meters), 250), 5000)
    warnings = []
    used_sources = []
    leads = []
    search_query = f"{category} in {location['resolved_city']}" if category else f"{keyword} in {location['resolved_city']}"
    zones = city_grid(location, req.city_wide)
    zone_limit = min(20, max(1, limit))

    # Provider registry: each provider can fail independently. Results are merged
    # and deduplicated, so adding a future authorized provider only needs a new
    # provider adapter/runner rather than a rewrite of the search endpoint.
    provider_runners = []
    if SERPAPI_API_KEY:
        provider_runners.append(("Google Maps via SerpApi", "serpapi"))
    if LATLNG_API_KEY:
        provider_runners.append(("LatLng Places", "latlng"))
    if not provider_runners:
        raise HTTPException(status_code=503, detail="No lead search provider is configured.")

    for zone in zones:
        remaining = max(limit * 2 - len(leads), 1)
        per_provider = min(zone_limit, remaining)
        for source_name, provider_name in provider_runners:
            try:
                if provider_name == "serpapi":
                    places = search_serpapi_maps(zone.get("lat"), zone.get("lon"), search_query, per_provider)
                    for place in places:
                        lead = extract_serpapi_lead(
                            place, category, location["resolved_city"], zone.get("lat"), zone.get("lon"), radius_meters
                        )
                        if lead:
                            leads.append(lead)
                else:
                    if zone.get("lat") is None or zone.get("lon") is None:
                        continue
                    places = search_nearby(zone["lat"], zone["lon"], category, per_provider, radius=radius_meters) if category else search_keyword(zone["lat"], zone["lon"], keyword, per_provider)
                    for place in places:
                        lead = extract_lead(
                            place, category, location["resolved_city"], zone["lat"], zone["lon"], radius_meters
                        )
                        if lead:
                            leads.append(lead)
                if places:
                    used_sources.append(source_name)
            except ProviderError as exc:
                warnings.append(f"{source_name}: {exc}")

        # Once enough raw results exist, stop adding expensive provider calls.
        if len(leads) >= limit * 2:
            break

    leads = dedupe_leads(leads, limit)

    if not leads and warnings:
        # If one provider failed but another returned no matches, expose the
        # provider issue without leaking API keys or backend internals.
        raise HTTPException(status_code=502, detail=warnings[0])
    if not leads:
        warnings.append(f"No matching businesses were returned for {category or keyword!r} in {location['resolved_city']}.")

    if req.enrich:
        for lead in leads:
            extra = enrich_public_website(lead.get("website"))
            for key, value in extra.items():
                if key != "sources" and not lead.get(key):
                    lead[key] = value
            lead["sources"] = list(dict.fromkeys([*(lead.get("sources") or []), *(extra.get("sources") or [])]))
            lead["source"] = ", ".join(lead["sources"])
            if extra.get("sources"):
                used_sources.extend(extra["sources"])

    used_sources = list(dict.fromkeys([x for x in used_sources if x]))
    saved = []
    duplicates = []
    for lead in leads:
        lead.setdefault("enrichment_source", None)
        lead.setdefault("enrichment_at", None)

    if DATABASE_URL:
        with db_conn() as conn:
            with conn.cursor() as cur:
                for lead in leads:
                    lead.update(gemini_analysis(lead))
                    # Smart DB merge: provider IDs are strongest; phone/domain/name+address
                    # are used as cross-provider fallback signals. Existing status/notes survive.
                    provider_id = lead.get("provider_place_id")
                    place_id = lead.get("place_id")
                    phone = phone_key(lead.get("phone"))
                    domain = domain_key(lead.get("website"))
                    name_key = norm(lead.get("name"))
                    address_key = norm(lead.get("address"))
                    cur.execute("""SELECT id,user_id,place_id,provider_place_id,name,address,phone,website,sources,status,notes,created_at,updated_at,assigned_to
                                   FROM leads WHERE (provider_place_id=%s OR place_id=%s OR phone=%s OR lower(name)=lower(%s))
                                   ORDER BY id LIMIT 50""",
                                (provider_id, place_id, phone or None, lead.get("name")))
                    candidates = cur.fetchall()
                    existing = None
                    for cand in candidates:
                        cand_domain = domain_key(cand[7])
                        cand_name = norm(cand[4])
                        cand_addr = norm(cand[5])
                        if (provider_id and cand[3] == provider_id) or (place_id and cand[2] == place_id) or (phone and phone_key(cand[6]) == phone) or (domain and cand_domain == domain) or (name_key and cand_name == name_key and address_key and cand_addr == address_key):
                            existing = cand
                            break

                    if existing:
                        existing_id = existing[0]
                        existing_sources = existing[8] or []
                        merged_sources = list(dict.fromkeys([*(existing_sources if isinstance(existing_sources, list) else []), *(lead.get("sources") or []), lead.get("source") or ""]))
                        # Update only with available non-conflicting data; preserve CRM fields.
                        updates = {
                            "provider_place_id": lead.get("provider_place_id") or existing[3],
                            "name": lead.get("name") or existing[4],
                            "address": lead.get("address") or existing[5],
                            "phone": lead.get("phone") or existing[6],
                            "website": lead.get("website") or existing[7],
                            "city": lead.get("city"), "region": lead.get("region"), "country": lead.get("country"),
                            "latitude": lead.get("latitude"), "longitude": lead.get("longitude"), "email": lead.get("email"),
                            "instagram": lead.get("instagram"), "facebook": lead.get("facebook"), "twitter": lead.get("twitter"),
                            "category": lead.get("category"), "rating": lead.get("rating"), "reviews": lead.get("reviews"),
                            "maps_url": lead.get("maps_url"), "source": ", ".join(merged_sources), "sources": json.dumps(merged_sources),
                            "lead_score": lead.get("lead_score"), "lead_type": lead.get("lead_type"),
                            "ai_recommendation": lead.get("ai_recommendation"), "ai_reason": lead.get("ai_reason"),
                            "enrichment_source": lead.get("enrichment_source"), "enrichment_at": lead.get("enrichment_at"),
                        }
                        row = None
                        if existing[1] == current_user["id"]:
                            cur.execute("""UPDATE leads SET provider_place_id=%(provider_place_id)s,name=%(name)s,address=%(address)s,city=%(city)s,region=%(region)s,country=%(country)s,
                                latitude=%(latitude)s,longitude=%(longitude)s,phone=%(phone)s,website=%(website)s,email=%(email)s,instagram=%(instagram)s,facebook=%(facebook)s,twitter=%(twitter)s,
                                category=%(category)s,rating=%(rating)s,reviews=%(reviews)s,maps_url=%(maps_url)s,source=%(source)s,sources=%(sources)s::jsonb,lead_score=%(lead_score)s,
                                lead_type=%(lead_type)s,ai_recommendation=%(ai_recommendation)s,ai_reason=%(ai_reason)s,enrichment_source=%(enrichment_source)s,enrichment_at=%(enrichment_at)s,updated_at=NOW()
                                WHERE id=%(id)s RETURNING id,status,notes,created_at,updated_at,assigned_to,last_contacted_date,next_followup_date,is_duplicate,duplicate_of""",
                                {**updates, "id": existing[0]})
                            row = cur.fetchone()
                    else:
                        cur.execute("""INSERT INTO leads(
                            user_id,place_id,provider_place_id,name,address,city,region,country,latitude,longitude,
                            phone,website,email,instagram,facebook,twitter,category,rating,reviews,maps_url,source,sources,
                            lead_score,lead_type,ai_recommendation,ai_reason,status,assigned_to,enrichment_source,enrichment_at,created_at,updated_at)
                            VALUES(%(user_id)s,%(place_id)s,%(provider_place_id)s,%(name)s,%(address)s,%(city)s,%(region)s,%(country)s,%(latitude)s,%(longitude)s,
                            %(phone)s,%(website)s,%(email)s,%(instagram)s,%(facebook)s,%(twitter)s,%(category)s,%(rating)s,%(reviews)s,%(maps_url)s,%(source)s,%(sources)s::jsonb,
                            %(lead_score)s,%(lead_type)s,%(ai_recommendation)s,%(ai_reason)s,'New',%(assigned_to)s,%(enrichment_source)s,%(enrichment_at)s,NOW(),NOW())
                            RETURNING id,status,notes,created_at,updated_at,assigned_to,last_contacted_date,next_followup_date,is_duplicate,duplicate_of""", {
                                **lead, "user_id": current_user["id"], "assigned_to": current_user["id"], "sources": json.dumps(lead.get("sources") or []),
                                "enrichment_source": lead.get("enrichment_source"), "enrichment_at": lead.get("enrichment_at")})
                        row = cur.fetchone()
                    if existing:
                        lead["is_duplicate"] = True
                        lead["duplicate_of"] = existing[0]
                        duplicates.append({"id": existing[0], "name": existing[4], "status": existing[9], "assigned_to": existing[13]})
                        log_activity(cur,current_user["id"],existing[0],"duplicate_skipped",{"matched_lead_id":existing[0],"matched_name":existing[4]})
                        # Existing CRM record is updated only when owned by this user; the repeated search result is never re-added.
                        continue
                    assigned_value = row[5] if row else current_user["id"]
                    lead.update({"id": row[0], "status": row[1], "notes": row[2], "created_at": str(row[3]), "updated_at": str(row[4]), "assigned_to": assigned_value})
                    saved.append(lead)

                cur.execute("""INSERT INTO search_history(user_id,city,category,keyword,city_wide,result_count,sources_used)
                               VALUES(%s,%s,%s,%s,%s,%s,%s::jsonb)""",
                            (current_user["id"], location["resolved_city"], category, keyword, req.city_wide, len(saved), json.dumps(used_sources)))
            conn.commit()
    else:
        for lead in leads:
            lead.update({"id": None, "status": "New", "notes": None, "created_at": None, "updated_at": None, "assigned_to": current_user["id"]})
            saved.append(lead)

    saved.sort(key=lambda x: (0 if x.get("lead_type")=="Hot Lead" else 1 if x.get("lead_type")=="Warm Lead" else 2, -(x.get("lead_score") or 0)))
    return {"city_resolved": location["resolved_city"], "category": category, "keyword": keyword,
            "city_wide": req.city_wide, "radius_meters": radius_meters, "count": len(saved),
            "duplicates_skipped": len(duplicates), "duplicates": duplicates,
            "sources_used": used_sources, "warnings": list(dict.fromkeys(warnings)), "leads": saved}
# ============================================================

# Saved leads / filters / updates

# ============================================================

LEAD_SELECT = """SELECT id,name,address,city,region,country,phone,website,email,instagram,facebook,twitter,category,rating,reviews,maps_url,source,lead_score,lead_type,ai_recommendation,ai_reason,status,place_id,created_at,updated_at,notes,latitude,longitude,provider_place_id,sources,enrichment_source,enrichment_at,assigned_to,last_contacted_date,next_followup_date,is_duplicate,duplicate_of FROM leads"""

def row_to_lead(row):

    return {

        "id": row[0], "name": row[1], "address": row[2], "city": row[3], "region": row[4], "country": row[5],

        "phone": row[6], "website": row[7], "email": row[8], "instagram": row[9], "facebook": row[10], "twitter": row[11],

        "category": row[12], "rating": row[13], "reviews": row[14], "maps_url": row[15], "source": row[16], "lead_score": row[17],

        "lead_type": row[18], "ai_recommendation": row[19], "ai_reason": row[20], "status": row[21], "place_id": row[22],

        "created_at": str(row[23]), "updated_at": str(row[24]), "notes": row[25], "latitude": row[26], "longitude": row[27], "provider_place_id": row[28],

        "sources": row[29] or [], "enrichment_source": row[30], "enrichment_at": str(row[31]) if row[31] else None,
        "whatsapp_url": whatsapp_url(row[6]), "assigned_to": row[32],
        "last_contacted_date": str(row[33]) if row[33] else None, "next_followup_date": str(row[34]) if row[34] else None,
        "is_duplicate": bool(row[35]), "duplicate_of": row[36],

    }

@app.get("/api/leads")
def get_saved_leads(has_website: Optional[bool]=None,has_phone: Optional[bool]=None,has_email: Optional[bool]=None,has_social: Optional[bool]=None,min_score: Optional[int]=Query(None,ge=0,le=100),lead_type: Optional[str]=None,min_rating: Optional[float]=Query(None,ge=0,le=5),category: Optional[str]=None,city: Optional[str]=None,source: Optional[str]=None,status: Optional[str]=None,assigned_to: Optional[int]=None,current_user: dict=Depends(get_current_user)):
    scope,params=lead_access_clause(current_user); clauses=[scope]; params=list(params)
    for flag,col in [(has_website,"website"),(has_phone,"phone"),(has_email,"email")]:
        if flag is True: clauses.append(f"NULLIF(TRIM({col}),'') IS NOT NULL")
        elif flag is False: clauses.append(f"NULLIF(TRIM({col}),'') IS NULL")
    if has_social is True: clauses.append("(NULLIF(TRIM(instagram),'') IS NOT NULL OR NULLIF(TRIM(facebook),'') IS NOT NULL OR NULLIF(TRIM(twitter),'') IS NOT NULL)")
    elif has_social is False: clauses.append("(NULLIF(TRIM(instagram),'') IS NULL AND NULLIF(TRIM(facebook),'') IS NULL AND NULLIF(TRIM(twitter),'') IS NULL)")
    if min_score is not None: clauses.append("COALESCE(lead_score,0)>=%s"); params.append(min_score)
    if lead_type: clauses.append("lead_type=%s"); params.append(lead_type)
    if min_rating is not None: clauses.append("COALESCE(rating,0)>=%s"); params.append(min_rating)
    for val,col in [(category,"category"),(city,"city"),(source,"source"),(status,"status")]:
        if val: clauses.append(f"{col} ILIKE %s"); params.append(f"%{val}%")
    if assigned_to is not None:
        if not is_admin(current_user) and assigned_to!=current_user["id"]: raise HTTPException(status_code=403,detail="Agents can only view their assigned leads.")
        clauses.append("assigned_to=%s"); params.append(assigned_to)
    order="CASE WHEN lead_type='Hot Lead' THEN 0 WHEN lead_type='Warm Lead' THEN 1 ELSE 2 END,COALESCE(lead_score,0) DESC,created_at DESC"
    with db_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(LEAD_SELECT+" WHERE "+" AND ".join(clauses)+" ORDER BY "+order,params); rows=cur.fetchall()
    return {"success":True,"total":len(rows),"leads":[row_to_lead(r) for r in rows]}

@app.patch("/api/team-members/{user_id}/role")
def update_team_member_role(user_id:int, role:str, current_user:dict=Depends(get_current_user)):
    require_admin(current_user)
    role=role.strip().title()
    if role not in VALID_ROLES: raise HTTPException(status_code=400,detail="Role must be Admin or Agent.")
    with db_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT id,COALESCE(role,'Agent') FROM users WHERE id=%s",(user_id,))
            target=cur.fetchone()
            if not target: raise HTTPException(status_code=404,detail="Team member not found.")
            if target[1] == "Admin" and role == "Agent":
                cur.execute("SELECT COUNT(*) FROM users WHERE role='Admin'")
                if (cur.fetchone() or [0])[0] <= 1:
                    raise HTTPException(status_code=400,detail="At least one Admin account must remain.")
            cur.execute("UPDATE users SET role=%s WHERE id=%s",(role,user_id))
            log_activity(cur,current_user["id"],None,"team_role_changed",{"user_id":user_id,"role":role})
        conn.commit()
    return {"success":True,"user_id":user_id,"role":role}

@app.get("/api/team-members")
def team_members(current_user: dict=Depends(get_current_user)):
    with db_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT id,email,name,COALESCE(role,'Agent') FROM users ORDER BY name NULLS LAST,email"); rows=cur.fetchall()
    return {"team_members":[{"id":r[0],"email":r[1],"name":r[2],"role":r[3]} for r in rows]}

@app.get("/api/dashboard/followups")
def dashboard_followups(current_user: dict=Depends(get_current_user)):
    scope,base_params=lead_access_clause(current_user,"l"); now=datetime.now(timezone.utc); start_utc,end_utc,today_name=local_today_utc_bounds()
    with db_conn() as conn:
        with conn.cursor() as cur:
            sql=LEAD_SELECT.replace("FROM leads","FROM leads l")+" WHERE "+scope+" AND next_followup_date IS NOT NULL AND next_followup_date < %s AND status NOT IN ('Converted','Lost') ORDER BY next_followup_date ASC"
            cur.execute(sql,base_params+[now]); missed=[row_to_lead(r) for r in cur.fetchall()]
            sql=LEAD_SELECT.replace("FROM leads","FROM leads l")+" WHERE "+scope+" AND next_followup_date >= %s AND next_followup_date < %s AND status NOT IN ('Converted','Lost') ORDER BY next_followup_date ASC"
            cur.execute(sql,base_params+[start_utc,end_utc]); today_rows=[row_to_lead(r) for r in cur.fetchall()]
    return {"today":today_rows,"missed":missed,"today_count":len(today_rows),"missed_count":len(missed),"date":today_name,"timezone":user_timezone_name()}

@app.get("/api/dashboard/summary")
def dashboard_summary(current_user: dict = Depends(get_current_user)):
    """Simple dashboard totals for the leads visible to the current user."""
    scope, params = lead_access_clause(current_user, "l")
    with db_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(f"""
                SELECT
                    COUNT(*) AS total_leads,
                    COUNT(*) FILTER (WHERE l.last_contacted_date IS NOT NULL) AS contacted_leads,
                    (SELECT COUNT(*) FROM call_logs c JOIN leads x ON x.id=c.lead_id WHERE {scope.replace('l.', 'x.')}) AS total_calls,
                    COUNT(*) FILTER (WHERE l.next_followup_date IS NOT NULL AND l.next_followup_date >= %s AND l.next_followup_date < %s AND l.status NOT IN ('Converted','Lost')) AS today_followups,
                    COUNT(*) FILTER (WHERE l.next_followup_date IS NOT NULL AND l.next_followup_date < %s AND l.status NOT IN ('Converted','Lost')) AS missed_followups
                FROM leads l
                WHERE {scope}
            """, params + list(local_today_utc_bounds()[:2]) + [datetime.now(timezone.utc)] + params)
            row = cur.fetchone()
    return {
        "total_leads": row[0] or 0,
        "contacted_leads": row[1] or 0,
        "total_calls": row[2] or 0,
        "today_followups": row[3] or 0,
        "missed_followups": row[4] or 0,
    }

@app.get("/api/dashboard/agent-performance")
def agent_performance(
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
    current_user: dict = Depends(get_current_user),
):
    """Professional agent scorecard with independently calculated metrics.

    Date filters are built dynamically so PostgreSQL never has to infer the
    datatype of a NULL parameter. This also prevents JOIN multiplication.
    """
    start = None
    end = None
    if date_from:
        try:
            parsed = datetime.fromisoformat(date_from)
            start = parsed if parsed.tzinfo else parsed.replace(tzinfo=timezone.utc)
        except ValueError:
            raise HTTPException(status_code=400, detail="Invalid date_from.")
    if date_to:
        try:
            parsed = datetime.fromisoformat(date_to)
            parsed = parsed if parsed.tzinfo else parsed.replace(tzinfo=timezone.utc)
            end = parsed + timedelta(days=1)
        except ValueError:
            raise HTTPException(status_code=400, detail="Invalid date_to.")

    def date_clause(column: str, start_value, end_value):
        parts = []
        args = []
        if start_value is not None:
            parts.append(f"{column} >= %s")
            args.append(start_value)
        if end_value is not None:
            parts.append(f"{column} < %s")
            args.append(end_value)
        return (" AND " + " AND ".join(parts)) if parts else "", args

    with db_conn() as conn:
        with conn.cursor() as cur:
            is_admin = str(current_user.get("role") or "Agent").strip().lower() == "admin"
            if is_admin:
                cur.execute("SELECT id, name, email FROM users ORDER BY name NULLS LAST, id")
            else:
                cur.execute("SELECT id, name, email FROM users WHERE id=%s", (current_user["id"],))
            users = cur.fetchall()

            out = []
            for uid, name, email in users:
                lead_date_sql, lead_date_args = date_clause("created_at", start, end)
                cur.execute(
                    f"SELECT COUNT(*) FROM leads WHERE assigned_to=%s{lead_date_sql}",
                    [uid, *lead_date_args],
                )
                assigned = cur.fetchone()[0] or 0

                call_date_sql, call_date_args = date_clause("call_date", start, end)
                cur.execute(
                    f"""
                    SELECT
                        COUNT(*),
                        COUNT(*) FILTER (WHERE lower(trim(COALESCE(call_outcome,''))) IN ('connected','interested','follow-up','converted')),
                        COUNT(*) FILTER (WHERE lower(trim(COALESCE(call_outcome,'')))='interested'),
                        COUNT(*) FILTER (WHERE lower(trim(COALESCE(call_outcome,'')))='follow-up'),
                        COUNT(*) FILTER (WHERE lower(trim(COALESCE(call_outcome,'')))='converted'),
                        COUNT(DISTINCT lead_id) FILTER (WHERE lower(trim(COALESCE(call_outcome,''))) IN ('connected','interested','follow-up','converted'))
                    FROM call_logs
                    WHERE called_by=%s{call_date_sql}
                    """,
                    [uid, *call_date_args],
                )
                c = cur.fetchone()
                calls, connected, interested, followups, converted_calls, contacted = [x or 0 for x in c]

                updated_sql, updated_args = date_clause("updated_at", start, end)
                cur.execute(
                    f"SELECT COUNT(*) FROM leads WHERE assigned_to=%s AND status='Converted'{updated_sql}",
                    [uid, *updated_args],
                )
                converted_status = cur.fetchone()[0] or 0
                converted = max(converted_calls, converted_status)

                followup_sql, followup_args = date_clause("next_followup_date", start, end)
                cur.execute(
                    f"""
                    SELECT COUNT(*) FROM leads
                    WHERE assigned_to=%s
                      AND next_followup_date IS NOT NULL
                      {followup_sql}
                      AND status NOT IN ('Converted','Lost')
                    """,
                    [uid, *followup_args],
                )
                scheduled_followups = cur.fetchone()[0] or 0

                contact_rate = round((connected / calls) * 100, 1) if calls else 0.0
                conversion_rate = round((converted / contacted) * 100, 1) if contacted else 0.0
                activity_rate = round(((connected + followups + interested) / calls) * 100, 1) if calls else 0.0

                out.append({
                    "id": uid, "name": name, "email": email,
                    "assigned_leads": assigned,
                    "contacts": contacted,
                    "calls": calls,
                    "connected": connected,
                    "interested": interested,
                    "followups": followups + scheduled_followups,
                    "scheduled_followups": scheduled_followups,
                    "converted": converted,
                    "converted_calls": converted_calls,
                    "contact_rate": contact_rate,
                    "conversion_rate": conversion_rate,
                    "activity_rate": activity_rate,
                })

    out.sort(key=lambda x: (x["converted"], x["contacts"], x["connected"], x["calls"]), reverse=True)
    for i, item in enumerate(out, 1):
        item["rank"] = i
        item["performance_score"] = round(
            min(item["contact_rate"], 100) * 0.30
            + min(item["conversion_rate"], 100) * 0.40
            + min((item["connected"] / max(item["calls"], 1)) * 100, 100) * 0.20
            + min(item["followups"] / max(item["assigned_leads"], 1) * 100, 100) * 0.10,
            1,
        )
    return {"performance": out, "date_from": date_from, "date_to": date_to}

@app.patch("/api/leads/{lead_id}/status")
def update_status(lead_id:int,req:StatusUpdateRequest,current_user:dict=Depends(get_current_user)):
    if req.status not in VALID_STATUSES: raise HTTPException(status_code=400,detail="Invalid status.")
    with db_conn() as conn:
        with conn.cursor() as cur:
            lead_owner_check(cur,lead_id,current_user,True); cur.execute("UPDATE leads SET status=%s,updated_at=NOW() WHERE id=%s",(req.status,lead_id)); log_activity(cur,current_user["id"],lead_id,"status_changed",{"status":req.status})
        conn.commit()
    return {"success":True,"lead_id":lead_id,"status":req.status}

@app.patch("/api/leads/{lead_id}/assignment")
def assign_lead(lead_id:int,req:AssignmentRequest,current_user:dict=Depends(get_current_user)):
    require_admin(current_user)
    with db_conn() as conn:
        with conn.cursor() as cur:
            lead_owner_check(cur,lead_id,current_user,True)
            if req.assigned_to is not None:
                cur.execute("SELECT id FROM users WHERE id=%s",(req.assigned_to,));
                if not cur.fetchone(): raise HTTPException(status_code=404,detail="Team member not found.")
            cur.execute("UPDATE leads SET assigned_to=%s,updated_at=NOW() WHERE id=%s",(req.assigned_to,lead_id)); log_activity(cur,current_user["id"],lead_id,"assigned",{"assigned_to":req.assigned_to})
        conn.commit()
    return {"success":True,"lead_id":lead_id,"assigned_to":req.assigned_to}

@app.patch("/api/leads/{lead_id}/followup")
def set_followup(lead_id:int,req:FollowupRequest,current_user:dict=Depends(get_current_user)):
    with db_conn() as conn:
        with conn.cursor() as cur:
            lead_owner_check(cur,lead_id,current_user,True); cur.execute("UPDATE leads SET next_followup_date=%s,updated_at=NOW() WHERE id=%s",(req.next_followup_date,lead_id)); log_activity(cur,current_user["id"],lead_id,"followup_changed",{"next_followup_date":req.next_followup_date.isoformat() if req.next_followup_date else None})
        conn.commit()
    return {"success":True,"lead_id":lead_id,"next_followup_date":req.next_followup_date}

@app.post("/api/leads/{lead_id}/calls")
def add_call_log(lead_id:int,req:CallLogRequest,current_user:dict=Depends(get_current_user)):
    with db_conn() as conn:
        with conn.cursor() as cur:
            lead_owner_check(cur,lead_id,current_user,True); call_date=req.call_date or datetime.now(timezone.utc)
            cur.execute("INSERT INTO call_logs(lead_id,called_by,call_date,call_outcome,notes) VALUES(%s,%s,%s,%s,%s) RETURNING id",(lead_id,current_user["id"],call_date,req.call_outcome.strip(),req.notes)); call_id=cur.fetchone()[0]
            cur.execute("UPDATE leads SET last_contacted_date=%s,status=CASE WHEN status='New' THEN 'Contacted' ELSE status END,updated_at=NOW() WHERE id=%s",(call_date,lead_id)); log_activity(cur,current_user["id"],lead_id,"call_logged",{"call_id":call_id,"outcome":req.call_outcome.strip()})
        conn.commit()
    return {"success":True,"call_id":call_id,"lead_id":lead_id,"last_contacted_date":str(call_date)}

@app.get("/api/leads/{lead_id}/calls")
def get_call_logs(lead_id:int,current_user:dict=Depends(get_current_user)):
    # Always authorize the lead before reading its call history.
    # This used to call lead_owner_check without current_user, which caused:
    # TypeError: lead_owner_check() missing 1 required positional argument: 'user'
    with db_conn() as conn:
        with conn.cursor() as cur:
            lead_owner_check(cur, lead_id, current_user)
            cur.execute(
                """
                SELECT c.id, c.called_by, c.call_date, c.call_outcome, c.notes,
                       u.name, u.email
                FROM call_logs c
                LEFT JOIN users u ON u.id = c.called_by
                WHERE c.lead_id = %s
                ORDER BY c.call_date DESC, c.id DESC
                """,
                (lead_id,),
            )
            rows = cur.fetchall()
    return {
        "calls": [
            {
                "id": r[0],
                "called_by": r[1],
                "call_date": str(r[2]),
                "call_outcome": r[3],
                "notes": r[4],
                "called_by_name": r[5],
                "called_by_email": r[6],
            }
            for r in rows
        ]
    }

@app.patch("/api/leads/{lead_id}/notes")
def update_notes(lead_id:int,req:NotesUpdateRequest,current_user:dict=Depends(get_current_user)):
    with db_conn() as conn:
        with conn.cursor() as cur:
            lead_owner_check(cur,lead_id, current_user,True); cur.execute("UPDATE leads SET notes=%s,updated_at=NOW() WHERE id=%s",((req.notes or "")[:5000],lead_id)); log_activity(cur,current_user["id"],lead_id,"notes_changed",{})
        conn.commit()
    return {"success":True,"lead_id":lead_id,"notes":req.notes or ""}

@app.delete("/api/leads/{lead_id}")
def delete_lead(lead_id:int,current_user:dict=Depends(get_current_user)):
    with db_conn() as conn:
        with conn.cursor() as cur:
            lead_owner_check(cur,lead_id,current_user,True); cur.execute("DELETE FROM leads WHERE id=%s",(lead_id,))
        conn.commit()
    return {"success":True}

@app.patch("/api/leads/{lead_id}")
def edit_lead(lead_id:int,req:LeadUpdateRequest,current_user:dict=Depends(get_current_user)):
    data=req.model_dump(exclude_unset=True)
    if "status" in data and data["status"] not in VALID_STATUSES: raise HTTPException(status_code=400,detail="Invalid status.")
    if "assigned_to" in data and not is_admin(current_user): raise HTTPException(status_code=403,detail="Only Admin can assign leads.")
    allowed={"name","address","city","region","country","phone","website","email","instagram","facebook","twitter","category","notes","status","assigned_to","last_contacted_date","next_followup_date"}; data={k:v for k,v in data.items() if k in allowed}
    if not data: raise HTTPException(status_code=400,detail="No changes supplied.")
    with db_conn() as conn:
        with conn.cursor() as cur:
            lead_owner_check(cur,lead_id,current_user,True)
            if data.get("assigned_to") is not None:
                cur.execute("SELECT id FROM users WHERE id=%s",(data["assigned_to"],));
                if not cur.fetchone(): raise HTTPException(status_code=404,detail="Team member not found.")
            cur.execute(f"UPDATE leads SET {','.join(f'{k}=%s' for k in data)},updated_at=NOW() WHERE id=%s",list(data.values())+[lead_id]); log_activity(cur,current_user["id"],lead_id,"lead_updated",data)
        conn.commit()
    return {"success":True,"lead_id":lead_id}

@app.post("/api/leads/{lead_id}/whatsapp")
def lead_whatsapp(lead_id:int,current_user:dict=Depends(get_current_user)):
    with db_conn() as conn:
        with conn.cursor() as cur:
            lead_owner_check(cur,lead_id,current_user); cur.execute("SELECT phone FROM leads WHERE id=%s",(lead_id,)); phone=cur.fetchone()[0]
    url=whatsapp_url(phone)
    if not url: raise HTTPException(status_code=400,detail="Lead has no usable phone number.")
    return {"success":True,"whatsapp_url":url,"manual_only":True}

@app.post("/api/leads/bulk")
def bulk_lead_action(req:BulkLeadActionRequest,current_user:dict=Depends(get_current_user)):
    ids=list(dict.fromkeys(req.lead_ids)); action=req.action.strip().lower()
    with db_conn() as conn:
        with conn.cursor() as cur:
            scope,params=lead_access_clause(current_user)
            if action=="status":
                if req.status not in VALID_STATUSES: raise HTTPException(status_code=400,detail="Invalid status.")
                cur.execute(f"UPDATE leads SET status=%s,updated_at=NOW() WHERE id=ANY(%s) AND {scope}",(req.status,ids,*params)); count=cur.rowcount
            elif action=="assign":
                require_admin(current_user)
                if req.assigned_to is None: raise HTTPException(status_code=400,detail="assigned_to is required.")
                cur.execute("SELECT id FROM users WHERE id=%s",(req.assigned_to,));
                if not cur.fetchone(): raise HTTPException(status_code=404,detail="Team member not found.")
                cur.execute("UPDATE leads SET assigned_to=%s,updated_at=NOW() WHERE id=ANY(%s)",(req.assigned_to,ids)); count=cur.rowcount
            elif action=="delete":
                cur.execute(f"DELETE FROM leads WHERE id=ANY(%s) AND {scope}",(ids,*params)); count=cur.rowcount
            else: raise HTTPException(status_code=400,detail="Unsupported bulk action.")
            if action == "status":
                cur.execute(f"SELECT id FROM leads WHERE id=ANY(%s) AND {scope}",(ids,*params))
                for (lead_id,) in cur.fetchall():
                    log_activity(cur,current_user["id"],lead_id,"bulk_status_changed",{"status":req.status})
            elif action == "assign":
                cur.execute("SELECT id FROM leads WHERE id=ANY(%s)",(ids,))
                for (lead_id,) in cur.fetchall():
                    log_activity(cur,current_user["id"],lead_id,"bulk_assigned",{"assigned_to":req.assigned_to})
            elif action == "delete":
                # Deletions cascade activity rows by design; record the action is not possible after the lead is deleted.
                pass
        conn.commit()
    return {"success":True,"updated":count} if action!="delete" else {"success":True,"deleted":count}

@app.get("/api/activity-log")
def activity_log(limit:int=Query(100,ge=1,le=500),lead_id:Optional[int]=None,current_user:dict=Depends(get_current_user)):
    scope,params=lead_access_clause(current_user,"l"); where=[scope]; params=list(params)
    if lead_id is not None: where.append("a.lead_id=%s"); params.append(lead_id)
    with db_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT a.id,a.user_id,a.lead_id,a.action,a.details,a.created_at,u.name,u.email,l.name FROM activity_logs a LEFT JOIN users u ON u.id=a.user_id LEFT JOIN leads l ON l.id=a.lead_id WHERE "+" AND ".join(where)+" ORDER BY a.created_at DESC LIMIT %s",params+[limit]); rows=cur.fetchall()
    return {"activity":[{"id":r[0],"user_id":r[1],"lead_id":r[2],"action":r[3],"details":r[4] or {},"created_at":str(r[5]),"user_name":r[6],"user_email":r[7],"lead_name":r[8]} for r in rows]}

# Groups

# ============================================================

def assert_group_owner(cur, group_id: int, user_id: int):

    cur.execute("SELECT id FROM lead_groups WHERE id=%s AND user_id=%s", (group_id, user_id))

    if not cur.fetchone(): raise HTTPException(status_code=404, detail="Group not found.")

@app.get("/api/groups")

def list_groups(current_user: dict = Depends(get_current_user)):

    with db_conn() as conn:

        with conn.cursor() as cur:

            cur.execute("SELECT id,name,created_at FROM lead_groups WHERE user_id=%s ORDER BY created_at DESC", (current_user["id"],))

            rows = cur.fetchall()

    return {"groups": [{"id": r[0], "name": r[1], "created_at": str(r[2])} for r in rows]}

@app.post("/api/groups")

def create_group(req: GroupCreateRequest, current_user: dict = Depends(get_current_user)):

    name = req.name.strip()

    if not name: raise HTTPException(status_code=400, detail="Group name is required.")

    with db_conn() as conn:

        with conn.cursor() as cur:

            cur.execute("INSERT INTO lead_groups(user_id,name) VALUES(%s,%s) RETURNING id,name,created_at", (current_user["id"], name[:150]))

            row = cur.fetchone()

        conn.commit()

    return {"group": {"id": row[0], "name": row[1], "created_at": str(row[2])}}

@app.post("/api/groups/{group_id}/leads")
def add_lead_to_group(group_id: int, req: GroupBulkAddRequest, current_user: dict = Depends(get_current_user)):
    if not req.lead_ids: raise HTTPException(status_code=400, detail="No lead selected.")
    added = 0
    with db_conn() as conn:
        with conn.cursor() as cur:
            assert_group_owner(cur, group_id, current_user["id"])
            for lead_id in req.lead_ids[:500]:
                lead_scope, lead_params = lead_access_clause(current_user)
                cur.execute(f"SELECT id FROM leads WHERE id=%s AND {lead_scope}", (lead_id, *lead_params))
                if cur.fetchone():
                    cur.execute("INSERT INTO group_leads(group_id,lead_id) VALUES(%s,%s) ON CONFLICT DO NOTHING", (group_id, lead_id))
                    added += cur.rowcount
        conn.commit()
    return {"success": True, "added": added}

@app.post("/api/groups/{group_id}/leads/bulk")

def add_leads_to_group_bulk(group_id: int, req: GroupBulkAddRequest, current_user: dict = Depends(get_current_user)):

    if not req.lead_ids: raise HTTPException(status_code=400, detail="No leads selected.")

    added = 0

    with db_conn() as conn:

        with conn.cursor() as cur:

            assert_group_owner(cur, group_id, current_user["id"])

            for lead_id in req.lead_ids[:500]:

                lead_scope, lead_params = lead_access_clause(current_user)
                cur.execute(f"SELECT id FROM leads WHERE id=%s AND {lead_scope}", (lead_id, *lead_params))

                if cur.fetchone():

                    cur.execute("INSERT INTO group_leads(group_id,lead_id) VALUES(%s,%s) ON CONFLICT DO NOTHING", (group_id, lead_id))

                    added += cur.rowcount

        conn.commit()

    return {"success": True, "added": added}

@app.delete("/api/groups/{group_id}/leads/{lead_id}")
def remove_lead_from_group(group_id: int, lead_id: int, current_user: dict = Depends(get_current_user)):
    with db_conn() as conn:
        with conn.cursor() as cur:
            assert_group_owner(cur, group_id, current_user["id"])
            cur.execute("DELETE FROM group_leads WHERE group_id=%s AND lead_id=%s", (group_id, lead_id))
            removed = cur.rowcount
        conn.commit()
    return {"success": True, "removed": removed}

@app.get("/api/groups/{group_id}/export-excel")
def export_group_excel(group_id:int,current_user:dict=Depends(get_current_user)):
    with db_conn() as conn:
        with conn.cursor() as cur:
            assert_group_owner(cur,group_id,current_user["id"])
            scope,params=lead_access_clause(current_user,"l")
            cur.execute("""SELECT l.name,l.phone,l.email,l.address,l.city,l.region,l.country,l.website,l.instagram,l.facebook,l.twitter,l.rating,l.reviews,l.category,l.lead_score,l.lead_type,l.ai_recommendation,l.ai_reason,l.status,l.assigned_to,l.last_contacted_date,l.next_followup_date,l.source,l.maps_url,l.notes,c.call_outcome,c.notes FROM leads l JOIN group_leads gl ON gl.lead_id=l.id LEFT JOIN LATERAL (SELECT call_outcome,notes FROM call_logs WHERE lead_id=l.id ORDER BY call_date DESC,id DESC LIMIT 1)c ON TRUE WHERE gl.group_id=%s AND """+scope+" ORDER BY l.created_at DESC",(group_id,*params))
            rows=cur.fetchall()
    wb=Workbook(); ws=wb.active; ws.title="Group Leads"; ws.append(EXPORT_HEADERS)
    for row in rows: ws.append(row)
    ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
    for col in ws.columns:
        width=min(max(max(len(str(c.value or "")) for c in col)+2,12),45); ws.column_dimensions[col[0].column_letter].width=width
    buf=BytesIO(); wb.save(buf); buf.seek(0)
    return StreamingResponse(buf,media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",headers={"Content-Disposition":f'attachment; filename="group_{group_id}.xlsx"'})


@app.get("/api/groups/{group_id}/export-excel/")
def export_group_excel_trailing(group_id:int,current_user:dict=Depends(get_current_user)):
    return export_group_excel(group_id, current_user)

@app.get("/api/groups/{group_id}/leads")

def get_group_leads(group_id: int, current_user: dict = Depends(get_current_user)):

    with db_conn() as conn:

        with conn.cursor() as cur:

            assert_group_owner(cur, group_id, current_user["id"])

            lead_scope, lead_params = lead_access_clause(current_user,"l")
            cur.execute("""SELECT l.id,l.name,l.address,l.city,l.region,l.country,l.phone,l.website,l.email,l.instagram,l.facebook,l.twitter,l.category,l.rating,l.reviews,l.maps_url,l.source,l.lead_score,l.lead_type,l.ai_recommendation,l.ai_reason,l.status,l.place_id,l.created_at,l.updated_at,l.notes,l.latitude,l.longitude,l.provider_place_id,l.sources,l.enrichment_source,l.enrichment_at,l.assigned_to,l.last_contacted_date,l.next_followup_date,l.is_duplicate,l.duplicate_of FROM leads l JOIN group_leads gl ON gl.lead_id=l.id WHERE gl.group_id=%s AND """+lead_scope+" ORDER BY l.created_at DESC", (group_id, *lead_params))

            rows = cur.fetchall()

    return {"leads": [row_to_lead(r) for r in rows]}

# ============================================================

# Search history / export / health

# ============================================================

@app.get("/api/search-history")
def search_history(current_user: dict = Depends(get_current_user)):
    with db_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT id,city,category,keyword,city_wide,result_count,sources_used,created_at FROM search_history WHERE user_id=%s ORDER BY created_at DESC LIMIT 50", (current_user["id"],))
            rows = cur.fetchall()
    return {"history": [{"id": r[0], "city": r[1], "category": r[2], "keyword": r[3], "city_wide": r[4], "result_count": r[5], "sources_used": r[6] or [], "created_at": str(r[7])} for r in rows]}

@app.post("/api/search-history/{history_id}/rerun")
def rerun_search(history_id: int, current_user: dict = Depends(get_current_user)):
    with db_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT city,category,keyword,city_wide FROM search_history WHERE id=%s AND user_id=%s", (history_id, current_user["id"]))
            row = cur.fetchone()
    if not row:
        raise HTTPException(status_code=404, detail="Search history item not found.")
    req = SearchRequest(city=row[0], category=row[1], keyword=row[2], city_wide=bool(row[3]), limit=10)
    return search(req, current_user)

def _export_lead_rows(current_user:dict,lead_ids:Optional[str]=None,has_website:Optional[bool]=None,has_phone:Optional[bool]=None,has_email:Optional[bool]=None,has_social:Optional[bool]=None,min_score:Optional[int]=None,lead_type:Optional[str]=None,min_rating:Optional[float]=None,category:Optional[str]=None,city:Optional[str]=None,source:Optional[str]=None,status:Optional[str]=None,assigned_to:Optional[int]=None,date_from:Optional[str]=None,date_to:Optional[str]=None):
    scope,params=lead_access_clause(current_user); clauses=[scope]; params=list(params)
    if lead_ids:
        try: ids=[int(x) for x in lead_ids.split(',') if x.strip()][:500]
        except ValueError: raise HTTPException(status_code=400,detail="Invalid lead selection.")
        clauses.append("l.id=ANY(%s)"); params.append(ids)
    for flag,col in [(has_website,'website'),(has_phone,'phone'),(has_email,'email')]:
        if flag is True: clauses.append(f"NULLIF(TRIM({col}),'') IS NOT NULL")
        elif flag is False: clauses.append(f"NULLIF(TRIM({col}),'') IS NULL")
    if has_social is True: clauses.append("(NULLIF(TRIM(instagram),'') IS NOT NULL OR NULLIF(TRIM(facebook),'') IS NOT NULL OR NULLIF(TRIM(twitter),'') IS NOT NULL)")
    if min_score is not None: clauses.append("COALESCE(lead_score,0)>=%s"); params.append(min_score)
    if lead_type: clauses.append("lead_type=%s"); params.append(lead_type)
    if min_rating is not None: clauses.append("COALESCE(rating,0)>=%s"); params.append(min_rating)
    for val,col in [(category,'category'),(city,'city'),(source,'source'),(status,'status')]:
        if val: clauses.append(f"{col} ILIKE %s"); params.append(f"%{val}%")
    if assigned_to is not None: clauses.append("assigned_to=%s"); params.append(assigned_to)
    if date_from: clauses.append("l.created_at >= %s::date"); params.append(date_from)
    if date_to: clauses.append("l.created_at < (%s::date + INTERVAL '1 day')"); params.append(date_to)
    sql="SELECT l.name,l.phone,l.email,l.address,l.city,l.region,l.country,l.website,l.instagram,l.facebook,l.twitter,l.rating,l.reviews,l.category,l.lead_score,l.lead_type,l.ai_recommendation,l.ai_reason,l.status,l.assigned_to,l.last_contacted_date,l.next_followup_date,l.source,l.maps_url,l.notes,c.call_outcome,c.notes FROM leads l LEFT JOIN LATERAL (SELECT call_outcome,notes FROM call_logs WHERE lead_id=l.id ORDER BY call_date DESC,id DESC LIMIT 1) c ON TRUE WHERE "+" AND ".join(clauses)+" ORDER BY l.created_at DESC"
    with db_conn() as conn:
        with conn.cursor() as cur: cur.execute(sql,params); return cur.fetchall()

EXPORT_HEADERS=["Business Name","Phone","Email","Address","City","Region","Country","Website","Instagram","Facebook","Twitter/X","Rating","Reviews","Category","Lead Score","Lead Priority","AI Recommendation","AI Reason","Status","Assigned To","Last Contacted Date","Next Follow-up Date","Source","Maps URL","Notes","Latest Call Outcome","Latest Call Notes"]

def export_rows(current_user,**kwargs): return _export_lead_rows(current_user,**kwargs)

@app.get("/api/export-excel")
def export_excel(lead_ids:Optional[str]=None,has_website:Optional[bool]=None,has_phone:Optional[bool]=None,has_email:Optional[bool]=None,has_social:Optional[bool]=None,min_score:Optional[int]=Query(None,ge=0,le=100),lead_type:Optional[str]=None,min_rating:Optional[float]=Query(None,ge=0,le=5),category:Optional[str]=None,city:Optional[str]=None,source:Optional[str]=None,status:Optional[str]=None,assigned_to:Optional[int]=None,date_from:Optional[str]=None,date_to:Optional[str]=None,current_user:dict=Depends(get_current_user)):
    rows=_export_lead_rows(current_user,lead_ids,has_website,has_phone,has_email,has_social,min_score,lead_type,min_rating,category,city,source,status,assigned_to,date_from,date_to)
    wb=Workbook(); ws=wb.active; ws.title="Leads"; ws.append(EXPORT_HEADERS)
    for row in rows: ws.append(row)
    buffer=BytesIO(); wb.save(buffer); buffer.seek(0)
    return StreamingResponse(buffer,media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",headers={"Content-Disposition":"attachment; filename=leads_export.xlsx"})

@app.get("/api/export-csv")
def export_csv(lead_ids:Optional[str]=None,has_website:Optional[bool]=None,has_phone:Optional[bool]=None,has_email:Optional[bool]=None,has_social:Optional[bool]=None,min_score:Optional[int]=Query(None,ge=0,le=100),lead_type:Optional[str]=None,min_rating:Optional[float]=Query(None,ge=0,le=5),category:Optional[str]=None,city:Optional[str]=None,source:Optional[str]=None,status:Optional[str]=None,assigned_to:Optional[int]=None,date_from:Optional[str]=None,date_to:Optional[str]=None,current_user:dict=Depends(get_current_user)):
    rows=_export_lead_rows(current_user,lead_ids,has_website,has_phone,has_email,has_social,min_score,lead_type,min_rating,category,city,source,status,assigned_to,date_from,date_to)
    import io; sio=io.StringIO(); writer=csv.writer(sio); writer.writerow(EXPORT_HEADERS)
    for row in rows: writer.writerow(row)
    return StreamingResponse(iter([sio.getvalue().encode('utf-8-sig')]),media_type='text/csv',headers={'Content-Disposition':'attachment; filename=leads_export.csv'})

def _google_service_account():
    raw = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if not raw:
        return None
    try:
        if raw.startswith("{"):
            return json.loads(raw)
        path = Path(raw)
        if path.exists():
            return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None
    return None

def _google_access_token(service_account: dict) -> str:
    private_key = service_account.get("private_key")
    client_email = service_account.get("client_email")
    token_uri = service_account.get("token_uri") or "https://oauth2.googleapis.com/token"
    if not private_key or not client_email:
        raise ProviderError("Google service-account JSON is missing private_key or client_email.")
    now = int(datetime.now(timezone.utc).timestamp())
    assertion = jwt.encode({
        "iss": client_email,
        "scope": "https://www.googleapis.com/auth/spreadsheets",
        "aud": token_uri,
        "iat": now,
        "exp": now + 3600,
    }, private_key, algorithm="RS256", headers={"kid": service_account.get("private_key_id")} if service_account.get("private_key_id") else None)
    r = requests.post(token_uri, data={"grant_type":"urn:ietf:params:oauth:grant-type:jwt-bearer","assertion":assertion}, timeout=15)
    if r.status_code != 200:
        raise ProviderError("Google authorization failed. Check the service-account credentials.")
    data = r.json()
    if not data.get("access_token"):
        raise ProviderError("Google did not return an access token.")
    return data["access_token"]

@app.get("/api/google-sheets/status")
def google_sheets_status(current_user:dict=Depends(get_current_user)):
    configured=bool(os.getenv('GOOGLE_SHEETS_SPREADSHEET_ID') and _google_service_account())
    return {"enabled":configured,"configured":configured,"priority":"low","mode":"optional"}

@app.post("/api/google-sheets/sync")
def google_sheets_sync(current_user:dict=Depends(get_current_user)):
    require_admin(current_user)
    spreadsheet_id = os.getenv("GOOGLE_SHEETS_SPREADSHEET_ID", "").strip()
    service_account = _google_service_account()
    if not spreadsheet_id or not service_account:
        raise HTTPException(status_code=501, detail="Google Sheets is optional. Configure GOOGLE_SHEETS_SPREADSHEET_ID and GOOGLE_SERVICE_ACCOUNT_JSON first.")
    sheet_range = os.getenv("GOOGLE_SHEETS_RANGE", "Leads!A1")
    try:
        access_token = _google_access_token(service_account)
        rows = _export_lead_rows(current_user)
        values = [EXPORT_HEADERS] + [list(r) for r in rows]
        headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
        clear_url = f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}/values/{quote_plus(sheet_range)}:clear"
        clear = requests.post(clear_url, headers=headers, json={}, timeout=20)
        if clear.status_code not in (200, 204):
            raise ProviderError("Google Sheets clear operation failed. Check spreadsheet sharing and range.")
        update_url = f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}/values/{quote_plus(sheet_range)}"
        update = requests.put(update_url, headers=headers, params={"valueInputOption":"RAW"}, json={"range":sheet_range,"majorDimension":"ROWS","values":values}, timeout=30)
        if update.status_code != 200:
            raise ProviderError("Google Sheets update failed. Make sure the service account has Editor access to the sheet.")
        return {"success":True,"message":f"Synced {len(rows)} lead(s) to Google Sheets.","rows":len(rows)}
    except ProviderError as exc:
        raise HTTPException(status_code=502, detail=str(exc))
    except requests.RequestException:
        raise HTTPException(status_code=502, detail="Google Sheets is temporarily unavailable.")
    except Exception as exc:
        print("Google Sheets sync error:", type(exc).__name__)
        raise HTTPException(status_code=502, detail="Google Sheets sync failed.")

@app.get("/api/health")

def health():

    db_connected = False

    if DATABASE_URL:

        try:

            with db_conn() as conn:

                with conn.cursor() as cur:

                    cur.execute("SELECT 1"); cur.fetchone()
                    cur.execute("SELECT to_regclass('public.users'), to_regclass('public.leads'), to_regclass('public.call_logs'), to_regclass('public.activity_logs')")
                    tables = cur.fetchone()
                    db_connected = all(tables)

        except Exception:

            pass

    return {"status": "ok" if db_connected else "degraded", "database_configured": bool(DATABASE_URL), "database_connected": db_connected, "latlng_configured": bool(LATLNG_API_KEY), "serpapi_configured": bool(SERPAPI_API_KEY), "gemini_configured": bool(GEMINI_API_KEY), "jwt_configured": bool(JWT_SECRET), "crm_statuses": sorted(VALID_STATUSES), "roles_enabled": True, "followups_enabled": True, "call_logging_enabled": True, "google_sheets_optional": True}