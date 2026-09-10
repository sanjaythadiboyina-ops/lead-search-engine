import os
import re
import base64
import hashlib
import hmac
import json
import secrets
from datetime import datetime, timedelta, timezone
from io import BytesIO
from pathlib import Path
from typing import Optional, List
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

app = FastAPI(title="Lead Searching Engine", version="3.0")
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
VALID_STATUSES = {"Not Contacted", "Contacted", "Interested", "Follow-up", "Converted"}

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
                created_at TIMESTAMPTZ DEFAULT NOW())""")
            _add_columns(cur, "users", {
                "name": "ALTER TABLE users ADD COLUMN name TEXT",
                "password_hash": "ALTER TABLE users ADD COLUMN password_hash TEXT",
                "salt": "ALTER TABLE users ADD COLUMN salt TEXT",
                "created_at": "ALTER TABLE users ADD COLUMN created_at TIMESTAMPTZ DEFAULT NOW()",
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
                status TEXT DEFAULT 'Not Contacted', notes TEXT,
                enrichment_source TEXT, enrichment_at TIMESTAMPTZ,
                created_at TIMESTAMPTZ DEFAULT NOW())""")
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
                "status": "ALTER TABLE leads ADD COLUMN status TEXT DEFAULT 'Not Contacted'",
                "notes": "ALTER TABLE leads ADD COLUMN notes TEXT",
                "enrichment_source": "ALTER TABLE leads ADD COLUMN enrichment_source TEXT",
                "enrichment_at": "ALTER TABLE leads ADD COLUMN enrichment_at TIMESTAMPTZ",
            })

            # Preserve old data and old group memberships. Never drop group_leads on startup.
            cur.execute("""CREATE TABLE IF NOT EXISTS search_history (
                id SERIAL PRIMARY KEY, user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
                city TEXT, category TEXT, keyword TEXT, city_wide BOOLEAN DEFAULT FALSE,
                result_count INTEGER, created_at TIMESTAMPTZ DEFAULT NOW())""")
            _add_columns(cur, "search_history", {
                "keyword": "ALTER TABLE search_history ADD COLUMN keyword TEXT",
                "city_wide": "ALTER TABLE search_history ADD COLUMN city_wide BOOLEAN DEFAULT FALSE",
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
            cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS leads_user_place_unique ON leads(user_id, place_id) WHERE user_id IS NOT NULL AND place_id IS NOT NULL")
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
            cur.execute("SELECT id, email, name FROM users WHERE id=%s", (user_id,))
            user = cur.fetchone()
    if not user:
        raise HTTPException(status_code=401, detail="User not found.")
    return {"id": user[0], "email": user[1], "name": user[2]}

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
    enrich: bool = False

class StatusUpdateRequest(BaseModel):
    status: str

class NotesUpdateRequest(BaseModel):
    notes: Optional[str] = None

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
                cur.execute("INSERT INTO users(email,name,password_hash,salt) VALUES(%s,%s,%s,%s) RETURNING id", (email, name or None, password_hash, salt))
                user_id = cur.fetchone()[0]
            conn.commit()
    except HTTPException:
        raise
    except Exception as exc:
        print(f"Signup database error: {exc}")
        raise HTTPException(status_code=500, detail=f"Signup database error: {type(exc).__name__}.")
    return {"token": create_token(user_id, email), "user": {"id": user_id, "email": email, "name": name}}

@app.post("/api/login")
def login(req: LoginRequest):
    email = req.email.strip().lower()
    try:
        with db_conn() as conn:
            with conn.cursor() as cur:
                cur.execute("SELECT id,email,name,password_hash,salt FROM users WHERE email=%s", (email,))
                user = cur.fetchone()
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Login database error: {type(exc).__name__}.")
    if not user or not verify_password(req.password, user[3], user[4]):
        raise HTTPException(status_code=401, detail="Invalid email or password.")
    return {"token": create_token(user[0], user[1]), "user": {"id": user[0], "email": user[1], "name": user[2]}}

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
        if distance_km(center_lat, center_lon, lat, lon) > SEARCH_RADIUS_METERS / 1000:
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


def city_grid(location: dict, enabled: bool):
    if not enabled:
        return [location]
    from math import cos, radians
    offsets = [(0,0),(0,1),(1,0),(0,-1),(-1,0),(1,1),(1,-1),(-1,1),(-1,-1)][:CITY_WIDE_MAX_ZONES]
    lat_delta = CITY_WIDE_GRID_SPACING_METERS / 111_320
    lon_delta = CITY_WIDE_GRID_SPACING_METERS / max(1, 111_320 * abs(cos(radians(location["lat"]))))
    return [{**location, "lat": location["lat"] + y * lat_delta, "lon": location["lon"] + x * lon_delta} for x, y in offsets]


def extract_lead(place: dict, requested_category: Optional[str], resolved_city: str, center_lat: float, center_lon: float):
    name, lat, lon = place.get("name"), place.get("lat"), place.get("lon")
    if not name or lat is None or lon is None:
        return None
    try:
        lat, lon = float(lat), float(lon)
    except (ValueError, TypeError):
        return None
    # Keep results geographically relevant to the requested city/grid zone.
    if distance_km(center_lat, center_lon, lat, lon) > SEARCH_RADIUS_METERS / 1000:
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


def gemini_analysis(lead: dict):
    fallback = fallback_analysis(lead)
    if not GEMINI_API_KEY:
        return fallback
    prompt = f'''Analyze this business lead for a digital marketing/web-services agency.\nBusiness: {lead.get("name")}\nCategory: {lead.get("category")}\nPhone: {"yes" if lead.get("phone") else "no"}\nWebsite: {"yes" if lead.get("website") else "no"}\nSocial: {"yes" if (lead.get("instagram") or lead.get("facebook") or lead.get("twitter")) else "no"}\nRating: {lead.get("rating", "not available")}\nReviews: {lead.get("reviews", "not available")}\nIMPORTANT: Missing fields only mean the source did not return that field. Do not invent facts. Return ONLY JSON with lead_score, lead_type, recommendation, reason.'''
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}"
    try:
        response = requests.post(url, json={"contents": [{"parts": [{"text": prompt}]}]}, timeout=12)
        if response.status_code != 200:
            return fallback
        text = response.json()["candidates"][0]["content"]["parts"][0]["text"].strip()
        text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text, flags=re.I)
        parsed = json.loads(text)
        allowed = {"High Potential", "Medium Potential", "Low Potential", "Needs More Information"}
        lead_type = parsed.get("lead_type") if parsed.get("lead_type") in allowed else fallback["lead_type"]
        return {"lead_score": max(0, min(100, int(parsed.get("lead_score", fallback["lead_score"])))),
                "lead_type": lead_type,
                "ai_recommendation": str(parsed.get("recommendation") or fallback["ai_recommendation"])[:500],
                "ai_reason": str(parsed.get("reason") or fallback["ai_reason"])[:1000]}
    except Exception:
        return fallback

# ============================================================
# Search
# ============================================================
@app.post("/api/search")
def search(req: SearchRequest, current_user: dict = Depends(get_current_user)):
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
    warnings = []
    raw_with_centers = []

    # PRIMARY: Google Maps via SerpApi. This gives us richer local-business
    # fields such as phone, website, rating, reviews and address.
    search_query = (
        f"{category} in {location['resolved_city']}"
        if category
        else f"{keyword} in {location['resolved_city']}"
    )

    leads = []
    try:
        serp_places = search_serpapi_maps(
            location["lat"], location["lon"], search_query, limit
        )
        for place in serp_places:
            lead = extract_serpapi_lead(
                place,
                category,
                location["resolved_city"],
                location["lat"],
                location["lon"],
            )
            if lead:
                leads.append(lead)
    except ProviderError as exc:
        warnings.append(str(exc))

    # FALLBACK: LatLng remains available if SerpApi is unavailable or returns
    # no geographically valid businesses.
    if not leads and location.get("lat") is not None and location.get("lon") is not None and LATLNG_API_KEY:
        try:
            if category:
                places = search_nearby(
                    location["lat"], location["lon"], category, limit
                )
            else:
                places = search_keyword(
                    location["lat"], location["lon"], keyword, limit
                )
            for place in places:
                lead = extract_lead(
                    place,
                    category,
                    location["resolved_city"],
                    location["lat"],
                    location["lon"],
                )
                if lead:
                    leads.append(lead)
        except ProviderError as exc:
            warnings.append(str(exc))

    leads = dedupe_leads(leads, limit)

    # Never hide a provider failure. But a genuine HTTP-200 empty response is
    # returned cleanly with a useful warning instead of pretending it succeeded.
    if not leads and warnings:
        raise HTTPException(status_code=502, detail=warnings[0])

    if not leads:
        warnings.append(
            f"No matching businesses were returned for {category or keyword!r} in {location['resolved_city']}."
        )

    if req.enrich:
        for lead in leads:
            extra = enrich_public_website(lead.get("website"))
            for key, value in extra.items():
                if key != "sources" and not lead.get(key):
                    lead[key] = value
            lead["sources"] = list(dict.fromkeys([
                *(lead.get("sources") or []),
                *(extra.get("sources") or []),
            ]))
            lead["source"] = ", ".join(lead["sources"])

    saved = []
    for lead in leads:
        lead.setdefault("enrichment_source", None)
        lead.setdefault("enrichment_at", None)

    if DATABASE_URL:
        with db_conn() as conn:
            with conn.cursor() as cur:
                for lead in leads:
                    lead.update(gemini_analysis(lead))
                    cur.execute("""INSERT INTO leads(
                        user_id,place_id,provider_place_id,name,address,city,region,country,latitude,longitude,
                        phone,website,email,instagram,facebook,twitter,category,rating,reviews,maps_url,source,sources,
                        lead_score,lead_type,ai_recommendation,ai_reason,status,enrichment_source,enrichment_at)
                        VALUES(%(user_id)s,%(place_id)s,%(provider_place_id)s,%(name)s,%(address)s,%(city)s,%(region)s,%(country)s,%(latitude)s,%(longitude)s,
                        %(phone)s,%(website)s,%(email)s,%(instagram)s,%(facebook)s,%(twitter)s,%(category)s,%(rating)s,%(reviews)s,%(maps_url)s,%(source)s,%(sources)s::jsonb,
                        %(lead_score)s,%(lead_type)s,%(ai_recommendation)s,%(ai_reason)s,'Not Contacted',%(enrichment_source)s,%(enrichment_at)s)
                        ON CONFLICT(user_id,place_id) DO UPDATE SET
                        provider_place_id=EXCLUDED.provider_place_id,name=EXCLUDED.name,address=EXCLUDED.address,city=EXCLUDED.city,region=EXCLUDED.region,country=EXCLUDED.country,
                        latitude=EXCLUDED.latitude,longitude=EXCLUDED.longitude,phone=EXCLUDED.phone,website=EXCLUDED.website,email=EXCLUDED.email,instagram=EXCLUDED.instagram,
                        facebook=EXCLUDED.facebook,twitter=EXCLUDED.twitter,category=EXCLUDED.category,rating=EXCLUDED.rating,reviews=EXCLUDED.reviews,maps_url=EXCLUDED.maps_url,
                        source=EXCLUDED.source,sources=EXCLUDED.sources,lead_score=EXCLUDED.lead_score,lead_type=EXCLUDED.lead_type,ai_recommendation=EXCLUDED.ai_recommendation,
                        ai_reason=EXCLUDED.ai_reason,enrichment_source=EXCLUDED.enrichment_source,enrichment_at=EXCLUDED.enrichment_at
                        RETURNING id,status,notes,created_at""", {
                            **lead,
                            "user_id": current_user["id"],
                            "sources": json.dumps(lead.get("sources") or []),
                            "enrichment_source": lead.get("enrichment_source"),
                            "enrichment_at": lead.get("enrichment_at"),
                        })
                    row = cur.fetchone()
                    lead.update({"id": row[0], "status": row[1], "notes": row[2], "created_at": str(row[3])})
                    saved.append(lead)

                cur.execute(
                    "INSERT INTO search_history(user_id,city,category,keyword,city_wide,result_count) VALUES(%s,%s,%s,%s,%s,%s)",
                    (current_user["id"], location["resolved_city"], category, keyword, req.city_wide, len(saved)),
                )
            conn.commit()
    else:
        for lead in leads:
            lead.update({"id": None, "status": "Not Contacted", "notes": None})
            saved.append(lead)

    return {
        "city_resolved": location["resolved_city"],
        "category": category,
        "keyword": keyword,
        "city_wide": req.city_wide,
        "count": len(saved),
        "warnings": list(dict.fromkeys(warnings)),
        "leads": saved,
    }

# ============================================================
# Saved leads / filters / updates
# ============================================================
LEAD_SELECT = """SELECT id,name,address,city,region,country,phone,website,email,instagram,facebook,twitter,category,rating,reviews,maps_url,source,lead_score,lead_type,ai_recommendation,ai_reason,status,place_id,created_at,notes,latitude,longitude,provider_place_id,sources,enrichment_source,enrichment_at FROM leads"""

def row_to_lead(row):
    return {
        "id": row[0], "name": row[1], "address": row[2], "city": row[3], "region": row[4], "country": row[5],
        "phone": row[6], "website": row[7], "email": row[8], "instagram": row[9], "facebook": row[10], "twitter": row[11],
        "category": row[12], "rating": row[13], "reviews": row[14], "maps_url": row[15], "source": row[16], "lead_score": row[17],
        "lead_type": row[18], "ai_recommendation": row[19], "ai_reason": row[20], "status": row[21], "place_id": row[22],
        "created_at": str(row[23]), "notes": row[24], "latitude": row[25], "longitude": row[26], "provider_place_id": row[27],
        "sources": row[28] or [], "enrichment_source": row[29], "enrichment_at": str(row[30]) if row[30] else None,
    }

@app.get("/api/leads")
def get_saved_leads(
    has_website: Optional[bool] = None, has_phone: Optional[bool] = None, has_email: Optional[bool] = None,
    has_social: Optional[bool] = None, min_score: Optional[int] = Query(None, ge=0, le=100), lead_type: Optional[str] = None,
    min_rating: Optional[float] = Query(None, ge=0, le=5), category: Optional[str] = None, city: Optional[str] = None,
    source: Optional[str] = None, status: Optional[str] = None, current_user: dict = Depends(get_current_user)):
    clauses, params = ["user_id=%s"], [current_user["id"]]
    bool_filters = [(has_website, "website"), (has_phone, "phone"), (has_email, "email")]
    for flag, col in bool_filters:
        if flag is True: clauses.append(f"NULLIF(TRIM({col}), '') IS NOT NULL")
        elif flag is False: clauses.append(f"NULLIF(TRIM({col}), '') IS NULL")
    if has_social is True: clauses.append("(NULLIF(TRIM(instagram),'') IS NOT NULL OR NULLIF(TRIM(facebook),'') IS NOT NULL OR NULLIF(TRIM(twitter),'') IS NOT NULL)")
    elif has_social is False: clauses.append("(NULLIF(TRIM(instagram),'') IS NULL AND NULLIF(TRIM(facebook),'') IS NULL AND NULLIF(TRIM(twitter),'') IS NULL)")
    if min_score is not None: clauses.append("COALESCE(lead_score,0)>=%s"); params.append(min_score)
    if lead_type: clauses.append("lead_type=%s"); params.append(lead_type)
    if min_rating is not None: clauses.append("COALESCE(rating,0)>=%s"); params.append(min_rating)
    for val, col in [(category,"category"),(city,"city"),(source,"source"),(status,"status")]:
        if val: clauses.append(f"{col} ILIKE %s"); params.append(f"%{val}%")
    with db_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(LEAD_SELECT + " WHERE " + " AND ".join(clauses) + " ORDER BY created_at DESC", params)
            rows = cur.fetchall()
    return {"success": True, "total": len(rows), "leads": [row_to_lead(r) for r in rows]}

@app.patch("/api/leads/{lead_id}/status")
def update_status(lead_id: int, req: StatusUpdateRequest, current_user: dict = Depends(get_current_user)):
    if req.status not in VALID_STATUSES:
        raise HTTPException(status_code=400, detail="Invalid status.")
    with db_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("UPDATE leads SET status=%s WHERE id=%s AND user_id=%s RETURNING id", (req.status, lead_id, current_user["id"]))
            if not cur.fetchone(): raise HTTPException(status_code=404, detail="Lead not found.")
        conn.commit()
    return {"success": True, "lead_id": lead_id, "status": req.status}

@app.patch("/api/leads/{lead_id}/notes")
def update_notes(lead_id: int, req: NotesUpdateRequest, current_user: dict = Depends(get_current_user)):
    with db_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("UPDATE leads SET notes=%s WHERE id=%s AND user_id=%s RETURNING id", ((req.notes or "")[:5000], lead_id, current_user["id"]))
            if not cur.fetchone(): raise HTTPException(status_code=404, detail="Lead not found.")
        conn.commit()
    return {"success": True, "lead_id": lead_id, "notes": req.notes or ""}

@app.delete("/api/leads/{lead_id}")
def delete_lead(lead_id: int, current_user: dict = Depends(get_current_user)):
    with db_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM leads WHERE id=%s AND user_id=%s RETURNING id", (lead_id, current_user["id"]))
            if not cur.fetchone(): raise HTTPException(status_code=404, detail="Lead not found.")
        conn.commit()
    return {"success": True}

# ============================================================
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
    with db_conn() as conn:
        with conn.cursor() as cur:
            assert_group_owner(cur, group_id, current_user["id"])
            cur.execute("SELECT id FROM leads WHERE id=%s AND user_id=%s", (req.lead_ids[0], current_user["id"]))
            if not cur.fetchone(): raise HTTPException(status_code=404, detail="Lead not found.")
            cur.execute("INSERT INTO group_leads(group_id,lead_id) VALUES(%s,%s) ON CONFLICT DO NOTHING", (group_id, req.lead_ids[0]))
        conn.commit()
    return {"success": True}

@app.post("/api/groups/{group_id}/leads/bulk")
def add_leads_to_group_bulk(group_id: int, req: GroupBulkAddRequest, current_user: dict = Depends(get_current_user)):
    if not req.lead_ids: raise HTTPException(status_code=400, detail="No leads selected.")
    added = 0
    with db_conn() as conn:
        with conn.cursor() as cur:
            assert_group_owner(cur, group_id, current_user["id"])
            for lead_id in req.lead_ids[:500]:
                cur.execute("SELECT id FROM leads WHERE id=%s AND user_id=%s", (lead_id, current_user["id"]))
                if cur.fetchone():
                    cur.execute("INSERT INTO group_leads(group_id,lead_id) VALUES(%s,%s) ON CONFLICT DO NOTHING", (group_id, lead_id))
                    added += cur.rowcount
        conn.commit()
    return {"success": True, "added": added}

@app.get("/api/groups/{group_id}/leads")
def get_group_leads(group_id: int, current_user: dict = Depends(get_current_user)):
    with db_conn() as conn:
        with conn.cursor() as cur:
            assert_group_owner(cur, group_id, current_user["id"])
            cur.execute(LEAD_SELECT.replace("FROM leads", "FROM leads l JOIN group_leads gl ON gl.lead_id=l.id") + " WHERE gl.group_id=%s AND l.user_id=%s ORDER BY l.created_at DESC", (group_id, current_user["id"]))
            rows = cur.fetchall()
    return {"leads": [row_to_lead(r) for r in rows]}

# ============================================================
# Search history / export / health
# ============================================================
@app.get("/api/search-history")
def search_history(current_user: dict = Depends(get_current_user)):
    with db_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT id,city,category,keyword,city_wide,result_count,created_at FROM search_history WHERE user_id=%s ORDER BY created_at DESC LIMIT 50", (current_user["id"],))
            rows = cur.fetchall()
    return {"history": [{"id": r[0], "city": r[1], "category": r[2], "keyword": r[3], "city_wide": r[4], "result_count": r[5], "created_at": str(r[6])} for r in rows]}

@app.get("/api/export-excel")
def export_excel(current_user: dict = Depends(get_current_user)):
    with db_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT name,phone,email,address,city,region,country,website,instagram,facebook,twitter,rating,reviews,category,lead_score,lead_type,ai_recommendation,ai_reason,status,source,maps_url,notes FROM leads WHERE user_id=%s ORDER BY created_at DESC", (current_user["id"],))
            rows = cur.fetchall()
    wb = Workbook(); ws = wb.active; ws.title = "Leads"
    ws.append(["Business Name","Phone","Email","Address","City","Region","Country","Website","Instagram","Facebook","Twitter","Rating","Reviews","Category","Lead Score","Lead Type","AI Recommendation","AI Reason","Status","Source","Maps URL","Notes"])
    for row in rows:
        ws.append([v if v not in (None, "") else "Not available from data source" for v in row])
    buffer = BytesIO(); wb.save(buffer); buffer.seek(0)
    return StreamingResponse(buffer, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition":"attachment; filename=leads_export.xlsx"})

@app.get("/api/health")
def health():
    db_connected = False
    if DATABASE_URL:
        try:
            with db_conn() as conn:
                with conn.cursor() as cur:
                    cur.execute("SELECT 1"); cur.fetchone()
            db_connected = True
        except Exception:
            pass
    return {"status": "ok" if db_connected else "degraded", "database_configured": bool(DATABASE_URL), "database_connected": db_connected, "latlng_configured": bool(LATLNG_API_KEY), "serpapi_configured": bool(SERPAPI_API_KEY), "gemini_configured": bool(GEMINI_API_KEY), "jwt_configured": bool(JWT_SECRET)}
