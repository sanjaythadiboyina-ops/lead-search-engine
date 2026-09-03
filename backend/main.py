import os
import hashlib
import secrets
from io import BytesIO
from pathlib import Path
from urllib.parse import quote

import psycopg
import requests
from dotenv import load_dotenv
from fastapi import Depends, FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer
from jose import JWTError, jwt
from openpyxl import Workbook
from pydantic import BaseModel


# =========================================================
# CONFIG
# =========================================================

load_dotenv()

GEOAPIFY_API_KEY = os.getenv("GEOAPIFY_API_KEY")
DATABASE_URL = os.getenv("DATABASE_URL")

JWT_SECRET = os.getenv(
    "JWT_SECRET",
    "change-this-secret"
)

JWT_ALGORITHM = "HS256"

BASE = Path(__file__).resolve().parent.parent


# =========================================================
# APP
# =========================================================

app = FastAPI(
    title="Lead Search Engine"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

security = HTTPBearer()


# =========================================================
# CATEGORY MAP
# =========================================================

CATEGORY_MAP = {

    "restaurant": "catering.restaurant",
    "restaurants": "catering.restaurant",

    "cafe": "catering.cafe",
    "cafes": "catering.cafe",

    "bakery": "catering.bakery",
    "bakeries": "catering.bakery",

    "bar": "catering.bar",
    "bars": "catering.bar",

    "gym": "sport.fitness",
    "gyms": "sport.fitness",

    "salon": "service.beauty",
    "salons": "service.beauty",
    "beauty salons": "service.beauty",

    "dentist": "healthcare.dentist",
    "dentists": "healthcare.dentist",
    "dental clinics": "healthcare.dentist",

    "clinic": "healthcare.clinic",
    "clinics": "healthcare.clinic",

    "pharmacy": "healthcare.pharmacy",
    "pharmacies": "healthcare.pharmacy",

    "hotel": "accommodation.hotel",
    "hotels": "accommodation.hotel",

    "real estate": "commercial.real_estate",
    "real estate agencies": "commercial.real_estate",

    "supermarket": "commercial.supermarket",
    "supermarkets": "commercial.supermarket",

    "jewellery": "commercial.jewelry",
    "jewellery shops": "commercial.jewelry",

    "furniture": "commercial.furniture",
    "furniture shops": "commercial.furniture",

    "car repair": "service.vehicle.repair",
    "car service centers": "service.vehicle.repair",

    "travel agency": "tourism.travel_agency",
    "travel agencies": "tourism.travel_agency",

    "photography": "service.photography",
    "photography studios": "service.photography",

    "event management": "service.event_management",
    "event management companies": "service.event_management",

    "interior design": "service.home_improvement",
    "interior designers": "service.home_improvement",

    "boutique": "commercial.clothing",
    "boutiques": "commercial.clothing",
}


# =========================================================
# REQUEST MODELS
# =========================================================

class SignupRequest(BaseModel):

    email: str
    password: str


class LoginRequest(BaseModel):

    email: str
    password: str


class StatusUpdate(BaseModel):

    status: str


# =========================================================
# PASSWORD HASHING
# =========================================================

def hash_password(password: str) -> str:

    salt = secrets.token_bytes(16)

    password_hash = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt,
        310000
    )

    return (
        salt.hex()
        + ":"
        + password_hash.hex()
    )


def verify_password(
    password: str,
    stored_hash: str
) -> bool:

    try:

        salt_hex, hash_hex = stored_hash.split(":", 1)

        salt = bytes.fromhex(salt_hex)

        new_hash = hashlib.pbkdf2_hmac(
            "sha256",
            password.encode("utf-8"),
            salt,
            310000
        )

        return secrets.compare_digest(
            new_hash.hex(),
            hash_hex
        )

    except (ValueError, TypeError):

        return False


# =========================================================
# JWT
# =========================================================

def create_token(
    user_id: int,
    email: str
):

    payload = {
        "user_id": user_id,
        "email": email
    }

    return jwt.encode(
        payload,
        JWT_SECRET,
        algorithm=JWT_ALGORITHM
    )


def get_current_user(
    credentials: HTTPAuthorizationCredentials = Depends(
        security
    )
):

    try:

        payload = jwt.decode(
            credentials.credentials,
            JWT_SECRET,
            algorithms=[JWT_ALGORITHM]
        )

        user_id = payload.get("user_id")
        email = payload.get("email")

        if not user_id or not email:

            raise HTTPException(
                status_code=401,
                detail="Invalid token"
            )

        return {
            "user_id": int(user_id),
            "email": str(email)
        }

    except (
        JWTError,
        ValueError,
        TypeError
    ):

        raise HTTPException(
            status_code=401,
            detail="Invalid or expired token"
        )


# =========================================================
# DATABASE
# =========================================================

def get_connection():

    if not DATABASE_URL:

        raise HTTPException(
            status_code=500,
            detail="DATABASE_URL is not configured"
        )

    return psycopg.connect(
        DATABASE_URL,
        connect_timeout=10
    )


def init_database():

    if not DATABASE_URL:

        print(
            "DATABASE_URL not configured"
        )

        return

    try:

        with get_connection() as conn:

            with conn.cursor() as cur:

                # USERS
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS users (
                        id SERIAL PRIMARY KEY,
                        email TEXT UNIQUE NOT NULL,
                        password_hash TEXT NOT NULL,
                        created_at TIMESTAMP
                            DEFAULT CURRENT_TIMESTAMP
                    )
                    """
                )

                # LEADS
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS leads (
                        id SERIAL PRIMARY KEY,
                        name TEXT NOT NULL,
                        phone TEXT,
                        address TEXT,
                        website TEXT,
                        rating TEXT,
                        reviews TEXT,
                        category TEXT,
                        lead_type TEXT,
                        status TEXT
                            DEFAULT 'Not Contacted',
                        maps TEXT,
                        place_id TEXT UNIQUE,
                        user_id INTEGER,
                        created_at TIMESTAMP
                            DEFAULT CURRENT_TIMESTAMP
                    )
                    """
                )

                # OLD DATABASE SUPPORT
                cur.execute(
                    """
                    ALTER TABLE leads
                    ADD COLUMN IF NOT EXISTS user_id INTEGER
                    """
                )

                # SEARCH HISTORY
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS search_history (
                        id SERIAL PRIMARY KEY,
                        user_id INTEGER,
                        city TEXT,
                        category TEXT,
                        result_count INTEGER,
                        created_at TIMESTAMP
                            DEFAULT CURRENT_TIMESTAMP
                    )
                    """
                )

        print(
            "Database initialized successfully"
        )

    except Exception as e:

        print(
            "Database initialization error:",
            e
        )


# =========================================================
# STARTUP
# =========================================================

@app.on_event("startup")
def startup():

    init_database()


# =========================================================
# HELPERS
# =========================================================

def normalize_url(value):

    if not value:
        return ""

    value = str(value).strip()

    if not value:
        return ""

    bad_prefixes = [
        "http://127.0.0.1:8000/",
        "http://localhost:8000/",
        "https://127.0.0.1:8000/",
        "https://localhost:8000/"
    ]

    for prefix in bad_prefixes:

        if value.startswith(prefix):

            value = value[len(prefix):]

            break

    if not value.startswith(
        ("http://", "https://")
    ):

        value = "https://" + value

    return value


# =========================================================
# SAVE LEAD
# =========================================================

def save_lead(
    lead: dict,
    user_id: int
):

    with get_connection() as conn:

        with conn.cursor() as cur:

            cur.execute(
                """
                INSERT INTO leads (
                    name,
                    phone,
                    address,
                    website,
                    rating,
                    reviews,
                    category,
                    lead_type,
                    status,
                    maps,
                    place_id,
                    user_id
                )
                VALUES (
                    %s, %s, %s, %s,
                    %s, %s, %s, %s,
                    %s, %s, %s, %s
                )
                ON CONFLICT(place_id)
                DO UPDATE SET
                    name = EXCLUDED.name,
                    phone = EXCLUDED.phone,
                    address = EXCLUDED.address,
                    website = EXCLUDED.website,
                    rating = EXCLUDED.rating,
                    reviews = EXCLUDED.reviews,
                    category = EXCLUDED.category,
                    lead_type = EXCLUDED.lead_type,
                    maps = EXCLUDED.maps
                RETURNING id
                """,
                (
                    lead.get("name"),
                    lead.get("phone"),
                    lead.get("address"),
                    lead.get("website"),
                    lead.get("rating"),
                    lead.get("reviews"),
                    lead.get("category"),
                    lead.get("lead_type"),
                    lead.get(
                        "status",
                        "Not Contacted"
                    ),
                    lead.get("maps"),
                    lead.get("place_id"),
                    user_id
                )
            )

            row = cur.fetchone()

            return (
                row[0]
                if row
                else None
            )


# =========================================================
# HOME
# =========================================================

@app.get("/")
def home():

    frontend = (
        BASE
        / "frontend"
        / "index.html"
    )

    if frontend.exists():

        return FileResponse(
            frontend
        )

    return {
        "message":
            "Lead Search Engine API is running"
    }


# =========================================================
# HEALTH
# =========================================================

@app.get("/api/health")
def health():

    database_connected = False

    if DATABASE_URL:

        try:

            with psycopg.connect(
                DATABASE_URL,
                connect_timeout=5
            ):

                database_connected = True

        except Exception:

            database_connected = False

    return {

        "status": "ok",

        "geoapify_configured":
            bool(GEOAPIFY_API_KEY),

        "database_configured":
            bool(DATABASE_URL),

        "database_connected":
            database_connected
    }


# =========================================================
# SIGNUP
# =========================================================

@app.post("/api/signup")
def signup(
    data: SignupRequest
):

    email = (
        str(data.email)
        .strip()
        .lower()
    )

    password = data.password

    if len(password) < 6:

        raise HTTPException(
            status_code=400,
            detail=
                "Password must be at least 6 characters"
        )

    try:

        with get_connection() as conn:

            with conn.cursor() as cur:

                cur.execute(
                    """
                    SELECT id
                    FROM users
                    WHERE email = %s
                    """,
                    (email,)
                )

                if cur.fetchone():

                    raise HTTPException(
                        status_code=400,
                        detail=
                            "Email already registered"
                    )

                cur.execute(
                    """
                    INSERT INTO users (
                        email,
                        password_hash
                    )
                    VALUES (%s, %s)
                    RETURNING id
                    """,
                    (
                        email,
                        hash_password(password)
                    )
                )

                user_id = cur.fetchone()[0]

        token = create_token(
            user_id,
            email
        )

        return {

            "message":
                "Signup successful",

            "token":
                token,

            "user": {

                "id":
                    user_id,

                "email":
                    email
            }
        }

    except HTTPException:

        raise

    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )


# =========================================================
# LOGIN
# =========================================================

@app.post("/api/login")
def login(
    data: LoginRequest
):

    email = (
        str(data.email)
        .strip()
        .lower()
    )

    try:

        with get_connection() as conn:

            with conn.cursor() as cur:

                cur.execute(
                    """
                    SELECT
                        id,
                        email,
                        password_hash
                    FROM users
                    WHERE email = %s
                    """,
                    (email,)
                )

                user = cur.fetchone()

        if not user:

            raise HTTPException(
                status_code=401,
                detail=
                    "Invalid email or password"
            )

        if not verify_password(
            data.password,
            user[2]
        ):

            raise HTTPException(
                status_code=401,
                detail=
                    "Invalid email or password"
            )

        token = create_token(
            user[0],
            user[1]
        )

        return {

            "message":
                "Login successful",

            "token":
                token,

            "user": {

                "id":
                    user[0],

                "email":
                    user[1]
            }
        }

    except HTTPException:

        raise

    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )


# =========================================================
# CURRENT USER
# =========================================================

@app.get("/api/me")
def get_me(
    current_user: dict = Depends(
        get_current_user
    )
):

    return {
        "user":
            current_user
    }


# =========================================================
# SEARCH LEADS
# =========================================================

@app.get("/api/search")
def search_leads(

    city: str,

    category: str,

    limit: int = 20,

    current_user: dict = Depends(
        get_current_user
    )
):

    if not GEOAPIFY_API_KEY:

        raise HTTPException(
            status_code=500,
            detail=
                "GEOAPIFY_API_KEY is not configured"
        )

    city = city.strip()
    category = category.strip()

    if not city:

        raise HTTPException(
            status_code=400,
            detail="City is required"
        )

    if not category:

        raise HTTPException(
            status_code=400,
            detail="Category is required"
        )

    limit = max(
        1,
        min(int(limit), 50)
    )

    category_key = category.lower()

    geo_category = CATEGORY_MAP.get(
        category_key,
        category_key
    )

    # -----------------------------------------------------
    # GEOCODING
    # -----------------------------------------------------

    try:

        response = requests.get(
            "https://api.geoapify.com/v1/geocode/search",
            params={
                "text": city,
                "apiKey":
                    GEOAPIFY_API_KEY,
                "limit": 1
            },
            timeout=20
        )

        response.raise_for_status()

        geo_data = response.json()

    except requests.RequestException as e:

        raise HTTPException(
            status_code=502,
            detail=
                f"Geocoding failed: {e}"
        )

    features = geo_data.get(
        "features",
        []
    )

    if not features:

        raise HTTPException(
            status_code=404,
            detail="City not found"
        )

    coordinates = (
        features[0]
        .get("geometry", {})
        .get("coordinates", [])
    )

    if len(coordinates) < 2:

        raise HTTPException(
            status_code=404,
            detail=
                "Location coordinates not found"
        )

    lon = coordinates[0]
    lat = coordinates[1]

    # -----------------------------------------------------
    # PLACES SEARCH
    # -----------------------------------------------------

    try:

        response = requests.get(
            "https://api.geoapify.com/v2/places",
            params={

                "categories":
                    geo_category,

                "filter":
                    f"circle:{lon},{lat},30000",

                "bias":
                    f"proximity:{lon},{lat}",

                "limit":
                    limit,

                "apiKey":
                    GEOAPIFY_API_KEY
            },
            timeout=30
        )

        response.raise_for_status()

        places_data = response.json()

    except requests.RequestException as e:

        raise HTTPException(
            status_code=502,
            detail=
                f"Places search failed: {e}"
        )

    results = []

    for feature in places_data.get(
        "features",
        []
    ):

        properties = feature.get(
            "properties",
            {}
        )

        name = properties.get(
            "name"
        )

        if not name:
            continue

        address = properties.get(
            "formatted",
            ""
        )

        datasource = (
            properties.get(
                "datasource"
            )
            or {}
        )

        raw = (
            datasource.get("raw")
            or {}
        )

        phone = (
            raw.get("phone")
            or properties.get("phone")
            or ""
        )

        website = (
            raw.get("website")
            or properties.get("website")
            or ""
        )

        website = normalize_url(
            website
        )

        rating = properties.get(
            "rating",
            ""
        )

        reviews = properties.get(
            "reviews",
            ""
        )

        maps = properties.get(
            "google_maps",
            ""
        )

        if not maps:

            maps = (
                "https://www.google.com/maps/search/"
                + quote(
                    f"{name} {address}"
                )
            )

        place_id = (
            properties.get(
                "place_id"
            )
            or feature.get("id")
            or f"{name}-{address}"
        )

        lead_type = (
            "Potential Lead"
            if not website
            else "Business Found"
        )

        lead = {

            "name":
                name,

            "phone":
                phone,

            "address":
                address,

            "website":
                website,

            "rating":
                rating,

            "reviews":
                reviews,

            "category":
                category,

            "lead_type":
                lead_type,

            "status":
                "Not Contacted",

            "maps":
                maps,

            "place_id":
                str(place_id)
        }

        try:

            saved_id = save_lead(
                lead,
                current_user[
                    "user_id"
                ]
            )

            # IMPORTANT:
            # Return the real database ID
            lead["id"] = saved_id

        except Exception as e:

            print(
                "Lead save error:",
                e
            )

            continue

        results.append(
            lead
        )

    # -----------------------------------------------------
    # SEARCH HISTORY
    # -----------------------------------------------------

    try:

        with get_connection() as conn:

            with conn.cursor() as cur:

                cur.execute(
                    """
                    INSERT INTO search_history (
                        user_id,
                        city,
                        category,
                        result_count
                    )
                    VALUES (
                        %s, %s, %s, %s
                    )
                    """,
                    (
                        current_user[
                            "user_id"
                        ],
                        city,
                        category,
                        len(results)
                    )
                )

    except Exception as e:

        print(
            "Search history error:",
            e
        )

    return {

        "city":
            city,

        "category":
            category,

        "total":
            len(results),

        "leads":
            results
    }


# =========================================================
# GET SAVED LEADS
# =========================================================

@app.get("/api/leads")
def get_leads(
    current_user: dict = Depends(
        get_current_user
    )
):

    try:

        with get_connection() as conn:

            with conn.cursor() as cur:

                cur.execute(
                    """
                    SELECT
                        id,
                        name,
                        phone,
                        address,
                        website,
                        rating,
                        reviews,
                        category,
                        lead_type,
                        status,
                        maps,
                        place_id,
                        created_at
                    FROM leads
                    WHERE user_id = %s
                    ORDER BY created_at DESC
                    """,
                    (
                        current_user[
                            "user_id"
                        ],
                    )
                )

                rows = cur.fetchall()

        leads = []

        for row in rows:

            leads.append({

                "id":
                    row[0],

                "name":
                    row[1],

                "phone":
                    row[2],

                "address":
                    row[3],

                "website":
                    normalize_url(row[4]),

                "rating":
                    row[5],

                "reviews":
                    row[6],

                "category":
                    row[7],

                "lead_type":
                    row[8],

                "status":
                    row[9]
                    or "Not Contacted",

                "maps":
                    row[10],

                "place_id":
                    row[11],

                "created_at":
                    str(row[12])
            })

        return {

            "total":
                len(leads),

            "leads":
                leads
        }

    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )


# =========================================================
# UPDATE LEAD STATUS
# =========================================================

@app.patch(
    "/api/leads/{lead_id}/status"
)
def update_status(

    lead_id: int,

    data: StatusUpdate,

    current_user: dict = Depends(
        get_current_user
    )
):

    allowed_statuses = {

        "Not Contacted",

        "Contacted",

        "Interested",

        "Follow-up",

        "Converted"
    }

    if data.status not in allowed_statuses:

        raise HTTPException(
            status_code=400,
            detail="Invalid status"
        )

    try:

        with get_connection() as conn:

            with conn.cursor() as cur:

                cur.execute(
                    """
                    UPDATE leads
                    SET status = %s
                    WHERE id = %s
                    AND user_id = %s
                    RETURNING id
                    """,
                    (
                        data.status,
                        lead_id,
                        current_user[
                            "user_id"
                        ]
                    )
                )

                updated = cur.fetchone()

        if not updated:

            raise HTTPException(
                status_code=404,
                detail="Lead not found"
            )

        return {

            "message":
                "Lead status updated",

            "lead_id":
                lead_id,

            "status":
                data.status
        }

    except HTTPException:

        raise

    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )


# =========================================================
# EXPORT EXCEL
# =========================================================

@app.get("/api/export-excel")
def export_excel(
    current_user: dict = Depends(
        get_current_user
    )
):

    try:

        with get_connection() as conn:

            with conn.cursor() as cur:

                cur.execute(
                    """
                    SELECT
                        id,
                        name,
                        phone,
                        address,
                        website,
                        rating,
                        reviews,
                        category,
                        lead_type,
                        status,
                        maps,
                        created_at
                    FROM leads
                    WHERE user_id = %s
                    ORDER BY created_at DESC
                    """,
                    (
                        current_user[
                            "user_id"
                        ],
                    )
                )

                rows = cur.fetchall()

        workbook = Workbook()

        worksheet = workbook.active

        worksheet.title = "Leads"

        headers = [

            "ID",

            "Business",

            "Phone",

            "Address",

            "Website",

            "Rating",

            "Reviews",

            "Category",

            "Lead Type",

            "Status",

            "Google Maps",

            "Created At"
        ]

        worksheet.append(
            headers
        )

        for row in rows:

            worksheet.append([

                row[0],

                row[1] or "",

                row[2] or "",

                row[3] or "",

                normalize_url(
                    row[4]
                ),

                row[5] or "",

                row[6] or "",

                row[7] or "",

                row[8] or "",

                row[9]
                or "Not Contacted",

                row[10] or "",

                str(row[11])
                if row[11]
                else ""
            ])

        worksheet.freeze_panes = "A2"

        worksheet.auto_filter.ref = (
            worksheet.dimensions
        )

        widths = {

            "A": 10,

            "B": 30,

            "C": 18,

            "D": 45,

            "E": 40,

            "F": 12,

            "G": 12,

            "H": 24,

            "I": 20,

            "J": 18,

            "K": 45,

            "L": 24
        }

        for column, width in widths.items():

            worksheet.column_dimensions[
                column
            ].width = width

        output = BytesIO()

        workbook.save(
            output
        )

        output.seek(0)

        return StreamingResponse(

            output,

            media_type=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),

            headers={

                "Content-Disposition":
                    (
                        'attachment; '
                        'filename="lead_search_results.xlsx"'
                    )
            }
        )

    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=
                f"Excel export failed: {e}"
        )


# =========================================================
# SEARCH HISTORY
# =========================================================

@app.get("/api/search-history")
def get_search_history(

    current_user: dict = Depends(
        get_current_user
    )
):

    try:

        with get_connection() as conn:

            with conn.cursor() as cur:

                cur.execute(
                    """
                    SELECT
                        id,
                        city,
                        category,
                        result_count,
                        created_at
                    FROM search_history
                    WHERE user_id = %s
                    ORDER BY created_at DESC
                    LIMIT 50
                    """,
                    (
                        current_user[
                            "user_id"
                        ],
                    )
                )

                rows = cur.fetchall()

        history = []

        for row in rows:

            history.append({

                "id":
                    row[0],

                "city":
                    row[1],

                "category":
                    row[2],

                "result_count":
                    row[3],

                "created_at":
                    str(row[4])
            })

        return {

            "history":
                history
        }

    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )