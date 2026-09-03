import os
from pathlib import Path
from urllib.parse import quote_plus

import requests
from fastapi import FastAPI, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from dotenv import load_dotenv

load_dotenv()

app = FastAPI(title="Lead Search Engine")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE = Path(__file__).resolve().parent

GEOAPIFY_API_KEY = os.getenv("GEOAPIFY_API_KEY")

CATEGORY_MAP = {
    "restaurant": "catering.restaurant",
    "restaurants": "catering.restaurant",
    "cafe": "catering.cafe",
    "cafes": "catering.cafe",
    "fast food": "catering.fast_food",
    "bakery": "commercial.food_and_drink.bakery",
    "bakeries": "commercial.food_and_drink.bakery",

    "salon": "service.beauty.hairdresser",
    "salons": "service.beauty.hairdresser",
    "beauty": "service.beauty",

    "gym": "sport.fitness.gym",
    "gyms": "sport.fitness.gym",
    "fitness": "sport.fitness",

    "pharmacy": "healthcare.pharmacy",
    "pharmacies": "healthcare.pharmacy",
    "hospital": "healthcare.hospital",
    "hospitals": "healthcare.hospital",
    "clinic": "healthcare.clinic_or_praxis",
    "clinics": "healthcare.clinic_or_praxis",
    "dentist": "healthcare.dentist",
    "dentists": "healthcare.dentist",

    "real estate": "service.estate_agent",
    "real estate agency": "service.estate_agent",

    "travel agency": "service.travel_agency",
    "travel agencies": "service.travel_agency",

    "photographer": "service.photographer",
    "photography": "service.photographer",

    "car service": "service.vehicle.repair",
    "car repair": "service.vehicle.repair.car",

    "furniture": "commercial.furniture",
    "furniture shops": "commercial.furniture",

    "supermarket": "commercial.supermarket",
    "supermarkets": "commercial.supermarket",

    "clothing": "commercial.clothing",
    "clothing store": "commercial.clothing",

    "florist": "commercial.florist",
    "flower shop": "commercial.florist",
}


@app.get("/")
def home():
   return FileResponse(BASE.parent / "frontend" / "index.html")

@app.get("/api/health")
def health():
    return {
        "status": "ok",
        "geoapify_configured": bool(GEOAPIFY_API_KEY)
    }


def get_city_coordinates(city: str):
    url = "https://api.geoapify.com/v1/geocode/search"

    params = {
        "text": city,
        "type": "city",
        "filter": "countrycode:in",
        "limit": 1,
        "format": "json",
        "apiKey": GEOAPIFY_API_KEY,
    }

    response = requests.get(url, params=params, timeout=20)
    response.raise_for_status()

    data = response.json()

    results = data.get("results", [])

    if not results:
        return None

    result = results[0]

    return {
        "lat": result.get("lat"),
        "lon": result.get("lon"),
        "name": result.get("formatted", city),
    }


@app.get("/api/search")
def search_leads(
    city: str = Query(...),
    category: str = Query("restaurants"),
    limit: int = Query(20, ge=1, le=100),
):
    if not GEOAPIFY_API_KEY:
        return {
            "success": False,
            "error": "GEOAPIFY_API_KEY is not configured."
        }

    try:
        location = get_city_coordinates(city)

        if not location:
            return {
                "success": False,
                "error": f"Could not find city: {city}"
            }

        category_key = category.strip().lower()

        geo_category = CATEGORY_MAP.get(
            category_key,
            category_key
        )

        lat = location["lat"]
        lon = location["lon"]

        url = "https://api.geoapify.com/v2/places"

        params = {
            "categories": geo_category,
            "filter": f"circle:{lon},{lat},10000",
            "bias": f"proximity:{lon},{lat}",
            "limit": limit,
            "apiKey": GEOAPIFY_API_KEY,
        }

        response = requests.get(url, params=params, timeout=30)

        if response.status_code != 200:
            return {
                "success": False,
                "error": f"Geoapify API error: {response.status_code}",
                "details": response.text[:500],
            }

        data = response.json()

        leads = []

        for feature in data.get("features", []):
            props = feature.get("properties", {})

            name = props.get("name")

            if not name:
                continue

            contact = props.get("contact") or {}

            phone = (
                contact.get("phone")
                or props.get("phone")
                or ""
            )

            website = (
                props.get("website")
                or contact.get("website")
                or ""
            )

            address = (
                props.get("formatted")
                or props.get("address_line1")
                or ""
            )

            categories = props.get("categories") or []

            maps_url = (
                "https://www.google.com/maps/search/?api=1&query="
                + quote_plus(f"{name}, {address}")
            )

            lead_type = (
                "Business Found"
                if website
                else "Potential Lead"
            )

            leads.append({
                "name": name,
                "phone": phone,
                "address": address,
                "website": website,
                "rating": props.get("rating", ""),
                "reviews": props.get("reviews", ""),
                "category": (
                    categories[0]
                    if categories
                    else category
                ),
                "lead_type": lead_type,
                "status": "Not Contacted",
                "maps": maps_url,
                "place_id": props.get("place_id", ""),
            })

        return {
            "success": True,
            "city": location["name"],
            "category": category,
            "total": len(leads),
            "leads": leads,
        }

    except requests.RequestException as e:
        return {
            "success": False,
            "error": f"Network error: {str(e)}"
        }

    except Exception as e:
        return {
            "success": False,
            "error": str(e)
        }
    