from __future__ import annotations

import json
import os
import sys
from datetime import datetime
from io import BytesIO
from pathlib import Path
from urllib.parse import quote_plus

BASE_DIR = Path(__file__).resolve().parent
VENDOR_DIR = BASE_DIR / "vendor"
if str(VENDOR_DIR) not in sys.path:
    sys.path.insert(0, str(VENDOR_DIR))

from flask import Flask, redirect, render_template, request, send_file, url_for
from openpyxl import Workbook

app = Flask(__name__)

def normalize_data(data):
    normalized = []
    for item in data:
        normalized.append({
            "name": item.get("name", "Unknown"),
            "type": item.get("category", "Unknown"),
            "city": item.get("city", ""),
            "visited": item.get("visited", False),
            "priority": item.get("priority", "Normal"),
            "contact": item.get("contact_person", ""),
            "phone": item.get("phone", ""),
            "notes": item.get("notes", "")
        })
    return normalized


DATA_FILE = BASE_DIR / "data.json"
ROUTE_ORIGIN = "Durg"
ROUTE = ["Rajnandgaon", "Dongargarh", "Bhandara", "Nagpur"]
ROUTE_LEGS = {
    ("Rajnandgaon", "Dongargarh"): 35,
    ("Dongargarh", "Bhandara"): 132,
    ("Bhandara", "Nagpur"): 62,
}
PRIORITY_SCORE = {"High": 3, "Medium": 2, "Low": 1}
TARGET_NEEDS = {"tool design", "die/mould", "cnc machining", "training"}
HIGH_TOOLING_VALUES = {"High"}
INSTITUTE_TYPE_SCORE = {"ITI": 4, "Diploma": 3, "Engineering": 2, "University": 1, "Science": 1}
COURSE_RELEVANCE = {"Mechanical", "Electrical", "Production", "Automobile", "Mechatronics", "Tool and Die", "Manufacturing"}


def load_records() -> list[dict]:
    with DATA_FILE.open("r", encoding="utf-8") as file:
        return json.load(file)


def save_records(records: list[dict]) -> None:
    with DATA_FILE.open("w", encoding="utf-8") as file:
        json.dump(records, file, indent=2)


def find_record(records: list[dict], record_id: int) -> dict | None:
    for record in records:
        if record["id"] == record_id:
            return record
    return None


def build_map_link(record: dict) -> str:
    query = quote_plus(f"{record['name']}, {record['city']}, India")
    return f"https://www.google.com/maps/search/?api=1&query={query}"


def route_index(city: str) -> int:
    try:
        return ROUTE.index(city)
    except ValueError:
        return len(ROUTE)


def get_status(record: dict) -> tuple[str, str]:
    if record["visited"]:
        return "visited", "Visited"
    if record["priority"] == "High":
        return "follow-up", "Priority follow-up"
    if record["notes"].strip():
        return "follow-up", "Follow-up"
    return "not-visited", "Not visited"


def decorate_record(record: dict) -> dict:
    enriched = dict(record)
    status_key, status_label = get_status(record)
    enriched["status_key"] = status_key
    enriched["status_label"] = status_label
    enriched["priority_score"] = PRIORITY_SCORE.get(record["priority"], 1)
    enriched["route_order"] = route_index(record["city"])
    enriched["map_url"] = build_map_link(record)
    enriched["priority_class"] = record["priority"].lower()
    enriched["tooling_class"] = record.get("tooling_requirement", "Low").lower()
    enriched["industry_segment"] = record.get("industry_segment", record.get("industry_type", ""))
    enriched["tooling_need"] = record.get(
        "tooling_need",
        f"{record.get('tooling_requirement', 'Low')} tooling demand",
    )
    enriched["approach_strategy"] = record.get(
        "approach_strategy",
        "Discuss tooling support, training collaboration, and vendor development.",
    )
    enriched["high_tooling_demand"] = record.get("tooling_requirement", "Low") in HIGH_TOOLING_VALUES
    enriched["phone"] = record.get("phone", record.get("contact_number", ""))
    enriched["email"] = record.get("email", "")
    enriched["website"] = record.get("website", "")
    enriched["contact_type"] = record.get("contact_type", "Directory")
    enriched["verified"] = bool(record.get("verified", False))
    enriched["verified_label"] = "Verified" if enriched["verified"] else "Unverified"
    enriched["verification_class"] = "verified" if enriched["verified"] else "unverified"
    enriched["confirm_note"] = (
        "" if enriched["verified"] else "Confirm contact during visit"
    )
    enriched["is_high_value_target"] = (
        record["category"] == "Industry" and record["priority"] == "High"
    )
    enriched["institute_type"] = record.get("institute_type", "")
    enriched["ownership"] = record.get("ownership", "")
    enriched["courses"] = record.get("courses", [])
    enriched["opportunity"] = record.get("opportunity", record.get("business_opportunity", ""))
    enriched["course_relevance_score"] = sum(
        1 for course in enriched["courses"] if course in COURSE_RELEVANCE
    )
    enriched["institute_priority_score"] = INSTITUTE_TYPE_SCORE.get(enriched["institute_type"], 0)
    return enriched


def percent(visited: int, total: int) -> int:
    if total == 0:
        return 0
    return round((visited / total) * 100)


def build_city_progress(records: list[dict]) -> list[dict]:
    progress = []
    for city in ROUTE:
        city_records = [record for record in records if record["city"] == city]
        visited_count = sum(1 for record in city_records if record["visited"])
        pending_count = len(city_records) - visited_count
        progress.append(
            {
                "city": city,
                "total": len(city_records),
                "visited": visited_count,
                "pending": pending_count,
                "completion": percent(visited_count, len(city_records)),
            }
        )
    return progress


def determine_active_city(city_progress: list[dict]) -> str:
    for city in city_progress:
        if city["pending"] > 0:
            return city["city"]
    for city in city_progress:
        if city["total"] > 0:
            return city["city"]
    return ROUTE[-1]


def build_city_sections(records: list[dict], cities: list[str]) -> list[dict]:
    sections = []
    for city in cities:
        city_records = [record for record in records if record["city"] == city]
        institutes = [record for record in city_records if record["category"] == "Institute"]
        institutes.sort(
            key=lambda record: (
                -record["institute_priority_score"],
                -record["course_relevance_score"],
                -record["priority_score"],
                record["name"],
            )
        )

        industries = [record for record in city_records if record["category"] == "Industry"]
        industries.sort(
            key=lambda record: (
                -record["priority_score"],
                -PRIORITY_SCORE.get(record.get("tooling_requirement", "Low"), 1),
                record["name"],
            )
        )

        area_map: dict[str, list[dict]] = {}
        for industry in industries:
            area = industry.get("industrial_area", "General")
            area_map.setdefault(area, []).append(industry)
        industry_clusters = [
            {"industrial_area": area, "industries": area_map[area]}
            for area in sorted(area_map)
        ]

        visited_count = sum(1 for record in city_records if record["visited"])
        sections.append(
            {
                "city": city,
                "records": city_records,
                "institutes": institutes,
                "industries": industries,
                "industry_clusters": industry_clusters,
                "completion": percent(visited_count, len(city_records)),
                "pending": len(city_records) - visited_count,
            }
        )
    return sections


def cluster_priority_score(cluster: list[dict]) -> tuple[int, int]:
    high_priority = sum(1 for record in cluster if record["priority"] == "High")
    total_priority = sum(record["priority_score"] for record in cluster)
    return high_priority, total_priority


def build_visit_order(records: list[dict], min_route_order: int = 0) -> list[dict]:
    candidates = [
        record
        for record in records
        if (
            record["category"] == "Industry"
            and not record["visited"]
            and record["route_order"] >= min_route_order
        )
    ]

    cluster_map: dict[tuple[int, str, str], list[dict]] = {}
    for record in candidates:
        key = (record["route_order"], record["city"], record.get("industrial_area", "General"))
        cluster_map.setdefault(key, []).append(record)

    ordered_clusters = sorted(
        cluster_map.items(),
        key=lambda item: (
            item[0][0],
            -cluster_priority_score(item[1])[0],
            -cluster_priority_score(item[1])[1],
            item[0][2],
        ),
    )

    visit_order = []
    rank = 1
    for (_, city, industrial_area), cluster_records in ordered_clusters:
        cluster_records.sort(
            key=lambda record: (
                -record["priority_score"],
                -PRIORITY_SCORE.get(record.get("tooling_requirement", "Low"), 1),
                record["name"],
            )
        )
        for record in cluster_records:
            visit_order.append(
                {
                    "rank": rank,
                    "record": record,
                    "cluster_label": f"{city} · {industrial_area}",
                }
            )
            rank += 1
    return visit_order[:8]


def build_focus_targets(records: list[dict], current_city: str, next_city: str | None) -> list[dict]:
    visible_cities = {current_city}
    if next_city:
        visible_cities.add(next_city)

    candidates = [
        record
        for record in records
        if not record["visited"] and record["city"] in visible_cities
    ]
    candidates.sort(
        key=lambda record: (
            record["route_order"],
            -record["priority_score"],
            -record.get("institute_priority_score", 0),
            -record.get("course_relevance_score", 0),
            -PRIORITY_SCORE.get(record.get("tooling_requirement", "Low"), 1),
            record["name"],
        )
    )
    return candidates[:5]


def build_suggestions(records: list[dict], min_route_order: int = 0) -> list[dict]:
    suggestions = [
        record
        for record in records
        if (
            not record["visited"]
            and record["route_order"] >= min_route_order
            and record["category"] == "Industry"
        )
    ]
    suggestions.sort(
        key=lambda record: (
            record["route_order"],
            -record["priority_score"],
            -PRIORITY_SCORE.get(record.get("tooling_requirement", "Low"), 1),
            record["category"],
            record["name"],
        )
    )
    top_suggestions = []
    for record in suggestions[:4]:
        matching_needs = [
            need for need in record.get("needs", []) if need.lower() in TARGET_NEEDS
        ]
        if matching_needs:
            reason = (
                f"{record['priority']} priority for "
                + ", ".join(matching_needs[:2])
                + f" in {record['industrial_area']}"
            )
        else:
            reason = (
                f"{record['priority']} priority and {record.get('tooling_requirement', 'Low')} "
                f"tooling requirement in {record['industrial_area']}"
            )
        top_suggestions.append({"record": record, "reason": reason})
    return top_suggestions


def build_institute_suggestions(records: list[dict], min_route_order: int = 0) -> list[dict]:
    suggestions = [
        record
        for record in records
        if (
            record["category"] == "Institute"
            and not record["visited"]
            and record["route_order"] >= min_route_order
        )
    ]
    suggestions.sort(
        key=lambda record: (
            record["route_order"],
            -record["institute_priority_score"],
            -record["course_relevance_score"],
            -record["priority_score"],
            record["name"],
        )
    )
    top = []
    for record in suggestions[:6]:
        reason = (
            f"{record['institute_type']} priority with "
            f"{', '.join(record['courses'][:2])} relevance in {record['city']}"
        )
        top.append({"record": record, "reason": reason})
    return top


def build_view_model(records: list[dict], high_tooling_only: bool = False) -> dict:
    filtered_records = [record for record in records if record.get("state") != "Chhattisgarh"]
    decorated = [decorate_record(record) for record in filtered_records]
    decorated.sort(key=lambda record: (record["route_order"], record["category"], record["name"]))

    city_progress = build_city_progress(decorated)
    active_city = determine_active_city(city_progress)
    active_index = route_index(active_city)
    next_city = ROUTE[active_index + 1] if active_index + 1 < len(ROUTE) else None
    visible_cities = [active_city] + ([next_city] if next_city else [])
    display_records = decorated
    if high_tooling_only:
        display_records = [
            record
            for record in decorated
            if record["category"] != "Industry" or record["high_tooling_demand"]
        ]
    city_sections = build_city_sections(display_records, visible_cities)

    total_records = len(decorated)
    visited_count = sum(1 for record in decorated if record["visited"])
    pending_count = total_records - visited_count
    next_leg_km = ROUTE_LEGS.get((active_city, next_city)) if next_city else None

    current_city_pending = sum(
        1 for record in decorated if record["city"] == active_city and not record["visited"]
    )
    next_city_pending = sum(
        1 for record in decorated if next_city and record["city"] == next_city and not record["visited"]
    )

    next_destination_index = route_index(next_city) if next_city else len(ROUTE)

    return {
        "records": decorated,
        "stats": {
            "total": total_records,
            "visited": visited_count,
            "pending": pending_count,
            "institutes": sum(1 for record in decorated if record["category"] == "Institute"),
            "industries": sum(1 for record in decorated if record["category"] == "Industry"),
            "overall_completion": percent(visited_count, total_records),
        },
        "origin_city": ROUTE_ORIGIN,
        "route": ROUTE,
        "city_progress": city_progress,
        "current_location": active_city,
        "next_destination": next_city,
        "next_leg_km": next_leg_km,
        "current_city_pending": current_city_pending,
        "next_city_pending": next_city_pending,
        "focus_targets": build_focus_targets(decorated, active_city, next_city),
        "city_sections": city_sections,
        "suggestions": build_suggestions(decorated, min_route_order=next_destination_index),
        "institute_suggestions": build_institute_suggestions(
            decorated, min_route_order=route_index(active_city)
        ),
        "visit_order": build_visit_order(decorated, min_route_order=next_destination_index),
        "high_tooling_only": high_tooling_only,
    }


@app.route("/")
def index():
    high_tooling_only = request.args.get("tooling") == "high"
    active_tab = request.args.get("tab", "industries")
    context = build_view_model(load_records(), high_tooling_only=high_tooling_only)
    context["active_tab"] = active_tab
    return render_template("index.html", **context)


@app.post("/toggle-visited/<int:record_id>")
def toggle_visited(record_id: int):
    records = load_records()
    record = find_record(records, record_id)
    if record is not None:
        record["visited"] = not record["visited"]
        record["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        save_records(records)
    return redirect(url_for("index"))


@app.post("/save-notes/<int:record_id>")
def save_notes(record_id: int):
    records = load_records()
    record = find_record(records, record_id)
    if record is not None:
        record["notes"] = request.form.get("notes", "").strip()
        record["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        save_records(records)
    return redirect(url_for("index"))


@app.get("/export")
def export():
    records = [record for record in load_records() if record.get("state") != "Chhattisgarh"]

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Outreach Visits"

    headers = [
        "ID",
        "State",
        "City",
        "Category",
        "Sub Type",
        "Institute Type",
        "Ownership",
        "Courses",
        "Opportunity",
        "Industrial Area",
        "Industry Segment",
        "Industry Type",
        "Tooling Requirement",
        "Tooling Need",
        "Business Opportunity",
        "Approach Strategy",
        "Name",
        "Contact Person",
        "Phone",
        "Email",
        "Website",
        "Contact Type",
        "Verified",
        "Priority",
        "Visited",
        "Notes",
        "Needs",
        "Google Maps Link",
        "Last Updated",
    ]
    worksheet.append(headers)

    for record in records:
        worksheet.append(
            [
                record["id"],
                record["state"],
                record["city"],
                record["category"],
                record["sub_type"],
                record.get("institute_type", ""),
                record.get("ownership", ""),
                ", ".join(record.get("courses", [])),
                record.get("opportunity", record.get("business_opportunity", "")),
                record.get("industrial_area", ""),
                record.get("industry_segment", ""),
                record.get("industry_type", ""),
                record.get("tooling_requirement", ""),
                record.get("tooling_need", ""),
                record.get("business_opportunity", ""),
                record.get("approach_strategy", ""),
                record["name"],
                record["contact_person"],
                record.get("phone", record.get("contact_number", "")),
                record.get("email", ""),
                record.get("website", ""),
                record.get("contact_type", ""),
                "Yes" if record.get("verified", False) else "No",
                record["priority"],
                "Yes" if record["visited"] else "No",
                record["notes"],
                ", ".join(record.get("needs", [])),
                build_map_link(record),
                record["updated_at"],
            ]
        )

    for column in worksheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            value = "" if cell.value is None else str(cell.value)
            max_length = max(max_length, len(value))
        worksheet.column_dimensions[column_letter].width = min(max_length + 4, 48)

    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    filename = f"outreach_route_tracker_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(debug=False, host="127.0.0.1", port=port)
