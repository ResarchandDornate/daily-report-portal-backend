"""
Demo dataset seeder — 4 departments + 11 employees + ~10 days of reports.

Mostly useful for spinning up a fresh dev environment.  Once you've run
import_employees.py (real Excel data), you don't need this any more.

Run from inside the FastAPI container:
    docker compose exec fastapi python -m scripts.seed_data

Idempotent — re-running upserts on (username) for users and (slug) for
departments.  Default password for every demo employee: `password123`.
"""
from __future__ import annotations

import random
import sys
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

# Make `from auth import ...` etc. resolvable when run as a module.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from auth import hash_password  # noqa: E402
from database import SessionLocal  # noqa: E402
from models import DailyReport, Department, User  # noqa: E402
import model_events  # noqa: F401, E402  -- registers the user invariants

random.seed(42)


DEPARTMENTS = [
    {
        "slug": "sales",
        "name": "Sales",
        "color": "rose",
        "report_fields": [
            {"key": "meeting", "label": "Meeting"},
            {"key": "revenue", "label": "Revenue"},
            {"key": "newCustomerOnboard", "label": "New Customer Onboard"},
        ],
    },
    {
        "slug": "marketing",
        "name": "Marketing",
        "color": "indigo",
        "report_fields": [
            {"key": "videoEditing", "label": "Video Making / Editing"},
            {"key": "creatives", "label": "Creatives"},
            {"key": "contentWriting", "label": "Content Writing"},
            {"key": "seo", "label": "SEO"},
            {"key": "websiteManagement", "label": "Website Management"},
            {"key": "reporting", "label": "Reporting"},
        ],
    },
    {
        "slug": "procurement",
        "name": "Procurement",
        "color": "emerald",
        "report_fields": [
            {"key": "enquiries", "label": "Enquiries for Pricing Done"},
            {"key": "negotiations", "label": "Comparison & Negotiations"},
            {"key": "vendorOnboarding", "label": "Vendor Onboarding / Meeting"},
            {"key": "purchaseOrder", "label": "Purchase Order Process"},
            {"key": "payment", "label": "Payment Process (NOPA)"},
            {"key": "dispatches", "label": "Dispatches Done"},
            {"key": "grn", "label": "GRN / Material Received"},
            {"key": "remarks", "label": "Remarks"},
        ],
    },
    {
        "slug": "insideSales",
        "name": "Inside Sales",
        "color": "amber",
        "report_fields": [
            {"key": "calling", "label": "Calling"},
            {"key": "newCustomerOnboard", "label": "New Customer Onboard"},
            {"key": "revenue", "label": "Revenue"},
            {"key": "callingList", "label": "Calling List"},
        ],
    },
]

EMPLOYEES = [
    {"username": "divya",  "email": "divya@ornatesolar.com",  "first_name": "Divya",  "last_name": "Nair",   "title": "Sales Head",         "dept": "sales"},
    {"username": "arjun",  "email": "arjun@ornatesolar.com",  "first_name": "Arjun",  "last_name": "Reddy",  "title": "Account Executive",  "dept": "sales"},
    {"username": "pooja",  "email": "pooja@ornatesolar.com",  "first_name": "Pooja",  "last_name": "Bhatt",  "title": "Sales Associate",    "dept": "sales"},
    {"username": "tarini", "email": "tarini@ornatesolar.com", "first_name": "Tarini", "last_name": "Sethi",  "title": "Inside Sales Lead",  "dept": "insideSales"},
    {"username": "naveen", "email": "naveen@ornatesolar.com", "first_name": "Naveen", "last_name": "Roy",    "title": "Inside Sales Rep",   "dept": "insideSales"},
    {"username": "ishita", "email": "ishita@ornatesolar.com", "first_name": "Ishita", "last_name": "Bansal", "title": "Marketing Lead",     "dept": "marketing"},
    {"username": "kabir",  "email": "kabir@ornatesolar.com",  "first_name": "Kabir",  "last_name": "Malik",  "title": "Content & SEO",      "dept": "marketing"},
    {"username": "riya",   "email": "riya@ornatesolar.com",   "first_name": "Riya",   "last_name": "Saxena", "title": "Creatives & Video",  "dept": "marketing"},
    {"username": "vivek",  "email": "vivek@ornatesolar.com",  "first_name": "Vivek",  "last_name": "Rao",    "title": "Procurement Lead",   "dept": "procurement"},
    {"username": "anjali", "email": "anjali@ornatesolar.com", "first_name": "Anjali", "last_name": "Sinha",  "title": "Vendor Manager",     "dept": "procurement"},
    {"username": "mohit",  "email": "mohit@ornatesolar.com",  "first_name": "Mohit",  "last_name": "Yadav",  "title": "Purchase Officer",   "dept": "procurement"},
]


FIELD_SAMPLES = {
    "meeting":            ["Met 2 EPC partners in Pune.", "Closed kickoff call with rooftop client.", "On-site visit to Indore plant.", "Demo for prospective dealer in Surat."],
    "revenue":            ["₹4.2 L invoiced.", "₹1.8 L PO received.", "₹6 L pipeline added.", "₹2.7 L closed this morning."],
    "newCustomerOnboard": ["1 new dealer onboarded — Surat.", "2 leads moved to onboarding.", "Onboarded Bhopal channel partner.", "—"],
    "calling":            ["38 outbound calls.", "27 calls / 6 connected.", "45 calls + 3 follow-ups.", "52 cold calls today."],
    "callingList":        ["Refreshed Maharashtra leads.", "Pulled new Tamil Nadu list (120 contacts).", "Cleaned bounced numbers.", "Added 30 fresh referrals."],
    "videoEditing":       ["Cut 30s reel for Instagram.", "Edited installation walkthrough.", "Color-graded testimonial video.", "Final cut for monsoon campaign."],
    "creatives":          ["Designed 4 LinkedIn posts.", "New product banner v2.", "Diwali campaign creatives.", "Refreshed homepage hero."],
    "contentWriting":     ["Drafted blog: 'Why poly modules?'", "Wrote landing-page copy.", "Edited case study.", "Newsletter draft for May."],
    "seo":                ["Tuned 6 meta titles.", "Backlink outreach (10 sites).", "Keyword cluster updated.", "Audited top-10 landing pages."],
    "websiteManagement":  ["Updated product specs page.", "Patched WordPress plugins.", "Pushed new contact form.", "Compressed 22 product images."],
    "reporting":          ["Sent weekly traffic snapshot.", "MoM lead-source breakdown.", "Updated GA4 conversions.", "Pulled paid-ads ROAS report."],
    "enquiries":          ["Sent RFQs for 12 modules.", "Got 3 fresh quotes for cables.", "Asked 5 vendors on inverters.", "Floated tender for MMS structures."],
    "negotiations":       ["Saved 7% on inverter PO.", "Renegotiated freight rates.", "Locked Tier-1 module price.", "Closed 4% discount on cables."],
    "vendorOnboarding":   ["Onboarded 1 cable vendor.", "Visit to Noida transformer plant.", "KYC done for 2 new vendors.", "—"],
    "purchaseOrder":      ["Raised 3 POs.", "PO #1248 released.", "Bulk PO for 200 panels.", "Released PO for inverters."],
    "payment":            ["NOPA cleared for Vendor X.", "Initiated 60% advance.", "Cleared 4 vendor payments.", "Released milestone payment."],
    "dispatches":         ["3 trucks dispatched to Jaipur.", "Container left Mundra port.", "Site dispatch to Coimbatore.", "Dispatch to Bhopal in transit."],
    "grn":                ["GRN #4421 booked.", "Inverters received & inspected.", "Modules QC passed.", "Material received with minor damage — flagged."],
    "remarks":            ["—", "Watch monsoon delays.", "Vendor X payment terms revised.", "Need to revise BoM for upcoming order."],
}


def _pick(key: str, idx: int) -> str:
    samples = FIELD_SAMPLES.get(key, [""])
    return samples[idx % len(samples)]


def main() -> None:
    print("=" * 60)
    print("Seeding demo dataset (4 departments, 11 employees, ~10 days of reports)")
    print("=" * 60)
    db = SessionLocal()
    try:
        # ---- departments ----
        slug_to_dept: dict[str, Department] = {}
        for d in DEPARTMENTS:
            existing = db.query(Department).filter(Department.slug == d["slug"]).first()
            if existing:
                existing.name = d["name"]
                existing.color = d["color"]
                existing.report_fields = d["report_fields"]
                slug_to_dept[d["slug"]] = existing
            else:
                obj = Department(
                    slug=d["slug"],
                    name=d["name"],
                    color=d["color"],
                    report_fields=d["report_fields"],
                )
                db.add(obj)
                db.flush()
                slug_to_dept[d["slug"]] = obj

        # ---- employees ----
        created_users = 0
        now = datetime.now(timezone.utc)
        for emp in EMPLOYEES:
            user = db.query(User).filter(User.username == emp["username"]).first()
            if user is None:
                user = User(
                    username=emp["username"],
                    email=emp["email"],
                    first_name=emp["first_name"],
                    last_name=emp["last_name"],
                    title=emp["title"],
                    role="employee",
                    department=slug_to_dept[emp["dept"]],
                    is_active=True,
                    is_staff=False,
                    is_superuser=False,
                    date_joined=now,
                    password=hash_password("password123"),
                )
                db.add(user)
                created_users += 1
                print(f"  created {emp['username']}")
            else:
                user.first_name = emp["first_name"]
                user.last_name = emp["last_name"]
                user.title = emp["title"]
                user.department = slug_to_dept[emp["dept"]]

        db.flush()  # ensure user.id is populated for the report-loop below

        # ---- reports — last 10 days, with some skipped ----
        today = date.today()
        for emp in EMPLOYEES:
            user = db.query(User).filter(User.username == emp["username"]).first()
            dept = slug_to_dept[emp["dept"]]
            fields = dept.report_fields or []
            for day_offset in range(1, 11):
                # Skip a day deterministically per employee for variety
                if (hash(emp["username"]) + day_offset) % 6 == 0:
                    continue
                d = today - timedelta(days=day_offset)
                data = {f["key"]: _pick(f["key"], hash(emp["username"]) + day_offset) for f in fields}
                existing = (
                    db.query(DailyReport)
                    .filter(DailyReport.user_id == user.id, DailyReport.date == d)
                    .first()
                )
                if existing:
                    existing.data = data
                else:
                    db.add(DailyReport(user_id=user.id, date=d, data=data))

        db.commit()

        print()
        print(f"  Users created this run: {created_users}")
        print(f"  Total active employees: {db.query(User).filter(User.role == 'employee', User.is_active.is_(True)).count()}")
        print(f"  Total reports:          {db.query(DailyReport).count()}")
        print()
        print("Login as any demo employee with password: password123")
        print("  e.g. divya@ornatesolar.com / password123")
    finally:
        db.close()


if __name__ == "__main__":
    main()
