"""
One-shot seed for the Daily Report Portal.

Creates 11 employees across the 4 departments, plus a mix of past daily
reports (with some missing days, so the "Missing today" panel has data).

Run from inside the Docker container:
    docker compose exec django python manage.py shell -c "exec(open('/app/seed_data.py').read())"

Or via the helper invocation in this repo. Idempotent — re-running won't
duplicate users (uses get_or_create) and report rows are upserted by
(user, date).

Default password for every seeded employee: `password123`
"""
import random
from datetime import date, timedelta

from django.contrib.auth import get_user_model

from departments.models import Department
from reports.models import DailyReport

User = get_user_model()
random.seed(42)  # deterministic so re-running yields the same numbers

EMPLOYEES = [
    # Sales
    {"username": "divya", "email": "divya@ornatesolar.com", "first_name": "Divya", "last_name": "Nair", "title": "Sales Head", "dept": "sales"},
    {"username": "arjun", "email": "arjun@ornatesolar.com", "first_name": "Arjun", "last_name": "Reddy", "title": "Account Executive", "dept": "sales"},
    {"username": "pooja", "email": "pooja@ornatesolar.com", "first_name": "Pooja", "last_name": "Bhatt", "title": "Sales Associate", "dept": "sales"},

    # Inside Sales
    {"username": "tarini", "email": "tarini@ornatesolar.com", "first_name": "Tarini", "last_name": "Sethi", "title": "Inside Sales Lead", "dept": "insideSales"},
    {"username": "naveen", "email": "naveen@ornatesolar.com", "first_name": "Naveen", "last_name": "Roy", "title": "Inside Sales Rep", "dept": "insideSales"},

    # Marketing
    {"username": "ishita", "email": "ishita@ornatesolar.com", "first_name": "Ishita", "last_name": "Bansal", "title": "Marketing Lead", "dept": "marketing"},
    {"username": "kabir", "email": "kabir@ornatesolar.com", "first_name": "Kabir", "last_name": "Malik", "title": "Content & SEO", "dept": "marketing"},
    {"username": "riya", "email": "riya@ornatesolar.com", "first_name": "Riya", "last_name": "Saxena", "title": "Creatives & Video", "dept": "marketing"},

    # Procurement
    {"username": "vivek", "email": "vivek@ornatesolar.com", "first_name": "Vivek", "last_name": "Rao", "title": "Procurement Lead", "dept": "procurement"},
    {"username": "anjali", "email": "anjali@ornatesolar.com", "first_name": "Anjali", "last_name": "Sinha", "title": "Vendor Manager", "dept": "procurement"},
    {"username": "mohit", "email": "mohit@ornatesolar.com", "first_name": "Mohit", "last_name": "Yadav", "title": "Purchase Officer", "dept": "procurement"},
]

FIELD_SAMPLES = {
    # Sales / Inside Sales
    "meeting": ["Met 2 EPC partners in Pune.", "Closed kickoff call with rooftop client.", "On-site visit to Indore plant.", "Demo for prospective dealer in Surat."],
    "revenue": ["₹4.2 L invoiced.", "₹1.8 L PO received.", "₹6 L pipeline added.", "₹2.7 L closed this morning."],
    "newCustomerOnboard": ["1 new dealer onboarded — Surat.", "2 leads moved to onboarding.", "Onboarded Bhopal channel partner.", "—"],
    "calling": ["38 outbound calls.", "27 calls / 6 connected.", "45 calls + 3 follow-ups.", "52 cold calls today."],
    "callingList": ["Refreshed Maharashtra leads.", "Pulled new Tamil Nadu list (120 contacts).", "Cleaned bounced numbers.", "Added 30 fresh referrals."],

    # Marketing
    "videoEditing": ["Cut 30s reel for Instagram.", "Edited installation walkthrough.", "Color-graded testimonial video.", "Final cut for monsoon campaign."],
    "creatives": ["Designed 4 LinkedIn posts.", "New product banner v2.", "Diwali campaign creatives.", "Refreshed homepage hero."],
    "contentWriting": ["Drafted blog: 'Why poly modules?'", "Wrote landing-page copy.", "Edited case study.", "Newsletter draft for May."],
    "seo": ["Tuned 6 meta titles.", "Backlink outreach (10 sites).", "Keyword cluster updated.", "Audited top-10 landing pages."],
    "websiteManagement": ["Updated product specs page.", "Patched WordPress plugins.", "Pushed new contact form.", "Compressed 22 product images."],
    "reporting": ["Sent weekly traffic snapshot.", "MoM lead-source breakdown.", "Updated GA4 conversions.", "Pulled paid-ads ROAS report."],

    # Procurement
    "enquiries": ["Sent RFQs for 12 modules.", "Got 3 fresh quotes for cables.", "Asked 5 vendors on inverters.", "Floated tender for MMS structures."],
    "negotiations": ["Saved 7% on inverter PO.", "Renegotiated freight rates.", "Locked Tier-1 module price.", "Closed 4% discount on cables."],
    "vendorOnboarding": ["Onboarded 1 cable vendor.", "Visit to Noida transformer plant.", "KYC done for 2 new vendors.", "—"],
    "purchaseOrder": ["Raised 3 POs.", "PO #1248 released.", "Bulk PO for 200 panels.", "Released PO for inverters."],
    "payment": ["NOPA cleared for Vendor X.", "Initiated 60% advance.", "Cleared 4 vendor payments.", "Released milestone payment."],
    "dispatches": ["3 trucks dispatched to Jaipur.", "Container left Mundra port.", "Site dispatch to Coimbatore.", "Dispatch to Bhopal in transit."],
    "grn": ["GRN #4421 booked.", "Inverters received & inspected.", "Modules QC passed.", "Material received with minor damage — flagged."],
    "remarks": ["—", "Watch monsoon delays.", "Vendor X payment terms revised.", "Need to revise BoM for upcoming order."],
}

DEFAULT_FIELD_SAMPLES = {
    "workDone": ["Completed module specs.", "Closed 3 customer tickets.", "Reconciled April invoices.", "Drafted Q2 budget sheet."],
    "workInProgress": ["Reviewing supplier datasheets.", "Refactoring auth flow.", "Tracking pending shipment.", "QA pass on dashboard module."],
    "upcomingPriorities": ["Prototype testing on Friday.", "GST filing on 25th.", "Inventory audit Monday.", "Site visit on Thursday."],
    "challenges": ["Awaiting approval on BoM revision.", "Pending PO numbers.", "Need staging DB credentials.", "Truck breakdown delayed Friday delivery."],
    "otherUpdate": ["Attended weekly sync.", "Onboarded new intern.", "—", "CRM cleanup completed."],
}


def pick(key, idx):
    samples = FIELD_SAMPLES.get(key) or DEFAULT_FIELD_SAMPLES.get(key) or [""]
    return samples[idx % len(samples)]


def main():
    print("=" * 60)
    print("Seeding employees + reports")
    print("=" * 60)

    # Build a slug -> Department map
    depts = {d.slug: d for d in Department.objects.all()}
    if not depts:
        print("ERROR: No departments found. Run the department-seed first.")
        return

    created_users = 0
    today = date.today()

    for emp in EMPLOYEES:
        dept = depts.get(emp["dept"])
        user, was_created = User.objects.get_or_create(
            username=emp["username"],
            defaults={
                "email": emp["email"],
                "first_name": emp["first_name"],
                "last_name": emp["last_name"],
                "title": emp["title"],
                "role": "employee",
                "department": dept,
                "is_active": True,
                "is_staff": False,
            },
        )
        if was_created:
            user.set_password("password123")
            user.save()
            created_users += 1
            print(f"  Created user: {user.username}  ({dept.name})")
        else:
            # Make sure existing seeded users still have the right dept/role
            user.department = dept
            user.title = emp["title"]
            user.first_name = emp["first_name"]
            user.last_name = emp["last_name"]
            user.role = "employee"
            user.save()
            print(f"  Updated user: {user.username}")

        # Seed reports — last 10 days, skip a couple for variety so missing-today has data
        fields = dept.report_fields if dept and dept.report_fields else [
            {"key": k, "label": k} for k in DEFAULT_FIELD_SAMPLES.keys()
        ]
        for day_offset in range(0, 10):
            # Skip a couple of days for some employees so the "missing" view has rows
            if (hash(emp["username"]) + day_offset) % 6 == 0:
                continue
            d = today - timedelta(days=day_offset)
            data = {f["key"]: pick(f["key"], hash(emp["username"]) + day_offset) for f in fields}
            DailyReport.objects.update_or_create(
                user=user,
                date=d,
                defaults={"data": data},
            )

    print()
    print(f"Users created this run: {created_users}")
    print(f"Total active employees: {User.objects.filter(role='employee', is_active=True).count()}")
    print(f"Total reports: {DailyReport.objects.count()}")
    print()
    print("Login as any employee with password: password123")
    print("Examples:")
    print("  divya@ornatesolar.com  -> Sales")
    print("  ishita@ornatesolar.com -> Marketing")
    print("  vivek@ornatesolar.com  -> Procurement")


main()
