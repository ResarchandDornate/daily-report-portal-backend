"""
Generate dummy daily reports for every existing employee so HR can preview
the weekly / monthly summary in the dashboard.

Run from inside the FastAPI container:
    docker compose exec fastapi python -m scripts.seed_reports

Re-running is safe — upserts on (user_id, date).  By default seeds the last
7 days, skipping ~1 in 6 days per employee so "missing today" has rows.
"""
from __future__ import annotations

import sys
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

# Make `from auth import ...` etc. resolvable when run as a module.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from database import SessionLocal  # noqa: E402
from models import DailyReport, User  # noqa: E402

DAYS_BACK = 7

FIELD_SAMPLES = {
    # Generic 5-field template
    "workDone":           ["Completed module specs.", "Closed 3 customer tickets.", "Reconciled April invoices.", "Drafted Q2 budget sheet.", "Finished daily QC pass.", "Reviewed inverter datasheets.", "Updated rooftop layout for Coimbatore site.", "Kicked off Bhopal install."],
    "workInProgress":     ["Reviewing supplier datasheets.", "Refactoring auth flow.", "Tracking pending shipment.", "QA pass on dashboard module.", "Drafting site SOP.", "Checking BoM for Pune project.", "Following up on cable RFQ.", "Auditing weekly leads."],
    "upcomingPriorities": ["Prototype testing on Friday.", "GST filing on 25th.", "Inventory audit Monday.", "Site visit on Thursday.", "Vendor evaluation next week.", "Tender deadline on 5th.", "Board review prep.", "Client demo on 12th."],
    "challenges":         ["Awaiting approval on BoM revision.", "Pending PO numbers.", "Need staging DB credentials.", "Truck breakdown delayed Friday delivery.", "Vendor X payment held up.", "Material delayed at Mundra port.", "Site team short on hands.", "—"],
    "otherUpdate":        ["Attended weekly sync.", "Onboarded new intern.", "—", "CRM cleanup completed.", "Filed expense claims.", "Refreshed lead-pipeline doc."],

    # Sales / Inside Sales
    "meeting":            ["Met 2 EPC partners in Pune.", "Closed kickoff call with rooftop client.", "On-site visit to Indore plant.", "Demo for prospective dealer in Surat.", "Quarterly review with channel partner."],
    "revenue":            ["₹4.2 L invoiced.", "₹1.8 L PO received.", "₹6 L pipeline added.", "₹2.7 L closed this morning.", "Closed Q4 forecast at ₹12L."],
    "newCustomerOnboard": ["1 new dealer onboarded — Surat.", "2 leads moved to onboarding.", "Onboarded Bhopal channel partner.", "Onboarded Tirupati EPC.", "—"],
    "calling":            ["38 outbound calls.", "27 calls / 6 connected.", "45 calls + 3 follow-ups.", "52 cold calls today.", "Reached 18 fresh leads."],
    "callingList":        ["Refreshed Maharashtra leads.", "Pulled new Tamil Nadu list (120 contacts).", "Cleaned bounced numbers.", "Added 30 fresh referrals.", "Synced with Salesforce export."],

    # Marketing
    "videoEditing":      ["Cut 30s reel for Instagram.", "Edited installation walkthrough.", "Color-graded testimonial video.", "Final cut for monsoon campaign.", "Trimmed event recap."],
    "creatives":         ["Designed 4 LinkedIn posts.", "New product banner v2.", "Diwali campaign creatives.", "Refreshed homepage hero.", "Drafted print ad layouts."],
    "contentWriting":    ["Drafted blog: 'Why poly modules?'", "Wrote landing-page copy.", "Edited case study.", "Newsletter draft for May.", "Wrote FAQ for new product."],
    "seo":               ["Tuned 6 meta titles.", "Backlink outreach (10 sites).", "Keyword cluster updated.", "Audited top-10 landing pages.", "Fixed broken internal links."],
    "websiteManagement": ["Updated product specs page.", "Patched WordPress plugins.", "Pushed new contact form.", "Compressed 22 product images.", "Set up redirect rules for renamed URLs."],
    "reporting":         ["Sent weekly traffic snapshot.", "MoM lead-source breakdown.", "Updated GA4 conversions.", "Pulled paid-ads ROAS report.", "Compiled creative-performance deck."],

    # Procurement
    "enquiries":        ["Sent RFQs for 12 modules.", "Got 3 fresh quotes for cables.", "Asked 5 vendors on inverters.", "Floated tender for MMS structures.", "Received 7 quotes for trackers."],
    "negotiations":     ["Saved 7% on inverter PO.", "Renegotiated freight rates.", "Locked Tier-1 module price.", "Closed 4% discount on cables.", "Negotiated extended payment terms."],
    "vendorOnboarding": ["Onboarded 1 cable vendor.", "Visit to Noida transformer plant.", "KYC done for 2 new vendors.", "Blacklisted unresponsive supplier.", "—"],
    "purchaseOrder":    ["Raised 3 POs.", "PO #1248 released.", "Bulk PO for 200 panels.", "Released PO for inverters.", "Amended PO #1267 for revised qty."],
    "payment":          ["NOPA cleared for Vendor X.", "Initiated 60% advance.", "Cleared 4 vendor payments.", "Released milestone payment.", "Held back final 10% pending GRN."],
    "dispatches":       ["3 trucks dispatched to Jaipur.", "Container left Mundra port.", "Site dispatch to Coimbatore.", "Dispatch to Bhopal in transit.", "Last-mile delivery to Pune complete."],
    "grn":              ["GRN #4421 booked.", "Inverters received & inspected.", "Modules QC passed.", "Material received with minor damage — flagged.", "Cables GRN booked, qty matched."],
    "remarks":          ["—", "Watch monsoon delays.", "Vendor X payment terms revised.", "Need to revise BoM for upcoming order.", "Weekly review scheduled Monday."],

    # Logistics (custom template)
    "task":         ["Coordinated dispatch to Pune site.", "Tracked 3 in-transit consignments.", "Resolved POD mismatch with carrier.", "Planned weekend shipment to Indore.", "Reviewed monthly freight invoices."],
    "workProgress": ["50% — pending vendor confirmation.", "On track for Friday cutoff.", "Delayed by 1 day, recovering.", "Awaiting GRN from site team.", "Closed."],
}


def _pick(key: str, idx: int) -> str:
    samples = FIELD_SAMPLES.get(key, [""])
    return samples[idx % len(samples)]


def main() -> None:
    today = date.today()
    # Django created the table without DB-level DEFAULTs on submitted_at /
    # created_at, so we set them explicitly per row.
    now = datetime.now(timezone.utc)
    db = SessionLocal()
    try:
        employees = (
            db.query(User)
            .filter(User.role == "employee", User.is_active.is_(True))
            .all()
        )
        # Force-load department + report_fields up front (single round-trip).
        for emp in employees:
            _ = emp.department  # triggers lazy-load

        print(f"Seeding ~{DAYS_BACK} days of reports for {len(employees)} employees…")

        upserts = 0
        skipped = 0
        no_dept = 0
        for emp_idx, emp in enumerate(employees):
            if not emp.department:
                no_dept += 1
                continue
            fields = emp.department.report_fields or []
            if not fields:
                continue
            for day_offset in range(DAYS_BACK):
                # Skip ~1 in 6 days per employee deterministically
                if (emp_idx + day_offset * 3) % 7 == 0:
                    skipped += 1
                    continue
                d = today - timedelta(days=day_offset)
                data = {
                    f["key"]: _pick(f["key"], emp_idx + day_offset + fi)
                    for fi, f in enumerate(fields)
                }
                existing = (
                    db.query(DailyReport)
                    .filter(DailyReport.user_id == emp.id, DailyReport.date == d)
                    .first()
                )
                if existing:
                    existing.data = data
                    existing.submitted_at = now
                else:
                    db.add(DailyReport(
                        user_id=emp.id,
                        date=d,
                        data=data,
                        submitted_at=now,
                        created_at=now,
                    ))
                upserts += 1

        db.commit()

        print(f"  Reports upserted:        {upserts}")
        print(f"  Days skipped on purpose: {skipped}")
        if no_dept:
            print(f"  Employees with no department: {no_dept}")
        print(f"  Total reports in DB now: {db.query(DailyReport).count()}")
        print(f"  Range: {today - timedelta(days=DAYS_BACK - 1)} to {today}")
    finally:
        db.close()


if __name__ == "__main__":
    main()
