"""
Seed dummy daily reports for a chosen set of departments only.

Usage (inside fastapi container):
    docker compose exec fastapi python -m scripts.seed_selected_depts \
        logistics procurement marketing sales

Each named department's employees get ~7 days of reports.  Re-running is safe
— rows are upserted on (user_id, date).
"""
from __future__ import annotations

import sys
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from database import SessionLocal  # noqa: E402
from models import DailyReport, Department, User  # noqa: E402

DAYS_BACK = 7

FIELD_SAMPLES: dict[str, list[str]] = {
    # ---- Generic 5-field template (BESS Sales, Design, Finance, HR, O&M,
    # Production, Project, R&D, R&D-BESS, Service, Support, Web Dev) ----
    "workDone": [
        "Completed module specs.",
        "Closed 3 customer tickets.",
        "Reconciled April invoices.",
        "Drafted Q2 budget sheet.",
        "Finished daily QC pass.",
        "Reviewed inverter datasheets.",
        "Updated rooftop layout for Coimbatore site.",
        "Kicked off Bhopal install.",
    ],
    "workInProgress": [
        "Reviewing supplier datasheets.",
        "Refactoring auth flow.",
        "Tracking pending shipment.",
        "QA pass on dashboard module.",
        "Drafting site SOP.",
        "Checking BoM for Pune project.",
        "Following up on cable RFQ.",
        "Auditing weekly leads.",
    ],
    "upcomingPriorities": [
        "Prototype testing on Friday.",
        "GST filing on 25th.",
        "Inventory audit Monday.",
        "Site visit on Thursday.",
        "Vendor evaluation next week.",
        "Tender deadline on 5th.",
        "Board review prep.",
        "Client demo on 12th.",
    ],
    "challenges": [
        "Awaiting approval on BoM revision.",
        "Pending PO numbers.",
        "Need staging DB credentials.",
        "Truck breakdown delayed Friday delivery.",
        "Vendor X payment held up.",
        "Material delayed at Mundra port.",
        "Site team short on hands.",
        "—",
    ],
    "otherUpdate": [
        "Attended weekly sync.",
        "Onboarded new intern.",
        "—",
        "CRM cleanup completed.",
        "Filed expense claims.",
        "Refreshed lead-pipeline doc.",
    ],

    # ---- Sales ----
    "meeting": [
        "Met 2 EPC partners in Pune.",
        "Closed kickoff call with rooftop client.",
        "On-site visit to Indore plant.",
        "Demo for prospective dealer in Surat.",
        "Quarterly review with channel partner.",
    ],
    "revenue": [
        "₹4.2 L invoiced.",
        "₹1.8 L PO received.",
        "₹6 L pipeline added.",
        "₹2.7 L closed this morning.",
        "Closed Q4 forecast at ₹12 L.",
    ],
    "newCustomerOnboard": [
        "1 new dealer onboarded — Surat.",
        "2 leads moved to onboarding.",
        "Onboarded Bhopal channel partner.",
        "Onboarded Tirupati EPC.",
        "—",
    ],

    # ---- Marketing ----
    "videoEditing": [
        "Cut 30s reel for Instagram.",
        "Edited installation walkthrough.",
        "Color-graded testimonial video.",
        "Final cut for monsoon campaign.",
        "Trimmed event recap.",
    ],
    "creatives": [
        "Designed 4 LinkedIn posts.",
        "New product banner v2.",
        "Diwali campaign creatives.",
        "Refreshed homepage hero.",
        "Drafted print ad layouts.",
    ],
    "contentWriting": [
        "Drafted blog: 'Why poly modules?'",
        "Wrote landing-page copy.",
        "Edited case study.",
        "Newsletter draft for May.",
        "Wrote FAQ for new product.",
    ],
    "seo": [
        "Tuned 6 meta titles.",
        "Backlink outreach (10 sites).",
        "Keyword cluster updated.",
        "Audited top-10 landing pages.",
        "Fixed broken internal links.",
    ],
    "websiteManagement": [
        "Updated product specs page.",
        "Patched WordPress plugins.",
        "Pushed new contact form.",
        "Compressed 22 product images.",
        "Set up redirect rules for renamed URLs.",
    ],
    "reporting": [
        "Sent weekly traffic snapshot.",
        "MoM lead-source breakdown.",
        "Updated GA4 conversions.",
        "Pulled paid-ads ROAS report.",
        "Compiled creative-performance deck.",
    ],

    # ---- Procurement ----
    "enquiries": [
        "Sent RFQs for 12 modules.",
        "Got 3 fresh quotes for cables.",
        "Asked 5 vendors on inverters.",
        "Floated tender for MMS structures.",
        "Received 7 quotes for trackers.",
    ],
    "negotiations": [
        "Saved 7% on inverter PO.",
        "Renegotiated freight rates.",
        "Locked Tier-1 module price.",
        "Closed 4% discount on cables.",
        "Negotiated extended payment terms.",
    ],
    "vendorOnboarding": [
        "Onboarded 1 cable vendor.",
        "Visit to Noida transformer plant.",
        "KYC done for 2 new vendors.",
        "Blacklisted unresponsive supplier.",
        "—",
    ],
    "purchaseOrder": [
        "Raised 3 POs.",
        "PO #1248 released.",
        "Bulk PO for 200 panels.",
        "Released PO for inverters.",
        "Amended PO #1267 for revised qty.",
    ],
    "payment": [
        "NOPA cleared for Vendor X.",
        "Initiated 60% advance.",
        "Cleared 4 vendor payments.",
        "Released milestone payment.",
        "Held back final 10% pending GRN.",
    ],
    "dispatches": [
        "3 trucks dispatched to Jaipur.",
        "Container left Mundra port.",
        "Site dispatch to Coimbatore.",
        "Dispatch to Bhopal in transit.",
        "Last-mile delivery to Pune complete.",
    ],
    "grn": [
        "GRN #4421 booked.",
        "Inverters received & inspected.",
        "Modules QC passed.",
        "Material received with minor damage — flagged.",
        "Cables GRN booked, qty matched.",
    ],
    "remarks": [
        "—",
        "Watch monsoon delays.",
        "Vendor X payment terms revised.",
        "Need to revise BoM for upcoming order.",
        "Weekly review scheduled Monday.",
    ],

    # ---- Inside Sales (numeric daily snapshot) ----
    "totalCalls":   ["752", "778", "697", "783", "1051", "640", "812", "905"],
    "picked":       ["412", "467", "312", "325", "581", "388", "455", "510"],
    "closed":       ["14", "7", "3", "2", "1", "9", "6", "11"],
    "lost":         ["7", "11", "16", "5", "14", "8", "12", "4"],
    "following":    ["21", "65", "45", "58", "40", "33", "52", "47"],
    "invoiceTotal": ["1857450", "816830", "331850", "240676", "90300", "612000", "1145200", "455900"],

    # ---- Logistics (new 12-field schema) ----
    "warehouseCoordination": [
        "Coordinated unloading of 200-panel shipment.",
        "Synced with site team on storage layout.",
        "Reorganized warehouse rack 4A.",
        "Spot-check audit of inbound material.",
        "Coordinated cycle count with night shift.",
    ],
    "truckTransportationCoordination": [
        "Arranged 3 trucks for Jaipur dispatch.",
        "Coordinated with TruckSeva for Indore route.",
        "Confirmed truck arrival ETA with site team.",
        "Re-routed truck due to highway closure.",
        "Scheduled return trip for empty trailers.",
    ],
    "courierDispatch": [
        "Dispatched 8 parcels via Bluedart.",
        "Booked 4 courier shipments for spare parts.",
        "Initiated overnight courier for site documents.",
        "Sent inverter accessories via DTDC.",
        "Dispatched marketing collateral to 6 dealers.",
    ],
    "shipmentTracking": [
        "Tracked Mundra container — ETA Friday.",
        "Following up on delayed Pune consignment.",
        "All 5 in-transit shipments on schedule.",
        "Resolved tracking mismatch for AWB 8821.",
        "Updated daily tracker dashboard.",
    ],
    "freightSharingDetails": [
        "Shared freight rates with finance for booking.",
        "Compared 3 transporter quotes for Tamil Nadu.",
        "Submitted weekly freight summary to manager.",
        "Forwarded LR copies to accounts.",
        "Sent freight ledger to vendor.",
    ],
    "flasherDataUpload": [
        "Uploaded flasher data for 120 modules.",
        "Pending flasher data from QA — chased.",
        "Processed flasher data batch #14.",
        "Reconciled flasher logs with shipment.",
        "Uploaded flasher CSV to portal.",
    ],
    "stockIn": [
        "Booked GRN for 200 modules.",
        "Inbound 50 inverters — verified & shelved.",
        "Stock-in 1,500 connectors received.",
        "Recorded 12 pallets of mounting structures.",
        "Booked spare parts from Vendor X.",
    ],
    "stockOut": [
        "Issued 50 modules for Surat site.",
        "Stock-out 8 inverters to Pune.",
        "Released 30 cables for Project A.",
        "Dispatched 100 mounting brackets.",
        "Stock-out for Site C — 12 inverters.",
    ],
    "portalOperations": [
        "Updated dispatch portal entries.",
        "Resolved portal sync error.",
        "Closed 7 portal tickets.",
        "Generated weekly portal report.",
        "Updated user access on Tally portal.",
    ],
    "courierIssueResolution": [
        "Resolved damaged-package claim with Bluedart.",
        "Re-shipped lost parcel for dealer.",
        "Filed POD complaint with DTDC.",
        "Escalated 2-day delay to courier ops.",
        "Recovered missing AWB tracking.",
    ],
    "inventoryManagement": [
        "Stock audit for module section completed.",
        "Reconciled 1 SKU mismatch.",
        "Updated inventory list in ERP.",
        "Performed monthly cycle count.",
        "Cleared dead stock from rack 6.",
    ],
    "otherOperationalWork": [
        "Attended weekly logistics review.",
        "Updated dispatch SOP draft.",
        "Hosted vendor visit for spot check.",
        "Onboarded new helper to night shift.",
        "Filed bi-weekly compliance report.",
    ],
}


def _pick(key: str, idx: int) -> str:
    samples = FIELD_SAMPLES.get(key, [""])
    return samples[idx % len(samples)]


def main() -> None:
    if len(sys.argv) < 2:
        print("Usage: python -m scripts.seed_selected_depts <slug> [<slug> ...]")
        sys.exit(1)
    target_slugs = sys.argv[1:]

    today = date.today()
    now = datetime.now(timezone.utc)
    db = SessionLocal()
    try:
        depts = db.query(Department).filter(Department.slug.in_(target_slugs)).all()
        found_slugs = {d.slug for d in depts}
        missing = [s for s in target_slugs if s not in found_slugs]
        if missing:
            print(f"  Unknown department slugs (skipped): {missing}")
        dept_ids = [d.id for d in depts]

        employees = (
            db.query(User)
            .filter(
                User.role == "employee",
                User.is_active.is_(True),
                User.department_id.in_(dept_ids),
            )
            .all()
        )
        for emp in employees:
            _ = emp.department

        print(f"Seeding {DAYS_BACK} days of reports for {len(employees)} employees "
              f"across {len(depts)} department(s): {sorted(found_slugs)}")

        upserts = 0
        skipped = 0
        for emp_idx, emp in enumerate(employees):
            fields = (emp.department.report_fields or []) if emp.department else []
            if not fields:
                continue
            for day_offset in range(DAYS_BACK):
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
        print(f"  Range: {today - timedelta(days=DAYS_BACK - 1)} to {today}")
    finally:
        db.close()


if __name__ == "__main__":
    main()
