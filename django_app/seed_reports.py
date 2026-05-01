"""
Generate dummy daily reports for the 102 real employees so HR can preview the
weekly / monthly summary in the dashboard.

  docker compose exec django python manage.py shell -c \
      "exec(open('/app/seed_reports.py').read())"

Re-running is safe — it upserts on (user, date).  By default seeds the last
7 days, skipping ~1 in 6 days per employee so "missing today" has rows.
"""
import random
from datetime import date, timedelta

from django.contrib.auth import get_user_model

from reports.models import DailyReport

User = get_user_model()
random.seed(42)  # deterministic re-runs

DAYS_BACK = 7

# Realistic content per field key.  Pulled from the original mock data so the
# summary text reads like a real Ornate roster.  The script picks an entry from
# the matching list deterministically per (employee, day, field).
FIELD_SAMPLES = {
    # Generic 5-field template (used by most departments by default)
    "workDone":           ["Completed module specs.", "Closed 3 customer tickets.", "Reconciled April invoices.", "Drafted Q2 budget sheet.", "Finished daily QC pass.", "Reviewed inverter datasheets.", "Updated rooftop layout for Coimbatore site.", "Kicked off Bhopal install."],
    "workInProgress":     ["Reviewing supplier datasheets.", "Refactoring auth flow.", "Tracking pending shipment.", "QA pass on dashboard module.", "Drafting site SOP.", "Checking BoM for Pune project.", "Following up on cable RFQ.", "Auditing weekly leads."],
    "upcomingPriorities": ["Prototype testing on Friday.", "GST filing on 25th.", "Inventory audit Monday.", "Site visit on Thursday.", "Vendor evaluation next week.", "Tender deadline on 5th.", "Board review prep.", "Client demo on 12th."],
    "challenges":         ["Awaiting approval on BoM revision.", "Pending PO numbers.", "Need staging DB credentials.", "Truck breakdown delayed Friday delivery.", "Vendor X payment held up.", "Material delayed at Mundra port.", "Site team short on hands.", "—"],
    "otherUpdate":        ["Attended weekly sync.", "Onboarded new intern.", "—", "CRM cleanup completed.", "Filed expense claims.", "Refreshed lead-pipeline doc."],

    # Sales
    "meeting":            ["Met 2 EPC partners in Pune.", "Closed kickoff call with rooftop client.", "On-site visit to Indore plant.", "Demo for prospective dealer in Surat.", "Quarterly review with channel partner."],
    "revenue":            ["₹4.2 L invoiced.", "₹1.8 L PO received.", "₹6 L pipeline added.", "₹2.7 L closed this morning.", "Closed Q4 forecast at ₹12L."],
    "newCustomerOnboard": ["1 new dealer onboarded — Surat.", "2 leads moved to onboarding.", "Onboarded Bhopal channel partner.", "Onboarded Tirupati EPC.", "—"],

    # Inside Sales
    "calling":     ["38 outbound calls.", "27 calls / 6 connected.", "45 calls + 3 follow-ups.", "52 cold calls today.", "Reached 18 fresh leads."],
    "callingList": ["Refreshed Maharashtra leads.", "Pulled new Tamil Nadu list (120 contacts).", "Cleaned bounced numbers.", "Added 30 fresh referrals.", "Synced with Salesforce export."],

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

    # Logistics (recently customised)
    "task":         ["Coordinated dispatch to Pune site.", "Tracked 3 in-transit consignments.", "Resolved POD mismatch with carrier.", "Planned weekend shipment to Indore.", "Reviewed monthly freight invoices."],
    "workProgress": ["50% — pending vendor confirmation.", "On track for Friday cutoff.", "Delayed by 1 day, recovering.", "Awaiting GRN from site team.", "Closed."],
}


def pick(key, idx):
    samples = FIELD_SAMPLES.get(key) or [""]
    return samples[idx % len(samples)]


def main():
    today = date.today()
    employees = list(User.objects.filter(role="employee", is_active=True).select_related("department"))
    print(f"Seeding ~{DAYS_BACK} days of reports for {len(employees)} employees...")

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
            # Skip ~1 in 6 days per employee for variety (so "missing today" has rows)
            if (emp_idx + day_offset * 3) % 7 == 0:
                skipped += 1
                continue
            d = today - timedelta(days=day_offset)
            data = {}
            for fi, f in enumerate(fields):
                data[f["key"]] = pick(f["key"], emp_idx + day_offset + fi)
            DailyReport.objects.update_or_create(
                user=emp,
                date=d,
                defaults={"data": data},
            )
            upserts += 1

    print(f"  Reports upserted:        {upserts}")
    print(f"  Days skipped on purpose: {skipped}")
    if no_dept:
        print(f"  Employees with no department (skipped): {no_dept}")
    print(f"  Total reports in DB now: {DailyReport.objects.count()}")
    print(f"  Range: {today - timedelta(days=DAYS_BACK - 1)} to {today}")


main()
