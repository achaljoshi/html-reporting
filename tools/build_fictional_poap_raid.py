#!/usr/bin/env python3
"""Build FICTIONAL sample POAP + RAID workbooks for the 'Self Assessment' sample portfolio.

Re-runnable:  python3 tools/build_fictional_poap_raid.py
Optional   :  RECALC_SCRIPT=/path/to/recalc.py  (LibreOffice recalc, run on the POAP afterwards)

Outputs (all content is invented; roles only, generic system names):
  samples/ATS_Program_SAMPLE/Self Assessment/ATS_POAP.xlsx      (from templates/ATS_POAP_Template.xlsx)
  samples/ATS_Program_SAMPLE/Self Assessment/ATS_RAID_Log.xlsx  (layout read by assets/js/raid.js)
"""
import copy
import datetime as dt
import os
import subprocess
import sys

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(ROOT, "samples", "ATS_Program_SAMPLE", "Self Assessment")
TEMPLATE = os.path.join(ROOT, "templates", "ATS_POAP_Template.xlsx")
POAP_OUT = os.path.join(OUT_DIR, "ATS_POAP.xlsx")
RAID_OUT = os.path.join(OUT_DIR, "ATS_RAID_Log.xlsx")

TODAY = dt.date(2026, 10, 2)
D = dt.date.fromisoformat
FONT = "Arial"
NAVY = "1B2A4A"

# ----------------------------------------------------------------------------
# POAP content
# ----------------------------------------------------------------------------
PILLAR_A, PILLAR_B = "Pillar A – Returns", "Pillar B – Payments"
PILLARS = [(PILLAR_A, "2563EB"), (PILLAR_B, "0EA5A0")]
PROGRAM = "HMRC ATS – Self Assessment (SAMPLE)"

# (type, item, start, end, scope)   -- dates for the base topic; each topic applies a shift in days
CHAIN = [
    ("Test Prep", "Test design & scripting", "2026-06-15", "2026-07-31", "Design and peer-review of release test scenarios."),
    ("Dependencies", "Test environment & data readiness", "2026-07-13", "2026-08-07", "Environment build, masked test data load and access set-up."),
    ("Smoke / P2P", "Smoke & connectivity checks", "2026-08-03", "2026-08-12", "Portal to API Gateway to Ledger connectivity smoke."),
    ("SIT", "SIT Cycle 1 execution", "2026-08-13", "2026-10-30", "Functional system integration test, first cycle."),
    ("Defect Retest", "Defect retest – Cycle 1", "2026-10-12", "2026-11-13", "Retest of fixes delivered during Cycle 1."),
    ("SIT", "SIT Cycle 2 execution", "2026-11-16", "2026-12-18", "Second cycle incl. fix verification and new scope."),
    ("Regression", "Regression pack execution", "2027-01-11", "2027-02-05", "Full regression of in-scope journeys."),
    ("TCR / Sign-off", "SIT exit report & TCR", "2027-02-08", "2027-02-19", "Test completion report and SIT sign-off."),
    ("E2E", "End-to-end journey testing", "2027-02-22", "2027-03-26", "Cross-system journeys across the portfolio."),
    ("NFR", "Performance & resilience testing", "2027-03-01", "2027-04-02", "Load, soak and failover tests on a production-like set-up."),
    ("Defect Retest", "Defect retest – E2E", "2027-03-29", "2027-04-16", "Retest of E2E and NFR defect fixes."),
    ("TCR / Sign-off", "E2E sign-off", "2027-04-19", "2027-04-30", "E2E test completion report and release recommendation."),
    ("Automation", "Automation pack refresh", "2027-05-03", "2027-05-28", "Refresh of the regression automation pack."),
]
FREEZE = ("Holiday / Freeze", "Christmas change freeze", "2026-12-21", "2027-01-04", "No change to test environments.")

# topic, pillar, id prefix, shift (days), tbc index (0-based in CHAIN) or None, delayed {index: (kind, extra_days)}
TOPICS = [
    ("Self Assessment – Release 1", PILLAR_A, "A", 0, None, {}),
    ("Self Assessment – Release 2", PILLAR_A, "A", 7, None, {3: ("Amber", 10)}),
    ("Self Assessment – Release 3", PILLAR_B, "B", 14, 9, {4: ("Red", 21)}),
    ("Self Assessment – Digital Filing", PILLAR_B, "B", 28, 8, {}),
]
OWNER_BY_TYPE = {"Dependencies": "Portfolio Test Lead", "TCR / Sign-off": "Portfolio Test Lead",
                 "Holiday / Freeze": "Portfolio Test Lead"}


def build_plan():
    rows, counters = [], {"A": 0, "B": 0}
    for topic, pillar, pre, shift, tbc, delayed in TOPICS:
        ids = []
        bars = [(i, b, False) for i, b in enumerate(CHAIN)] + [(None, FREEZE, True)]
        for idx, (typ, item, s, e, scope), is_freeze in bars:
            counters[pre] += 1
            bid = f"{pre}-{counters[pre]:03d}"
            start = end = None
            if not (idx is not None and idx == tbc):
                sh = 0 if is_freeze else shift
                start, end = D(s) + dt.timedelta(days=sh), D(e) + dt.timedelta(days=sh)
            base_s, base_e = start, end
            rag, notes = "Green", None
            if start is None:
                status, pct = "Dates TBC", 0.0
            elif end < TODAY:
                status, pct = "Complete", 1.0
            elif start <= TODAY <= end:
                status = "In Progress"
                pct = round(((TODAY - start).days + 1) / ((end - start).days + 1), 2)
            else:
                status, pct = "Not Started", 0.0
            if idx in delayed:
                rag, extra = delayed[idx]
                base_e = end - dt.timedelta(days=extra)
                status, notes = "Delayed", "SAMPLE slip for demo"
            if is_freeze:
                dep = ""
            elif ids:
                dep = ids[-1]
            else:
                dep = ""
            row = dict(id=bid, pillar=pillar, topic=topic, type=typ, item=item, start=start, end=end,
                       bs=base_s, be=base_e, pct=pct, status=status, rag=rag, owner=OWNER_BY_TYPE.get(typ, "Test Manager"),
                       dep=dep, scope=scope, notes=notes)
            rows.append(row)
            if not is_freeze:
                ids.append(bid)
            else:
                freeze_id = bid
        # regression depends on SIT cycle 2 and the freeze
        for r in rows:
            if r["topic"] == topic and r["type"] == "Regression":
                r["dep"] = f"{r['dep']}, {freeze_id}"
    return rows


def milestones(plan):
    by = {}
    for r in plan:
        by[(r["topic"], r["item"])] = r

    def sit(topic, n):
        return by[(topic, f"SIT Cycle {n} execution")]

    out, n = [], [0]

    def add(pillar, topic, name, date, typ, status, notes=""):
        n[0] += 1
        out.append([f"M-{n[0]:02d}", pillar, topic, name, date, typ, status, notes])

    r1, r2, r3, df = (t[0] for t in TOPICS)
    add(PILLAR_A, r1, "Release 1 SIT start", sit(r1, 1)["start"], "SIT Start", "Achieved")
    add(PILLAR_A, r1, "Release 1 SIT Cycle 1 end", sit(r1, 1)["end"], "SIT End", "Planned")
    add(PILLAR_A, r2, "Release 2 SIT start", sit(r2, 1)["start"], "SIT Start", "Achieved")
    add(PILLAR_A, r2, "Release 2 SIT Cycle 1 end", sit(r2, 1)["end"], "SIT End", "At Risk", "Slipping – see POAP_Plan")
    add(PILLAR_B, r3, "Release 3 SIT start", sit(r3, 1)["start"], "SIT Start", "Achieved")
    add(PILLAR_B, r3, "Release 3 SIT Cycle 1 end", sit(r3, 1)["end"], "SIT End", "Planned")
    add(PILLAR_B, df, "Digital Filing SIT start", sit(df, 1)["start"], "SIT Start", "Achieved")
    add(PILLAR_A, r1, "Release 1 SIT exit sign-off", by[(r1, "SIT exit report & TCR")]["end"], "Sign-off / TCR", "Planned")
    add("", "", "Test environment refresh complete", D("2026-10-16"), "Environment", "Planned", "Programme-wide")
    add(PILLAR_B, r3, "Vendor drop 3.2 delivered", D("2026-10-09"), "Dependency", "At Risk", "Awaiting vendor confirmation")
    add("", "", "Christmas change freeze begins", D("2026-12-21"), "Gate", "Planned", "Programme-wide")
    add(PILLAR_A, r1, "Release 1 go-live gate", D("2027-05-14"), "Go-Live", "Planned", "Go / no-go decision")
    return out


P2P = [
    # id, topic, sub process, path, test no, parent, doc, date, status, 2xx, 3xx, 4xx, 5xx, notes
    ["F-01", "Self Assessment – Release 1", "Return submission", "Portal → API Gateway → Ledger", "Test 1", "P2P – Flow 01", "FLOW-SPEC-01", D("2026-08-07"), "Complete", "Pass", "Pass", "Pass", "Pass", ""],
    ["F-02", "Self Assessment – Release 1", "Payment on account", "Ledger → Notification Service", "Test 2", "P2P – Flow 02", "FLOW-SPEC-02", D("2026-08-12"), "Complete", "Pass", "Pass", "Pass", "Pass", ""],
    ["F-03", "Self Assessment – Release 2", "Amend a return", "Portal → API Gateway → Document Store", "Test 3", "P2P – Flow 03", "FLOW-SPEC-03", D("2026-09-18"), "In Progress", "Pass", "Pass", "Fail", "Not Run", "Defect raised for 4xx validation message"],
    ["F-04", "Self Assessment – Release 2", "Supporting documents", "API Gateway → Document Store", "Test 4", "P2P – Flow 04", "FLOW-SPEC-04", D("2026-09-25"), "In Progress", "Pass", "Not Run", "Not Run", "Not Run", ""],
    ["F-05", "Self Assessment – Release 3", "Instalment plan set-up", "Portal → API Gateway → Ledger → Notification Service", "Test 5", "P2P – Flow 05", "FLOW-SPEC-05", D("2026-10-14"), "Blocked", "Blocked", "Not Run", "Not Run", "Not Run", "Blocked by test environment outage"],
    ["F-06", "Self Assessment – Release 3", "Refund notification", "Ledger → Notification Service", "Test 6", "P2P – Flow 06", "FLOW-SPEC-06", D("2026-10-28"), "Not Started", "Not Run", "Not Run", "Not Run", "Not Run", ""],
    ["F-07", "Self Assessment – Digital Filing", "Online filing hand-off", "Portal → API Gateway", "Test 7", "P2P – Flow 07", "FLOW-SPEC-07", D("2026-10-21"), "Not Started", "Not Run", "N/A", "Not Run", "Not Run", ""],
    ["F-08", "Self Assessment – Digital Filing", "Receipt retrieval", "Document Store → API Gateway → Portal", "Test 8", "P2P – Flow 08", "FLOW-SPEC-08", None, "Dates TBC", "Not Run", "Not Run", "Not Run", "Not Run", ""],
]


def copy_style(src, dst):
    dst.font = copy.copy(src.font)
    dst.fill = copy.copy(src.fill)
    dst.border = copy.copy(src.border)
    dst.alignment = copy.copy(src.alignment)
    dst.number_format = src.number_format
    dst.protection = copy.copy(src.protection)


def restyle_row_like(ws, row, like_row, ncols):
    for c in range(1, ncols + 1):
        copy_style(ws.cell(like_row, c), ws.cell(row, c))


def build_poap():
    wb = openpyxl.load_workbook(TEMPLATE)
    plan = build_plan()

    # Guide: SAMPLE note
    g = wb["Guide"]
    g["B4"] = "SAMPLE – fictional data (dates as at 2026-10-02). System names, topics and slips are invented for demonstration."
    g["B4"].font = Font(name=FONT, size=10, bold=True, color="B54708")

    # Config
    c = wb["POAP_Config"]
    c["A2"] = "SAMPLE – fictional data. Global settings, pillar colours and activity-type colours used by the dashboard."
    c["B5"] = PROGRAM
    c["B6"] = D("2026-06-01")
    c["B7"] = D("2027-06-30")
    c["B8"] = TODAY
    for ref in ("B6", "B7", "B8"):
        c[ref].number_format = "yyyy-mm-dd"
        c[ref].alignment = Alignment(horizontal="left")
    c["B9"] = "SAMPLE"
    # pillar table rows 13..17 (template): 2 pillars, clear the rest
    blank = c["A18"]
    for i in range(5):
        r = 13 + i
        if i < len(PILLARS):
            c.cell(r, 1, PILLARS[i][0])
            c.cell(r, 2, PILLARS[i][1])
        else:
            c.cell(r, 1).value = None
            c.cell(r, 2).value = None
            copy_style(blank, c.cell(r, 1))
            copy_style(blank, c.cell(r, 2))

    # Lists: pillars
    ls = wb["Lists"]
    for r in range(2, 8):
        ls.cell(r, 1).value = PILLARS[r - 2][0] if r - 2 < len(PILLARS) else None

    # Plan
    ws = wb["POAP_Plan"]
    restyle_row_like(ws, 6, 7, 18)
    for i, p in enumerate(plan):
        r = 6 + i
        vals = [p["id"], p["pillar"], p["topic"], p["type"], p["item"], p["start"], p["end"], p["bs"], p["be"], p["pct"],
                p["status"], p["rag"], p["owner"], p["dep"] or None, p["scope"]]
        for col, v in enumerate(vals, start=1):
            ws.cell(r, col).value = v
        ws.cell(r, 18).value = p["notes"]
        ws.cell(r, 16).value = f'=IF(AND(F{r}<>"",G{r}<>""),G{r}-F{r}+1,"")'
        ws.cell(r, 17).value = f'=IF(AND(G{r}<>"",I{r}<>""),G{r}-I{r},"")'
    ws["A2"] = "SAMPLE – fictional data. Each row is a bar on the roadmap: Pillar > Topic swimlane, coloured by Activity Type."

    # Milestones
    wm = wb["POAP_Milestones"]
    restyle_row_like(wm, 6, 7, 8)
    for i, m in enumerate(milestones(plan)):
        for col, v in enumerate(m, start=1):
            wm.cell(6 + i, col).value = v if v != "" else None

    # P2P
    wp = wb["P2P_Matrix"]
    restyle_row_like(wp, 6, 7, 14)
    wp["A3"] = "SAMPLE – fictional data. HOW TO FILL: System Path uses ' → ' between systems (e.g. Portal → API Gateway → Ledger). Results from the dropdown."
    for i, f in enumerate(P2P):
        for col, v in enumerate(f, start=1):
            wp.cell(6 + i, col).value = v if v != "" else None

    os.makedirs(OUT_DIR, exist_ok=True)
    wb.save(POAP_OUT)
    return len(plan)


# ----------------------------------------------------------------------------
# RAID content
# ----------------------------------------------------------------------------
LIK_T = {1: "Very Low", 2: "Low", 3: "Possible", 4: "Likely", 5: "Very Likely"}
IMP_T = {1: "Negligible", 2: "Minor", 3: "Moderate", 4: "Major", 5: "Severe"}
PRI_T = {1: "Very Low", 2: "Low", 3: "Medium", 4: "High", 5: "Very High"}


def band(score):
    return "Very Low" if score < 5 else "Low" if score < 10 else "Medium" if score < 15 else "High" if score < 20 else "Very High"


# id, title, desc, category, L, prevOcc, I, status, mitigation, owner, review, where, prog, progref
RISKS = [
    ("Test environment instability delays SIT", "Shared test environments suffer unplanned outages during peak execution weeks.", "Environment", 4, "Yes", 4, "Open", "Daily environment health check; agree outage windows; escalate repeated outages to the platform lead.", "Portfolio Test Lead", "2026-10-09", "Weekly test review", "Yes", "PR-0101"),
    ("Test data not representative or delivered late", "Masked data sets arrive after cycle start, limiting scenario coverage.", "Test Data", 3, "Yes", 4, "Open", "Raise data requests two sprints ahead; keep a minimum reusable data baseline.", "Test Manager", "2026-10-16", "Planning workshop", "No", ""),
    ("Vendor drop delivered late", "Vendor build for the next release is not delivered by the agreed date, compressing SIT.", "Supplier", 5, "Yes", 4, "Open", "Weekly vendor checkpoint; agree entry criteria; re-plan SIT window if drop slips beyond 5 days.", "Programme Test Manager", "2026-10-09", "Steering group", "Yes", "PR-0102"),
    ("Loss of key test analyst", "Single points of knowledge for integration scenarios; attrition would slow execution.", "Resourcing", 3, "No", 3, "Open", "Pair working and documented scenario ownership; identify backup analysts.", "Test Manager", "2026-10-23", "Resourcing review", "No", ""),
    ("Late requirements change", "Scope changes after test design baseline force rework of scripts and data.", "Requirements", 4, "Yes", 3, "Open", "Change control with test impact assessment; freeze scope two weeks before cycle start.", "Portfolio Test Lead", "2026-10-16", "Change board", "No", ""),
    ("Regression scope growth", "Each release adds journeys to the regression pack faster than capacity grows.", "Scope", 4, "No", 4, "Open", "Risk-based regression selection; invest in automation for stable journeys.", "Test Manager", "2026-10-30", "Planning workshop", "No", ""),
    ("Defect fix turnaround exceeds target", "Fix turnaround above five working days delays retest and sign-off.", "Defects", 4, "Yes", 3, "Mitigated", "Daily defect triage with delivery leads; fast-track severity 1 and 2 items.", "Test Manager", "2026-10-09", "Defect triage", "No", ""),
    ("Third-party sandbox unavailable", "External sandbox used for notification testing has limited availability.", "Environment", 2, "No", 4, "Open", "Stub fallback agreed; book sandbox slots in advance.", "Portfolio Test Lead", "2026-11-06", "Weekly test review", "No", ""),
    ("Automation framework maintenance backlog", "Framework upgrades are deferred, increasing script flakiness.", "Automation", 2, "No", 2, "Open", "Allocate fixed weekly maintenance time.", "Test Manager", "2026-11-13", "Retrospective", "No", ""),
    ("Christmas freeze compresses test window", "The change freeze removes two weeks of execution capacity.", "Planning", 3, "No", 3, "Closed", "Plan agreed: Cycle 2 completes before the freeze; regression starts after.", "Portfolio Test Lead", "2026-09-18", "Planning workshop", "No", ""),
    ("Performance environment not production-like", "NFR results may not be representative if sizing differs from production.", "Environment", 3, "Yes", 4, "Mitigated", "Sizing comparison documented; results caveated and extrapolated by the platform team.", "Test Manager", "2026-10-23", "NFR working group", "Yes", "PR-0103"),
    ("New joiner access delays", "Account and tooling access takes more than a week for new starters.", "Resourcing", 2, "Yes", 2, "Closed", "Access pre-requested during onboarding.", "Test Manager", "2026-09-04", "Resourcing review", "No", ""),
    ("Unclear ownership of E2E scenarios", "Cross-team journeys lack a named scenario owner, risking gaps.", "Governance", 3, "No", 2, "Open", "Agree scenario ownership matrix at the next test forum.", "Portfolio Test Lead", "2026-10-30", "Test forum", "No", ""),
    ("Test tooling licence expiry", "Licences for the test management tool expire during the programme.", "Tooling", 1, "No", 3, "Closed", "Renewal confirmed to end of programme.", "Programme Test Manager", "2026-09-11", "Weekly test review", "No", ""),
]

# id, title, desc, tracking, reporter, owner, reported, pri, sev, status, origRisk, target, actual, resolution, notes, where, prog, progref
ISSUES = [
    ("Test environment outage blocks SIT execution", "Unplanned outage of the shared integration environment has halted SIT execution for several days.", "TRK-2001", "Test Manager", "Portfolio Test Lead", "2026-09-28", 5, 5, "Open", "R-001", "2026-10-09", None, None, "Environment team investigating; daily call in place.", "Weekly test review", "Yes", "PR-0201"),
    ("Defect fix backlog above target", "Open defects awaiting fix have exceeded the agreed turnaround target.", "TRK-2002", "Test Manager", "Portfolio Test Lead", "2026-09-14", 4, 4, "Resolution in progress", "R-007", "2026-10-16", None, None, "Extra delivery capacity agreed; backlog reducing weekly.", "Defect triage", "No", ""),
    ("Test data load failure in Cycle 1", "Masked data load failed validation and had to be re-run.", "TRK-2003", "Test Manager", "Test Manager", "2026-09-07", 3, 3, "Resolved", "R-002", "2026-09-18", "2026-09-16", "Data load corrected and re-run successfully.", "Closed after re-run.", "Weekly test review", "No", ""),
    ("Interface contract mismatch between Portal and API Gateway", "Field length mismatch caused submission failures in SIT.", "TRK-2004", "Test Manager", "Portfolio Test Lead", "2026-09-10", 4, 3, "Resolved", "", "2026-09-25", "2026-09-23", "Contract aligned and fix deployed.", "Regression test added.", "Defect triage", "No", ""),
    ("Notification Service sandbox intermittent", "Sandbox responses are intermittent, producing false failures.", "TRK-2005", "Test Manager", "Portfolio Test Lead", "2026-09-21", 3, 3, "Resolution in progress", "R-008", "2026-10-23", None, None, "Stub fallback in use while the sandbox is stabilised.", "Weekly test review", "No", ""),
]

# id, desc, raisedBy, logged, confidence, impact, validation, due, status, notes, where, prog, progref
ASSUMPTIONS = [
    ("Production-like test data will be available before each SIT cycle starts.", "Test Manager", "2026-07-06", "Medium", "Reduced coverage and delayed execution.", "Confirm data refresh schedule with the data team.", "2026-10-16", "Unconfirmed", "", "Planning workshop", "No", ""),
    ("The vendor will deliver builds on the agreed fortnightly cadence.", "Programme Test Manager", "2026-07-06", "Low", "SIT window compressed; sign-off dates at risk.", "Weekly vendor checkpoint.", "2026-10-09", "Unconfirmed", "", "Steering group", "Yes", "PR-0301"),
    ("Test environments are available 24x7 except for agreed maintenance windows.", "Portfolio Test Lead", "2026-07-13", "High", "Execution blocked; schedule impact.", "Review environment availability report monthly.", "2026-11-06", "Confirmed Correct", "Availability report reviewed.", "Planning workshop", "No", ""),
    ("Regression pack scope is limited to in-scope journeys agreed at baseline.", "Test Manager", "2026-09-04", "Medium", "Additional regression effort required.", "Confirm scope with the product owner.", "2026-10-30", "Unconfirmed", "", "Weekly test review", "No", ""),
    ("No change to non-functional requirements after baseline.", "Test Manager", "2026-09-14", "High", "NFR re-planning and rework.", "Sign-off of NFR catalogue.", "2026-09-25", "Confirmed Correct", "NFR catalogue signed off.", "NFR working group", "No", ""),
    ("Existing automation covers at least 60% of regression journeys.", "Test Manager", "2026-09-21", "Medium", "Manual regression effort increases.", "Automation coverage review.", "2026-11-13", "Unconfirmed", "", "Retrospective", "No", ""),
]

# id, desc, for, from, type, required, requestor, owner, priority, impact, status, notes, where, prog, progref
DEPS = [
    ("Vendor build 3.2 for Release 3 SIT entry", "SIT team", "Vendor delivery team", "Delivery", "2026-10-09", "Test Manager", "Programme Test Manager", 5, "SIT start delayed by the same duration.", "Open", "Checkpoint twice weekly.", "Steering group", "Yes", "PR-0401"),
    ("Refreshed test environment for Release 2", "SIT team", "Environment team", "Environment", "2026-10-16", "Test Manager", "Portfolio Test Lead", 4, "Cycle 1 completion slips.", "Open", "Refresh scheduled for the weekend.", "Weekly test review", "No", ""),
    ("Masked data set for Digital Filing", "Test design", "Data team", "Data", "2026-10-23", "Test Manager", "Test Manager", 4, "Scenarios cannot be executed.", "Open", "", "Planning workshop", "No", ""),
    ("Notification Service sandbox slots for Release 3", "Smoke / P2P", "Third-party integration team", "Third Party", "2026-10-28", "Test Manager", "Portfolio Test Lead", 3, "Fallback to stubs reduces confidence.", "Open", "Slots to be booked.", "Weekly test review", "No", ""),
    ("Approved NFR catalogue", "NFR testing", "Architecture", "Requirements", "2026-11-20", "Portfolio Test Lead", "Portfolio Test Lead", 3, "NFR test design delayed.", "Open", "", "NFR working group", "No", ""),
    ("Two additional test analysts onboarded", "Test execution", "Resourcing", "Resource", "2026-12-04", "Test Manager", "Programme Test Manager", 4, "Execution capacity shortfall in Cycle 2.", "Open", "", "Resourcing review", "No", ""),
    ("Test management tool upgrade", "All test teams", "Tooling team", "Tooling", "2026-09-18", "Test Manager", "Portfolio Test Lead", 3, "Reporting gaps.", "Closed", "Upgrade completed.", "Weekly test review", "No", ""),
    ("Interface contract baseline for Portal and API Gateway", "Test design", "Solution design", "Requirements", "2026-08-14", "Test Manager", "Portfolio Test Lead", 5, "Test design cannot be baselined.", "Closed", "Contract baselined.", "Planning workshop", "No", ""),
]
DEP_TYPES = ["Delivery", "Environment", "Data", "Third Party", "Requirements", "Resource", "Tooling"]

RISK_H = ["ID", "Summary Title", "Description", "Category", "Likelihood", "Likelihood Rating", "Previous Occurrence", "Impact", "Impact Rating", "Risk Score", "Status", "Mitigation Details / Narrative", "Owner", "Review Date", "Where Raised", "Prog RAID", "Prog RAID Reference", "Archived"]
ISSUE_H = ["ID", "Summary Title", "Full Description", "Tracking ID Reference (Link)", "Reporter", "Owner", "Date Reported", "Priority", "Priority Rating", "Severity", "Severity Rating", "Overall Issue Rating", "Status", "Original Risk ID", "Target Resolution Date", "Actual Resolution Date", "Resolution Summary", "Notes/Actions", "Where Raised", "Prog RAID", "Prog RAID Reference", "Archived"]
ASSUMP_H = ["ID", "Description", "Raised By", "Date Logged", "Confidence Level", "Impact if Assumption is Incorrect", "Validation Action / Narrative", "Validation Due Date", "Status", "Notes/Actions", "Where Raised", "Prog RAID", "Prog RAID Reference", "Archived"]
DEP_H = ["ID", "Description", "Dependency For (Who Needs This)", "Dependency From (Who Delivers This)", "Type", "Date Required", "Requestor", "Owner", "Priority", "Impact If Not Met", "Status", "Notes/Actions", "Where Raised", "Prog RAID", "Prog RAID Reference", "Archived"]

THIN = Side(style="thin", color="D0D5DD")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor=NAVY)
IN_FILL = PatternFill("solid", fgColor="FFF6CC")
F_HDR = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_IN = Font(name=FONT, size=10, color="0000FF")
F_TITLE = Font(name=FONT, size=14, bold=True, color=NAVY)
F_SUB = Font(name=FONT, size=10, color="6B7280")
F_NOTE = Font(name=FONT, size=10, bold=True, color="B54708")
F_BOLD = Font(name=FONT, size=10, bold=True, color=NAVY)
F_BASE = Font(name=FONT, size=10)
DATE_FMT = "yyyy-mm-dd"
FIRST_COL, HDR_ROW, FIRST_ROW, PREFORMAT = 3, 10, 11, 60


def d(s):
    return D(s) if s else None


def write_table(ws, title, hint, headers, widths, rows, date_cols, dvs):
    ws.sheet_view.showGridLines = False
    ws["C2"], ws["C2"].font = title, F_TITLE
    ws["C3"], ws["C3"].font = hint, F_SUB
    ws["C5"], ws["C5"].font = "SAMPLE – fictional data for demonstration only. Yellow cells are inputs.", F_NOTE
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 2
    for i, (h, w) in enumerate(zip(headers, widths)):
        col = FIRST_COL + i
        c = ws.cell(HDR_ROW, col, h)
        c.font, c.fill, c.border = F_HDR, HDR_FILL, BORDER
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = w
    ws.row_dimensions[HDR_ROW].height = 32
    for r in range(FIRST_ROW, FIRST_ROW + max(PREFORMAT, len(rows))):
        for i in range(len(headers)):
            c = ws.cell(r, FIRST_COL + i)
            c.font, c.fill, c.border = F_IN, IN_FILL, BORDER
            c.alignment = Alignment(vertical="top", wrap_text=True)
            if i in date_cols:
                c.number_format = DATE_FMT
    for r, row in enumerate(rows, start=FIRST_ROW):
        for i, v in enumerate(row):
            ws.cell(r, FIRST_COL + i).value = v
    last = FIRST_ROW + 199
    for header, items in dvs.items():
        col = openpyxl.utils.get_column_letter(FIRST_COL + headers.index(header))
        dv = DataValidation(type="list", formula1='"' + ",".join(items) + '"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"{col}{FIRST_ROW}:{col}{last}")
    ws.freeze_panes = ws.cell(FIRST_ROW, FIRST_COL + 1)
    ws.auto_filter.ref = f"C{HDR_ROW}:{openpyxl.utils.get_column_letter(FIRST_COL + len(headers) - 1)}{FIRST_ROW + len(rows) - 1}"


def build_raid():
    wb = openpyxl.Workbook()
    s = wb.active
    s.title = "Summary Sheet"
    s.sheet_view.showGridLines = False
    s.column_dimensions["A"].width = 2
    s.column_dimensions["B"].width = 2
    s.column_dimensions["C"].width = 30
    s.column_dimensions["D"].width = 4
    s.column_dimensions["E"].width = 34
    s.column_dimensions["F"].width = 10
    s.column_dimensions["G"].width = 4
    s.column_dimensions["H"].width = 30
    s["C2"], s["C2"].font = "RAID Log – Summary", F_TITLE
    s["C3"], s["C3"].font = "SAMPLE – fictional data for demonstration only.", F_NOTE
    info = [("Project Name", "Self Assessment"), ("TSR ID", "TSR_SAMPLE-000001"), ("Programme Test Manager", "TBC"),
            ("Start Date", D("2026-07-01")), ("End Date", D("2027-06-30")), ("Version", 1)]
    for i, (k, v) in enumerate(info):
        r = 5 + i
        s.cell(r, 3, k).font = F_BOLD
        c = s.cell(r, 5, v)
        c.font, c.fill, c.border = F_IN, IN_FILL, BORDER
        c.alignment = Alignment(horizontal="left")
        if isinstance(v, dt.date):
            c.number_format = DATE_FMT
    s["C13"], s["C13"].font = "Rating legend (likelihood x impact)", F_BOLD
    for i, h in enumerate(["Rating", "Score range"]):
        c = s.cell(14, 3 + 2 * i, h)
        c.font, c.fill = F_HDR, HDR_FILL
    legend = [("Very Low", "1 – 4", "9AD8B0"), ("Low", "5 – 9", "C9E58F"), ("Medium", "10 – 14", "FFD66B"),
              ("High", "15 – 19", "FF9D5C"), ("Very High", "20+", "EF5B5B")]
    for i, (a, b, col) in enumerate(legend):
        c = s.cell(15 + i, 3, a)
        c.font, c.fill, c.border = F_BASE, PatternFill("solid", fgColor=col), BORDER
        s.cell(15 + i, 5, b).font = F_BASE

    risk_rows = []
    for n, r in enumerate(RISKS, start=1):
        (title, desc, cat, l, prev, i, status, mit, owner, review, where, prog, pref) = r
        risk_rows.append([f"R-{n:03d}", title, desc, cat, LIK_T[l], l, prev, IMP_T[i], i, band(l * i), status, mit, owner, d(review), where, prog, pref or None, "No"])
    issue_rows = []
    for n, r in enumerate(ISSUES, start=1):
        (title, desc, trk, rep, own, reported, p, sv, status, orig, target, actual, res, notes, where, prog, pref) = r
        issue_rows.append([f"I-{n:03d}", title, desc, trk, rep, own, d(reported), PRI_T[p], p, IMP_T[sv], sv, band(p * sv), status, orig or None, d(target), d(actual), res, notes or None, where, prog, pref or None, "No"])
    assump_rows = []
    for n, r in enumerate(ASSUMPTIONS, start=1):
        (desc, by, logged, conf, imp, val, due, status, notes, where, prog, pref) = r
        assump_rows.append([f"A-{n:03d}", desc, by, d(logged), conf, imp, val, d(due), status, notes or None, where, prog, pref or None, "No"])
    dep_rows = []
    for n, r in enumerate(DEPS, start=1):
        (desc, forw, frm, typ, req, requestor, owner, p, impact, status, notes, where, prog, pref) = r
        dep_rows.append([f"D-{n:03d}", desc, forw, frm, typ, d(req), requestor, owner, PRI_T[p], impact, status, notes or None, where, prog, pref or None, "No"])

    yn = ["Yes", "No"]
    sheets = [
        ("Risks", "Risk Register", "One row per risk. Likelihood x Impact gives the score band used by the dashboard.", RISK_H,
         [9, 34, 48, 14, 13, 11, 12, 12, 10, 12, 12, 48, 22, 13, 20, 9, 14, 10], risk_rows, {5, 8, 13},
         {"Likelihood": list(LIK_T.values()), "Impact": list(IMP_T.values()), "Status": ["Open", "Mitigated", "Closed"], "Previous Occurrence": yn, "Prog RAID": yn, "Archived": yn}),
        ("Issues", "Issue Log", "One row per issue. Priority x Severity gives the overall rating.", ISSUE_H,
         [9, 34, 48, 14, 20, 20, 13, 11, 10, 11, 10, 12, 20, 11, 13, 13, 36, 36, 20, 9, 14, 10], issue_rows, {6, 14, 15},
         {"Priority": list(PRI_T.values()), "Severity": list(IMP_T.values()), "Status": ["Open", "Resolution in progress", "Resolved", "Closed"], "Prog RAID": yn, "Archived": yn}),
        ("Assumptions", "Assumptions Log", "One row per assumption with a validation action and due date.", ASSUMP_H,
         [9, 50, 22, 13, 12, 36, 36, 13, 18, 30, 20, 9, 14, 10], assump_rows, {3, 7},
         {"Confidence Level": ["High", "Medium", "Low"], "Status": ["Unconfirmed", "Confirmed Correct", "Confirmed Incorrect"], "Prog RAID": yn, "Archived": yn}),
        ("Dependencies", "Dependency Log", "One row per dependency; the dashboard lists open items due in the next 30 days.", DEP_H,
         [9, 46, 22, 26, 14, 13, 20, 22, 11, 36, 10, 30, 20, 9, 14, 10], dep_rows, {5},
         {"Type": DEP_TYPES, "Priority": list(PRI_T.values()), "Status": ["Open", "Closed"], "Prog RAID": yn, "Archived": yn}),
    ]
    for name, title, hint, headers, widths, rows, dcols, dvs in sheets:
        ws = wb.create_sheet(name)
        write_table(ws, title, hint, headers, widths, rows, dcols, dvs)

    g = wb.create_sheet("User Guide")
    g.sheet_view.showGridLines = False
    g.column_dimensions["B"].width = 26
    g.column_dimensions["C"].width = 100
    g["B2"], g["B2"].font = "RAID Log – User Guide", F_TITLE
    g["B3"], g["B3"].font = "SAMPLE – fictional data for demonstration only.", F_NOTE
    lines = [
        ("Summary Sheet", "Project details in column E (Project Name, TSR ID, Programme Test Manager, Start Date, End Date, Version) and the rating legend."),
        ("Risks / Issues / Assumptions / Dependencies", "Header row is row 10, data from row 11, ID in column C. Keep header text exactly as written - the dashboard finds columns by name."),
        ("Ratings", "Risk score = Likelihood rating x Impact rating: <5 Very Low, <10 Low, <15 Medium, <20 High, otherwise Very High. Issues use Priority x Severity."),
        ("Status", "Risks: Open / Mitigated / Closed. Issues: Open / Resolution in progress / Resolved / Closed. Dependencies: Open / Closed. Archived = Yes hides a row from the dashboard."),
        ("Colours", "Yellow cells with blue text are inputs. Use the dropdowns where provided."),
    ]
    for i, (a, b) in enumerate(lines):
        g.cell(5 + i, 2, a).font = F_BOLD
        c = g.cell(5 + i, 3, b)
        c.font = F_BASE
        c.alignment = Alignment(wrap_text=True, vertical="top")
    wb.save(RAID_OUT)
    return len(risk_rows), len(issue_rows), len(dep_rows), len(assump_rows)


def main():
    n = build_poap()
    print("POAP bars:", n, "->", POAP_OUT)
    print("RAID (risks, issues, deps, assumptions):", build_raid(), "->", RAID_OUT)
    rc = os.environ.get("RECALC_SCRIPT")
    if rc:
        subprocess.run([sys.executable, rc, POAP_OUT, "90"], check=False)
    else:
        print("Tip: set RECALC_SCRIPT to recalculate the POAP formulas with LibreOffice.")


if __name__ == "__main__":
    main()
