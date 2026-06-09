# =============================================================================
# FILE: ops_dashboard.py
# PROJECT: eVidyaloka Operations & Impact Command Center
#
# WHAT THIS FILE DOES:
#   This file builds a multi-tab, interactive web dashboard for eVidyaloka,
#   an NGO that places volunteers to teach in rural schools across India.
#   It reads two Excel workbooks (VRM and DRM), cleans and transforms the data,
#   then renders charts, KPI cards, tables, and heatmaps inside a Streamlit app.
#
# HOW IT IS STRUCTURED (top to bottom):
#   1. IMPORTS & CONSTANTS  — libraries, file paths, colour palettes,
#                             and lookup dictionaries for data normalisation.
#   2. DATA LOADING         — two cached functions that read Excel sheets and
#                             return clean DataFrames ready for analysis.
#   3. HELPER FUNCTIONS     — reusable utilities for chart styling and HTML cards.
#   4. MAIN RENDER FUNCTION — the single entry point that builds the entire UI:
#                             sidebar filters → 4 tabs → charts → tables.
#
# KEY CONCEPTS FOR BEGINNERS:
#   • Streamlit  — a Python library that turns a plain .py script into a web app.
#                  Every `st.*` call you see renders something in the browser.
#   • pandas     — Python's spreadsheet library. A "DataFrame" is like an
#                  in-memory Excel table you can filter, group, and compute on.
#   • Plotly     — a charting library. `px` (plotly.express) gives one-liner
#                  charts; `go` (plotly.graph_objects) gives lower-level control.
#   • @st.cache_data — a decorator that stores the result of an expensive
#                  function call (like reading an Excel file) so it isn't
#                  re-run every time the user interacts with the page.
#
# OVERARCHING LOGIC:
#   load data → apply sidebar filters → for each tab: compute metrics → draw charts
# =============================================================================

import streamlit as st   # The web-app framework — renders every visual element
import pandas as pd       # Data manipulation: reading Excel, filtering, grouping
import plotly.express as px          # High-level charting (most charts here)
import plotly.graph_objects as go    # Low-level charting (used for multi-trace figures)
import os                            # Used to check whether a file exists on disk


# =============================================================================
# SECTION 1 — FILE PATHS & SHEET NAMES
#
# WHY: We define these as module-level constants (UPPERCASE by convention) so
# they are easy to update in one place if the filenames ever change.  Any
# function below that needs a file path imports it from here rather than
# hard-coding the string inside the function.
# =============================================================================

DATA_PATH     = "VRM_May__2026.xlsx"   # Volunteer Relationship Management workbook
DRM_DATA_PATH = "DRM_May_2026.xlsx"    # Donor Relationship Management workbook

# Names of the three sheets we read from the VRM workbook.
# Note the trailing space in SHEET_NEW_REG — this matches the actual tab name in
# the Excel file exactly (a common real-world data quirk).
SHEET_ACTIVE  = "Active VT"
SHEET_DROPPED = "Dropped VT"
SHEET_NEW_REG = "Newly Registered VT "   # ← trailing space is intentional!


# =============================================================================
# SECTION 2 — COLOUR PALETTES
#
# WHY: Defining colours once here means every chart automatically shares the
# same brand identity.  If a designer asks us to change "teal" we update it
# once and every chart updates automatically.
#
# HOW: Python dictionaries (key → value pairs) map colour names to hex codes.
# Think of a hex code like "#0094c9" as a precise paint-chip reference: it
# tells the browser exactly what shade of blue-green to display.
# =============================================================================

# Primary palette — used across the VRM tabs (Volunteers, Centres, Academic)
P = {
    "teal":    "#0094c9",   # Bright blue-green — primary accent
    "green":   "#00964d",   # Forest green — positive/success metrics
    "orange":  "#f27c48",   # Warm orange — general accent
    "red":     "#ed1c2d",   # Alert red — cancellations / at-risk
    "violet":  "#6c5ce7",   # Purple — CLH and learning metrics
    "amber":   "#fdcb6e",   # Golden yellow — offline / partial completion
    "sky":     "#74b9ff",   # Light blue — secondary bars
    "mint":    "#00b894",   # Soft green — completion rates
    "coral":   "#e17055",   # Salmon-orange — female/girl metrics, dropped vols
    "lavender":"#a29bfe",   # Soft violet — enrolled student charts
    "salmon":  "#fab1a0",   # Light salmon — gradient starts
    "aqua":    "#55efc4",   # Bright aqua — accent
}

# SEQ is a flat list of all colour values, used wherever Plotly needs a
# sequence of colours to assign automatically (e.g. one colour per category).
SEQ = list(P.values())

# BG = fully transparent background — we want charts to inherit the page colour
# rather than showing a white or grey box behind them.
BG   = "rgba(0,0,0,0)"

# GRID = a very light grey used for horizontal grid lines on charts,
# subtle enough not to distract from the data itself.
GRID = "#e9ecef"

# Secondary palette D — richer, slightly darker tones used exclusively in the
# premium DRM Client Report tab so it has a visually distinct, more formal look.
D = {
    "teal":   "#0f8a6e",   # Deeper teal
    "blue":   "#185fa5",   # Navy blue — Microsoft donor colour
    "amber":  "#ba7517",   # Deep amber — offline / warnings
    "coral":  "#993c1d",   # Dark red-brown — cancellations / alerts
    "purple": "#534ab7",   # Deep purple — CLH on client tab
    "green":  "#3b6d11",   # Deep green — positive metrics
    "red":    "#a32d2d",   # Dark red — at-risk indicators
    "gray":   "#5f5e5a",   # Neutral grey — supporting text / unknown donors
}


# =============================================================================
# SECTION 3 — DATA NORMALISATION LOOKUP TABLES
#
# WHY: Real-world data is messy.  Volunteers enter their profession in free text,
# so we might see "Housewife", "house wife", "Home maker", and "Home Maker" all
# meaning the same thing.  These dictionaries ("maps") act as a translation
# table: raw messy value → clean standardised label.
#
# HOW: Python's dict.map() method replaces every value in a DataFrame column
# with its corresponding clean label.  If no match is found, the original
# value is kept (handled by .fillna() later in the loading functions).
# =============================================================================

# SUBJECT_MAP: maps the verbose subject names used in the Excel file to the
# short, chart-friendly labels we display in the dashboard.
# Example: "Conceptual Learning - Math" and "Maths - Worksheet" both → "Math"
SUBJECT_MAP = {
    "Conceptual Learning - Math":     "Math",
    "Conceptual Learning Math-HM":    "Math",
    "Maths - Worksheet":              "Math",
    "Conceptual Learning - Science":  "Science",
    "Conceptual Learning Science-HM": "Science",
    "English":                        "English",
    "English - Worksheet":            "English",
    "Spoken English: Level 1":        "Spoken English",
    "Spoken English: Level 2":        "Spoken English",
    "Basic Digital Literacy":         "Digital Literacy",
    "Artificial Intelligence (AI)":   "AI",
    "Explore Coding":                 "Coding",
    "Concise content 1":              "Concise Content",
    "Concise Content 1":              "Concise Content",   # capitalisation variant
    "Concise content 2":              "Concise Content",
    "Guest Sessions":                 "Guest Sessions",
    "Scholarship":                    "Scholarship",
    "Reading Program":                "Reading Program",
}

# PROFESSION_MAP: normalises the 30+ free-text profession entries that
# volunteers type into registration forms into 10 clean categories.
# This is used in Tab 1 to draw the "Volunteers by Profession" pie chart.
PROFESSION_MAP = {
    "corporates":                              "Corporate Professional",
    "others":                                  "Others",
    "student_ug":                              "Student (UG)",
    "home_makers":                             "Home Maker",
    "Housewife":                               "Home Maker",   # variant spellings
    "house wife":                              "Home Maker",
    "Home Maker":                              "Home Maker",
    "Home maker":                              "Home Maker",
    "teacher_school":                          "School Teacher",
    "teacher_university":                      "University Teacher",
    "Teaching":                                "School Teacher",
    "teaching":                                "School Teacher",
    "Educator":                                "School Teacher",
    "Principal of a Primary school":           "School Teacher",
    "retired teacher":                         "Retired Professional",
    "student_pg":                              "Student (PG)",
    "Student":                                 "Student (UG)",
    "Student - PG":                            "Student (PG)",
    "business/self_employed":                  "Business / Self-Employed",
    "Self-employed":                           "Business / Self-Employed",
    "self employed":                           "Business / Self-Employed",
    "Service":                                 "Business / Self-Employed",
    "retired_professional":                    "Retired Professional",
    "Retd. Professional":                      "Retired Professional",
    "top_managemenet":                         "Management",   # typo in source data
    "Management":                              "Management",
    "finance":                                 "Finance",
    "legal":                                   "Legal",
    "Legal":                                   "Legal",
    "Education, Training, and Library":        "School Teacher",
    "Computer and Mathematical":               "Corporate Professional",
    "Business and Financial Operations":       "Finance",
    "Healthcare Practitioners and Technical":  "Healthcare",
    "Medical Professional":                    "Healthcare",
    "Life, Physical, and Social Science":      "Others",
    "Office and Administrative Support":       "Others",
    "Production/Manufacturing":                "Others",
    "Sales and Related":                       "Business / Self-Employed",
    "Community and Social Service":            "Others",
    "Farming, Fishing, and Forestry":          "Others",
    "Trainer, manufacturers":                  "Others",
}


# =============================================================================
# SECTION 4 — REFERENCE / ACQUISITION CHANNEL BUCKETING FUNCTION
#
# WHY: The "Reference" column (how a volunteer heard about eVidyaloka) contains
# hundreds of unique free-text strings.  We need to collapse them into a
# handful of meaningful categories to draw a readable bar chart.
#
# HOW: This function accepts a single string and uses Python's `in` operator
# and string containment tests to decide which bucket it belongs to.
# It is applied to every row via .apply(_group_ref) in the data loading step.
#
# Think of it like a filing clerk: given a note that says "I heard about this
# from my colleague at Infosys", the clerk places it in the
# "Corporate / Partner Referral" folder.
# =============================================================================
def _group_ref(ref: str) -> str:
    # Convert to string (handles NaN/None gracefully) and remove surrounding whitespace
    r = str(ref).strip()

    # Exact matches for common digital / direct channels
    if r in ("Internet Search", "Facebook", "Community Outreach"):
        return "Online / Direct"

    # eVidyaloka's own internal campaigns and mailing lists
    if r in ("Emailer", "DCP Campaign", "eVidyaloka", "eVidyaloka Trust"):
        return "eVidyaloka Campaign"

    # Word of mouth — kept as its own category because it's a strong signal
    if r in ("Word of Mouth",):
        return "Word of Mouth"

    # Check if any known corporate name appears anywhere in the reference string.
    # `any(k in r for k in (...))` returns True if at least one keyword matches.
    # This handles cases like "Referred by EY colleague" or "Through Microsoft CSR".
    if any(k in r for k in (
        "KPMG", "Infosys", "Tech Mahindra", "HPInc", "HPE", "L&T",
        "HSBC", "EY", "Adobe", "Brillio", "Broadridge", "Accenture",
        "CISCO", "Fidelity", "ConnectFor", "Microsoft", "Cognizant",
        "Firstsource", "Pricewaterhousecoopers", "Reliance Foundation",
        "Scaler", "Joy of Reading",
    )):
        return "Corporate / Partner Referral"

    # Academic institutions — colleges, universities, IITs, IIMs, NITs
    if any(k in r for k in (
        "College", "University", "BMS", "NIT", "IIT", "IIM",
    )):
        return "Academic Institution"

    # NGOs and community organisations that partner with eVidyaloka
    if any(k in r for k in (
        "Bhumi", "Udaan", "Swabhiman", "Foundation", "NGO",
    )):
        return "NGO / Community Org"

    # Anything that doesn't match any rule above falls into "Other"
    return "Other"


# =============================================================================
# SECTION 5 — VRM DATA LOADING FUNCTION
#
# WHY: Reading and cleaning Excel data is expensive (can take several seconds).
# We define this as a separate function decorated with @st.cache_data so
# Streamlit only runs it once and then stores ("caches") the result in memory.
# Every subsequent page interaction reuses the cached result instantly.
#
# HOW: The function reads three sheets from the VRM workbook, renames columns
# to shorter internal names, fixes encoding bugs, coerces data types, and
# applies our normalisation maps above.  It returns a dict (dictionary —
# like a labelled container) so the caller can access each sheet by name.
#
# ANALOGY: Think of this function as a data-prep kitchen.  Raw ingredients
# (messy Excel rows) go in; clean, labelled dishes come out on a tray (dict).
# The @st.cache_data decorator is like a warming cabinet — once prepared,
# dishes sit ready to serve instantly without re-cooking.
# =============================================================================
@st.cache_data(show_spinner=False)
def load_data(path: str) -> dict:
    """
    Reads the VRM Excel workbook and returns three cleaned DataFrames:
      active   — Active VT sheet (one row per volunteer-offering pair)
      dropped  — Dropped VT sheet (volunteers who left their offering)
      new_reg  — Newly Registered VT sheet (all new signups this period)
    """

    # Guard clause: if the file doesn't exist, return an empty dict.
    # The calling code checks for this and shows a friendly error message
    # to the user instead of crashing with a Python exception.
    if not os.path.exists(path):
        return {}

    # ── STEP 1: Read the "Active VT" sheet ───────────────────────────────────
    # pd.read_excel loads the sheet into a DataFrame — picture a spreadsheet
    # stored in Python memory with rows and named columns.
    active = pd.read_excel(path, sheet_name=SHEET_ACTIVE)

    # Strip leading/trailing whitespace from every column header.
    # Excel files often have hidden spaces in column names ("Name " vs "Name")
    # that would cause column lookups to fail silently.
    active.columns = active.columns.str.strip()

    # ── STEP 2: Rename columns to shorter internal names ─────────────────────
    # The Excel headers are long and verbose ("Volunteer ID", "Center State").
    # We rename them to the short names used throughout the rest of the dashboard
    # so chart code stays concise.  inplace=True modifies the DataFrame directly.
    active.rename(columns={
        "Volunteer ID":               "Volunteer id",
        "Volunteer Name":             "Volunteer name",
        "EV Joined Date":             "Joined(ev)",
        "Volunteer State":            "Residence state",
        "Volunteer City":             "Residence city",
        "Center Name":                "Center name",
        "Center State":               "State",
        "Offering Status":            "Offering status",
        "Total Hours (Comp + Offline)":"Total hours(Comp+Offline)",
        "Attendance %":               "Attendance%",
    }, inplace=True)

    # ── STEP 3: Fix Unicode encoding artefacts in State columns ──────────────
    # Some Excel exports produce "mangled" characters when non-breaking spaces
    # (\u00a0) or other Unicode control characters appear in cell values.
    # For example, "Uttar Pradesh" might come through as "Uttar¬\u00a0Pradesh".
    # We replace those garbled character sequences with a plain space, then
    # collapse any double-spaces and trim the edges.
    for col in ["State", "Residence state"]:
        if col in active.columns:   # Only process the column if it actually exists
            active[col] = (
                active[col].astype(str)                             # Ensure everything is a string
                .str.replace("\u00ac\u00a0", " ", regex=False)      # Replace the specific garbled sequence
                .str.replace("\u00a0", " ", regex=False)            # Replace standalone non-breaking space
                .str.replace(r"\s+", " ", regex=True)               # Collapse multiple spaces into one
                .str.strip()                                         # Remove leading/trailing whitespace
            )

    # ── STEP 4: Convert "Attendance%" to a floating-point number ─────────────
    # Excel sometimes stores percentages as strings like "84.5%".
    # pd.to_numeric(..., errors="coerce") converts what it can, and turns
    # anything that isn't a number (like the "%" sign itself or empty cells)
    # into NaN ("Not a Number") instead of crashing.  This is safe because
    # chart functions handle NaN gracefully by skipping those values.
    active["Attendance%"] = pd.to_numeric(active["Attendance%"], errors="coerce")

    # Convert the "Joined(ev)" date column from text to a proper datetime object.
    # errors="coerce" turns any unparseable date string into NaT (Not a Time),
    # the datetime equivalent of NaN.
    active["Joined(ev)"] = pd.to_datetime(active["Joined(ev)"], errors="coerce")

    # ── STEP 5: Apply normalisation maps to create new "clean" columns ────────
    # .map(SUBJECT_MAP) replaces each value in the "Subject" column using the
    # lookup dictionary defined at the top of the file.
    # .fillna(active["Subject"]) means: if a subject wasn't in the map, keep
    # the original raw value (so nothing is lost — it just won't be renamed).
    # We store the result in a new column "Subject_clean" so the original
    # raw data is preserved alongside the cleaned version.
    active["Subject_clean"]    = active["Subject"].map(SUBJECT_MAP).fillna(active["Subject"])
    active["Profession_clean"] = active["Profession"].map(PROFESSION_MAP).fillna("Others")

    # Apply the reference-bucketing function to every row of the "Reference" column.
    # .apply(fn) calls the function once per row, passing the cell value.
    # The result is a new column "Ref_group" with the bucketed category.
    active["Ref_group"] = active["Reference"].apply(_group_ref)

    # ── STEP 6: Fill remaining missing text values with "Unknown" ─────────────
    # For categorical text columns, NaN would break groupby/filter operations.
    # We replace NaN with the string "Unknown" as a safe placeholder.
    for col in ["Donor", "State", "Subject_clean", "Center name",
                "Residence state", "Residence city", "Ref_group"]:
        if col in active.columns:
            active[col] = active[col].fillna("Unknown")

    # ── STEP 7: Convert count columns to integers ─────────────────────────────
    # Excel numeric columns sometimes arrive as floats (e.g., 42.0) or strings.
    # We coerce everything to numeric, replace any NaN with 0, then cast to int
    # so KPI sums like "total enrolled" display as clean whole numbers.
    for col in ["Registered", "Reg Boys", "Reg Girls",
                "Enrolled",   "En Boys",  "En Girls",
                "Planned",    "Scheduled", "Completed",
                "Offline",    "Cancelled",
                "Total hours(Comp+Offline)", "CLH"]:
        if col in active.columns:
            active[col] = pd.to_numeric(active[col], errors="coerce").fillna(0).astype(int)

    # ── STEP 8: Load the "Dropped VT" sheet ──────────────────────────────────
    # This sheet lists volunteers who exited their offering.  We only need it
    # for a count and a reason-breakdown chart, so minimal cleaning is needed.
    dropped = pd.read_excel(path, sheet_name=SHEET_DROPPED)
    dropped.columns = dropped.columns.str.strip()
    # Apply the same Unicode fix to the State column if it exists
    if "State" in dropped.columns:
        dropped["State"] = (
            dropped["State"].astype(str)
            .str.replace("\u00ac\u00a0", " ", regex=False)
            .str.replace("\u00a0", " ", regex=False)
            .str.strip()
        )

    # ── STEP 9: Load the "Newly Registered VT" sheet ─────────────────────────
    # This sheet records everyone who signed up during the reporting period,
    # whether they are currently active or not.
    new_reg = pd.read_excel(path, sheet_name=SHEET_NEW_REG)
    new_reg.columns = new_reg.columns.str.strip()

    # Parse the join date so we can draw a monthly trend line in Tab 1.
    new_reg["Date Joined"] = pd.to_datetime(new_reg["Date Joined"], errors="coerce")

    # Bucket the reference channel for use in the gender-by-channel chart.
    new_reg["Ref_group"] = new_reg["Reference"].apply(_group_ref)

    # Standardise gender values: strip whitespace and title-case
    # so "male", "MALE", " Male " all become "Male".
    if "Gender" in new_reg.columns:
        new_reg["Gender"] = new_reg["Gender"].astype(str).str.strip().str.title()

    # ── STEP 10: Return all three cleaned DataFrames in a single dictionary ───
    # The caller uses this dict like:  data["active"], data["dropped"], etc.
    return {"active": active, "dropped": dropped, "new_reg": new_reg}


# =============================================================================
# SECTION 6 — DRM-SPECIFIC SUBJECT NORMALISATION
#
# WHY: The DRM workbook uses a slightly different (smaller) set of subject names
# than the VRM workbook.  We define a separate map here rather than reusing
# SUBJECT_MAP, because the DRM sheet doesn't have worksheet variants and
# additional subjects like "AI" or "Coding" that appear in VRM data.
#
# This keeps each workbook's cleaning logic self-contained and avoids
# accidentally mapping subjects that don't exist in the DRM data.
# =============================================================================
SUBJECT_NORM_DRM = {
    "Conceptual Learning - Math":     "Math",
    "Conceptual Learning Math-HM":    "Math",
    "Conceptual Learning - Science":  "Science",
    "Conceptual Learning Science-HM": "Science",
    "Concise Content 1":              "Concise Content",
    "Concise content 1":              "Concise Content",   # lowercase variant
    "English":                        "English",
    "Spoken English: Level 1":        "Spoken English",
    "Guest Sessions":                 "Guest Sessions",
}


# =============================================================================
# SECTION 7 — DRM DATA LOADING FUNCTION
#
# WHY: The DRM workbook has a completely different structure from the VRM one —
# it contains session-level logs, centre summaries, offering details, and
# enrolment records.  We load it in its own cached function so its logic
# is isolated and so it doesn't re-run unless the file path changes.
#
# HOW: We load five sheets, apply the same encoding-fix and type-coercion
# patterns used in load_data(), and additionally COMPUTE derived metrics
# (Completion%, Dropout%, Girl%) from the raw columns, because these
# calculated columns are needed in multiple chart sections.
# =============================================================================
@st.cache_data(show_spinner=False)
def load_drm_data(path: str) -> dict:
    """
    Reads the DRM Excel workbook and returns five cleaned DataFrames:
      sess — SESSION DUMP: one row per individual teaching session
      drm  — DRM: one row per centre (aggregated monthly summary)
      ac   — Active centers: master list of centre metadata
      od   — Offering Details: one row per subject offering per centre
      ne   — New Enrolled student: mid-month new enrolment records
    """

    if not os.path.exists(path):
        return {}

    # ── Inner helper: fix the same Unicode encoding artefacts ─────────────────
    # Defined locally (inside the function) because it's only used here.
    # str() handles NaN/None; .replace() removes the two garbled sequences;
    # .strip() removes leading/trailing whitespace.
    def _fix_state(s):
        return (str(s)
                .replace("\u00ac\u00a0", " ")
                .replace("\u00a0", " ")
                .strip())

    # ── Load SESSION DUMP ──────────────────────────────────────────────────────
    # This is the most detailed sheet — every teaching session gets one row.
    # It powers the weekly trend chart, day-of-week patterns, subject analytics,
    # volunteer performance table, and cancellation analysis in Tab 4.
    sess = pd.read_excel(path, sheet_name="SESSION DUMP")
    sess.columns = sess.columns.str.strip()

    # Parse the session start timestamp — used to extract week number and
    # day-of-week for trend/pattern charts.
    sess["Session_start"] = pd.to_datetime(sess["Session_start"], errors="coerce")

    # Coerce attendance to a float; non-numeric values become NaN.
    sess["Attendance%"] = pd.to_numeric(sess["Attendance%"], errors="coerce")

    # Apply the state encoding fix to each row using .apply()
    sess["State"] = sess["State"].apply(_fix_state)

    # Replace any missing donor names with "Unknown" so groupby doesn't drop rows
    sess["Donor"] = sess["Donor"].fillna("Unknown")

    # Normalise subject names using the DRM-specific map defined above.
    # .str.strip() first removes whitespace, then .map() translates the value,
    # then .fillna() keeps the original if no translation was found.
    sess["Subject_clean"] = sess["Subject"].str.strip().map(SUBJECT_NORM_DRM).fillna(sess["Subject"])

    # Coerce numeric columns to integers — CLH and student counts must be whole numbers.
    for col in ["Present/CLH", "Total students", "Boys", "Girls"]:
        sess[col] = pd.to_numeric(sess[col], errors="coerce").fillna(0).astype(int)

    # ── Derive time-based columns from the session timestamp ──────────────────
    # .dt accesses the "datetime accessor" — a Pandas feature that lets you
    # extract parts of a datetime value like you'd look at a calendar.

    # .to_period("W") groups the date into its calendar week (e.g., "2026-W18").
    # .astype(str) converts the Period object to a readable string so Plotly
    # can use it as a chart axis label.
    sess["week"] = sess["Session_start"].dt.to_period("W").astype(str)

    # .day_name() returns the full English day name: "Monday", "Tuesday", etc.
    # This is used to create the day-of-week bar chart in Section 5.
    sess["dow"]  = sess["Session_start"].dt.day_name()

    # .hour returns the hour of day (0-23).  Stored for potential future use.
    sess["hour"] = sess["Session_start"].dt.hour

    # ── Load DRM (centre-level summary sheet) ─────────────────────────────────
    # This sheet has one row per teaching centre and contains aggregated counts
    # (Planned sessions, Completed, Enrolled students, etc.) for the month.
    # It powers the scatter plot, state analysis, and full scorecard table.
    drm = pd.read_excel(path, sheet_name="DRM")
    drm.columns = drm.columns.str.strip()
    drm["State"]      = drm["State"].apply(_fix_state)
    drm["Donor Name"] = drm["Donor Name"].fillna("Unknown")

    # Coerce all count/numeric columns to numbers; missing values become 0.
    for col in ["Planned","Scheduled","Completed","Offline","Cancelled",
                "Live Volunteers","Live_CLH","Registered","Reg Boys","Reg Girls",
                "Enrolled","En Boys","En Girls"]:
        if col in drm.columns:
            drm[col] = pd.to_numeric(drm[col], errors="coerce").fillna(0)

    # Coerce attendance separately — leave as float (not int) because it's a %
    drm["Attendance %"] = pd.to_numeric(drm["Attendance %"], errors="coerce")

    # ── Compute derived (calculated) percentage columns ───────────────────────
    # WHY: The Excel sheet stores raw counts (Completed=14, Planned=15).
    # The dashboard needs percentages.  We calculate them once here so
    # every chart below can just reference the column directly.
    #
    # .replace(0, float("nan")): If Planned=0, dividing by 0 would crash Python.
    # Replacing 0 with NaN first means the division returns NaN, which is safe.
    # .round(1): keep one decimal place (e.g., 93.3%).

    # Completion% = (Completed + Offline) / Planned × 100
    # WHY include Offline: "Offline" sessions were delivered, just asynchronously.
    # They count toward completion because learning happened.
    drm["Completion%"]   = ((drm["Completed"] + drm["Offline"])
                            / drm["Planned"].replace(0, float("nan")) * 100).round(1)

    # Cancellation% = Cancelled / Planned × 100
    drm["Cancellation%"] = (drm["Cancelled"]
                            / drm["Planned"].replace(0, float("nan")) * 100).round(1)

    # Dropout% = students who registered but never enrolled / total registered × 100
    # This measures how many students signed up but dropped before the programme started.
    drm["Dropout%"]      = ((drm["Registered"] - drm["Enrolled"])
                            / drm["Registered"].replace(0, float("nan")) * 100).round(1)

    # Girl% = enrolled girls / total enrolled × 100
    # eVidyaloka has a strong girl-student focus; this metric is prominent in reports.
    drm["Girl%"]         = (drm["En Girls"]
                            / (drm["En Boys"] + drm["En Girls"]).replace(0, float("nan")) * 100).round(1)

    # ── Load the remaining three sheets — minimal cleaning needed ─────────────

    # Active centers: master list with centre metadata (location, coordinator, etc.)
    ac = pd.read_excel(path, sheet_name="Active centers")
    ac.columns = ac.columns.str.strip()

    # Offering Details: one row per subject-grade combination offered at each centre.
    # We add Subject_clean for consistency with the session-level analysis.
    od = pd.read_excel(path, sheet_name="Offering Details")
    od.columns = od.columns.str.strip()
    od["Subject_clean"] = od["Subject"].str.strip().map(SUBJECT_NORM_DRM).fillna(od["Subject"])

    # New Enrolled student: records students who joined mid-month after the programme started.
    ne = pd.read_excel(path, sheet_name="New Enrolled student")
    ne.columns = ne.columns.str.strip()

    # Return all five DataFrames in a single dictionary.
    # In Tab 4 (DRM Client Report) we unpack this with:
    #   sess = drm_data["sess"], drm = drm_data["drm"], etc.
    return {"sess": sess, "drm": drm, "ac": ac, "od": od, "ne": ne}


# =============================================================================
# SECTION 8 — SHARED HELPER FUNCTIONS
#
# WHY: We have dozens of charts, and they all need the same visual style
# (transparent background, specific fonts, grid colour, etc.).  Instead of
# copy-pasting those 6 lines into every chart, we write them once here.
# This follows the programming principle "DRY" — Don't Repeat Yourself.
# =============================================================================

def _layout(fig, height=380, legend_bottom=False, margin=None):
    """
    Applies the standard visual theme to any Plotly figure.

    Parameters:
      fig           — the Plotly figure object to style
      height        — chart height in pixels (default 380)
      legend_bottom — if True, move the legend below the chart (useful when
                      bar labels would otherwise overlap the legend)
      margin        — dict of pixel margins {l, r, t, b}; defaults to
                      tight margins that let charts fill their column

    Returns the styled figure so callers can chain further updates.
    """
    # If no margin was passed in, use a tight default
    m = margin or dict(l=0, r=0, t=40, b=0)

    fig.update_layout(
        height=height,
        # Transparent backgrounds so charts inherit the page/card background
        plot_bgcolor=BG, paper_bgcolor=BG,
        # Consistent font across all charts
        font=dict(family="Inter, Helvetica, sans-serif", size=12, color="#2d3436"),
        margin=m,
    )
    # X-axis: no vertical grid lines (they clutter horizontal bar charts).
    # A subtle border line replaces the heavy grid.
    fig.update_xaxes(showgrid=False, linecolor="#dee2e6", linewidth=1)

    # Y-axis: light horizontal grid lines help the eye read values precisely.
    # zeroline=False removes the heavy "0" baseline that Plotly shows by default.
    fig.update_yaxes(showgrid=True, gridcolor=GRID, zeroline=False)

    # Optionally move the legend below the chart to prevent it overlapping bars.
    # orientation="h" makes the legend items line up horizontally.
    if legend_bottom:
        fig.update_layout(legend=dict(
            orientation="h", yanchor="top", y=-0.18,
            xanchor="center", x=0.5, title_text="",
        ))
    return fig


def _kpi(label: str, value: str, color: str, sub: str = "") -> str:
    """
    Builds an HTML string for a KPI metric card used in the VRM tabs.

    WHY HTML: Streamlit's built-in metric widget doesn't support custom colours
    or the left border accent.  We write raw HTML and inject it with
    st.markdown(..., unsafe_allow_html=True).

    Parameters:
      label — small uppercase text at the top (e.g., "TOTAL CLH")
      value — large bold number in the middle (e.g., "32,502")
      color — the left-border and number colour (from palette P or D)
      sub   — small grey subtext at the bottom (e.g., "child learning hrs")

    Returns: an HTML div string.  The caller renders it with st.markdown().
    """
    return (
        # Outer container: light grey background, rounded corners, coloured left border
        f"<div style='background:#f8f9fa;border-radius:10px;padding:14px 10px;"
        f"border-left:4px solid {color};text-align:center;'>"
        # Label: small, grey, uppercase
        f"<div style='font-size:0.72em;color:#636e72;font-weight:600;"
        f"letter-spacing:0.05em;text-transform:uppercase;margin-bottom:4px;'>{label}</div>"
        # Value: large, bold, coloured
        f"<div style='font-size:1.7rem;font-weight:800;color:{color};"
        f"line-height:1.15;'>{value}</div>"
        # Sub-label: tiny, muted grey
        f"<div style='font-size:0.7em;color:#b2bec3;margin-top:3px;'>{sub}</div>"
        f"</div>"
    )


def _drm_kpi(label: str, value: str, color: str, sub: str = "") -> str:
    """
    A larger, more polished KPI card for the premium DRM Client Report tab.

    WHY DIFFERENT FROM _kpi(): The DRM tab is client-facing and needs a more
    elevated visual appearance — white background, box shadow, slightly larger
    numbers, and monospace font to make numeric values look crisp and precise.

    Same parameters as _kpi().
    """
    return (
        # White background + subtle drop shadow gives a "card" feel on client reports
        f"<div style='background:white;border-radius:12px;padding:18px 14px;"
        f"border-left:5px solid {color};text-align:center;"
        f"box-shadow:0 2px 8px rgba(0,0,0,0.06);'>"
        f"<div style='font-size:0.68em;color:#636e72;font-weight:700;"
        f"letter-spacing:0.07em;text-transform:uppercase;margin-bottom:5px;'>{label}</div>"
        # Monospace font ("DM Mono") makes large numbers easier to scan at a glance
        f"<div style='font-size:1.9rem;font-weight:800;color:{color};"
        f"line-height:1.1;font-family:\"DM Mono\",monospace;'>{value}</div>"
        f"<div style='font-size:0.68em;color:#b2bec3;margin-top:4px;'>{sub}</div>"
        f"</div>"
    )


def _insight_box(text: str, color: str = "#ba7517", bg: str = "#faeeda") -> str:
    """
    Renders a highlighted "Key Insight" callout box below a chart.

    WHY: After showing a cancellation pie chart, for example, we want to add
    a data-driven written insight ("69% of cancellations are volunteer-side").
    A visually distinct box draws the reader's eye and signals editorial commentary.

    Parameters:
      text  — the insight text (HTML allowed inside)
      color — accent colour for the label and left border (default: amber)
      bg    — background colour for the box (default: pale amber)
    """
    return (
        f"<div style='background:{bg};border-left:4px solid {color};"
        f"border-radius:8px;padding:12px 16px;margin-top:12px;'>"
        # "KEY INSIGHT" label in the accent colour
        f"<span style='color:{color};font-weight:700;font-size:0.8em;'>KEY INSIGHT&nbsp;&nbsp;</span>"
        # The insight text in dark near-black
        f"<span style='color:#2d3436;font-size:0.82em;'>{text}</span>"
        f"</div>"
    )


# =============================================================================
# SECTION 9 — MAIN RENDER FUNCTION
#
# WHY ONE BIG FUNCTION: All UI construction lives inside render_ops_dashboard().
# This makes it easy to call from a parent app.py file that may handle routing
# between multiple dashboards.  The function has no return value — it renders
# Streamlit elements as a side effect.
#
# EXECUTION ORDER:
#   1. Render page title and subtitle
#   2. Load VRM data (cached)
#   3. Build sidebar filter widgets
#   4. Apply filters to produce the working DataFrame `df`
#   5. Create the tab structure (4 tabs)
#   6. Inside each `with tab_xxx:` block, compute metrics and render charts
#
# IMPORTANT NOTE ON STREAMLIT'S EXECUTION MODEL:
#   Every time the user interacts (clicks a filter, changes a tab), Streamlit
#   re-runs the ENTIRE script from top to bottom.  @st.cache_data ensures
#   data loading doesn't re-execute; the filter code below runs fresh each time.
# =============================================================================
def render_ops_dashboard():

    # ── Page header ───────────────────────────────────────────────────────────
    # st.title() renders the main H1 heading of the page.
    st.title("🏢 Operations & Impact Command Center")

    # st.markdown() renders arbitrary Markdown or HTML.
    # unsafe_allow_html=True is required when the string contains HTML tags.
    # The negative margin-top pulls the subtitle closer to the title.
    st.markdown(
        "<p style='color:gray;font-size:1.05em;margin-top:-12px;'>"
        "Volunteer Relationship Management  ·  Centre Operations  ·  Academic Health"
        "  —  May 2026</p>",
        unsafe_allow_html=True,
    )
    # A horizontal rule (grey divider line) visually separates the header from content
    st.markdown("---")

    # ── Load VRM data ─────────────────────────────────────────────────────────
    # st.spinner() shows a "loading…" animation while the indented block runs.
    # Because load_data() is cached, this spinner only appears on the very first
    # load — subsequent renders return instantly from cache.
    with st.spinner("Loading VRM dataset…"):
        data = load_data(DATA_PATH)

    # If load_data() returned an empty dict (file not found), show an error
    # and exit the function early.  `return` inside a Streamlit function stops
    # rendering the rest of the page — nothing below this point is executed.
    if not data:
        st.error(
            f"⚠️ Data file not found at `{DATA_PATH}`. "
            "Place `VRM_May__2026.xlsx` alongside `app.py`."
        )
        return

    # Unpack the three DataFrames from the dictionary for convenient access below.
    # Think of this as taking labelled jars off the shelf and placing them on
    # the workbench where we'll actually use them.
    df_raw  = data["active"]    # All active volunteer-offering rows
    dropped = data["dropped"]   # Volunteers who left
    new_reg = data["new_reg"]   # All new sign-ups this period

    # ── Sidebar filter widgets ────────────────────────────────────────────────
    # `with st.sidebar:` places everything inside the indented block into the
    # collapsible panel on the left side of the Streamlit page.
    with st.sidebar:
        st.markdown("---")
        st.header("🎯 VRM Filters")

        # st.multiselect() renders a dropdown that allows selecting multiple values.
        # options=sorted(...) provides the full sorted list of unique values found
        # in the data, so new donors/states/subjects appear automatically.
        # default=... pre-selects all values so the dashboard shows everything
        # on first load without the user needing to make a selection.
        # key= gives each widget a unique internal ID to prevent Streamlit from
        # mixing up widget states when multiple multiselects exist on the page.

        sel_donors = st.multiselect(
            "Donor",
            options=sorted(df_raw["Donor"].dropna().unique()),   # All unique non-null donors
            default=sorted(df_raw["Donor"].dropna().unique()),   # All selected by default
            key="ops_donor",
        )
        sel_states = st.multiselect(
            "State (Centre)",
            options=sorted(df_raw["State"].dropna().unique()),
            default=sorted(df_raw["State"].dropna().unique()),
            key="ops_state",
        )
        sel_subjects = st.multiselect(
            "Subject",
            options=sorted(df_raw["Subject_clean"].dropna().unique()),
            default=sorted(df_raw["Subject_clean"].dropna().unique()),
            key="ops_subject",
        )
        # Inform users which tabs these filters affect (not Tab 4 — DRM is independent)
        st.caption("Filters apply to the VRM tabs (Volunteers, Centres, Academic Health). "
                   "The DRM Client Report tab uses the full DRM dataset independently.")

    # ── Apply sidebar filters to produce the working DataFrame ────────────────
    # This is a boolean mask filter — we keep only rows where ALL three conditions
    # are true simultaneously.  The .isin() method checks whether each row's value
    # is in the user's current selection list.  .copy() creates an independent copy
    # so we don't accidentally modify the cached df_raw.
    #
    # Think of it like putting three sieves on top of each other:
    # only data that passes through all three (donor AND state AND subject) survives.
    df = df_raw[
        df_raw["Donor"].isin(sel_donors) &
        df_raw["State"].isin(sel_states) &
        df_raw["Subject_clean"].isin(sel_subjects)
    ].copy()

    # Guard: if the combination of filters produces zero rows (e.g., user
    # deselects everything), show a friendly warning and stop rendering.
    if df.empty:
        st.warning("⚠️ No data for the current filter combination. Broaden your selection.")
        return

    # ── Volunteer-level deduplicated DataFrame ────────────────────────────────
    # The main DataFrame df has one row per volunteer-OFFERING pair, meaning a
    # volunteer who teaches 3 subjects appears 3 times.  For volunteer-level
    # metrics (e.g., "how many unique volunteers are there?") we need to count
    # each volunteer once.  drop_duplicates(subset="Volunteer id") keeps only
    # the first occurrence of each unique Volunteer id.
    # This vol_df is used in Tab 1 for profession pie charts and state bar charts.
    vol_df = df.drop_duplicates(subset="Volunteer id").copy()

    # ── Create the 4-tab layout ───────────────────────────────────────────────
    # st.tabs() returns one Python variable per tab name.  Code placed inside
    # `with tab_vol:` only renders when the user clicks the "Volunteers" tab.
    # The variable names (tab_vol, etc.) are used below to enter each tab context.
    tab_vol, tab_ctr, tab_aca, tab_drm = st.tabs([
        "🙋 Volunteers",
        "🏫 Centres",
        "📚 Academic Health",
        "📊 DRM Client Report",
    ])

    # =========================================================================
    # TAB 1 — VOLUNTEERS
    # Shows: volunteer counts, acquisition channels, profession mix,
    #        geographic residence, monthly registration trend, gender split,
    #        and dropped volunteer reasons.
    # Data source: df (filtered active), vol_df (deduplicated), new_reg, dropped
    # =========================================================================
    with tab_vol:

        # ── KPI calculations ──────────────────────────────────────────────────
        # These numbers are computed from the filtered DataFrame df and will be
        # displayed as coloured metric cards in the next block.

        # Total rows in df = number of active vol-offering pairs (not unique vols)
        total_vols_offering = len(df)

        # nunique() = "number unique" — counts distinct values in the column.
        # This gives us the actual volunteer headcount, not offering count.
        unique_vols = vol_df["Volunteer id"].nunique()

        # Dropped and new registrations come from separate sheets, so they are
        # independent of the sidebar filters (they don't have the same columns).
        dropped_vols = len(dropped)
        new_vols     = new_reg["User ID"].nunique()

        # Count how many unique centre names exist in the filtered data
        active_centres = df["Center name"].nunique()

        # WHY drop_duplicates here: Enrolled is a centre-level number, but df has
        # multiple rows per centre (one per volunteer-offering at that centre).
        # Summing directly would multiply-count the same centre's enrolled students.
        # By deduplicating on centre name first, we get the true programme total.
        total_enrolled = int(df.drop_duplicates("Center name")["Enrolled"].sum())

        # Sum total volunteer-hours and Child Learning Hours (CLH) across all offerings
        total_vol_hrs = int(df["Total hours(Comp+Offline)"].sum())
        total_clh     = int(df["CLH"].sum())

        # Mean of the Attendance% column across all filtered rows.
        # pd.notna() guards against the case where all values are NaN.
        avg_att         = df["Attendance%"].mean()
        avg_att_display = f"{avg_att:.1f}%" if pd.notna(avg_att) else "N/A"

        # Completion rate: what fraction of planned sessions were actually completed?
        # Guard against division by zero (if Planned sums to 0, default to 0%).
        completion_rt = (
            df["Completed"].sum() / df["Planned"].sum() * 100
            if df["Planned"].sum() > 0 else 0
        )

        # ── KPI Row 1: Volunteer counts ───────────────────────────────────────
        # st.columns(4) creates a 4-column horizontal grid.
        # Each column variable (kc1..kc4) is a context in which we render content.
        # kc1.markdown(...) places content inside the first column only.
        st.markdown("#### Volunteer Overview")
        kc1, kc2, kc3, kc4 = st.columns(4)
        kc1.markdown(_kpi("Total Volunteers",   f"{total_vols_offering:,}", P["teal"],   "vol-offering rows"),   unsafe_allow_html=True)
        kc2.markdown(_kpi("Unique Volunteers",  f"{unique_vols:,}",         P["green"],  "deduplicated"),        unsafe_allow_html=True)
        kc3.markdown(_kpi("Newly Registered",   f"{new_vols:,}",            P["violet"], "from registrations sheet"), unsafe_allow_html=True)
        kc4.markdown(_kpi("Dropped Volunteers", f"{dropped_vols:,}",        P["coral"],  "from dropped sheet"),  unsafe_allow_html=True)

        # ── KPI Row 2: Impact counts ──────────────────────────────────────────
        # We add a blank line for visual breathing room between the two KPI rows.
        st.markdown("<br>", unsafe_allow_html=True)

        # 6-column layout for the second row of impact metrics
        ki1, ki2, ki3, ki4, ki5, ki6 = st.columns(6)
        ki1.markdown(_kpi("Active Centres",   f"{active_centres:,}",  P["teal"],   "unique"),             unsafe_allow_html=True)
        ki2.markdown(_kpi("Enrolled Students",f"{total_enrolled:,}",  P["green"],  "total seats"),        unsafe_allow_html=True)
        ki3.markdown(_kpi("Total Vol Hrs",    f"{total_vol_hrs:,}",   P["orange"], "Comp + Offline hrs"), unsafe_allow_html=True)
        ki4.markdown(_kpi("Total CLH",        f"{total_clh:,}",       P["violet"], "child learning hrs"), unsafe_allow_html=True)
        ki5.markdown(_kpi("Avg Attendance",   avg_att_display,        P["amber"],  "across sessions"),    unsafe_allow_html=True)
        ki6.markdown(_kpi("Class Completion", f"{completion_rt:.1f}%",P["mint"],   "completed / planned"),unsafe_allow_html=True)

        st.markdown("---")

        # ── Volunteer Distribution: Acquisition Channel + Profession ──────────
        # Two charts side by side: a horizontal bar chart (left) and a donut (right).
        st.markdown("#### Volunteer Distribution")
        col_b1, col_b2 = st.columns(2)

        with col_b1:
            st.markdown("##### By Reference / Acquisition Channel")

            # .groupby("Ref_group") groups rows by the bucketed channel name.
            # ["Volunteer id"].nunique() counts distinct volunteers per group.
            # .reset_index() converts the grouped result back into a flat DataFrame
            # (required for Plotly to read it as a table with named columns).
            # .rename() gives the count column a human-readable name.
            # .sort_values() orders bars shortest-to-tallest (ascending) so the
            # longest bar is at the top of a horizontal chart — most readable.
            ref_data = (
                vol_df.groupby("Ref_group")["Volunteer id"].nunique()
                .reset_index()
                .rename(columns={"Volunteer id": "Volunteers"})
                .sort_values("Volunteers", ascending=True)
            )

            # px.bar() creates a bar chart.
            # orientation="h" makes it horizontal (categories on Y, values on X).
            # text="Volunteers" displays the exact count at the end of each bar.
            # color="Ref_group" with color_discrete_sequence assigns a different
            # colour from SEQ to each channel automatically.
            fig_ref = px.bar(
                ref_data, x="Volunteers", y="Ref_group",
                orientation="h", text="Volunteers",
                color="Ref_group", color_discrete_sequence=SEQ,
            )

            # update_traces modifies styling of the data elements (bars, lines, etc.)
            # textposition="outside" places the count labels just beyond the bar end.
            # showlegend=False hides the colour legend (redundant here since the
            # Y-axis labels already name each channel).
            # marker_line_width=0 removes the thin border Plotly adds around each bar.
            # hovertemplate controls what appears in the tooltip on mouse-over.
            # %{y} = the Y-axis value, %{x:,} = the X value with thousands comma.
            # <extra></extra> removes Plotly's default trace-name suffix in tooltips.
            fig_ref.update_traces(
                textposition="outside", showlegend=False,
                marker_line_width=0,
                hovertemplate="<b>%{y}</b><br>Volunteers: %{x:,}<extra></extra>",
            )
            _layout(fig_ref, height=320)  # Apply our standard visual theme
            fig_ref.update_layout(xaxis_title="Volunteers", yaxis_title="", showlegend=False)
            fig_ref.update_xaxes(showgrid=True, gridcolor=GRID)
            fig_ref.update_yaxes(showgrid=False)

            # Render the Plotly chart inside the Streamlit page.
            # use_container_width=True makes the chart fill its column width.
            st.plotly_chart(fig_ref, use_container_width=True)

        with col_b2:
            st.markdown("##### By Profession")

            # Same groupby pattern as above, but sorted descending (largest first)
            # because the donut chart assigns colours from 12 o'clock clockwise.
            prof_data = (
                vol_df.groupby("Profession_clean")["Volunteer id"].nunique()
                .reset_index()
                .rename(columns={"Volunteer id": "Volunteers"})
                .sort_values("Volunteers", ascending=False)
            )

            # px.pie() with hole=0.45 creates a donut chart (45% hole in the centre).
            # The centre will be used for an annotation showing the total count.
            fig_prof = px.pie(
                prof_data, names="Profession_clean", values="Volunteers",
                hole=0.45, color_discrete_sequence=SEQ,
            )
            fig_prof.update_traces(
                textposition="outside",
                # texttemplate controls the label on each slice.
                # %{label} = profession name, %{percent:.1%} = formatted percentage.
                texttemplate="<b>%{label}</b><br>%{percent:.1%}",
                hovertemplate="<b>%{label}</b><br>Volunteers: %{value:,}<extra></extra>",
                # pull slightly explodes each slice outward for visual clarity
                pull=[0.02] * len(prof_data),
            )
            fig_prof.update_layout(
                height=320, showlegend=False,
                plot_bgcolor=BG, paper_bgcolor=BG,
                margin=dict(l=10, r=10, t=40, b=10),
                # annotations place text inside the donut hole.
                # x=0.5, y=0.5 centres it; showarrow=False hides the pointer arrow.
                annotations=[dict(
                    text=f"<b>{unique_vols:,}</b><br>Vols",
                    x=0.5, y=0.5, font_size=15, showarrow=False,
                    font=dict(color="#2d3436"),
                )],
            )
            st.plotly_chart(fig_prof, use_container_width=True)

        st.markdown("---")

        # ── Volunteer Residence by State ──────────────────────────────────────
        # Shows where volunteers physically live (vs. where they teach).
        # A tall chart is needed when there are many states, so we use
        # max(300, len(res_state)*28) to set a minimum of 300px but grow
        # proportionally for large datasets.
        st.markdown("#### Volunteer Residence — by State")
        res_state = (
            vol_df.groupby("Residence state")["Volunteer id"].nunique()
            .reset_index()
            .rename(columns={"Volunteer id": "Volunteers"})
            .sort_values("Volunteers", ascending=True)
        )
        fig_res = px.bar(
            res_state, x="Volunteers", y="Residence state",
            orientation="h", text="Volunteers",
            color="Volunteers",
            # A continuous colour scale makes bars darker where values are higher,
            # giving an instant visual sense of magnitude without needing to read labels.
            color_continuous_scale=[P["sky"], P["teal"]],
        )
        fig_res.update_traces(
            textposition="outside", marker_line_width=0,
            hovertemplate="<b>%{y}</b><br>Volunteers: %{x:,}<extra></extra>",
        )
        # showscale=False hides the colour gradient legend bar (redundant here)
        fig_res.update_coloraxes(showscale=False)
        _layout(fig_res, height=max(300, len(res_state) * 28),
                margin=dict(l=0, r=60, t=20, b=0))
        fig_res.update_layout(xaxis_title="Volunteers", yaxis_title="")
        fig_res.update_xaxes(showgrid=True, gridcolor=GRID)
        fig_res.update_yaxes(showgrid=False)
        st.plotly_chart(fig_res, use_container_width=True)

        st.markdown("---")

        # ── Monthly registration trend ────────────────────────────────────────
        # Shows how many new volunteers signed up each calendar month.
        # Note: This uses new_reg (all registrations ever), NOT the filtered df.
        # Reason: the sidebar filters don't apply to the new_reg sheet (different
        # columns), and it's more informative to show the full historical trend.
        st.markdown("#### New Volunteer Registrations — Monthly Trend")
        st.caption("Source: Newly Registered VT sheet — all registrations regardless of sidebar filters.")

        trend_df = new_reg.dropna(subset=["Date Joined"]).copy()

        # .dt.to_period("M") converts each date to its calendar month (e.g. "2025-11").
        # .astype(str) makes it a plain string so Plotly can use it as a chart label.
        trend_df["YM"] = trend_df["Date Joined"].dt.to_period("M").astype(str)

        monthly = (
            trend_df.groupby("YM")["User ID"].nunique()
            .reset_index()
            .rename(columns={"User ID": "New Vols"})
            .sort_values("YM")  # Sort chronologically (works because YM is "YYYY-MM" format)
        )

        # px.area() is like a line chart with the area below shaded — good for
        # showing cumulative growth trends over time.
        fig_trend = px.area(
            monthly, x="YM", y="New Vols",
            markers=True,  # Show dots at each data point
            color_discrete_sequence=[P["teal"]],
        )
        fig_trend.update_traces(
            line=dict(width=2.5),
            marker=dict(size=6, color=P["teal"]),
            # rgba with alpha=0.12 makes the fill colour very translucent
            fillcolor="rgba(0,148,201,0.12)",
            hovertemplate="<b>%{x}</b><br>New Vols: %{y:,}<extra></extra>",
        )
        _layout(fig_trend, height=240, margin=dict(l=0, r=0, t=20, b=60))
        fig_trend.update_layout(xaxis_title="", yaxis_title="New Volunteers")
        # tickangle=-45 rotates month labels diagonally so they don't overlap
        fig_trend.update_xaxes(tickangle=-45, showgrid=False)
        st.plotly_chart(fig_trend, use_container_width=True)

        st.markdown("---")

        # ── Gender split of newly registered volunteers ───────────────────────
        # We only show this section if the Gender column exists in the data.
        # This is a defensive check — if the sheet schema changes, we gracefully
        # degrade to an info message rather than crashing with a KeyError.
        st.markdown("#### Newly Registered Volunteers — Gender Split")
        if "Gender" in new_reg.columns:
            gen_new = (
                # Filter to only Male/Female (exclude blank, "Unknown", etc.)
                new_reg[new_reg["Gender"].isin(["Male", "Female"])]
                .groupby("Gender")["User ID"].nunique()
                .reset_index()
                .rename(columns={"User ID": "Volunteers"})
            )

            # Two columns: left = donut chart showing M/F split, right = stacked
            # horizontal bar chart showing M/F split per acquisition channel.
            # [1, 2] means the right column is twice as wide as the left.
            col_g1, col_g2 = st.columns([1, 2])

            with col_g1:
                # Gender donut chart — same technique as profession donut above
                fig_gen = px.pie(
                    gen_new, names="Gender", values="Volunteers",
                    hole=0.5,
                    color="Gender",
                    # Explicit colour mapping: Male=teal, Female=coral (consistent
                    # with the same encoding used in the Centres and DRM tabs)
                    color_discrete_map={"Male": P["teal"], "Female": P["coral"]},
                )
                fig_gen.update_traces(
                    texttemplate="<b>%{label}</b><br>%{percent:.1%}",
                    textposition="outside",
                    hovertemplate="<b>%{label}</b><br>Volunteers: %{value:,}<extra></extra>",
                )
                fig_gen.update_layout(
                    height=280, showlegend=False,
                    plot_bgcolor=BG, paper_bgcolor=BG,
                    margin=dict(l=10, r=10, t=30, b=10),
                )
                st.plotly_chart(fig_gen, use_container_width=True)

            with col_g2:
                # Stacked bar: for each acquisition channel, how many M vs F?
                # .assign() adds a new column inline using a lambda (anonymous function).
                # lambda x: x["Reference"].apply(_group_ref) runs our bucketing function
                # on the Reference column of the filtered new_reg rows.
                gen_ref = (
                    new_reg[new_reg["Gender"].isin(["Male", "Female"])]
                    .assign(Ref_group=lambda x: x["Reference"].apply(_group_ref))
                    .groupby(["Ref_group", "Gender"])["User ID"].nunique()
                    .reset_index()
                    .rename(columns={"User ID": "Volunteers"})
                )
                fig_gen_ref = px.bar(
                    gen_ref, x="Volunteers", y="Ref_group", color="Gender",
                    orientation="h",
                    # barmode="stack" places Male and Female bars end-to-end per channel,
                    # making proportions easy to compare visually.
                    barmode="stack",
                    color_discrete_map={"Male": P["teal"], "Female": P["coral"]},
                    text="Volunteers",
                )
                fig_gen_ref.update_traces(
                    textposition="inside", texttemplate="%{text}",
                    hovertemplate="<b>%{y}</b> · %{fullData.name}: %{x}<extra></extra>",
                )
                _layout(fig_gen_ref, height=280, legend_bottom=False,
                        margin=dict(l=0, r=0, t=30, b=0))
                fig_gen_ref.update_layout(
                    xaxis_title="Volunteers", yaxis_title="",
                    legend=dict(orientation="h", yanchor="top", y=-0.15,
                                xanchor="center", x=0.5, title_text=""),
                )
                fig_gen_ref.update_xaxes(showgrid=True, gridcolor=GRID)
                fig_gen_ref.update_yaxes(showgrid=False)
                st.plotly_chart(fig_gen_ref, use_container_width=True)
        else:
            st.info("Gender data not available in the Newly Registered sheet.")

        st.markdown("---")

        # ── Dropped volunteer reason breakdown ───────────────────────────────
        # A horizontal bar chart of the "Reasons" column in the Dropped VT sheet.
        # We first check that the sheet is non-empty AND has a "Reasons" column.
        st.markdown("#### Dropped Volunteers — Reason Breakdown")
        st.caption("Source: Dropped VT sheet — volunteers who exited their offering(s).")
        if not dropped.empty and "Reasons" in dropped.columns:
            # .value_counts() returns a Series sorted by frequency, highest first.
            # .reset_index() converts it to a two-column DataFrame (reason, count).
            reason_counts = (
                dropped["Reasons"].fillna("Not specified")
                .value_counts().reset_index()
            )
            reason_counts.columns = ["Reason", "Count"]

            fig_drop = px.bar(
                reason_counts, x="Count", y="Reason",
                orientation="h", text="Count",
                color="Count",
                # Gradient from light salmon to bright red — higher count = more alarming
                color_continuous_scale=[P["salmon"], P["red"]],
            )
            fig_drop.update_traces(
                textposition="outside", marker_line_width=0,
                hovertemplate="<b>%{y}</b><br>Count: %{x}<extra></extra>",
            )
            fig_drop.update_coloraxes(showscale=False)
            _layout(fig_drop, height=max(220, len(reason_counts) * 42),
                    margin=dict(l=0, r=60, t=20, b=0))
            fig_drop.update_layout(xaxis_title="Volunteers", yaxis_title="")
            fig_drop.update_xaxes(showgrid=True, gridcolor=GRID)
            fig_drop.update_yaxes(showgrid=False)
            st.plotly_chart(fig_drop, use_container_width=True)
        else:
            st.info("No dropped volunteer records in the current data.")

    # =========================================================================
    # TAB 2 — CENTRES
    # Shows: Centres + volunteers per donor, enrolled students, CLH by donor,
    #        gender split by donor, states by active centres, class completion
    #        split by subject.
    # Data source: df (filtered active VRM data)
    # =========================================================================
    with tab_ctr:

        # ── Donor summary aggregation ─────────────────────────────────────────
        # .groupby("Donor") groups all rows by donor name.
        # .agg() computes multiple statistics simultaneously — think of it like
        # building a pivot table: each row in the result is one donor, each
        # column is one aggregated metric.
        # Named aggregation syntax: NewColName=("SourceCol", "aggregation_function")
        st.markdown("#### Centres & Volunteers by Donor")
        donor_summary = (
            df.groupby("Donor").agg(
                Centres    =("Center name",  "nunique"),   # Unique centre count
                Enrolled   =("Enrolled",     "sum"),       # Total enrolled students
                CLH        =("CLH",          "sum"),       # Total child learning hours
                Volunteers =("Volunteer id", "nunique"),   # Unique volunteer count
                Completed  =("Completed",    "sum"),
                Planned    =("Planned",      "sum"),
                En_Boys    =("En Boys",      "sum"),
                En_Girls   =("En Girls",     "sum"),
            ).reset_index()
        )

        # Compute Completion % from the aggregated sums.
        # .replace(0, pd.NA) prevents division-by-zero (same pattern as load_drm_data).
        # .fillna(0) replaces any resulting NaN back to 0 for display.
        donor_summary["Completion %"] = (
            donor_summary["Completed"]
            / donor_summary["Planned"].replace(0, pd.NA) * 100
        ).fillna(0).round(1)
        donor_summary = donor_summary.sort_values("Centres", ascending=False)

        # ── Grouped bar chart: Centres AND Volunteers side by side ────────────
        # We use go.Figure() (Graph Objects) instead of px.bar() here because
        # we need TWO separate bar series plotted on the same chart.
        # px.bar() with barmode="group" would work too, but go.Figure() gives us
        # explicit control over each trace's colour and label.
        fig_donor = go.Figure()
        fig_donor.add_trace(go.Bar(
            name="Centres", x=donor_summary["Donor"], y=donor_summary["Centres"],
            marker_color=P["teal"], text=donor_summary["Centres"],
            textposition="outside",
            hovertemplate="<b>%{x}</b><br>Centres: %{y}<extra></extra>",
        ))
        fig_donor.add_trace(go.Bar(
            name="Volunteers", x=donor_summary["Donor"], y=donor_summary["Volunteers"],
            marker_color=P["green"], text=donor_summary["Volunteers"],
            textposition="outside",
            hovertemplate="<b>%{x}</b><br>Volunteers: %{y}<extra></extra>",
        ))
        fig_donor.update_layout(
            barmode="group",   # Side-by-side bars (vs "stack" for stacked bars)
            height=380,
            plot_bgcolor=BG, paper_bgcolor=BG,
            margin=dict(l=0, r=0, t=30, b=80),
            legend=dict(orientation="h", yanchor="top", y=-0.22,
                        xanchor="center", x=0.5, title_text=""),
            xaxis=dict(tickangle=-20, showgrid=False, linecolor="#dee2e6"),
            yaxis=dict(showgrid=True, gridcolor=GRID, zeroline=False),
            font=dict(family="Inter, Helvetica, sans-serif", size=12),
        )
        st.plotly_chart(fig_donor, use_container_width=True)

        # ── Two smaller charts side by side ───────────────────────────────────
        col_d1, col_d2 = st.columns(2)

        with col_d1:
            st.markdown("##### Enrolled Students by Donor")
            # Sort ascending so the largest bar is at the top of the horizontal chart
            fig_enr = px.bar(
                donor_summary.sort_values("Enrolled", ascending=True),
                x="Enrolled", y="Donor", orientation="h", text="Enrolled",
                color="Enrolled",
                color_continuous_scale=[P["lavender"], P["violet"]],
            )
            fig_enr.update_traces(textposition="outside",
                                  texttemplate="%{text:,}", marker_line_width=0)
            fig_enr.update_coloraxes(showscale=False)
            _layout(fig_enr, height=280, margin=dict(l=0, r=70, t=20, b=0))
            fig_enr.update_layout(xaxis_title="Enrolled Students", yaxis_title="")
            fig_enr.update_xaxes(showgrid=True, gridcolor=GRID)
            fig_enr.update_yaxes(showgrid=False)
            st.plotly_chart(fig_enr, use_container_width=True)

        with col_d2:
            st.markdown("##### CLH by Donor")
            fig_clh_d = px.bar(
                donor_summary.sort_values("CLH", ascending=True),
                x="CLH", y="Donor", orientation="h", text="CLH",
                color="CLH",
                color_continuous_scale=[P["salmon"], P["coral"]],
            )
            fig_clh_d.update_traces(textposition="outside",
                                    texttemplate="%{text:,}", marker_line_width=0)
            fig_clh_d.update_coloraxes(showscale=False)
            _layout(fig_clh_d, height=280, margin=dict(l=0, r=70, t=20, b=0))
            fig_clh_d.update_layout(xaxis_title="Child Learning Hours", yaxis_title="")
            fig_clh_d.update_xaxes(showgrid=True, gridcolor=GRID)
            fig_clh_d.update_yaxes(showgrid=False)
            st.plotly_chart(fig_clh_d, use_container_width=True)

        st.markdown("---")

        # ── Gender split by donor ─────────────────────────────────────────────
        # .melt() "unpivots" wide data into long format, which Plotly needs for
        # grouped/stacked bar charts.
        # Before melt:  one row per donor, columns En_Boys and En_Girls
        # After melt:   two rows per donor, one column "Gender", one column "Students"
        # This is like folding a two-column spreadsheet into a single list of entries.
        st.markdown("#### Enrolled Students — Gender Split by Donor")
        st.caption("En Boys and En Girls columns are available in this dataset.")
        gen_donor_melt = donor_summary[["Donor", "En_Boys", "En_Girls"]].melt(
            id_vars="Donor",                          # Keep Donor as the row identifier
            value_vars=["En_Boys", "En_Girls"],       # These two columns become values
            var_name="Gender",                        # The column name goes to "Gender"
            value_name="Students"                     # The cell values go to "Students"
        )
        # Replace the raw column names with readable labels
        gen_donor_melt["Gender"] = gen_donor_melt["Gender"].map(
            {"En_Boys": "Boys", "En_Girls": "Girls"}
        )
        fig_gen_enr = px.bar(
            gen_donor_melt, x="Donor", y="Students",
            color="Gender", barmode="group",
            text="Students",
            color_discrete_map={"Boys": P["teal"], "Girls": P["coral"]},
        )
        fig_gen_enr.update_traces(
            textposition="outside", texttemplate="%{text:,}",
            hovertemplate="<b>%{x}</b> · %{fullData.name}: %{y:,}<extra></extra>",
        )
        _layout(fig_gen_enr, height=340, legend_bottom=True,
                margin=dict(l=0, r=0, t=20, b=60))
        fig_gen_enr.update_layout(xaxis_title="", yaxis_title="Enrolled Students")
        fig_gen_enr.update_xaxes(tickangle=-15, showgrid=False)
        st.plotly_chart(fig_gen_enr, use_container_width=True)

        st.markdown("---")

        # ── States by active centres ──────────────────────────────────────────
        # Shows which states have the most teaching centres, with enrolled
        # students, CLH, and volunteers available on hover.
        # hover_data={"Enrolled": ":,"} adds formatted enrolled count to the tooltip.
        st.markdown("#### Top States by Active Centres")
        state_centres = (
            df.groupby("State").agg(
                Centres    =("Center name",  "nunique"),
                Enrolled   =("Enrolled",     "sum"),
                CLH        =("CLH",          "sum"),
                Volunteers =("Volunteer id", "nunique"),
            ).reset_index()
            .sort_values("Centres", ascending=True)  # Ascending = longest bar on top
        )
        fig_sc = px.bar(
            state_centres, x="Centres", y="State",
            orientation="h", text="Centres",
            color="Centres",
            color_continuous_scale=[P["sky"], P["teal"]],
            hover_data={"Enrolled": ":,", "CLH": ":,", "Volunteers": True},
        )
        fig_sc.update_traces(textposition="outside", marker_line_width=0)
        fig_sc.update_coloraxes(showscale=False)
        # Dynamically set chart height: at least 260px, or 42px per state row
        _layout(fig_sc, height=max(260, len(state_centres) * 42),
                margin=dict(l=0, r=60, t=20, b=0))
        fig_sc.update_layout(xaxis_title="Active Centres", yaxis_title="")
        fig_sc.update_xaxes(showgrid=True, gridcolor=GRID)
        fig_sc.update_yaxes(showgrid=False)
        st.plotly_chart(fig_sc, use_container_width=True)

        st.markdown("---")

        # ── Class completion split by subject ─────────────────────────────────
        # A stacked horizontal bar chart: for each subject, shows how many
        # sessions were Completed, Offline, or Cancelled.
        # This tells the programme team WHERE completion is weakest.
        st.markdown("#### Class Completion Split by Subject")
        st.caption("🟢 Completed  🟡 Offline (rescheduled / async)  🔴 Cancelled")

        # Sum the three session-status columns, grouped by subject
        exec_df = (
            df.groupby("Subject_clean")[["Completed", "Offline", "Cancelled"]]
            .sum().reset_index()
            .sort_values("Completed", ascending=False)  # Most-completed subjects first
        )

        # Melt from wide (3 status columns) to long (1 Status column + 1 Sessions column).
        # This is the same unpivot technique used for the gender chart above.
        exec_melt = exec_df.melt(
            id_vars="Subject_clean",
            value_vars=["Completed", "Offline", "Cancelled"],
            var_name="Status", value_name="Sessions",
        )
        fig_exec = px.bar(
            exec_melt, x="Sessions", y="Subject_clean",
            color="Status", orientation="h", barmode="stack",
            text="Sessions",
            # Explicit colour mapping: green=good, amber=partial, red=bad
            color_discrete_map={
                "Completed": P["green"],
                "Offline":   P["amber"],
                "Cancelled": P["red"],
            },
        )
        fig_exec.update_traces(
            textposition="inside", texttemplate="%{text:,}",
            hovertemplate="<b>%{y}</b><br>%{fullData.name}: %{x:,}<extra></extra>",
        )
        _layout(fig_exec, height=max(320, len(exec_df) * 42),
                legend_bottom=True, margin=dict(l=0, r=0, t=20, b=60))
        fig_exec.update_layout(xaxis_title="Sessions", yaxis_title="")
        fig_exec.update_xaxes(showgrid=True, gridcolor=GRID)
        fig_exec.update_yaxes(showgrid=False)
        st.plotly_chart(fig_exec, use_container_width=True)

    # =========================================================================
    # TAB 3 — ACADEMIC HEALTH
    # Shows: CLH by subject, attendance heatmap (state × subject),
    #        attendance histogram, full filtered data table.
    # Data source: df (filtered active VRM data)
    # =========================================================================
    with tab_aca:

        # ── CLH by subject ────────────────────────────────────────────────────
        # Simple horizontal bar chart: how much learning (in child-hours)
        # does each subject deliver?  Sorted ascending so the highest bar is
        # at the top of the horizontal chart.
        st.markdown("#### CLH by Subject")
        clh_subj = (
            df.groupby("Subject_clean")["CLH"].sum()
            .reset_index()
            .sort_values("CLH", ascending=True)
        )
        fig_clh = px.bar(
            clh_subj, x="CLH", y="Subject_clean",
            orientation="h", text="CLH",
            color="CLH",
            color_continuous_scale=[P["salmon"], P["coral"]],
        )
        fig_clh.update_traces(
            textposition="outside", texttemplate="%{text:,}",
            marker_line_width=0,
            hovertemplate="<b>%{y}</b><br>CLH: %{x:,}<extra></extra>",
        )
        fig_clh.update_coloraxes(showscale=False)
        _layout(fig_clh, height=max(300, len(clh_subj) * 40),
                margin=dict(l=0, r=70, t=20, b=0))
        fig_clh.update_layout(xaxis_title="Child Learning Hours", yaxis_title="")
        fig_clh.update_xaxes(showgrid=True, gridcolor=GRID)
        fig_clh.update_yaxes(showgrid=False)
        st.plotly_chart(fig_clh, use_container_width=True)

        st.markdown("---")

        # ── Attendance heatmap: State × Subject ───────────────────────────────
        # A heatmap is a colour-coded table: rows = states, columns = subjects,
        # cell colour = average attendance %.  Dark = high attendance, light = low.
        # This lets the reader instantly spot which subjects struggle in which states.
        st.markdown("#### Average Attendance % — State × Subject Heatmap")
        st.caption("Scale clamped 50–100%. White = no sessions for that pairing.")

        # Compute mean attendance for every State × Subject combination
        heat_df = (
            df.groupby(["State", "Subject_clean"])["Attendance%"]
            .mean().reset_index()
        )

        # .pivot() reshapes the long DataFrame into a matrix:
        # rows = State, columns = Subject_clean, values = Attendance%.
        # .fillna(0) fills missing cells (no sessions for that combo) with 0.
        heat_pivot = (
            heat_df.pivot(index="State", columns="Subject_clean", values="Attendance%")
            .fillna(0)
        )

        # Sort rows so the highest-average-attendance state is at the top.
        # .replace(0, pd.NA) converts zeros back to NaN first so they don't
        # drag down the row average (zero means "no sessions", not "0% attendance").
        # .mean(axis=1) computes the row mean; .sort_values().index gives the order.
        sort_order = (
            heat_pivot.replace(0, pd.NA).mean(axis=1)
            .sort_values(ascending=False).index
        )
        heat_pivot = heat_pivot.loc[sort_order]   # Reindex rows in that order

        # px.imshow() renders a 2D array as a colour-coded heatmap.
        # text_auto=".0f" displays rounded integers inside each cell.
        # zmin/zmax clamps the colour scale to 50-100% — cells below 50 and
        # zero-cells are both shown as the lightest colour.
        fig_heat = px.imshow(
            heat_pivot,
            color_continuous_scale=["#dfe6e9", P["teal"]],
            aspect="auto",
            text_auto=".0f",
            zmin=50, zmax=100,
        )
        fig_heat.update_traces(
            hovertemplate=(
                "<b>%{y}</b> · <b>%{x}</b><br>"
                "Avg Attendance: %{z:.1f}%<extra></extra>"
            )
        )
        fig_heat.update_layout(
            height=max(300, len(heat_pivot) * 50),
            plot_bgcolor=BG, paper_bgcolor=BG,
            margin=dict(l=0, r=0, t=20, b=120),
            # Show a vertical colour bar on the right as a legend
            coloraxis_colorbar=dict(
                title="Att%", ticksuffix="%", len=0.6, thickness=12,
            ),
            font=dict(family="Inter, Helvetica, sans-serif", size=12),
        )
        fig_heat.update_xaxes(tickangle=-40, side="bottom", showgrid=False)
        fig_heat.update_yaxes(showgrid=False)
        st.plotly_chart(fig_heat, use_container_width=True)

        st.markdown("---")

        # ── Attendance distribution histogram ─────────────────────────────────
        # A histogram divides the attendance % values into 25 equal-width buckets
        # and counts how many sessions fall into each bucket.
        # This tells us whether attendance is clustered high, spread out, or bimodal.
        st.markdown("#### Attendance Distribution Across All Sessions")
        fig_hist = px.histogram(
            df, x="Attendance%", nbins=25,
            color_discrete_sequence=[P["violet"]], opacity=0.82,
        )
        fig_hist.update_layout(
            xaxis_title="Attendance %", yaxis_title="Number of Sessions",
            height=240,
            plot_bgcolor=BG, paper_bgcolor=BG,
            margin=dict(l=0, r=0, t=20, b=0),
            bargap=0.05,  # Small gap between histogram bars for readability
            font=dict(family="Inter, Helvetica, sans-serif", size=12),
        )
        fig_hist.update_xaxes(showgrid=False)
        fig_hist.update_yaxes(showgrid=True, gridcolor=GRID)
        st.plotly_chart(fig_hist, use_container_width=True)

        st.markdown("---")

        # ── Full filtered data table ──────────────────────────────────────────
        # st.dataframe() renders a sortable, scrollable table.
        # We select only the most useful columns and rename verbose ones.
        st.markdown("#### 🗂️ Filtered Operational Data")
        st.caption(
            f"**{len(df):,} rows** match current filters · sorted by CLH · "
            "click any column header to re-sort."
        )
        display_cols = [
            "Volunteer name", "Center name", "State", "Donor",
            "Subject_clean", "Registered", "Reg Boys", "Reg Girls",
            "Enrolled", "En Boys", "En Girls",
            "Attendance%", "CLH", "Total hours(Comp+Offline)",
            "Planned", "Completed", "Offline", "Cancelled",
        ]
        # Only include columns that actually exist in the DataFrame
        # (defensive coding in case the schema changes between uploads)
        available_display = [c for c in display_cols if c in df.columns]
        display_df = df[available_display].rename(columns={
            "Subject_clean":             "Subject",
            "Total hours(Comp+Offline)": "Vol Hrs",
        }).copy()
        display_df["Attendance%"] = display_df["Attendance%"].round(1)
        st.dataframe(
            display_df.sort_values("CLH", ascending=False),
            use_container_width=True,
            hide_index=True,   # Don't show the pandas row index numbers
            height=420,
        )

    # =========================================================================
    # TAB 4 — DRM CLIENT REPORT
    #
    # This is the premium client-facing tab.  It loads the separate DRM workbook
    # (which contains session-level data, not just volunteer-level data),
    # and presents a polished, branded report across 8 sections.
    #
    # IMPORTANT: This tab does NOT use the sidebar filters.  It always shows
    # the full DRM dataset so clients see all centres and all donors.
    #
    # Data source: DRM_May_2026.xlsx — all 5 sheets loaded via load_drm_data()
    # =========================================================================
    with tab_drm:

        # Load the DRM workbook (cached — only reads from disk on first call)
        with st.spinner("Loading DRM data…"):
            drm_data = load_drm_data(DRM_DATA_PATH)

        # If the DRM file doesn't exist, show an error and stop this tab's rendering.
        # `return` exits the entire render_ops_dashboard() function here.
        if not drm_data:
            st.error(
                f"⚠️ DRM file not found at `{DRM_DATA_PATH}`. "
                "Place `DRM_May_2026.xlsx` alongside `app.py`."
            )
            return

        # Unpack all five DataFrames from the loaded dictionary.
        # These variables are used throughout the rest of Tab 4.
        sess = drm_data["sess"]   # Session-level log (1 row per session)
        drm  = drm_data["drm"]   # Centre-level monthly summary
        ac   = drm_data["ac"]    # Active centres master list
        od   = drm_data["od"]    # Offering details (grade × subject)
        ne   = drm_data["ne"]    # Newly enrolled students

        # ── Branded report header banner ──────────────────────────────────────
        # Renders an HTML gradient banner with the report title and summary badges.
        # This gives the tab a professional, client-ready appearance.
        st.markdown(
            """
            <div style='background:linear-gradient(135deg,#0f8a6e 0%,#185fa5 100%);
                        border-radius:14px;padding:26px 32px;margin-bottom:24px;'>
                <div style='display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px;'>
                    <div>
                        <h2 style='color:white;margin:0;font-size:1.7rem;font-weight:700;letter-spacing:-0.02em;'>
                            📊 eVidyaloka — Programme Impact Report
                        </h2>
                        <p style='color:rgba(255,255,255,0.82);margin:6px 0 0 0;font-size:0.95em;'>
                            Donor Relationship Management &nbsp;·&nbsp; Session Analytics
                            &nbsp;·&nbsp; Centre Performance &nbsp;·&nbsp; May 2026
                        </p>
                    </div>
                    <div style='display:flex;gap:8px;flex-wrap:wrap;'>
                        <span style='background:rgba(255,255,255,0.15);color:white;border-radius:20px;
                                     padding:5px 14px;font-size:0.78em;font-weight:600;'>58 Centres</span>
                        <span style='background:rgba(255,255,255,0.15);color:white;border-radius:20px;
                                     padding:5px 14px;font-size:0.78em;font-weight:600;'>4 States</span>
                        <span style='background:rgba(255,255,255,0.15);color:white;border-radius:20px;
                                     padding:5px 14px;font-size:0.78em;font-weight:600;'>4 Donors</span>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        # =====================================================================
        # DRM SECTION 1 — HEADLINE KPI METRICS
        # 10 KPI cards in 2 rows of 5, showing the programme's top-line numbers
        # for the month.  These are the first numbers a donor or board member sees.
        # =====================================================================

        # ── Compute all headline metrics from the session and centre DataFrames ──

        # Total rows in sess = total sessions this month (planned + already completed)
        total_sessions = len(sess)

        # Count sessions where the status column equals "Completed" / "Offline" / "Cancelled".
        # .sum() on a boolean Series counts the True values (True = 1, False = 0).
        comp_sessions  = int((sess["Session_status"] == "Completed").sum())
        off_sessions   = int((sess["Session_status"] == "Offline").sum())
        canc_sessions  = int((sess["Session_status"] == "Cancelled").sum())

        # "Present/CLH" = Child Learning Hours (CLH) per session.  Sum across all sessions.
        total_clh_drm  = int(sess["Present/CLH"].sum())

        # Total enrolled from the centre-level DRM sheet (not the session sheet)
        total_students = int(drm["Enrolled"].sum())

        # Average attendance % across all sessions (can return NaN if all are null)
        avg_att_drm    = sess["Attendance%"].mean()

        # Unique centre count from the DRM sheet (avoids double-counting)
        total_centres  = drm["Center Name"].nunique()

        # Unique volunteer IDs who appear in the session log this month
        drm_total_vols = sess["Volunteer_id"].nunique()

        # Completion rate: (Completed + Offline) / total planned × 100
        # Offline sessions are counted as completed because teaching happened.
        comp_rate      = (comp_sessions + off_sessions) / total_sessions * 100 if total_sessions else 0

        # Total enrolled girls and boys from the centre summary
        total_girls    = int(drm["En Girls"].sum())
        total_boys     = int(drm["En Boys"].sum())

        # Percentage of enrolled students who are girls
        girl_pct       = total_girls / (total_girls + total_boys) * 100 if (total_girls + total_boys) else 0

        # Count how many unique states have active centres this month
        num_states     = drm["State"].nunique()

        st.markdown("#### 🎯 Programme Snapshot — May 2026")

        # Create two rows of 5 columns each for the 10 KPI cards.
        # r1 = row 1, c1..c5 = columns 1-5.
        r1c1, r1c2, r1c3, r1c4, r1c5 = st.columns(5)
        r2c1, r2c2, r2c3, r2c4, r2c5 = st.columns(5)

        # Render 10 KPI cards using a loop over a list of tuples.
        # Each tuple contains: (column_object, label, value, colour, sub_label).
        # The loop calls _drm_kpi() for each entry and renders it in its column.
        # This is more concise than writing 10 identical col.markdown(_drm_kpi(...))
        # lines individually.
        for col_w, lbl, val, color, sub in [
            (r1c1, "Active Centres",    f"{total_centres}",           D["teal"],   f"across {num_states} states"),
            (r1c2, "Active Volunteers", f"{drm_total_vols}",          D["green"],  "this month"),
            (r1c3, "Total Sessions",    f"{total_sessions:,}",        D["blue"],   "planned this month"),
            (r1c4, "Completed",         f"{comp_sessions:,}",         D["teal"],   f"{comp_rate:.1f}% rate"),
            (r1c5, "Cancelled",         f"{canc_sessions}",           D["coral"],  f"{canc_sessions/total_sessions*100:.1f}% of planned"),
            (r2c1, "Total CLH",         f"{total_clh_drm:,}",         D["purple"], "child learning hours"),
            (r2c2, "Students Enrolled", f"{total_students:,}",        D["teal"],   "across all centres"),
            (r2c3, "Girl Students",     f"{total_girls:,}",           D["coral"],  f"{girl_pct:.0f}% of enrolled"),
            (r2c4, "Avg Attendance",    f"{avg_att_drm:.1f}%",        D["amber"],  "session attendance"),
            (r2c5, "Completion Rate",   f"{comp_rate:.1f}%",          D["green"],  "completed + offline"),
        ]:
            col_w.markdown(_drm_kpi(lbl, val, color, sub), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("---")

        # =====================================================================
        # DRM SECTION 2 — WEEKLY SESSION TREND + CLH BY DONOR
        # Left (wider): stacked bar chart of sessions per week with attendance
        #               percentage overlaid as a dotted line on a second Y-axis.
        # Right:        donut chart showing CLH split by donor.
        # =====================================================================
        st.markdown("#### 📅 Session Activity")
        # [1.6, 1] means the left column is 1.6x as wide as the right column
        col_t1, col_t2 = st.columns([1.6, 1])

        with col_t1:
            st.markdown("##### Weekly sessions with attendance overlay")

            # Aggregate session-level data by calendar week.
            # We create one row per week with counts of each session status and
            # the average attendance % for that week.
            # lambda x: (x == "Completed").sum() — a small anonymous function that
            # counts how many rows in the group have status "Completed".
            weekly = (
                sess.groupby("week").agg(
                    Completed =("Session_status", lambda x: (x == "Completed").sum()),
                    Offline   =("Session_status", lambda x: (x == "Offline").sum()),
                    Cancelled =("Session_status", lambda x: (x == "Cancelled").sum()),
                    CLH       =("Present/CLH",    "sum"),
                    Attendance=("Attendance%",    "mean"),
                ).reset_index().sort_values("week")
            )

            # Melt the three status columns into long format for the stacked bar chart
            weekly_melt = weekly.melt(
                id_vars="week",
                value_vars=["Completed", "Offline", "Cancelled"],
                var_name="Status", value_name="Sessions"
            )

            # Create the stacked bar chart for session counts
            fig_week = px.bar(
                weekly_melt, x="week", y="Sessions", color="Status",
                barmode="stack", text="Sessions",
                color_discrete_map={
                    "Completed": D["teal"],
                    "Offline":   D["amber"],
                    "Cancelled": D["red"],
                },
            )
            fig_week.update_traces(
                textposition="inside", texttemplate="%{text}",
                hovertemplate="<b>%{x}</b><br>%{fullData.name}: %{y}<extra></extra>",
            )

            # Overlay a dotted line showing average attendance % on a secondary Y-axis.
            # WHY secondary axis: Attendance % (0-100) and session counts (0-500)
            # are on completely different scales.  A second Y-axis lets both be shown
            # accurately on the same chart without distorting either scale.
            # yaxis="y2" tells Plotly to use the second axis for this trace.
            fig_week.add_trace(go.Scatter(
                x=weekly["week"], y=weekly["Attendance"].round(1),
                name="Avg Attendance %",
                mode="lines+markers+text",    # Show both a line and dot markers
                text=weekly["Attendance"].round(1).astype(str) + "%",
                textposition="top center",
                line=dict(color="#2d3436", width=2.5, dash="dot"),
                marker=dict(size=8, color="#2d3436"),
                yaxis="y2",                   # ← this trace uses the right-hand axis
                hovertemplate="Week: %{x}<br>Attendance: %{y:.1f}%<extra></extra>",
            ))
            fig_week.update_layout(
                height=380,
                plot_bgcolor=BG, paper_bgcolor=BG,
                margin=dict(l=0, r=60, t=30, b=60),
                font=dict(family="Inter, Helvetica, sans-serif", size=12),
                legend=dict(orientation="h", yanchor="top", y=-0.18,
                            xanchor="center", x=0.5, title_text=""),
                yaxis =dict(title="Sessions", showgrid=True, gridcolor=GRID),
                # yaxis2 is the secondary axis on the right for the attendance line
                yaxis2=dict(title="Attendance %", overlaying="y", side="right",
                            range=[60, 105], showgrid=False, ticksuffix="%"),
                xaxis =dict(title="", showgrid=False, linecolor="#dee2e6", tickangle=-20),
            )
            st.plotly_chart(fig_week, use_container_width=True)

        with col_t2:
            st.markdown("##### CLH split by donor")

            # Group session-level CLH by donor and sort to put largest slice first
            donor_clh = (
                sess.groupby("Donor")["Present/CLH"].sum()
                .reset_index()
                .rename(columns={"Present/CLH": "CLH"})
                .sort_values("CLH", ascending=False)
            )

            # Donut chart showing each donor's share of total CLH.
            # hole=0.58 means 58% of the chart area is the hollow centre.
            # The centre annotation shows the grand total CLH.
            fig_donut = px.pie(
                donor_clh, names="Donor", values="CLH",
                hole=0.58,
                color_discrete_sequence=[D["blue"], D["teal"], D["purple"], D["amber"], D["gray"]],
            )
            fig_donut.update_traces(
                textposition="outside",
                texttemplate="<b>%{label}</b><br>%{percent:.1%}",
                hovertemplate="<b>%{label}</b><br>CLH: %{value:,}<extra></extra>",
                pull=[0.03] * len(donor_clh),  # Slightly explode all slices outward
            )
            fig_donut.update_layout(
                height=380, showlegend=False,
                plot_bgcolor=BG, paper_bgcolor=BG,
                margin=dict(l=10, r=10, t=30, b=10),
                annotations=[dict(
                    text=f"<b>{total_clh_drm:,}</b><br>Total CLH",
                    x=0.5, y=0.5, font_size=14, showarrow=False,
                    font=dict(color="#2d3436"),
                )],
            )
            st.plotly_chart(fig_donut, use_container_width=True)

        st.markdown("---")

        # =====================================================================
        # DRM SECTION 3 — CENTRE PERFORMANCE SCATTER + DROPOUT BAR
        # Left: Bubble scatter plot — each bubble is one centre, positioned by
        #       attendance (X) and completion (Y), sized by CLH, coloured by donor.
        # Right: Horizontal bar chart showing the 15 centres with highest dropout.
        # =====================================================================
        st.markdown("#### 🏫 Centre Performance")
        col_m1, col_m2 = st.columns([1.3, 1])

        with col_m1:
            st.markdown("##### Attendance vs completion — bubble size = CLH")
            st.caption(
                "Top-right quadrant = high attendance AND high completion. "
                "Bubble size = CLH delivered."
            )

            # Compute the median for both axes to draw reference lines.
            # The median divides centres into "above average" and "below average" halves.
            med_att  = drm["Attendance %"].median()
            med_comp = drm["Completion%"].median()

            # px.scatter() creates a scatter plot.
            # size="Live_CLH" scales each dot's radius by that centre's CLH total.
            # color="Donor Name" assigns a colour per donor automatically.
            # hover_name="Center Name" makes the centre name the tooltip title.
            # hover_data adds extra fields to the tooltip, formatted with :.1f (1 decimal).
            # size_max=48 caps the maximum bubble size to prevent huge bubbles obscuring others.
            fig_scatter = px.scatter(
                drm,
                x="Attendance %", y="Completion%",
                size="Live_CLH", color="Donor Name",
                hover_name="Center Name",
                hover_data={
                    "State":         True,
                    "Enrolled":      True,
                    "Dropout%":      ":.1f",
                    "Cancellation%": ":.1f",
                    "Live_CLH":      ":,",
                    "Donor Name":    False,   # Already shown as the bubble colour legend
                },
                color_discrete_sequence=[D["blue"], D["teal"], D["purple"], D["amber"]],
                size_max=48,
            )

            # Draw dashed reference lines at the medians to create four quadrants:
            # Top-right = high attendance AND high completion (best centres)
            # Bottom-left = low attendance AND low completion (centres needing support)
            fig_scatter.add_vline(x=med_att,  line_dash="dot", line_color="#b2bec3",
                                  annotation_text=f"Med att: {med_att:.0f}%",
                                  annotation_position="top right",
                                  annotation_font_size=10)
            fig_scatter.add_hline(y=med_comp, line_dash="dot", line_color="#b2bec3",
                                  annotation_text=f"Med comp: {med_comp:.0f}%",
                                  annotation_position="top right",
                                  annotation_font_size=10)
            fig_scatter.update_layout(
                height=460,
                plot_bgcolor=BG, paper_bgcolor=BG,
                margin=dict(l=0, r=0, t=20, b=0),
                font=dict(family="Inter, Helvetica, sans-serif", size=12),
                xaxis=dict(title="Attendance %", showgrid=True, gridcolor=GRID,
                           ticksuffix="%", range=[30, 105]),
                yaxis=dict(title="Completion %", showgrid=True, gridcolor=GRID,
                           ticksuffix="%", range=[30, 110]),
                legend=dict(title="Donor", orientation="h",
                            yanchor="top", y=-0.12, xanchor="center", x=0.5),
            )
            st.plotly_chart(fig_scatter, use_container_width=True)

        with col_m2:
            st.markdown("##### Student drop-off — top 15 centres")
            st.caption("Drop-off = (Registered − Enrolled) / Registered × 100")

            # Find the 15 centres with the highest dropout and sort ascending
            # (lowest dropout at bottom) so the worst centre appears at the top
            # of the horizontal bar chart.
            dropout_df = (
                drm[["Center Name", "Dropout%", "Enrolled", "Registered", "State"]]
                .sort_values("Dropout%", ascending=False)
                .head(15)
                .sort_values("Dropout%", ascending=True)  # Reverse for horizontal display
            )
            fig_drop2 = px.bar(
                dropout_df, x="Dropout%", y="Center Name",
                orientation="h", text="Dropout%",
                color="Dropout%",
                # Three-stop gradient: amber → orange → red (escalating concern)
                color_continuous_scale=["#fdcb6e", "#e17055", "#ed1c2d"],
                hover_data={"Registered": True, "Enrolled": True, "State": True},
            )
            fig_drop2.update_traces(
                texttemplate="%{text:.1f}%", textposition="outside",
                marker_line_width=0,
            )
            fig_drop2.update_coloraxes(showscale=False)
            _layout(fig_drop2, height=460, margin=dict(l=0, r=70, t=20, b=0))
            fig_drop2.update_layout(xaxis_title="Drop-off %", yaxis_title="")
            fig_drop2.update_xaxes(showgrid=True, gridcolor=GRID, ticksuffix="%")
            fig_drop2.update_yaxes(showgrid=False)
            st.plotly_chart(fig_drop2, use_container_width=True)

        st.markdown("---")

        # =====================================================================
        # DRM SECTION 4 — STATE CLH + GENDER SPLIT
        # Left: Horizontal bar chart of CLH per state.
        # Right: Stacked horizontal bar chart of boys vs girls per state.
        # Together these tell donors which states have the most impact and
        # whether girl students are being reached in each geography.
        # =====================================================================
        st.markdown("#### 🗺️ State & Gender Analysis")
        col_sg1, col_sg2 = st.columns(2)

        with col_sg1:
            st.markdown("##### State-wise CLH and completion")

            # Aggregate all centre-level metrics by state.
            # This gives us one summary row per state for the chart.
            state_stats = (
                drm.groupby("State").agg(
                    Centres    =("Center Name",    "count"),
                    Volunteers =("Live Volunteers","sum"),
                    CLH        =("Live_CLH",       "sum"),
                    Enrolled   =("Enrolled",        "sum"),
                    Completed  =("Completed",       "sum"),
                    Planned    =("Planned",         "sum"),
                    Girls      =("En Girls",        "sum"),
                    Boys       =("En Boys",         "sum"),
                ).reset_index()
            )
            # Compute derived metrics for this state-level summary
            state_stats["Comp%"] = (
                state_stats["Completed"] / state_stats["Planned"].replace(0, pd.NA) * 100
            ).round(1).fillna(0)
            state_stats["Girl%"] = (
                state_stats["Girls"] / (state_stats["Girls"] + state_stats["Boys"]).replace(0, pd.NA) * 100
            ).round(1).fillna(0)
            state_stats = state_stats.sort_values("CLH", ascending=True)

            # Use go.Figure() for explicit control over bar colour and formatting.
            # lambda v: f"{v:,.0f}" formats CLH values with thousands commas.
            fig_state_clh = go.Figure()
            fig_state_clh.add_trace(go.Bar(
                name="CLH", x=state_stats["CLH"], y=state_stats["State"],
                orientation="h", marker_color=D["teal"],
                text=state_stats["CLH"].apply(lambda v: f"{v:,.0f}"),
                textposition="outside",
                hovertemplate="<b>%{y}</b><br>CLH: %{x:,}<extra></extra>",
            ))
            fig_state_clh.update_layout(
                height=320,
                plot_bgcolor=BG, paper_bgcolor=BG,
                margin=dict(l=0, r=80, t=10, b=0),
                font=dict(family="Inter, Helvetica, sans-serif", size=12),
                xaxis=dict(showgrid=True, gridcolor=GRID, title="Child Learning Hours"),
                yaxis=dict(showgrid=False, title=""),
                showlegend=False,
            )
            st.plotly_chart(fig_state_clh, use_container_width=True)

        with col_sg2:
            st.markdown("##### Gender split by state (enrolled students)")

            # Aggregate boys and girls by state from the centre-level DRM data
            state_gen = drm.groupby("State")[["En Boys", "En Girls"]].sum().reset_index()

            # Melt into long format: one row per state-gender combination
            state_gen_melt = state_gen.melt(
                id_vars="State",
                value_vars=["En Boys", "En Girls"],
                var_name="Gender", value_name="Students"
            )
            state_gen_melt["Gender"] = state_gen_melt["Gender"].map(
                {"En Boys": "Boys", "En Girls": "Girls"}
            )
            fig_state_gen = px.bar(
                state_gen_melt, x="Students", y="State",
                color="Gender", barmode="stack", orientation="h",
                text="Students",
                color_discrete_map={"Boys": D["blue"], "Girls": D["coral"]},
            )
            fig_state_gen.update_traces(
                textposition="inside", texttemplate="%{text:,}",
                hovertemplate="<b>%{y}</b> · %{fullData.name}: %{x:,}<extra></extra>",
            )
            _layout(fig_state_gen, height=320, legend_bottom=True,
                    margin=dict(l=0, r=0, t=10, b=60))
            fig_state_gen.update_layout(xaxis_title="Students", yaxis_title="")
            fig_state_gen.update_xaxes(showgrid=True, gridcolor=GRID)
            fig_state_gen.update_yaxes(showgrid=False)
            st.plotly_chart(fig_state_gen, use_container_width=True)

        st.markdown("---")

        # =====================================================================
        # DRM SECTION 5 — DAY-OF-WEEK PATTERNS + CANCELLATION ANALYSIS
        # Left: Stacked bar (Mon-Sat) with attendance overlay line.
        # Right: Donut chart of cancellation reasons + insight callout box.
        # =====================================================================
        st.markdown("#### 📆 Session Patterns & Cancellations")
        col_d1, col_d2 = st.columns([1.4, 1])

        with col_d1:
            st.markdown("##### Activity by day of week")

            # We define the correct calendar order explicitly because groupby sorts
            # alphabetically by default ("Friday", "Monday", "Saturday"…) which
            # would make a nonsensical chart.  pd.Categorical with ordered=True
            # tells pandas to treat this as an ordered categorical variable.
            dow_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
            dow_stats = (
                sess.groupby("dow").agg(
                    Completed=("Session_status", lambda x: (x == "Completed").sum()),
                    Offline  =("Session_status", lambda x: (x == "Offline").sum()),
                    Cancelled=("Session_status", lambda x: (x == "Cancelled").sum()),
                    Avg_Att  =("Attendance%",    "mean"),
                ).reset_index()
            )
            # Convert to ordered categorical and sort so Monday comes first
            dow_stats["dow"] = pd.Categorical(
                dow_stats["dow"], categories=dow_order, ordered=True
            )
            dow_stats = dow_stats.sort_values("dow")
            dow_stats["dow"] = dow_stats["dow"].astype(str)  # Back to string for Plotly

            # Melt session status columns for stacked bar
            dow_melt = dow_stats.melt(
                id_vars=["dow", "Avg_Att"],
                value_vars=["Completed", "Offline", "Cancelled"],
                var_name="Status", value_name="Sessions"
            )
            fig_dow = px.bar(
                dow_melt, x="dow", y="Sessions", color="Status",
                barmode="stack", text="Sessions",
                color_discrete_map={
                    "Completed": D["teal"],
                    "Offline":   D["amber"],
                    "Cancelled": D["red"],
                },
                # category_orders tells Plotly to display days in our specified order
                category_orders={"dow": dow_order},
            )
            fig_dow.update_traces(
                textposition="inside", texttemplate="%{text}",
                hovertemplate="<b>%{x}</b><br>%{fullData.name}: %{y}<extra></extra>",
            )
            # Overlay attendance % as a dotted line on a second Y-axis (same technique as weekly chart)
            fig_dow.add_trace(go.Scatter(
                x=dow_stats["dow"], y=dow_stats["Avg_Att"].round(1),
                name="Avg Attendance %",
                mode="lines+markers",
                line=dict(color="#2d3436", width=2, dash="dot"),
                marker=dict(size=7, color="#2d3436"),
                yaxis="y2",
                hovertemplate="%{x}<br>Avg Att: %{y:.1f}%<extra></extra>",
            ))
            fig_dow.update_layout(
                height=340,
                plot_bgcolor=BG, paper_bgcolor=BG,
                margin=dict(l=0, r=60, t=20, b=60),
                font=dict(family="Inter, Helvetica, sans-serif", size=12),
                legend=dict(orientation="h", yanchor="top", y=-0.2,
                            xanchor="center", x=0.5, title_text=""),
                yaxis =dict(title="Sessions", showgrid=True, gridcolor=GRID, zeroline=False),
                yaxis2=dict(title="Att %", overlaying="y", side="right",
                            range=[70, 100], showgrid=False, ticksuffix="%"),
                xaxis =dict(title="", showgrid=False, linecolor="#dee2e6"),
            )
            st.plotly_chart(fig_dow, use_container_width=True)

        with col_d2:
            st.markdown("##### Cancellation reasons")

            # Filter to only cancelled sessions, then count by reason.
            # .fillna("Not specified") handles rows where Cancel_reason is blank.
            canc_df = (
                sess[sess["Session_status"] == "Cancelled"]
                ["Cancel_reason"].fillna("Not specified")
                .value_counts().reset_index()
            )
            canc_df.columns = ["Reason", "Count"]

            fig_canc = px.pie(
                canc_df, names="Reason", values="Count",
                hole=0.5,
                color_discrete_sequence=[D["coral"], D["amber"], D["red"]],
            )
            total_canc = canc_df["Count"].sum()
            fig_canc.update_traces(
                texttemplate="<b>%{label}</b><br>%{value} sessions<br>%{percent:.0%}",
                textposition="outside",
                hovertemplate="<b>%{label}</b><br>Sessions: %{value}<extra></extra>",
                pull=[0.04] * len(canc_df),
            )
            fig_canc.update_layout(
                height=340, showlegend=False,
                plot_bgcolor=BG, paper_bgcolor=BG,
                margin=dict(l=10, r=10, t=20, b=10),
                annotations=[dict(
                    text=f"<b>{total_canc}</b><br>cancelled",
                    x=0.5, y=0.5, font_size=14, showarrow=False,
                    font=dict(color="#2d3436"),
                )],
            )
            st.plotly_chart(fig_canc, use_container_width=True)

            # Programmatically generate an insight callout from the actual data.
            # We only show this if there are any cancellations to analyse.
            # .iloc[0] accesses the first row — the most common reason (already sorted
            # descending by value_counts() above).
            if not canc_df.empty:
                top_reason  = canc_df.iloc[0]["Reason"]
                top_pct     = canc_df.iloc[0]["Count"] / total_canc * 100
                st.markdown(
                    _insight_box(
                        f"{top_pct:.0f}% of cancellations are due to <b>{top_reason}</b>. "
                        "Proactive volunteer scheduling support could recover these sessions."
                    ),
                    unsafe_allow_html=True,
                )

        st.markdown("---")

        # =====================================================================
        # DRM SECTION 6 — ACADEMIC COVERAGE
        # Three charts showing the curriculum landscape:
        # 1. Session heatmap: subject (row) × grade (column) = session count
        # 2. Average attendance % per subject (horizontal bar)
        # 3. Enrolled student heatmap: grade × subject from Offering Details
        # =====================================================================
        st.markdown("#### 📚 Academic Coverage & Attendance")
        col_s1, col_s2 = st.columns([1.3, 1])

        with col_s1:
            st.markdown("##### Sessions by subject & grade")

            # Count how many sessions occurred for each subject-grade combination
            subj_grade = (
                sess.groupby(["Subject_clean", "Grade"])["Session_id"]
                .count().reset_index()
                .rename(columns={"Session_id": "Sessions"})
            )

            # Pivot into a 2D matrix: subject rows, grade columns.
            # .fillna(0) fills missing combinations with zero.
            pivot_sg = subj_grade.pivot(
                index="Subject_clean", columns="Grade", values="Sessions"
            ).fillna(0)

            fig_sg = px.imshow(
                pivot_sg,
                color_continuous_scale=["#dfe6e9", D["teal"]],
                aspect="auto", text_auto=".0f",
            )
            fig_sg.update_traces(
                hovertemplate="<b>%{y}</b> · Grade %{x}<br>Sessions: %{z:.0f}<extra></extra>"
            )
            fig_sg.update_layout(
                height=320,
                plot_bgcolor=BG, paper_bgcolor=BG,
                margin=dict(l=0, r=0, t=20, b=0),
                coloraxis_showscale=False,
                xaxis=dict(title="Grade", showgrid=False),
                yaxis=dict(title="", showgrid=False),
                font=dict(family="Inter, Helvetica, sans-serif", size=12),
            )
            st.plotly_chart(fig_sg, use_container_width=True)

        with col_s2:
            st.markdown("##### Avg attendance % by subject")

            # Average attendance per subject across all sessions
            att_subj = (
                sess.groupby("Subject_clean")["Attendance%"]
                .mean().reset_index()
                .rename(columns={"Attendance%": "Avg Attendance %"})
                .sort_values("Avg Attendance %", ascending=True)  # Lowest at bottom
            )

            # A colour-gradient bar chart that visually highlights which subjects
            # have the lowest attendance (red) vs highest (teal).
            # range_color=[40, 100] anchors the colour scale so 40%=red, 100%=teal.
            fig_att = px.bar(
                att_subj, x="Avg Attendance %", y="Subject_clean",
                orientation="h",
                # Format the text label as a percentage string (e.g. "84.1%")
                text=att_subj["Avg Attendance %"].apply(lambda x: f"{x:.1f}%"),
                color="Avg Attendance %",
                color_continuous_scale=["#e17055", D["amber"], D["teal"]],
                range_color=[40, 100],
            )
            fig_att.update_traces(
                textposition="outside", marker_line_width=0,
                hovertemplate="<b>%{y}</b><br>Avg Attendance: %{x:.1f}%<extra></extra>",
            )
            fig_att.update_coloraxes(showscale=False)
            _layout(fig_att, height=320, margin=dict(l=0, r=70, t=20, b=0))
            fig_att.update_layout(
                xaxis_title="Avg Attendance %", yaxis_title="",
                xaxis=dict(range=[30, 115], ticksuffix="%",
                           showgrid=True, gridcolor=GRID),
            )
            fig_att.update_yaxes(showgrid=False)
            st.plotly_chart(fig_att, use_container_width=True)

        # ── Grade × Subject enrolment heatmap (from Offering Details sheet) ───
        # This uses od (Offering Details) rather than sess (Session Dump).
        # od has the "Enrolled_count" column which sess does not.
        # This shows WHERE students are enrolled (grade + subject), not just
        # where sessions happened.
        st.markdown("##### Enrolled students — grade × subject heatmap")
        st.caption("Cell value = total students enrolled in that grade-subject combination.")
        grade_subj_enr = (
            od.groupby(["Grade", "Subject_clean"])["Enrolled_count"]
            .sum().reset_index()
        )
        # Pivot: rows = Grade, columns = Subject, values = enrolled count
        pivot_gs = grade_subj_enr.pivot(
            index="Grade", columns="Subject_clean", values="Enrolled_count"
        ).fillna(0)

        # A lighter teal gradient for this heatmap to distinguish it from the
        # session heatmap above.
        fig_gs_heat = px.imshow(
            pivot_gs,
            color_continuous_scale=["#e1f5ee", D["teal"]],
            aspect="auto", text_auto=".0f",
        )
        fig_gs_heat.update_traces(
            hovertemplate="Grade %{y} · <b>%{x}</b><br>Enrolled: %{z:.0f}<extra></extra>"
        )
        fig_gs_heat.update_layout(
            height=280,
            plot_bgcolor=BG, paper_bgcolor=BG,
            margin=dict(l=0, r=0, t=10, b=60),
            coloraxis_showscale=False,
            xaxis=dict(title="", showgrid=False, tickangle=-30),
            yaxis=dict(title="Grade", showgrid=False),
            font=dict(family="Inter, Helvetica, sans-serif", size=12),
        )
        st.plotly_chart(fig_gs_heat, use_container_width=True)

        st.markdown("---")

        # =====================================================================
        # DRM SECTION 7 — VOLUNTEER PERFORMANCE TABLE + NEW ENROLMENTS
        # Left: Interactive sortable table of all volunteers with CLH,
        #       session count, completion %, attendance %, and a health flag.
        # Right: Visual cards for each newly enrolled centre (from ne sheet).
        # =====================================================================
        st.markdown("#### 🙋 Volunteer Performance & New Enrolments")
        col_v1, col_v2 = st.columns([1.4, 1])

        with col_v1:
            st.markdown("##### Volunteer session scorecard")

            # Build a volunteer-level summary from session-level data.
            # Each volunteer may appear in hundreds of session rows.
            # We group by volunteer ID and name to collapse to one row per volunteer.
            vol_perf = (
                sess.groupby(["Volunteer_id", "Teacher_name"]).agg(
                    Sessions  =("Session_id",     "count"),           # Total sessions assigned
                    Completed =("Session_status", lambda x: (x == "Completed").sum()),  # Sessions completed
                    CLH       =("Present/CLH",    "sum"),             # Total CLH delivered
                    Avg_Att   =("Attendance%",    "mean"),            # Mean attendance across sessions
                    Students  =("Total students", "sum"),             # Total student-sessions
                    # Join all unique subjects this volunteer has taught, sorted alphabetically
                    Subjects  =("Subject_clean",  lambda x: ", ".join(sorted(x.unique()))),
                ).reset_index()
            )
            vol_perf["Comp %"]  = (vol_perf["Completed"] / vol_perf["Sessions"] * 100).round(1)
            vol_perf["Avg_Att"] = vol_perf["Avg_Att"].round(1)

            # _flag() assigns a health status based on completion rate.
            # Applied row-by-row with .apply(fn, axis=1).
            # axis=1 means "pass each row as a Series to the function".
            def _flag(row):
                if row["Comp %"] < 60:  return "🔴 At Risk"
                if row["Comp %"] < 85:  return "🟡 Moderate"
                return "🟢 Strong"

            vol_perf["Status"] = vol_perf.apply(_flag, axis=1)

            # Select and rename columns for display — hide internal IDs
            vol_display = vol_perf[[
                "Teacher_name", "Subjects", "Sessions", "Comp %",
                "Avg_Att", "CLH", "Students", "Status"
            ]].rename(columns={
                "Teacher_name": "Volunteer",
                "Avg_Att":      "Att %",
            }).sort_values("CLH", ascending=False)  # Highest CLH first

            # st.dataframe() with column_config gives Streamlit's native progress bars
            # and number formatting inside the table.  ProgressColumn renders a mini
            # bar chart inside the cell showing the percentage visually.
            st.dataframe(
                vol_display,
                use_container_width=True,
                hide_index=True,
                height=400,
                column_config={
                    "CLH":    st.column_config.NumberColumn("CLH",    format="%d"),
                    "Comp %": st.column_config.ProgressColumn(
                        "Comp %", min_value=0, max_value=100, format="%.1f%%"
                    ),
                    "Att %":  st.column_config.ProgressColumn(
                        "Att %",  min_value=0, max_value=100, format="%.1f%%"
                    ),
                },
            )

        with col_v2:
            st.markdown("##### 🆕 Newly enrolled students")
            st.caption("Students enrolled mid-month at centres with new registrations.")

            # Select only the columns we want to display and copy to avoid modifying ne
            ne_display = ne[[
                "Center_name", "State", "Donor", "Grade", "Enrolled", "Boys", "Girls"
            ]].copy()

            # Compute Girl % for display.
            # .replace(0, pd.NA) prevents 0÷0 = NaN instead of a crash.
            # .round(0) rounds to nearest integer.  .astype(int).astype(str) + "%" gives
            # a clean "100%" string rather than "100.0%".
            ne_display["Girl %"] = (
                ne_display["Girls"] / ne_display["Enrolled"].replace(0, pd.NA) * 100
            ).round(0).fillna(0).astype(int).astype(str) + "%"

            # Loop over each new-enrolment row and render a custom HTML card.
            # _ is a Python convention for "I don't need the index value".
            # row is a pandas Series (like a dict of column: value pairs).
            for _, row in ne_display.iterrows():
                # Calculate the width % for the mini girl-% progress bar (0-100)
                girl_bar_w = int(row["Girls"] / row["Enrolled"] * 100) if row["Enrolled"] > 0 else 0

                # f-strings (f"...") allow embedding Python expressions {like_this}
                # directly inside the HTML string.  D["teal"] injects the hex colour.
                st.markdown(
                    f"""<div style='background:#f8f9fa;border-radius:10px;
                            padding:12px 16px;margin-bottom:10px;
                            border-left:5px solid {D["teal"]};'>
                        <div style='font-weight:700;font-size:0.93em;color:#2d3436;'>
                            {row['Center_name']}
                        </div>
                        <div style='color:#636e72;font-size:0.78em;margin:2px 0 8px;'>
                            {row['State']} &nbsp;·&nbsp; {row['Donor']} &nbsp;·&nbsp; Grade {row['Grade']}
                        </div>
                        <div style='display:flex;gap:16px;font-size:0.83em;margin-bottom:8px;'>
                            <span style='color:{D["teal"]};font-weight:700;'>
                                {row['Enrolled']} enrolled
                            </span>
                            <span style='color:#636e72;'>👦 {row['Boys']}</span>
                            <span style='color:#636e72;'>👧 {row['Girls']} ({row['Girl %']})</span>
                        </div>
                        <!-- Mini progress bar showing girl %; width is computed above -->
                        <div style='height:4px;background:#dee2e6;border-radius:2px;'>
                            <div style='height:100%;width:{girl_bar_w}%;
                                        background:{D["coral"]};border-radius:2px;'></div>
                        </div>
                        <div style='font-size:0.7em;color:#b2bec3;margin-top:3px;'>girl %</div>
                    </div>""",
                    unsafe_allow_html=True,
                )

        st.markdown("---")

        # =====================================================================
        # DRM SECTION 8 — FULL CENTRE SCORECARD TABLE
        # A sortable, scrollable table showing every centre's key metrics.
        # Uses Streamlit's column_config to render progress bars for % columns.
        # Sorted by CLH descending by default (highest-impact centres first).
        # =====================================================================
        st.markdown("#### 🗂️ Full Centre Scorecard")
        st.caption(
            f"All {total_centres} active centres ranked by CLH delivered. &nbsp;"
            "🟢 ≥ 90% completion &nbsp; 🟡 70–90% &nbsp; 🔴 < 70% &nbsp;|&nbsp; "
            "Click any column header to re-sort."
        )

        # Select the columns we want to show and copy to avoid mutating the cached drm
        drm_display = drm[[
            "Center Name", "State", "Donor Name",
            "Live Volunteers", "Planned", "Completed", "Offline", "Cancelled",
            "Completion%", "Cancellation%",
            "Live_CLH", "Enrolled", "En Boys", "En Girls",
            "Attendance %", "Dropout%", "Girl%"
        ]].copy().rename(columns={
            # Map to shorter, display-friendly column headers
            "Donor Name":     "Donor",
            "Live Volunteers":"Vols",
            "Completion%":    "Comp %",
            "Cancellation%":  "Canc %",
            "Live_CLH":       "CLH",
            "Attendance %":   "Att %",
            "Dropout%":       "Drop %",
            "Girl%":          "Girl %",
        })

        # Round all percentage columns to 1 decimal for consistent display
        drm_display["Att %"]  = drm_display["Att %"].round(1)
        drm_display["Comp %"] = drm_display["Comp %"].round(1)
        drm_display["Drop %"] = drm_display["Drop %"].round(1)
        drm_display["Girl %"] = drm_display["Girl %"].round(1)
        drm_display["Canc %"] = drm_display["Canc %"].round(1)

        # Sort by CLH descending so the most impactful centres appear at the top
        drm_display = drm_display.sort_values("CLH", ascending=False)

        # Render the dataframe with rich column configurations.
        # ProgressColumn renders a mini horizontal bar chart inside the cell.
        # NumberColumn controls the display format (e.g. "93.3%" vs "93.3").
        st.dataframe(
            drm_display,
            use_container_width=True,
            hide_index=True,
            height=520,
            column_config={
                "Comp %": st.column_config.ProgressColumn(
                    "Comp %", min_value=0, max_value=100, format="%.1f%%"
                ),
                "Att %":  st.column_config.ProgressColumn(
                    "Att %",  min_value=0, max_value=100, format="%.1f%%"
                ),
                "CLH":    st.column_config.NumberColumn("CLH",    format="%d"),
                "Drop %": st.column_config.NumberColumn("Drop %", format="%.1f%%"),
                "Girl %": st.column_config.NumberColumn("Girl %", format="%.1f%%"),
                "Canc %": st.column_config.NumberColumn("Canc %", format="%.1f%%"),
            },
        )

        # ── Report footer ──────────────────────────────────────────────────────
        # A centred, muted footer line summarising the report scope.
        # The f-string injects the live computed values so it always reflects
        # the actual data rather than being a hardcoded string.
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            f"""<div style='text-align:center;color:#b2bec3;font-size:0.75em;
                    padding:16px;border-top:1px solid #dee2e6;margin-top:8px;'>
                eVidyaloka Programme Impact Report &nbsp;·&nbsp; May 2026
                &nbsp;·&nbsp; {total_centres} centres &nbsp;·&nbsp;
                {drm_total_vols} volunteers &nbsp;·&nbsp;
                {total_students:,} students &nbsp;·&nbsp;
                {total_clh_drm:,} child learning hours
            </div>""",
            unsafe_allow_html=True,
        )
# =============================================================================
# END OF FILE
#
# To run this dashboard, ensure both Excel files are in the same directory
# as this script, then run:
#
#   streamlit run app.py
#
# where app.py imports and calls render_ops_dashboard() from this module.
# The function builds the full UI each time it is called.  Streamlit re-calls
# it on every user interaction (filter change, tab click, etc.).
# =============================================================================
