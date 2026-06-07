"""
ops_dashboard.py
================
Operations & Impact Command Center  —  eVidyaloka VRM
Data source: VRM_May__2026.xlsx (multi-sheet Excel)

Sheets used
───────────
  Active VT            — primary operational data (one row per vol-offering)
  Dropped VT           — volunteers who dropped their offering(s)
  Newly Registered VT  — all volunteers who registered in the current period

Column mapping (new file → dashboard internal name)
────────────────────────────────────────────────────
  Volunteer ID              → vol_id
  Volunteer Name            → vol_name
  EV Joined Date            → joined_dt
  Volunteer State           → res_state
  Volunteer City            → res_city
  Center Name               → center
  Center State              → state
  Offering Status           → status
  Total Hours (Comp+Offline)→ vol_hrs
  Attendance %              → attendance  (already float in this file)
  En Boys / En Girls        → gender split (NEW — not in old CSV)
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────
DATA_PATH = "VRM_May__2026.xlsx"

# Sheets
SHEET_ACTIVE  = "Active VT"
SHEET_DROPPED = "Dropped VT"
SHEET_NEW_REG = "Newly Registered VT "   # trailing space is in the file

# ── Colour palette ────────────────────────────────────────────────────────────
P = {
    "teal":    "#0094c9",
    "green":   "#00964d",
    "orange":  "#f27c48",
    "red":     "#ed1c2d",
    "violet":  "#6c5ce7",
    "amber":   "#fdcb6e",
    "sky":     "#74b9ff",
    "mint":    "#00b894",
    "coral":   "#e17055",
    "lavender":"#a29bfe",
    "salmon":  "#fab1a0",
    "aqua":    "#55efc4",
}
SEQ  = list(P.values())
BG   = "rgba(0,0,0,0)"
GRID = "#e9ecef"

# ── Subject normalisation ─────────────────────────────────────────────────────
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
    "Concise Content 1":              "Concise Content",
    "Concise content 2":              "Concise Content",
    "Guest Sessions":                 "Guest Sessions",
    "Scholarship":                    "Scholarship",
    "Reading Program":                "Reading Program",
}

# ── Profession normalisation ──────────────────────────────────────────────────
PROFESSION_MAP = {
    "corporates":                              "Corporate Professional",
    "others":                                  "Others",
    "student_ug":                              "Student (UG)",
    "home_makers":                             "Home Maker",
    "Housewife":                               "Home Maker",
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
    "top_managemenet":                         "Management",
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

# ── Reference channel bucketing ───────────────────────────────────────────────
def _group_ref(ref: str) -> str:
    r = str(ref).strip()
    if r in ("Internet Search", "Facebook", "Community Outreach"):
        return "Online / Direct"
    if r in ("Emailer", "DCP Campaign", "eVidyaloka", "eVidyaloka Trust"):
        return "eVidyaloka Campaign"
    if r in ("Word of Mouth",):
        return "Word of Mouth"
    if any(k in r for k in (
        "KPMG", "Infosys", "Tech Mahindra", "HPInc", "HPE", "L&T",
        "HSBC", "EY", "Adobe", "Brillio", "Broadridge", "Accenture",
        "CISCO", "Fidelity", "ConnectFor", "Microsoft", "Cognizant",
        "Firstsource", "Pricewaterhousecoopers", "Reliance Foundation",
        "Scaler", "Joy of Reading",
    )):
        return "Corporate / Partner Referral"
    if any(k in r for k in (
        "College", "University", "BMS", "NIT", "IIT", "IIM",
    )):
        return "Academic Institution"
    if any(k in r for k in (
        "Bhumi", "Udaan", "Swabhiman", "Foundation", "NGO",
    )):
        return "NGO / Community Org"
    return "Other"


# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_data(path: str) -> dict:
    """
    Returns a dict with three clean DataFrames:
      active   — Active VT sheet, fully normalised
      dropped  — Dropped VT sheet
      new_reg  — Newly Registered VT sheet
    """
    if not os.path.exists(path):
        return {}

    # ── Active VT ─────────────────────────────────────────────────────────────
    active = pd.read_excel(path, sheet_name=SHEET_ACTIVE)
    active.columns = active.columns.str.strip()

    # Rename to internal names matching rest of dashboard
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

    # Fix State encoding artefacts (mangled non-breaking spaces)
    for col in ["State", "Residence state"]:
        if col in active.columns:
            active[col] = (
                active[col].astype(str)
                .str.replace("\u00ac\u00a0", " ", regex=False)
                .str.replace("\u00a0", " ", regex=False)
                .str.replace(r"\s+", " ", regex=True)
                .str.strip()
            )

    # Attendance% is already a float in this file — just coerce to be safe
    active["Attendance%"] = pd.to_numeric(active["Attendance%"], errors="coerce")

    # Parse join date
    active["Joined(ev)"] = pd.to_datetime(active["Joined(ev)"], errors="coerce")

    # Normalise Subject, Profession, Reference
    active["Subject_clean"]    = active["Subject"].map(SUBJECT_MAP).fillna(active["Subject"])
    active["Profession_clean"] = active["Profession"].map(PROFESSION_MAP).fillna("Others")
    active["Ref_group"]        = active["Reference"].apply(_group_ref)

    # Fill string NaNs
    for col in ["Donor", "State", "Subject_clean", "Center name",
                "Residence state", "Residence city", "Ref_group"]:
        if col in active.columns:
            active[col] = active[col].fillna("Unknown")

    # Numeric coercion
    for col in ["Registered", "Reg Boys", "Reg Girls",
                "Enrolled",   "En Boys",  "En Girls",
                "Planned",    "Scheduled", "Completed",
                "Offline",    "Cancelled",
                "Total hours(Comp+Offline)", "CLH"]:
        if col in active.columns:
            active[col] = pd.to_numeric(active[col], errors="coerce").fillna(0).astype(int)

    # ── Dropped VT ────────────────────────────────────────────────────────────
    dropped = pd.read_excel(path, sheet_name=SHEET_DROPPED)
    dropped.columns = dropped.columns.str.strip()
    # Fix encoding in State
    if "State" in dropped.columns:
        dropped["State"] = (
            dropped["State"].astype(str)
            .str.replace("\u00ac\u00a0", " ", regex=False)
            .str.replace("\u00a0", " ", regex=False)
            .str.strip()
        )

    # ── Newly Registered VT ───────────────────────────────────────────────────
    new_reg = pd.read_excel(path, sheet_name=SHEET_NEW_REG)
    new_reg.columns = new_reg.columns.str.strip()
    new_reg["Date Joined"] = pd.to_datetime(new_reg["Date Joined"], errors="coerce")
    new_reg["Ref_group"]   = new_reg["Reference"].apply(_group_ref)
    # Normalise gender
    if "Gender" in new_reg.columns:
        new_reg["Gender"] = new_reg["Gender"].astype(str).str.strip().str.title()

    return {"active": active, "dropped": dropped, "new_reg": new_reg}


# ─────────────────────────────────────────────────────────────────────────────
# SHARED HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def _layout(fig, height=380, legend_bottom=False, margin=None):
    m = margin or dict(l=0, r=0, t=40, b=0)
    fig.update_layout(
        height=height,
        plot_bgcolor=BG, paper_bgcolor=BG,
        font=dict(family="Inter, Helvetica, sans-serif", size=12, color="#2d3436"),
        margin=m,
    )
    fig.update_xaxes(showgrid=False, linecolor="#dee2e6", linewidth=1)
    fig.update_yaxes(showgrid=True, gridcolor=GRID, zeroline=False)
    if legend_bottom:
        fig.update_layout(legend=dict(
            orientation="h", yanchor="top", y=-0.18,
            xanchor="center", x=0.5, title_text="",
        ))
    return fig


def _kpi(label: str, value: str, color: str, sub: str = "") -> str:
    return (
        f"<div style='background:#f8f9fa;border-radius:10px;padding:14px 10px;"
        f"border-left:4px solid {color};text-align:center;'>"
        f"<div style='font-size:0.72em;color:#636e72;font-weight:600;"
        f"letter-spacing:0.05em;text-transform:uppercase;margin-bottom:4px;'>{label}</div>"
        f"<div style='font-size:1.7rem;font-weight:800;color:{color};"
        f"line-height:1.15;'>{value}</div>"
        f"<div style='font-size:0.7em;color:#b2bec3;margin-top:3px;'>{sub}</div>"
        f"</div>"
    )


# ─────────────────────────────────────────────────────────────────────────────
# MAIN RENDER
# ─────────────────────────────────────────────────────────────────────────────
def render_ops_dashboard():
    st.title("🏢 Operations & Impact Command Center")
    st.markdown(
        "<p style='color:gray;font-size:1.05em;margin-top:-12px;'>"
        "Volunteer Relationship Management  ·  Centre Operations  ·  Academic Health"
        "  —  May 2026</p>",
        unsafe_allow_html=True,
    )
    st.markdown("---")

    # ── Load data ─────────────────────────────────────────────────────────────
    with st.spinner("Loading VRM dataset…"):
        data = load_data(DATA_PATH)

    if not data:
        st.error(
            f"⚠️ Data file not found at `{DATA_PATH}`. "
            "Place `VRM_May__2026.xlsx` alongside `app.py`."
        )
        return

    df_raw   = data["active"]
    dropped  = data["dropped"]
    new_reg  = data["new_reg"]

    # ── Sidebar filters ───────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("---")
        st.header("🎯 VRM Filters")

        sel_donors = st.multiselect(
            "Donor",
            options=sorted(df_raw["Donor"].dropna().unique()),
            default=sorted(df_raw["Donor"].dropna().unique()),
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
        st.caption("Filters apply to all three tabs simultaneously.")

    df = df_raw[
        df_raw["Donor"].isin(sel_donors) &
        df_raw["State"].isin(sel_states) &
        df_raw["Subject_clean"].isin(sel_subjects)
    ].copy()

    if df.empty:
        st.warning("⚠️ No data for the current filter combination. Broaden your selection.")
        return

    # Volunteer-level dedup (for vol-level metrics)
    vol_df = df.drop_duplicates(subset="Volunteer id").copy()

    # ─────────────────────────────────────────────────────────────────────────
    # TABS
    # ─────────────────────────────────────────────────────────────────────────
    tab_vol, tab_ctr, tab_aca = st.tabs([
        "🙋 Volunteers",
        "🏫 Centres",
        "📚 Academic Health",
    ])

    # ═════════════════════════════════════════════════════════════════════════
    # TAB 1 — VOLUNTEERS
    # ═════════════════════════════════════════════════════════════════════════
    with tab_vol:

        # ── KPI calculations ──────────────────────────────────────────────────
        total_vols    = len(df)
        unique_vols   = vol_df["Volunteer id"].nunique()

        # Dropped — directly from the Dropped VT sheet (exact count)
        dropped_vols  = len(dropped)

        # Newly Registered — from the Newly Registered VT sheet (exact count)
        new_vols      = new_reg["User ID"].nunique()

        # Impact metrics
        active_centres  = df["Center name"].nunique()
        total_enrolled  = int(df["Enrolled"].sum())
        total_vol_hrs   = int(df["Total hours(Comp+Offline)"].sum())
        total_clh       = int(df["CLH"].sum())
        avg_att         = df["Attendance%"].mean()
        completion_rt   = (
            df["Completed"].sum() / df["Planned"].sum() * 100
            if df["Planned"].sum() > 0 else 0
        )

        # ── KPI Row 1: Volunteer counts ───────────────────────────────────────
        st.markdown("#### Volunteer Overview")
        kc1, kc2, kc3, kc4 = st.columns(4)
        kc1.markdown(_kpi("Total Volunteers",    f"{total_vols:,}",    P["teal"],    "vol-offering rows"),   unsafe_allow_html=True)
        kc2.markdown(_kpi("Unique Volunteers",   f"{unique_vols:,}",   P["green"],   "deduplicated"),        unsafe_allow_html=True)
        kc3.markdown(_kpi("Newly Registered",    f"{new_vols:,}",      P["violet"],  "from registrations sheet"), unsafe_allow_html=True)
        kc4.markdown(_kpi("Dropped Volunteers",  f"{dropped_vols:,}",  P["coral"],   "from dropped sheet"),  unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ── KPI Row 2: Impact counts ──────────────────────────────────────────
        ki1, ki2, ki3, ki4, ki5, ki6 = st.columns(6)
        ki1.markdown(_kpi("Active Centres",    f"{active_centres:,}",  P["teal"],    "unique"),             unsafe_allow_html=True)
        ki2.markdown(_kpi("Enrolled Students", f"{total_enrolled:,}",  P["green"],   "total seats"),        unsafe_allow_html=True)
        ki3.markdown(_kpi("Total Vol Hrs",     f"{total_vol_hrs:,}",   P["orange"],  "Comp + Offline hrs"), unsafe_allow_html=True)
        ki4.markdown(_kpi("Total CLH",         f"{total_clh:,}",       P["violet"],  "child learning hrs"), unsafe_allow_html=True)
        ki5.markdown(_kpi("Avg Attendance",    f"{avg_att:.1f}%",      P["amber"],   "across sessions"),    unsafe_allow_html=True)
        ki6.markdown(_kpi("Class Completion",  f"{completion_rt:.1f}%",P["mint"],    "completed / planned"),unsafe_allow_html=True)

        st.markdown("---")

        # ── Acquisition channel + Profession ─────────────────────────────────
        st.markdown("#### Volunteer Distribution")
        col_b1, col_b2 = st.columns(2)

        with col_b1:
            st.markdown("##### By Reference / Acquisition Channel")
            ref_data = (
                vol_df.groupby("Ref_group")["Volunteer id"].nunique()
                .reset_index()
                .rename(columns={"Volunteer id": "Volunteers"})
                .sort_values("Volunteers", ascending=True)
            )
            fig_ref = px.bar(
                ref_data, x="Volunteers", y="Ref_group",
                orientation="h", text="Volunteers",
                color="Ref_group", color_discrete_sequence=SEQ,
            )
            fig_ref.update_traces(
                textposition="outside", showlegend=False,
                marker_line_width=0,
                hovertemplate="<b>%{y}</b><br>Volunteers: %{x:,}<extra></extra>",
            )
            _layout(fig_ref, height=320)
            fig_ref.update_layout(xaxis_title="Volunteers", yaxis_title="", showlegend=False)
            fig_ref.update_xaxes(showgrid=True, gridcolor=GRID)
            fig_ref.update_yaxes(showgrid=False)
            st.plotly_chart(fig_ref, use_container_width=True)

        with col_b2:
            st.markdown("##### By Profession")
            prof_data = (
                vol_df.groupby("Profession_clean")["Volunteer id"].nunique()
                .reset_index()
                .rename(columns={"Volunteer id": "Volunteers"})
                .sort_values("Volunteers", ascending=False)
            )
            fig_prof = px.pie(
                prof_data, names="Profession_clean", values="Volunteers",
                hole=0.45, color_discrete_sequence=SEQ,
            )
            fig_prof.update_traces(
                textposition="outside",
                texttemplate="<b>%{label}</b><br>%{percent:.1%}",
                hovertemplate="<b>%{label}</b><br>Volunteers: %{value:,}<extra></extra>",
                pull=[0.02] * len(prof_data),
            )
            fig_prof.update_layout(
                height=320, showlegend=False,
                plot_bgcolor=BG, paper_bgcolor=BG,
                margin=dict(l=10, r=10, t=40, b=10),
                annotations=[dict(
                    text=f"<b>{unique_vols:,}</b><br>Vols",
                    x=0.5, y=0.5, font_size=15, showarrow=False,
                    font=dict(color="#2d3436"),
                )],
            )
            st.plotly_chart(fig_prof, use_container_width=True)

        st.markdown("---")

        # ── Residence by State ────────────────────────────────────────────────
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
            color_continuous_scale=[P["sky"], P["teal"]],
        )
        fig_res.update_traces(
            textposition="outside", marker_line_width=0,
            hovertemplate="<b>%{y}</b><br>Volunteers: %{x:,}<extra></extra>",
        )
        fig_res.update_coloraxes(showscale=False)
        _layout(fig_res, height=max(300, len(res_state) * 28),
                margin=dict(l=0, r=60, t=20, b=0))
        fig_res.update_layout(xaxis_title="Volunteers", yaxis_title="")
        fig_res.update_xaxes(showgrid=True, gridcolor=GRID)
        fig_res.update_yaxes(showgrid=False)
        st.plotly_chart(fig_res, use_container_width=True)

        st.markdown("---")

        # ── Newly Registered monthly trend ────────────────────────────────────
        st.markdown("#### New Volunteer Registrations — Monthly Trend")
        st.caption("Source: Newly Registered VT sheet — all registrations regardless of sidebar filters.")
        trend_df = new_reg.dropna(subset=["Date Joined"]).copy()
        trend_df["YM"] = trend_df["Date Joined"].dt.to_period("M").astype(str)
        monthly = (
            trend_df.groupby("YM")["User ID"].nunique()
            .reset_index()
            .rename(columns={"User ID": "New Vols"})
            .sort_values("YM")
        )
        fig_trend = px.area(
            monthly, x="YM", y="New Vols",
            markers=True, color_discrete_sequence=[P["teal"]],
        )
        fig_trend.update_traces(
            line=dict(width=2.5),
            marker=dict(size=6, color=P["teal"]),
            fillcolor="rgba(0,148,201,0.12)",
            hovertemplate="<b>%{x}</b><br>New Vols: %{y:,}<extra></extra>",
        )
        _layout(fig_trend, height=240, margin=dict(l=0, r=0, t=20, b=60))
        fig_trend.update_layout(xaxis_title="", yaxis_title="New Volunteers")
        fig_trend.update_xaxes(tickangle=-45, showgrid=False)
        st.plotly_chart(fig_trend, use_container_width=True)

        # ── Newly Registered gender split (available in new_reg sheet) ────────
        st.markdown("---")
        st.markdown("#### Newly Registered Volunteers — Gender Split")
        if "Gender" in new_reg.columns:
            gen_new = (
                new_reg[new_reg["Gender"].isin(["Male", "Female"])]
                .groupby("Gender")["User ID"].nunique()
                .reset_index()
                .rename(columns={"User ID": "Volunteers"})
            )
            col_g1, col_g2 = st.columns([1, 2])
            with col_g1:
                fig_gen = px.pie(
                    gen_new, names="Gender", values="Volunteers",
                    hole=0.5,
                    color="Gender",
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
                # Gender breakdown by Reference channel
                gen_ref = (
                    new_reg[new_reg["Gender"].isin(["Male", "Female"])]
                    .assign(Ref_group=lambda x: x["Reference"].apply(_group_ref))
                    .groupby(["Ref_group", "Gender"])["User ID"].nunique()
                    .reset_index()
                    .rename(columns={"User ID": "Volunteers"})
                )
                fig_gen_ref = px.bar(
                    gen_ref, x="Volunteers", y="Ref_group", color="Gender",
                    orientation="h", barmode="stack",
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

        # ── Dropped volunteers detail ─────────────────────────────────────────
        st.markdown("#### Dropped Volunteers — Reason Breakdown")
        st.caption("Source: Dropped VT sheet — volunteers who exited their offering(s).")
        if not dropped.empty and "Reasons" in dropped.columns:
            reason_counts = (
                dropped["Reasons"].fillna("Not specified")
                .value_counts().reset_index()
            )
            reason_counts.columns = ["Reason", "Count"]
            fig_drop = px.bar(
                reason_counts, x="Count", y="Reason",
                orientation="h", text="Count",
                color="Count",
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

    # ═════════════════════════════════════════════════════════════════════════
    # TAB 2 — CENTRES
    # ═════════════════════════════════════════════════════════════════════════
    with tab_ctr:

        # ── Centres by Donor ──────────────────────────────────────────────────
        st.markdown("#### Centres & Volunteers by Donor")
        donor_summary = (
            df.groupby("Donor").agg(
                Centres    =("Center name",  "nunique"),
                Enrolled   =("Enrolled",     "sum"),
                CLH        =("CLH",          "sum"),
                Volunteers =("Volunteer id", "nunique"),
                Completed  =("Completed",    "sum"),
                Planned    =("Planned",      "sum"),
                En_Boys    =("En Boys",      "sum"),
                En_Girls   =("En Girls",     "sum"),
            ).reset_index()
        )
        donor_summary["Completion %"] = (
            donor_summary["Completed"]
            / donor_summary["Planned"].replace(0, pd.NA) * 100
        ).fillna(0).round(1)
        donor_summary = donor_summary.sort_values("Centres", ascending=False)

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
            barmode="group", height=380,
            plot_bgcolor=BG, paper_bgcolor=BG,
            margin=dict(l=0, r=0, t=30, b=80),
            legend=dict(orientation="h", yanchor="top", y=-0.22,
                        xanchor="center", x=0.5, title_text=""),
            xaxis=dict(tickangle=-20, showgrid=False, linecolor="#dee2e6"),
            yaxis=dict(showgrid=True, gridcolor=GRID, zeroline=False),
            font=dict(family="Inter, Helvetica, sans-serif", size=12),
        )
        st.plotly_chart(fig_donor, use_container_width=True)

        # Enrolled by Donor + CLH by Donor side by side
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            st.markdown("##### Enrolled Students by Donor")
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

        # ── Enrolled Students Gender Split (now available!) ───────────────────
        st.markdown("#### Enrolled Students — Gender Split by Donor")
        st.caption("En Boys and En Girls columns are available in this dataset.")
        gen_donor = donor_summary[["Donor", "En_Boys", "En_Girls"]].copy()
        gen_donor_melt = gen_donor.melt(
            id_vars="Donor", value_vars=["En_Boys", "En_Girls"],
            var_name="Gender", value_name="Students"
        )
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

        # ── Top States by Active Centres ──────────────────────────────────────
        st.markdown("#### Top States by Active Centres")
        state_centres = (
            df.groupby("State").agg(
                Centres    =("Center name",  "nunique"),
                Enrolled   =("Enrolled",     "sum"),
                CLH        =("CLH",          "sum"),
                Volunteers =("Volunteer id", "nunique"),
            ).reset_index()
            .sort_values("Centres", ascending=True)
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
        _layout(fig_sc, height=max(260, len(state_centres) * 42),
                margin=dict(l=0, r=60, t=20, b=0))
        fig_sc.update_layout(xaxis_title="Active Centres", yaxis_title="")
        fig_sc.update_xaxes(showgrid=True, gridcolor=GRID)
        fig_sc.update_yaxes(showgrid=False)
        st.plotly_chart(fig_sc, use_container_width=True)

        st.markdown("---")

        # ── Class Completion Split ────────────────────────────────────────────
        st.markdown("#### Class Completion Split by Subject")
        st.caption(
            "🟢 Completed  🟡 Offline (rescheduled / async)  🔴 Cancelled"
        )
        exec_df = (
            df.groupby("Subject_clean")[["Completed", "Offline", "Cancelled"]]
            .sum().reset_index()
            .sort_values("Completed", ascending=False)
        )
        exec_melt = exec_df.melt(
            id_vars="Subject_clean",
            value_vars=["Completed", "Offline", "Cancelled"],
            var_name="Status", value_name="Sessions",
        )
        fig_exec = px.bar(
            exec_melt, x="Sessions", y="Subject_clean",
            color="Status", orientation="h", barmode="stack",
            text="Sessions",
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

    # ═════════════════════════════════════════════════════════════════════════
    # TAB 3 — ACADEMIC HEALTH
    # ═════════════════════════════════════════════════════════════════════════
    with tab_aca:

        # ── CLH by Subject ────────────────────────────────────────────────────
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

        # ── Avg Attendance % by State & Subject heatmap ───────────────────────
        st.markdown("#### Average Attendance % — State × Subject Heatmap")
        st.caption("Scale clamped 50–100%. White = no sessions for that pairing.")
        heat_df = (
            df.groupby(["State", "Subject_clean"])["Attendance%"]
            .mean().reset_index()
        )
        heat_pivot = (
            heat_df.pivot(index="State", columns="Subject_clean", values="Attendance%")
            .fillna(0)
        )
        sort_order = (
            heat_pivot.replace(0, pd.NA).mean(axis=1)
            .sort_values(ascending=False).index
        )
        heat_pivot = heat_pivot.loc[sort_order]

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
            bargap=0.05,
            font=dict(family="Inter, Helvetica, sans-serif", size=12),
        )
        fig_hist.update_xaxes(showgrid=False)
        fig_hist.update_yaxes(showgrid=True, gridcolor=GRID)
        st.plotly_chart(fig_hist, use_container_width=True)

        st.markdown("---")

        # ── Raw data table ────────────────────────────────────────────────────
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
        available_display = [c for c in display_cols if c in df.columns]
        display_df = df[available_display].rename(columns={
            "Subject_clean":             "Subject",
            "Total hours(Comp+Offline)": "Vol Hrs",
        }).copy()
        display_df["Attendance%"] = display_df["Attendance%"].round(1)
        st.dataframe(
            display_df.sort_values("CLH", ascending=False),
            use_container_width=True,
            hide_index=True,
            height=420,
        )
