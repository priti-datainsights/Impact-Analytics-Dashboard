"""
ops_dashboard.py
================
Operations & Impact Command Center  —  eVidyaloka VRM
Restructured to match the Volunteer Relationship Management team's workflow.

Layout
──────
Tab 1 │ Volunteers       — KPIs + acquisition + profession + residence
Tab 2 │ Centres          — Donor × Centre matrix + state ranking + completion
Tab 3 │ Academic Health  — CLH by subject + attendance heatmap + class split

Data expected at DATA_PATH (relative to app.py).
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────
DATA_PATH = "Copy of VRM FEB 2026 - Active VT.csv"

# Academic Year start — volunteers registered on or after this date
# count as "Newly Registered" in the current AY.
AY_START = pd.Timestamp("2025-08-01")

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
SEQ   = list(P.values())        # full sequential palette
BG    = "rgba(0,0,0,0)"
GRID  = "#e9ecef"

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
    "Completed my graduation. Currently not working.": "Others",
}

# ── Reference channel grouping ────────────────────────────────────────────────
def _group_ref(ref: str) -> str:
    r = str(ref).strip()
    if r == "Internet Search":    return "Internet Search"
    if r == "Emailer":            return "Emailer"
    if r == "Word of Mouth":      return "Word of Mouth"
    if r == "DCP Campaign":       return "DCP Campaign"
    if r == "NEAID":              return "NEAID / NGO Partner"
    if r == "eVidyaloka":         return "eVidyaloka Direct"
    if any(k in r for k in (
        "Cognizant","Infosys","Tech Mahindra","KPMG","HPInc","HPE","CGI",
        "L&T","Sanrakshan","HSBC","EY","Adobe","Brillio","Broadridge",
        "Accenture","CISCO","Fidelity","ConnectFor","Microsoft","LTTS",
        "Lowe","CAMS","Atlassian","WisdomCircle","Udaan",
    )):
        return "Corporate / Partner Referral"
    if any(k in r for k in (
        "College","University","BMS","NIT","IIT","IIM",
    )):
        return "Academic Institution"
    if any(k in r for k in (
        "Community","Foundation","NGO","Bhumi","Kaivalya",
        "Pratham","Teach","Impact",
    )):
        return "NGO / Community Org"
    return "Other"


# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_data(path: str) -> pd.DataFrame:
    if not os.path.exists(path):
        return pd.DataFrame()

    df = pd.read_csv(path, low_memory=False)
    df.columns = df.columns.str.strip()

    # Attendance% → float
    df["Attendance%"] = pd.to_numeric(
        df["Attendance%"].astype(str)
          .str.replace("%", "", regex=False)
          .str.strip(),
        errors="coerce",
    )

    # Fix State encoding artefacts (mangled non-breaking spaces)
    df["State"] = (
        df["State"].astype(str)
        .str.replace("\u00ac\u00a0", " ", regex=False)
        .str.replace("\u00a0", " ", regex=False)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )

    # Normalise Subject, Profession, Reference
    df["Subject_clean"]   = df["Subject"].map(SUBJECT_MAP).fillna(df["Subject"])
    df["Profession_clean"]= df["Profession"].map(PROFESSION_MAP).fillna("Others")
    df["Ref_group"]       = df["Reference"].apply(_group_ref)

    # Parse join date + flag newly registered vols
    df["Joined_dt"] = pd.to_datetime(df["Joined(ev)"], errors="coerce")
    df["Is_new"]    = df["Joined_dt"] >= AY_START

    # Fill string NaNs so dropdowns stay clean
    for col in ["Donor", "State", "Subject_clean", "Center name",
                "Residence state", "Residence city", "Ref_group"]:
        df[col] = df[col].fillna("Unknown")

    # Numeric coercion safety net
    for col in ["Registered", "Enrolled", "CLH", "Planned",
                "Completed", "Offline", "Cancelled",
                "Total hours(Comp+Offline)"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    return df


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
    """HTML metric card — consistent styling across both rows."""
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
        "Volunteer Relationship Management  ·  Centre Operations  ·  Academic Health</p>",
        unsafe_allow_html=True,
    )
    st.markdown("---")

    # ── Load data ─────────────────────────────────────────────────────────────
    with st.spinner("Loading VRM dataset…"):
        df_raw = load_data(DATA_PATH)

    if df_raw.empty:
        st.error(
            f"⚠️ Data file not found at `{DATA_PATH}`. "
            "Place `Copy of VRM FEB 2026 - Active VT.csv` alongside `app.py`."
        )
        return

    # ── Sidebar filters ───────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("---")
        st.header("🎯 VRM Filters")

        sel_donors   = st.multiselect(
            "Donor",
            options=sorted(df_raw["Donor"].dropna().unique()),
            default=sorted(df_raw["Donor"].dropna().unique()),
            key="ops_donor",
        )
        sel_states   = st.multiselect(
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

    # One row per volunteer (for volunteer-level metrics)
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
        new_vols      = vol_df[vol_df["Is_new"]]["Volunteer id"].nunique()

        # "Dropped" = volunteer whose every offering is Inactive or Completed
        # (most honest proxy in the absence of an explicit 'dropped' flag)
        vol_status_agg = df.groupby("Volunteer id")["Offering status"].apply(
            lambda x: all(s in ("Inactive", "Completed") for s in x)
        )
        dropped_vols   = int(vol_status_agg.sum())

        total_enrolled = int(df["Enrolled"].sum())
        total_vol_hrs  = int(df["Total hours(Comp+Offline)"].sum())
        total_clh      = int(df["CLH"].sum())
        avg_att        = df["Attendance%"].mean()
        completion_rt  = (
            df["Completed"].sum() / df["Planned"].sum() * 100
            if df["Planned"].sum() > 0 else 0
        )
        active_centres = df["Center name"].nunique()

        # ── Row 1: volunteer counts ───────────────────────────────────────────
        st.markdown("#### Volunteer Overview")
        kc1, kc2, kc3, kc4 = st.columns(4)
        kc1.markdown(_kpi("Total Volunteers",    f"{total_vols:,}",    P["teal"],    "vol-offering rows"),    unsafe_allow_html=True)
        kc2.markdown(_kpi("Unique Volunteers",   f"{unique_vols:,}",   P["green"],   "deduplicated"),         unsafe_allow_html=True)
        kc3.markdown(_kpi("Newly Registered",    f"{new_vols:,}",      P["violet"],  f"since {AY_START.strftime('%b %Y')}"), unsafe_allow_html=True)
        kc4.markdown(_kpi("Dropped / Completed", f"{dropped_vols:,}",  P["coral"],   "all offerings done"),   unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Row 2: impact counts ──────────────────────────────────────────────
        ki1, ki2, ki3, ki4, ki5, ki6 = st.columns(6)
        ki1.markdown(_kpi("Active Centres",   f"{active_centres:,}",  P["teal"],   "unique"),             unsafe_allow_html=True)
        ki2.markdown(_kpi("Enrolled Students",f"{total_enrolled:,}",  P["green"],  "total seats"),        unsafe_allow_html=True)
        ki3.markdown(_kpi("Total Vol Hrs",    f"{total_vol_hrs:,}",   P["orange"], "Comp + Offline hrs"), unsafe_allow_html=True)
        ki4.markdown(_kpi("Total CLH",        f"{total_clh:,}",       P["violet"], "child learning hrs"), unsafe_allow_html=True)
        ki5.markdown(_kpi("Avg Attendance",   f"{avg_att:.1f}%",      P["amber"],  "across sessions"),    unsafe_allow_html=True)
        ki6.markdown(_kpi("Class Completion", f"{completion_rt:.1f}%",P["mint"],   "completed / planned"),unsafe_allow_html=True)

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
            _layout(fig_ref, height=340)
            fig_ref.update_layout(
                xaxis_title="Volunteers", yaxis_title="", showlegend=False
            )
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
                height=340, showlegend=False,
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
        st.markdown("#### Volunteer Residence — Top 20 States")
        st.caption(
            "Restricted to state-level (city detail removed for clarity). "
            "Each bar = unique volunteers residing in that state."
        )
        res_state = (
            vol_df.groupby("Residence state")["Volunteer id"].nunique()
            .reset_index()
            .rename(columns={"Volunteer id": "Volunteers"})
            .sort_values("Volunteers", ascending=True)
            .tail(20)
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
        _layout(fig_res, height=500, margin=dict(l=0, r=60, t=30, b=0))
        fig_res.update_layout(xaxis_title="Volunteers", yaxis_title="")
        fig_res.update_xaxes(showgrid=True, gridcolor=GRID)
        fig_res.update_yaxes(showgrid=False)
        st.plotly_chart(fig_res, use_container_width=True)

        st.markdown("---")

        # ── Monthly joining trend ─────────────────────────────────────────────
        st.markdown("#### New Volunteer Registrations Over Time (last 30 months)")
        trend_df = vol_df.dropna(subset=["Joined_dt"]).copy()
        trend_df["YM"] = trend_df["Joined_dt"].dt.to_period("M").astype(str)
        monthly = (
            trend_df.groupby("YM")["Volunteer id"].nunique()
            .reset_index()
            .rename(columns={"Volunteer id": "New Vols"})
            .sort_values("YM")
            .tail(30)
        )
        fig_trend = px.area(
            monthly, x="YM", y="New Vols",
            markers=True, color_discrete_sequence=[P["teal"]],
        )
        fig_trend.update_traces(
            line=dict(width=2.5),
            marker=dict(size=5, color=P["teal"]),
            fillcolor="rgba(0,148,201,0.12)",
            hovertemplate="<b>%{x}</b><br>New Vols: %{y:,}<extra></extra>",
        )
        _layout(fig_trend, height=260, margin=dict(l=0, r=0, t=20, b=60))
        fig_trend.update_layout(xaxis_title="", yaxis_title="New Volunteers")
        fig_trend.update_xaxes(tickangle=-45, showgrid=False)
        st.plotly_chart(fig_trend, use_container_width=True)

    # ═════════════════════════════════════════════════════════════════════════
    # TAB 2 — CENTRES
    # ═════════════════════════════════════════════════════════════════════════
    with tab_ctr:

        # ── Centres by Donor ──────────────────────────────────────────────────
        st.markdown("#### Centres by Donor")
        st.caption(
            "Grouped bars: Centres (teal) and Volunteers (green) per Donor. "
            "Top 12 donors by centre count shown."
        )
        donor_summary = (
            df.groupby("Donor").agg(
                Centres    =("Center name",  "nunique"),
                Enrolled   =("Enrolled",     "sum"),
                CLH        =("CLH",          "sum"),
                Volunteers =("Volunteer id", "nunique"),
                Completed  =("Completed",    "sum"),
                Planned    =("Planned",      "sum"),
            ).reset_index()
        )
        donor_summary["Completion %"] = (
            donor_summary["Completed"]
            / donor_summary["Planned"].replace(0, pd.NA) * 100
        ).fillna(0).round(1)
        donor_summary = donor_summary.sort_values("Centres", ascending=False)
        top12 = donor_summary.head(12)

        fig_donor = go.Figure()
        fig_donor.add_trace(go.Bar(
            name="Centres",
            x=top12["Donor"], y=top12["Centres"],
            marker_color=P["teal"],
            text=top12["Centres"], textposition="outside",
            hovertemplate="<b>%{x}</b><br>Centres: %{y}<extra></extra>",
        ))
        fig_donor.add_trace(go.Bar(
            name="Volunteers",
            x=top12["Donor"], y=top12["Volunteers"],
            marker_color=P["green"],
            text=top12["Volunteers"], textposition="outside",
            hovertemplate="<b>%{x}</b><br>Volunteers: %{y}<extra></extra>",
        ))
        fig_donor.update_layout(
            barmode="group", height=420,
            plot_bgcolor=BG, paper_bgcolor=BG,
            margin=dict(l=0, r=0, t=30, b=90),
            legend=dict(orientation="h", yanchor="top", y=-0.28,
                        xanchor="center", x=0.5, title_text=""),
            xaxis=dict(tickangle=-35, showgrid=False, linecolor="#dee2e6"),
            yaxis=dict(showgrid=True, gridcolor=GRID, zeroline=False),
            font=dict(family="Inter, Helvetica, sans-serif", size=12),
        )
        st.plotly_chart(fig_donor, use_container_width=True)

        # Enrolled + CLH side-by-side horizontal bars
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            st.markdown("##### Enrolled Students by Donor")
            enr_d = donor_summary.sort_values("Enrolled", ascending=True).tail(12)
            fig_enr = px.bar(
                enr_d, x="Enrolled", y="Donor", orientation="h",
                text="Enrolled",
                color="Enrolled",
                color_continuous_scale=[P["lavender"], P["violet"]],
            )
            fig_enr.update_traces(textposition="outside",
                                  texttemplate="%{text:,}", marker_line_width=0)
            fig_enr.update_coloraxes(showscale=False)
            _layout(fig_enr, height=380, margin=dict(l=0, r=70, t=20, b=0))
            fig_enr.update_layout(xaxis_title="Enrolled Students", yaxis_title="")
            fig_enr.update_xaxes(showgrid=True, gridcolor=GRID)
            fig_enr.update_yaxes(showgrid=False)
            st.plotly_chart(fig_enr, use_container_width=True)

        with col_d2:
            st.markdown("##### CLH by Donor")
            clh_d = donor_summary.sort_values("CLH", ascending=True).tail(12)
            fig_clh_d = px.bar(
                clh_d, x="CLH", y="Donor", orientation="h",
                text="CLH",
                color="CLH",
                color_continuous_scale=[P["salmon"], P["coral"]],
            )
            fig_clh_d.update_traces(textposition="outside",
                                    texttemplate="%{text:,}", marker_line_width=0)
            fig_clh_d.update_coloraxes(showscale=False)
            _layout(fig_clh_d, height=380, margin=dict(l=0, r=70, t=20, b=0))
            fig_clh_d.update_layout(xaxis_title="Child Learning Hours", yaxis_title="")
            fig_clh_d.update_xaxes(showgrid=True, gridcolor=GRID)
            fig_clh_d.update_yaxes(showgrid=False)
            st.plotly_chart(fig_clh_d, use_container_width=True)

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
        _layout(fig_sc, height=420, margin=dict(l=0, r=60, t=20, b=0))
        fig_sc.update_layout(xaxis_title="Active Centres", yaxis_title="")
        fig_sc.update_xaxes(showgrid=True, gridcolor=GRID)
        fig_sc.update_yaxes(showgrid=False)
        st.plotly_chart(fig_sc, use_container_width=True)

        st.markdown("---")

        # ── Class completion split ────────────────────────────────────────────
        st.markdown("#### Class Completion Split — Completed · Offline · Cancelled")
        st.caption(
            "🟢 Completed = held as scheduled  "
            "🟡 Offline = rescheduled / async  "
            "🔴 Cancelled = did not happen"
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
        _layout(fig_exec, height=440, legend_bottom=True,
                margin=dict(l=0, r=0, t=20, b=60))
        fig_exec.update_layout(xaxis_title="Sessions", yaxis_title="")
        fig_exec.update_xaxes(showgrid=True, gridcolor=GRID)
        fig_exec.update_yaxes(showgrid=False)
        st.plotly_chart(fig_exec, use_container_width=True)

        # Gender placeholder
        st.info(
            "**ℹ️ Enrolled Students Gender Split** — The current VRM extract does not "
            "include a student gender column. Once the updated sheet with gender data "
            "is shared, this chart will be activated automatically."
        )

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
        _layout(fig_clh, height=420, margin=dict(l=0, r=70, t=20, b=0))
        fig_clh.update_layout(xaxis_title="Child Learning Hours", yaxis_title="")
        fig_clh.update_xaxes(showgrid=True, gridcolor=GRID)
        fig_clh.update_yaxes(showgrid=False)
        st.plotly_chart(fig_clh, use_container_width=True)

        st.markdown("---")

        # ── Average Attendance % by State & Subject ───────────────────────────
        st.markdown("#### Average Attendance % — State × Subject Heatmap")
        st.caption(
            "Cells = mean attendance % for each State × Subject pairing. "
            "Empty cells (no sessions) shown as white. Scale clamped 60–100%."
        )
        heat_df = (
            df.groupby(["State", "Subject_clean"])["Attendance%"]
            .mean().reset_index()
        )
        heat_pivot = (
            heat_df.pivot(
                index="State", columns="Subject_clean", values="Attendance%"
            ).fillna(0)
        )
        # Sort states by mean attendance (non-zero cells only)
        sort_order = (
            heat_pivot.replace(0, pd.NA)
            .mean(axis=1)
            .sort_values(ascending=False)
            .index
        )
        heat_pivot = heat_pivot.loc[sort_order]

        fig_heat = px.imshow(
            heat_pivot,
            color_continuous_scale=["#dfe6e9", P["teal"]],
            aspect="auto",
            text_auto=".0f",
            zmin=60, zmax=100,
        )
        fig_heat.update_traces(
            hovertemplate=(
                "<b>%{y}</b> · <b>%{x}</b><br>"
                "Avg Attendance: %{z:.1f}%<extra></extra>"
            )
        )
        fig_heat.update_layout(
            height=500,
            plot_bgcolor=BG, paper_bgcolor=BG,
            margin=dict(l=0, r=0, t=20, b=110),
            coloraxis_colorbar=dict(
                title="Att%", ticksuffix="%",
                len=0.6, thickness=12,
            ),
            font=dict(family="Inter, Helvetica, sans-serif", size=12),
        )
        fig_heat.update_xaxes(tickangle=-40, side="bottom", showgrid=False)
        fig_heat.update_yaxes(showgrid=False)
        st.plotly_chart(fig_heat, use_container_width=True)

        st.markdown("---")

        # ── Attendance distribution ───────────────────────────────────────────
        st.markdown("#### Attendance Distribution Across All Sessions")
        fig_hist = px.histogram(
            df, x="Attendance%", nbins=30,
            color_discrete_sequence=[P["violet"]], opacity=0.82,
        )
        fig_hist.update_layout(
            xaxis_title="Attendance %", yaxis_title="Number of Sessions",
            height=260,
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
            "Subject_clean", "Registered", "Enrolled", "Attendance%",
            "CLH", "Total hours(Comp+Offline)",
            "Planned", "Completed", "Offline", "Cancelled",
        ]
        display_df = df[display_cols].rename(columns={
            "Subject_clean":            "Subject",
            "Total hours(Comp+Offline)": "Vol Hrs",
        }).copy()
        display_df["Attendance%"] = display_df["Attendance%"].round(1)
        st.dataframe(
            display_df.sort_values("CLH", ascending=False),
            use_container_width=True,
            hide_index=True,
            height=420,
        )
