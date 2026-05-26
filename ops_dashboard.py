"""
ops_dashboard.py
================
Operations & Impact Command Center
eVidyaloka — Volunteer Management Analytics

Drop this file next to app.py and call render_ops_dashboard() from the
longitudinal module's home-page routing block (see integration note at
the bottom of this file).

Data expected at:  DATA_PATH  (see constant below).
All heavy lifting is cached via @st.cache_data so reruns are instant.
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS & COLOUR PALETTE
# ─────────────────────────────────────────────────────────────────────────────

# Absolute path to the CSV — adjust to match your deployment layout.
DATA_PATH = "Copy of VRM FEB 2026 - Active VT.csv"

# Premium palette — avoids Plotly defaults.
# Ordered from darkest → lightest teal/slate family.
PALETTE = [
    "#0094c9",   # eVidyaloka teal (primary)
    "#00964d",   # evergreen
    "#f27c48",   # warm orange
    "#ed1c2d",   # alert red
    "#6c5ce7",   # violet
    "#fdcb6e",   # amber
    "#a29bfe",   # lavender
    "#00b894",   # mint
    "#e17055",   # terra cotta
    "#74b9ff",   # sky
    "#55efc4",   # aqua
    "#fab1a0",   # salmon
]

ACCENT      = "#0094c9"
BG_CARD     = "#f8f9fa"
PLOT_BG     = "rgba(0,0,0,0)"
GRID_COLOR  = "#e9ecef"

# ─────────────────────────────────────────────────────────────────────────────
# PROFESSION NORMALISATION MAP
# Raw codes in the CSV → clean human-readable labels
# ─────────────────────────────────────────────────────────────────────────────
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
    "Trainer, manufacturers":                  "Others",
    "top_managemenet":                         "Management",
    "Management":                              "Management",
    "finance":                                 "Finance",
    "legal":                                   "Legal",
    "Legal":                                   "Legal",
    "Others":                                  "Others",
    "Education, Training, and Library":        "School Teacher",
    "Computer and Mathematical":               "Corporate Professional",
    "Life, Physical, and Social Science":      "Others",
    "Business and Financial Operations":       "Finance",
    "Healthcare Practitioners and Technical":  "Healthcare",
    "Medical Professional":                    "Healthcare",
    "Office and Administrative Support":       "Others",
    "Production/Manufacturing":                "Others",
    "Sales and Related":                       "Business / Self-Employed",
    "Community and Social Service":            "Others",
    "Farming, Fishing, and Forestry":          "Others",
    "Completed my graduation. Currently not working.": "Others",
}

# ─────────────────────────────────────────────────────────────────────────────
# SUBJECT NORMALISATION MAP
# Collapses near-duplicate subject names into clean buckets
# ─────────────────────────────────────────────────────────────────────────────
SUBJECT_MAP = {
    "Conceptual Learning - Math":    "Math",
    "Conceptual Learning Math-HM":   "Math",
    "Maths - Worksheet":             "Math",
    "Conceptual Learning - Science": "Science",
    "Conceptual Learning Science-HM":"Science",
    "English":                       "English",
    "English - Worksheet":           "English",
    "Spoken English: Level 1":       "Spoken English",
    "Spoken English: Level 2":       "Spoken English",
    "Basic Digital Literacy":        "Digital Literacy",
    "Concise content 1":             "Concise Content",
    "Concise Content 1":             "Concise Content",
    "Concise content 2":             "Concise Content",
    "Artificial Intelligence (AI)":  "AI",
    "Explore Coding":                "Coding",
    "Guest Sessions":                "Guest Sessions",
    "Scholarship":                   "Scholarship",
    "Reading Program":               "Reading Program",
}

# Reference channel grouping — 80 raw values → clean acquisition buckets
def _group_reference(ref: str) -> str:
    r = str(ref).strip()
    if r in ("Internet Search", "Emailer", "Word of Mouth", "DCP Campaign",
             "eVidyaloka", "NEAID"):
        return r if r != "NEAID" else "NEAID / NGO Partner"
    if any(k in r for k in ("Cognizant", "Infosys", "Tech Mahindra", "KPMG",
                             "HPInc", "HPE", "CGI", "L&T", "Sanrakshan",
                             "WisdomCircle", "Udaan", "HSBC", "EY", "Adobe",
                             "Brillio", "Broadridge", "Accenture", "CISCO",
                             "Fidelity", "ConnectFor", "Microsoft", "LTTS",
                             "Lowe", "CAMS", "Atlassian")):
        return "Corporate / Partner Referral"
    if any(k in r for k in ("College", "University", "BMS", "NIT", "IIT", "IIM")):
        return "Academic Institution"
    if any(k in r for k in ("Community", "Foundation", "NGO", "Bhumi",
                             "Kaivalya", "Pratham", "Teach", "Impact")):
        return "NGO / Community Org"
    return "Other"


# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADING & CLEANING
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_data(path: str) -> pd.DataFrame:
    """
    Load the volunteer-matching CSV, apply all normalisations, and return a
    clean DataFrame.  The function is cached so it only runs once per session.
    """
    if not os.path.exists(path):
        return pd.DataFrame()          # caller handles empty case gracefully

    df = pd.read_csv(path, low_memory=False)

    # ── 1. Strip column-name whitespace ──────────────────────────────────────
    df.columns = df.columns.str.strip()

    # ── 2. Attendance% → float ───────────────────────────────────────────────
    df["Attendance%"] = (
        df["Attendance%"]
        .astype(str)
        .str.replace("%", "", regex=False)
        .str.strip()
    )
    df["Attendance%"] = pd.to_numeric(df["Attendance%"], errors="coerce")

    # ── 3. Fix State encoding artefact (Jammu\u00a0and\u00a0Kashmir) ─────────
    df["State"] = df["State"].astype(str).str.replace("\u00ac\u00a0", " ", regex=False)
    df["State"] = df["State"].str.replace(r"\s+", " ", regex=True).str.strip()

    # ── 4. Normalise Subject ─────────────────────────────────────────────────
    df["Subject_clean"] = df["Subject"].map(SUBJECT_MAP).fillna(df["Subject"])

    # ── 5. Normalise Profession ──────────────────────────────────────────────
    df["Profession_clean"] = df["Profession"].map(PROFESSION_MAP).fillna("Others")

    # ── 6. Group Reference / Acquisition channel ─────────────────────────────
    df["Reference_group"] = df["Reference"].apply(_group_reference)

    # ── 7. Fill obvious string NaNs so dropdowns don't show 'nan' ───────────
    for col in ["Donor", "State", "Subject_clean", "Center name",
                "Residence state", "Residence city", "Reference_group"]:
        df[col] = df[col].fillna("Unknown")

    # ── 8. Numeric coercion safety net ───────────────────────────────────────
    for col in ["Registered", "Enrolled", "CLH", "Planned",
                "Completed", "Offline", "Cancelled"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    return df


# ─────────────────────────────────────────────────────────────────────────────
# HELPER: compact plotly layout defaults
# ─────────────────────────────────────────────────────────────────────────────
def _base_layout(**kwargs) -> dict:
    base = dict(
        plot_bgcolor=PLOT_BG,
        paper_bgcolor=PLOT_BG,
        font=dict(family="Inter, Helvetica, sans-serif", size=12, color="#2d3436"),
        margin=dict(l=0, r=0, t=40, b=0),
        colorway=PALETTE,
    )
    base.update(kwargs)
    return base


def _style(fig, legend_bottom: bool = False, height: int = 380):
    fig.update_layout(height=height, **_base_layout())
    fig.update_xaxes(showgrid=False, linecolor="#dee2e6", linewidth=1)
    fig.update_yaxes(showgrid=True, gridcolor=GRID_COLOR, zeroline=False)
    if legend_bottom:
        fig.update_layout(
            legend=dict(
                orientation="h", yanchor="top", y=-0.18,
                xanchor="center", x=0.5, title_text="",
            )
        )
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# MAIN RENDER FUNCTION — called from app.py
# ─────────────────────────────────────────────────────────────────────────────
def render_ops_dashboard():
    """
    Render the full Operations & Impact Command Center.
    Expects Streamlit page config already set (layout='wide').
    """
    st.title("🏢 Operations & Impact Command Center")
    st.markdown(
        "<p style='color:gray;font-size:1.05em;margin-top:-12px;'>"
        "Volunteer Management · Center Operations · Academic Health  —  AY 25-26</p>",
        unsafe_allow_html=True,
    )
    st.markdown("---")

    # ── Load data ─────────────────────────────────────────────────────────────
    with st.spinner("Loading volunteer dataset…"):
        df_raw = load_data(DATA_PATH)

    if df_raw.empty:
        st.error(
            f"⚠️ Data file not found at `{DATA_PATH}`. "
            "Place `Copy_of_VRM_FEB_2026_-_Active_VT.csv` in the same folder as `app.py`."
        )
        return

    # ─────────────────────────────────────────────────────────────────────────
    # SIDEBAR — GLOBAL FILTERS
    # ─────────────────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("---")
        st.header("🎯 Ops Filters")

        # Donor
        all_donors = sorted(df_raw["Donor"].dropna().unique())
        sel_donors = st.multiselect(
            "Donor", options=all_donors, default=all_donors, key="ops_donor"
        )

        # State (center location)
        all_states = sorted(df_raw["State"].dropna().unique())
        sel_states = st.multiselect(
            "State (Center)", options=all_states, default=all_states, key="ops_state"
        )

        # Subject (normalised)
        all_subjects = sorted(df_raw["Subject_clean"].dropna().unique())
        sel_subjects = st.multiselect(
            "Subject", options=all_subjects, default=all_subjects, key="ops_subject"
        )

        st.caption(
            "Filters apply to all three tabs. "
            "Deselect items to narrow the analysis."
        )

    # Apply filters
    df = df_raw[
        df_raw["Donor"].isin(sel_donors) &
        df_raw["State"].isin(sel_states) &
        df_raw["Subject_clean"].isin(sel_subjects)
    ].copy()

    if df.empty:
        st.warning("⚠️ No data matches the current filter combination. Broaden your selection.")
        return

    # ─────────────────────────────────────────────────────────────────────────
    # TABS
    # ─────────────────────────────────────────────────────────────────────────
    tab_exec, tab_vol, tab_ops = st.tabs([
        "📊 Executive Impact",
        "🙋 Volunteer Sourcing & Demographics",
        "🏫 Center Operations & Academic Health",
    ])

    # ═════════════════════════════════════════════════════════════════════════
    # TAB 1 — EXECUTIVE IMPACT
    # ═════════════════════════════════════════════════════════════════════════
    with tab_exec:
        # ── KPI ROW ──────────────────────────────────────────────────────────
        total_vols    = df["Volunteer id"].nunique()
        total_enroll  = int(df["Enrolled"].sum())
        total_clh     = int(df["CLH"].sum())
        avg_attend    = df["Attendance%"].mean()
        total_centers = df["Center name"].nunique()
        completion_rt = (
            df["Completed"].sum() / df["Planned"].sum() * 100
            if df["Planned"].sum() > 0 else 0
        )

        k1, k2, k3, k4, k5, k6 = st.columns(6)
        _metric_style = """
            <div style='background:#f8f9fa;border-radius:10px;padding:14px 10px;
                        border-left:4px solid {color};text-align:center;'>
                <div style='font-size:0.78em;color:#636e72;font-weight:600;
                            letter-spacing:0.04em;text-transform:uppercase;
                            margin-bottom:4px;'>{label}</div>
                <div style='font-size:1.75rem;font-weight:800;color:{color};
                            line-height:1.1;'>{value}</div>
                <div style='font-size:0.72em;color:#b2bec3;margin-top:2px;'>{sub}</div>
            </div>
        """
        k1.markdown(_metric_style.format(
            color=ACCENT, label="Active Volunteers",
            value=f"{total_vols:,}", sub="unique VTs"), unsafe_allow_html=True)
        k2.markdown(_metric_style.format(
            color="#00964d", label="Students Enrolled",
            value=f"{total_enroll:,}", sub="total seats"), unsafe_allow_html=True)
        k3.markdown(_metric_style.format(
            color="#6c5ce7", label="Child Learning Hours",
            value=f"{total_clh:,}", sub="CLH delivered"), unsafe_allow_html=True)
        k4.markdown(_metric_style.format(
            color="#f27c48", label="Avg Attendance",
            value=f"{avg_attend:.1f}%", sub="across all sessions"), unsafe_allow_html=True)
        k5.markdown(_metric_style.format(
            color="#00b894", label="Active Centers",
            value=f"{total_centers:,}", sub="unique schools"), unsafe_allow_html=True)
        k6.markdown(_metric_style.format(
            color="#e17055", label="Class Completion",
            value=f"{completion_rt:.1f}%", sub="completed / planned"), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ── ROW 1: Donut (Donor CLH) + Horizontal bar (Top States) ──────────
        col_l, col_r = st.columns([1, 1])

        with col_l:
            st.markdown("#### CLH Distribution by Donor")
            donor_clh = (
                df.groupby("Donor")["CLH"].sum()
                .reset_index()
                .sort_values("CLH", ascending=False)
            )
            # Collapse small donors into "Other" for readability
            top_n = 8
            if len(donor_clh) > top_n:
                top    = donor_clh.head(top_n)
                others = pd.DataFrame([{
                    "Donor": "Other Donors",
                    "CLH":   donor_clh.iloc[top_n:]["CLH"].sum()
                }])
                donor_clh = pd.concat([top, others], ignore_index=True)

            fig_donut = px.pie(
                donor_clh, names="Donor", values="CLH",
                hole=0.52,
                color_discrete_sequence=PALETTE,
            )
            fig_donut.update_traces(
                textposition="outside",
                texttemplate="<b>%{label}</b><br>%{percent:.1%}",
                hovertemplate="<b>%{label}</b><br>CLH: %{value:,}<br>Share: %{percent:.1%}<extra></extra>",
                pull=[0.03] * len(donor_clh),
            )
            fig_donut.update_layout(
                height=400,
                showlegend=False,
                plot_bgcolor=PLOT_BG, paper_bgcolor=PLOT_BG,
                margin=dict(l=10, r=10, t=40, b=10),
                annotations=[dict(
                    text=f"<b>{total_clh:,}</b><br>CLH",
                    x=0.5, y=0.5, font_size=17, showarrow=False,
                    font=dict(color="#2d3436"),
                )],
            )
            st.plotly_chart(fig_donut, use_container_width=True)

        with col_r:
            st.markdown("#### Top States by Active Centers")
            state_centers = (
                df.groupby("State")["Center name"].nunique()
                .reset_index()
                .rename(columns={"Center name": "Centers"})
                .sort_values("Centers", ascending=True)   # ascending for horiz bar
                .tail(10)
            )
            fig_states = px.bar(
                state_centers, x="Centers", y="State",
                orientation="h",
                text="Centers",
                color="Centers",
                color_continuous_scale=["#74b9ff", ACCENT],
            )
            fig_states.update_traces(
                textposition="outside",
                hovertemplate="<b>%{y}</b><br>Centers: %{x}<extra></extra>",
                marker_line_width=0,
            )
            fig_states.update_coloraxes(showscale=False)
            fig_states.update_layout(
                height=400,
                xaxis_title="Number of Centers",
                yaxis_title="",
                plot_bgcolor=PLOT_BG, paper_bgcolor=PLOT_BG,
                margin=dict(l=0, r=50, t=40, b=0),
            )
            fig_states.update_xaxes(showgrid=True, gridcolor=GRID_COLOR)
            fig_states.update_yaxes(showgrid=False)
            st.plotly_chart(fig_states, use_container_width=True)

        # ── ROW 2: Enrolled students by Donor (bar) + CLH by Subject ─────────
        st.markdown("---")
        col_m1, col_m2 = st.columns(2)

        with col_m1:
            st.markdown("#### Enrolled Students by Donor")
            donor_enroll = (
                df.groupby("Donor")["Enrolled"].sum()
                .reset_index()
                .sort_values("Enrolled", ascending=False)
                .head(10)
            )
            fig_de = px.bar(
                donor_enroll, x="Donor", y="Enrolled",
                text="Enrolled",
                color="Enrolled",
                color_continuous_scale=["#a29bfe", "#6c5ce7"],
            )
            fig_de.update_traces(textposition="outside", texttemplate="%{text:,}", marker_line_width=0)
            fig_de.update_coloraxes(showscale=False)
            fig_de.update_layout(
                xaxis_title="", yaxis_title="Enrolled Students",
                plot_bgcolor=PLOT_BG, paper_bgcolor=PLOT_BG,
                margin=dict(l=0, r=0, t=40, b=60),
                height=380,
            )
            fig_de.update_xaxes(tickangle=-35, showgrid=False)
            fig_de.update_yaxes(showgrid=True, gridcolor=GRID_COLOR)
            st.plotly_chart(fig_de, use_container_width=True)

        with col_m2:
            st.markdown("#### CLH by Subject")
            subj_clh = (
                df.groupby("Subject_clean")["CLH"].sum()
                .reset_index()
                .sort_values("CLH", ascending=False)
            )
            fig_sc = px.bar(
                subj_clh, x="CLH", y="Subject_clean",
                orientation="h",
                text="CLH",
                color="CLH",
                color_continuous_scale=["#fab1a0", "#e17055"],
            )
            fig_sc.update_traces(
                textposition="outside", texttemplate="%{text:,}", marker_line_width=0
            )
            fig_sc.update_coloraxes(showscale=False)
            fig_sc.update_layout(
                xaxis_title="Child Learning Hours", yaxis_title="",
                plot_bgcolor=PLOT_BG, paper_bgcolor=PLOT_BG,
                margin=dict(l=0, r=60, t=40, b=0),
                height=380,
            )
            fig_sc.update_xaxes(showgrid=True, gridcolor=GRID_COLOR)
            fig_sc.update_yaxes(showgrid=False)
            st.plotly_chart(fig_sc, use_container_width=True)

    # ═════════════════════════════════════════════════════════════════════════
    # TAB 2 — VOLUNTEER SOURCING & DEMOGRAPHICS
    # ═════════════════════════════════════════════════════════════════════════
    with tab_vol:
        # Volunteer-level deduplicated frame (one row per volunteer)
        vol_df = df.drop_duplicates(subset="Volunteer id").copy()

        # ── ROW 1: Acquisition funnel + Residence treemap ─────────────────────
        col_v1, col_v2 = st.columns([1, 1.4])

        with col_v1:
            st.markdown("#### Volunteer Acquisition Channels")
            ref_counts = (
                vol_df.groupby("Reference_group")["Volunteer id"].nunique()
                .reset_index()
                .rename(columns={"Volunteer id": "Volunteers"})
                .sort_values("Volunteers", ascending=False)
            )
            fig_ref = px.bar(
                ref_counts, x="Volunteers", y="Reference_group",
                orientation="h",
                text="Volunteers",
                color="Reference_group",
                color_discrete_sequence=PALETTE,
            )
            fig_ref.update_traces(
                textposition="outside",
                hovertemplate="<b>%{y}</b><br>Volunteers: %{x:,}<extra></extra>",
                marker_line_width=0,
                showlegend=False,
            )
            fig_ref.update_layout(
                xaxis_title="Number of Volunteers", yaxis_title="",
                plot_bgcolor=PLOT_BG, paper_bgcolor=PLOT_BG,
                margin=dict(l=0, r=60, t=40, b=0),
                height=360,
                showlegend=False,
            )
            fig_ref.update_xaxes(showgrid=True, gridcolor=GRID_COLOR)
            fig_ref.update_yaxes(showgrid=False)
            st.plotly_chart(fig_ref, use_container_width=True)

        with col_v2:
            st.markdown("#### Volunteer Talent Pool — Residence Geography")
            # Treemap: Residence state → Residence city, sized by volunteer count
            tree_df = (
                vol_df.groupby(["Residence state", "Residence city"])["Volunteer id"]
                .nunique()
                .reset_index()
                .rename(columns={"Volunteer id": "Volunteers"})
            )
            # Filter to top 15 states by volunteer volume to keep chart readable
            top_states_res = (
                tree_df.groupby("Residence state")["Volunteers"].sum()
                .nlargest(15).index
            )
            tree_df = tree_df[tree_df["Residence state"].isin(top_states_res)]

            fig_tree = px.treemap(
                tree_df,
                path=[px.Constant("All"), "Residence state", "Residence city"],
                values="Volunteers",
                color="Volunteers",
                color_continuous_scale=["#dfe6e9", ACCENT],
                hover_data={"Volunteers": ":,"},
            )
            fig_tree.update_traces(
                texttemplate="<b>%{label}</b><br>%{value:,}",
                hovertemplate="<b>%{label}</b><br>Volunteers: %{value:,}<extra></extra>",
                marker_line_width=0.5,
                marker_line_color="white",
            )
            fig_tree.update_coloraxes(showscale=False)
            fig_tree.update_layout(
                height=360,
                plot_bgcolor=PLOT_BG, paper_bgcolor=PLOT_BG,
                margin=dict(l=0, r=0, t=40, b=0),
            )
            st.plotly_chart(fig_tree, use_container_width=True)

        st.markdown("---")

        # ── ROW 2: Profession pie + Monthly joining trend ─────────────────────
        col_v3, col_v4 = st.columns([1, 1.2])

        with col_v3:
            st.markdown("#### Volunteer Profession Breakdown")
            prof_counts = (
                vol_df.groupby("Profession_clean")["Volunteer id"].nunique()
                .reset_index()
                .rename(columns={"Volunteer id": "Volunteers"})
                .sort_values("Volunteers", ascending=False)
            )
            fig_prof = px.pie(
                prof_counts, names="Profession_clean", values="Volunteers",
                hole=0.0,
                color_discrete_sequence=PALETTE,
            )
            fig_prof.update_traces(
                textposition="inside",
                texttemplate="<b>%{label}</b><br>%{percent:.1%}",
                hovertemplate="<b>%{label}</b><br>Volunteers: %{value:,}<extra></extra>",
                pull=[0.02] * len(prof_counts),
            )
            fig_prof.update_layout(
                height=380,
                showlegend=True,
                plot_bgcolor=PLOT_BG, paper_bgcolor=PLOT_BG,
                margin=dict(l=0, r=0, t=40, b=0),
                legend=dict(
                    orientation="v", yanchor="middle", y=0.5,
                    xanchor="left", x=1.02, font=dict(size=10),
                ),
            )
            st.plotly_chart(fig_prof, use_container_width=True)

        with col_v4:
            st.markdown("#### New Volunteer Joins Over Time")
            # Parse join date — handle multiple formats gracefully
            vol_df2 = vol_df.copy()
            vol_df2["Join_dt"] = pd.to_datetime(
                vol_df2["Joined(ev)"], errors="coerce"
            )
            vol_df2 = vol_df2.dropna(subset=["Join_dt"])
            vol_df2["YearMonth"] = vol_df2["Join_dt"].dt.to_period("M").astype(str)
            monthly = (
                vol_df2.groupby("YearMonth")["Volunteer id"].nunique()
                .reset_index()
                .rename(columns={"Volunteer id": "New Volunteers"})
                .sort_values("YearMonth")
                .tail(24)           # last 24 months
            )
            fig_trend = px.area(
                monthly, x="YearMonth", y="New Volunteers",
                markers=True,
                color_discrete_sequence=[ACCENT],
            )
            fig_trend.update_traces(
                line=dict(width=2.5),
                marker=dict(size=5, color=ACCENT),
                fillcolor="rgba(0,148,201,0.12)",
                hovertemplate="<b>%{x}</b><br>New VTs: %{y:,}<extra></extra>",
            )
            fig_trend.update_layout(
                xaxis_title="Month", yaxis_title="New Volunteers",
                height=380,
                plot_bgcolor=PLOT_BG, paper_bgcolor=PLOT_BG,
                margin=dict(l=0, r=0, t=40, b=60),
            )
            fig_trend.update_xaxes(tickangle=-45, showgrid=False)
            fig_trend.update_yaxes(showgrid=True, gridcolor=GRID_COLOR)
            st.plotly_chart(fig_trend, use_container_width=True)

        st.markdown("---")

        # ── ROW 3: Attendance distribution histogram ──────────────────────────
        st.markdown("#### Attendance Distribution Across Volunteers")
        fig_hist = px.histogram(
            df, x="Attendance%", nbins=30,
            color_discrete_sequence=["#6c5ce7"],
            opacity=0.85,
        )
        fig_hist.update_layout(
            xaxis_title="Attendance %", yaxis_title="Count of Sessions",
            height=280,
            plot_bgcolor=PLOT_BG, paper_bgcolor=PLOT_BG,
            margin=dict(l=0, r=0, t=30, b=0),
            bargap=0.05,
        )
        fig_hist.update_xaxes(showgrid=False)
        fig_hist.update_yaxes(showgrid=True, gridcolor=GRID_COLOR)
        st.plotly_chart(fig_hist, use_container_width=True)

    # ═════════════════════════════════════════════════════════════════════════
    # TAB 3 — CENTER OPERATIONS & ACADEMIC HEALTH
    # ═════════════════════════════════════════════════════════════════════════
    with tab_ops:
        # ── ROW 1: Scatter (Registered vs Enrolled) ───────────────────────────
        st.markdown("#### Registered vs Enrolled Students — Center-Level Drop-off")
        st.caption(
            "Each bubble is one center. Size = CLH. Points **above** the diagonal = "
            "enrollment exceeded registration (re-enrollments / walk-ins). "
            "Points **below** = student drop-off."
        )
        center_df = (
            df.groupby(["Center name", "State", "Donor"])
            .agg(
                Registered=("Registered", "sum"),
                Enrolled=("Enrolled", "sum"),
                CLH=("CLH", "sum"),
                Attendance=("Attendance%", "mean"),
            )
            .reset_index()
        )
        center_df["Drop-off %"] = (
            (center_df["Registered"] - center_df["Enrolled"])
            / center_df["Registered"].replace(0, pd.NA) * 100
        ).fillna(0).round(1)

        max_val = max(center_df["Registered"].max(), center_df["Enrolled"].max()) + 20

        fig_scatter = px.scatter(
            center_df,
            x="Registered", y="Enrolled",
            size="CLH",
            color="State",
            hover_name="Center name",
            hover_data={
                "Donor": True,
                "Drop-off %": True,
                "CLH": ":,",
                "Registered": True,
                "Enrolled": True,
                "State": False,
            },
            color_discrete_sequence=PALETTE,
            size_max=40,
        )
        # Diagonal reference line (Registered == Enrolled)
        fig_scatter.add_shape(
            type="line",
            x0=0, y0=0, x1=max_val, y1=max_val,
            line=dict(color="#b2bec3", width=1.5, dash="dash"),
        )
        fig_scatter.add_annotation(
            x=max_val * 0.85, y=max_val * 0.88,
            text="No drop-off line",
            showarrow=False, font=dict(size=10, color="#b2bec3"),
        )
        fig_scatter.update_layout(
            xaxis_title="Registered Students",
            yaxis_title="Enrolled Students",
            height=480,
            plot_bgcolor=PLOT_BG, paper_bgcolor=PLOT_BG,
            margin=dict(l=0, r=0, t=20, b=0),
            legend=dict(
                title="State", orientation="v",
                yanchor="top", y=1, xanchor="left", x=1.01,
                font=dict(size=10),
            ),
        )
        fig_scatter.update_xaxes(showgrid=True, gridcolor=GRID_COLOR, zeroline=False)
        fig_scatter.update_yaxes(showgrid=True, gridcolor=GRID_COLOR, zeroline=False)
        st.plotly_chart(fig_scatter, use_container_width=True)

        st.markdown("---")

        # ── ROW 2: Stacked bar (Class Execution) + Attendance heatmap ─────────
        col_o1, col_o2 = st.columns([1.2, 1])

        with col_o1:
            st.markdown("#### Class Execution by Subject")
            exec_df = (
                df.groupby("Subject_clean")[["Completed", "Offline", "Cancelled"]]
                .sum()
                .reset_index()
                .sort_values("Completed", ascending=False)
            )
            exec_melt = exec_df.melt(
                id_vars="Subject_clean",
                value_vars=["Completed", "Offline", "Cancelled"],
                var_name="Status", value_name="Sessions",
            )
            exec_color = {
                "Completed": "#00964d",
                "Offline":   "#fdcb6e",
                "Cancelled": "#ed1c2d",
            }
            fig_exec = px.bar(
                exec_melt,
                x="Sessions", y="Subject_clean",
                color="Status",
                orientation="h",
                color_discrete_map=exec_color,
                text="Sessions",
                barmode="stack",
            )
            fig_exec.update_traces(
                textposition="inside",
                texttemplate="%{text:,}",
                hovertemplate="<b>%{y}</b><br>%{fullData.name}: %{x:,}<extra></extra>",
            )
            fig_exec.update_layout(
                xaxis_title="Sessions", yaxis_title="",
                height=440,
                plot_bgcolor=PLOT_BG, paper_bgcolor=PLOT_BG,
                margin=dict(l=0, r=0, t=40, b=0),
                legend=dict(
                    orientation="h", yanchor="top", y=-0.12,
                    xanchor="center", x=0.5, title_text="",
                ),
            )
            fig_exec.update_xaxes(showgrid=True, gridcolor=GRID_COLOR)
            fig_exec.update_yaxes(showgrid=False)
            st.plotly_chart(fig_exec, use_container_width=True)

        with col_o2:
            st.markdown("#### Average Attendance % by State & Subject")
            heat_df = (
                df.groupby(["State", "Subject_clean"])["Attendance%"]
                .mean()
                .reset_index()
            )
            heat_pivot = heat_df.pivot(
                index="State", columns="Subject_clean", values="Attendance%"
            ).fillna(0)

            fig_heat = px.imshow(
                heat_pivot,
                color_continuous_scale=["#dfe6e9", "#0094c9"],
                aspect="auto",
                text_auto=".0f",
                zmin=50, zmax=100,
            )
            fig_heat.update_traces(
                hovertemplate="<b>%{y}</b> | %{x}<br>Avg Attendance: %{z:.1f}%<extra></extra>"
            )
            fig_heat.update_layout(
                xaxis_title="", yaxis_title="",
                height=440,
                plot_bgcolor=PLOT_BG, paper_bgcolor=PLOT_BG,
                margin=dict(l=0, r=0, t=40, b=80),
                coloraxis_colorbar=dict(title="Att%", ticksuffix="%"),
            )
            fig_heat.update_xaxes(tickangle=-35, side="bottom")
            st.plotly_chart(fig_heat, use_container_width=True)

        st.markdown("---")

        # ── ROW 3: Planned vs Completed completion rate by State ──────────────
        st.markdown("#### Class Completion Rate by State")
        state_exec = (
            df.groupby("State").agg(
                Planned=("Planned", "sum"),
                Completed=("Completed", "sum"),
                Offline=("Offline", "sum"),
                Cancelled=("Cancelled", "sum"),
            ).reset_index()
        )
        state_exec["Completion %"] = (
            (state_exec["Completed"] + state_exec["Offline"])
            / state_exec["Planned"].replace(0, pd.NA) * 100
        ).fillna(0).round(1)
        state_exec = state_exec.sort_values("Completion %", ascending=True)

        fig_comp = px.bar(
            state_exec, x="Completion %", y="State",
            orientation="h",
            text="Completion %",
            color="Completion %",
            color_continuous_scale=["#ed1c2d", "#fdcb6e", "#00964d"],
            range_color=[60, 100],
        )
        fig_comp.update_traces(
            texttemplate="%{text:.1f}%",
            textposition="outside",
            hovertemplate="<b>%{y}</b><br>Completion: %{x:.1f}%<extra></extra>",
            marker_line_width=0,
        )
        fig_comp.update_coloraxes(showscale=False)
        fig_comp.update_layout(
            xaxis_title="Completion Rate (%)", yaxis_title="",
            height=380,
            plot_bgcolor=PLOT_BG, paper_bgcolor=PLOT_BG,
            margin=dict(l=0, r=80, t=20, b=0),
        )
        fig_comp.update_xaxes(showgrid=True, gridcolor=GRID_COLOR, range=[0, 115])
        fig_comp.update_yaxes(showgrid=False)
        st.plotly_chart(fig_comp, use_container_width=True)

        st.markdown("---")

        # ── DATA TABLE ────────────────────────────────────────────────────────
        st.markdown("#### 🗂️ Filtered Operational Data")
        st.caption(
            f"Showing **{len(df):,} rows** matching current filters. "
            "Click a column header to sort."
        )

        # Select the most operationally useful columns for the table
        display_cols = [
            "Volunteer name", "Center name", "State", "Donor",
            "Subject_clean", "Registered", "Enrolled", "Attendance%",
            "CLH", "Planned", "Completed", "Offline", "Cancelled",
        ]
        display_df = df[display_cols].copy()
        display_df.rename(columns={"Subject_clean": "Subject"}, inplace=True)
        display_df["Attendance%"] = display_df["Attendance%"].round(1)

        st.dataframe(
            display_df.sort_values("CLH", ascending=False),
            use_container_width=True,
            hide_index=True,
            height=400,
        )
