"""
ops_dashboard.py
================
Operations & Impact Command Center  —  eVidyaloka VRM
Data source: VRM_May__2026.xlsx (multi-sheet Excel)

Sheets used (VRM)
─────────────────
  Active VT            — primary operational data (one row per vol-offering)
  Dropped VT           — volunteers who dropped their offering(s)
  Newly Registered VT  — all volunteers who registered in the current period

Sheets used (DRM)
─────────────────
  SESSION DUMP         — session-level log (1 row per session)
  DRM                  — centre-level summary
  DRM1                 — alternate centre summary with total attendance
  Offering Details     — one row per offering (grade × subject × centre)
  Active centers       — centre master with registration data
  New Enrolled student — mid-month new enrolments

Column mapping (VRM new file → dashboard internal name)
────────────────────────────────────────────────────────
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
