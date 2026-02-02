"""
FLC Portfolio Optimization Dashboard - Data Module
===================================================
Contains all institutional data extracted from project documents.
Edit the dictionaries/lists below to update dashboard data without touching app code.

DATA SOURCE LEGEND:
  [INTERNAL] = Extracted from FLC project documents
  [METHODOLOGY: Internet] = Framework methodology sourced from published research;
                            applied to FLC internal data where available
"""

import pandas as pd
from datetime import date

# ============================================================================
# INSTITUTIONAL OVERVIEW  [INTERNAL - Fact Sheet & Enrollment Overview]
# ============================================================================

INSTITUTION = {
    "name": "Fort Lewis College",
    "location": "Durango, Colorado",
    "type": "Public Liberal Arts",
    "total_enrollment_f24": 3544,
    "total_enrollment_f25": 3457,
    "undergrad_degree_seeking_f25": 3021,
    "graduate_f25": 160,
    "faculty_fte": 239,
    "staff_fte": 358,
    "student_faculty_ratio": "15:1",
    "avg_class_size": 19,
    "retention_rate_f24": 66.10,
    "pell_eligible_pct": 42,
    "first_gen_pct": 43,
    "minority_pct": 52,
    "native_american_pct": 24,
    "in_state_pct": 42.1,
    "out_of_state_pct": 57.9,
}

# ============================================================================
# ENROLLMENT TRENDS  [INTERNAL - Enrollment Overview PDF]
# ============================================================================

ENROLLMENT_HISTORY = pd.DataFrame({
    "Year": [2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025],
    "Total_Headcount": [3595, 3356, 3335, 3308, 3442, 3550, 3360, 3425, 3544, 3457],
    "UG_Degree_Seeking": [3498, 3204, 3144, 3092, 3207, 3263, 3114, 3075, 3077, 3021],
    "Graduate": [None, None, 10, 32, 47, 79, 94, 107, 92, 105],
    "FY_Students": [775, 705, 754, 760, 812, 960, 850, 815, 751, 777],
    "Continuing": [2307, 2132, 2024, 2010, 2036, 1960, 2021, 2009, 2055, 2009],
})

# Graduate enrollment extended series
GRADUATE_ENROLLMENT = pd.DataFrame({
    "Year": [2014, 2015, 2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025],
    "Graduate": [25, 15, 10, 32, 47, 79, 94, 107, 92, 105, 152, 160],
})

RETENTION_HISTORY = pd.DataFrame({
    "Year": [2015, 2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024],
    "Retention_Rate": [66.34, 58.88, 61.90, 62.33, 68.10, 54.27, 59.96, 62.85, 66.84, 66.10],
})

RETENTION_BY_DEMO = pd.DataFrame({
    "Group": ["Total Population", "First Generation", "Pell Eligible", "Students of Color"],
    "Retention_Rate": [66.1, 60.9, 61.7, 62.6],
})

# ============================================================================
# TOP PROGRAMS  [INTERNAL - Fact Sheet]
# ============================================================================

TOP_MAJORS_ENROLLMENT = pd.DataFrame({
    "Program": [
        "Business Administration", "Psychology", "Engineering",
        "Exercise Physiology", "Environmental Conservation & Mgmt",
        "Criminology & Justice Studies", "Environmental Science",
        "Health Sciences", "Adventure Education", "Computer Information Systems",
    ],
    "Enrollment": [298, 227, 207, 163, 133, 103, 87, 86, 78, 77],
    "Pct_of_Total": [9, 7, 6, 5, 4, 3, 3, 3, 2, 2],
})

DEGREES_AWARDED = pd.DataFrame({
    "Program": [
        "Psychology", "Business Administration", "Environmental Studies",
        "Biology", "Exercise Science", "Art", "Public Health",
        "Accounting", "Engineering", "Criminology",
        "Environmental Science", "Outdoor Education",
        "Computer & Info Technology", "Economics",
        "English", "Elementary Teacher Education",
        "Chemistry", "Sports & Fitness Mgmt",
        "Sociology", "Am Indian/Native Amer Studies",
    ],
    "Degrees_Awarded": [59, 55, 37, 33, 26, 24, 24, 15, 14, 13, 13, 13, 12, 12, 11, 11, 11, 11, 22, 8],
})

# ============================================================================
# BCG GROWTH-SHARE MATRIX  [INTERNAL - BCG-growthMatrixDepts.png]
# ============================================================================

BCG_DATA = pd.DataFrame({
    "Department": [
        "English", "Mathematics", "Health & Human Performance",
        "Biology", "Sociology", "Performing Arts", "Teacher Education",
        "Physics & Engineering", "Chemistry",
        "Business Administration", "Psychology",
        "Accounting", "History",
        "Geosciences", "Art & Design", "Economics", "Political Science",
        "Anthropology", "Environment & Sustainability", "Marketing",
        "Adventure Education", "Philosophy",
    ],
    "SCH_Pct": [
        10.0, 8.5, 7.5, 7.0, 6.0, 5.5, 5.8,
        5.0, 4.5,
        5.8, 6.0,
        1.8, 3.8,
        4.2, 2.5, 2.5, 1.5,
        2.5, 3.0, 1.8,
        2.0, 2.2,
    ],
    "Two_Year_Change": [
        -6.0, -11.0, -3.0, -14.5, -5.5, -3.0, -2.0,
        -1.0, -4.0,
        4.0, 3.0,
        12.0, 3.0,
        -15.5, -18.0, -24.0, -26.0,
        -7.5, -8.0, -8.0,
        -10.0, -11.0,
    ],
    "Quadrant": [
        "Cash Cow", "Cash Cow", "Cash Cow",
        "Cash Cow", "Cash Cow", "Cash Cow", "Cash Cow",
        "Cash Cow", "Cash Cow",
        "Star", "Star",
        "Question Mark", "Question Mark",
        "Concern", "Concern", "Concern", "Concern",
        "Concern", "Concern", "Concern",
        "Concern", "Concern",
    ],
})

BCG_QUADRANT_COLORS = {
    "Star": "#2ecc71",
    "Cash Cow": "#3498db",
    "Question Mark": "#f1c40f",
    "Concern": "#e74c3c",
}

BCG_INSIGHTS = [
    "Stars (High Share, Growing): Business Administration and Psychology show strong market share AND growth - invest to maintain.",
    "Cash Cows (High Share, Declining): English, Math, Biology, HHP generate significant SCH but are declining - optimize efficiency.",
    "Question Marks (Low Share, Growing): Accounting and History show growth potential but small share - evaluate investment.",
    "Concern (Low Share, Declining): Political Science, Economics, Art & Design face both low share and steep declines - restructure or sunset.",
]

# ============================================================================
# PESTLE ANALYSIS  [INTERNAL - PESTLE_Report_FLC.docx & External Forces deck]
# ============================================================================

PESTLE_DATA = {
    "Political": {
        "impact": "High",
        "impact_score": 4,
        "trend": "Negative",
        "factors": [
            "Colorado state funding volatility for higher education",
            "Federal financial aid policy changes (Pell Grant, Title IV)",
            "Native American tuition waiver mandate (federal obligation)",
            "State performance-based funding models",
            "Political pressure on DEI programs in public institutions",
        ],
        "opportunities": [
            "Leverage federal tribal education funding",
            "Advocate for rural institution support in state legislature",
        ],
    },
    "Economic": {
        "impact": "High",
        "impact_score": 5,
        "trend": "Mixed",
        "factors": [
            "Declining state appropriations per student",
            "Rising tuition sensitivity among families",
            "Durango cost of living affecting faculty recruitment",
            "Native American tuition waiver revenue impact (~37% of students)",
            "Economic diversification in Four Corners region",
            "Student debt burden concerns nationally",
        ],
        "opportunities": [
            "Grow graduate programs for additional revenue",
            "Expand dual enrollment pipeline",
            "Develop workforce-aligned certificates",
        ],
    },
    "Social": {
        "impact": "High",
        "impact_score": 4,
        "trend": "Mixed",
        "factors": [
            "Declining college-going rates nationally",
            "Changing student expectations (career-focused outcomes)",
            "Growing demand for flexible/hybrid learning",
            "FLC unique mission serving Native American students (166 tribes)",
            "First-generation students (43%) need additional support",
            "Mental health and wellness demands increasing",
        ],
        "opportunities": [
            "Outdoor recreation lifestyle as recruitment differentiator",
            "Indigenous education leadership positioning",
            "Experiential learning emphasis",
        ],
    },
    "Technological": {
        "impact": "Medium",
        "impact_score": 3,
        "trend": "Opportunity",
        "factors": [
            "AI disruption in curriculum and pedagogy",
            "Need for technology infrastructure upgrades",
            "Online/hybrid program delivery expectations",
            "Data analytics for student success and retention",
            "AI Institute at FLC as emerging strength",
        ],
        "opportunities": [
            "AI Institute partnerships and growth",
            "Technology-enhanced experiential learning",
            "Online graduate program expansion",
        ],
    },
    "Legal": {
        "impact": "Medium",
        "impact_score": 3,
        "trend": "Stable",
        "factors": [
            "Accreditation compliance requirements (HLC)",
            "Title IX and student safety regulations",
            "Federal reporting mandates (IPEDS)",
            "Employment law for faculty/staff",
            "Tribal sovereignty considerations in partnerships",
        ],
        "opportunities": [
            "Streamlined accreditation through proactive compliance",
            "Tribal education partnership agreements",
        ],
    },
    "Environmental": {
        "impact": "Medium",
        "impact_score": 3,
        "trend": "Opportunity",
        "factors": [
            "Climate change impacts on Durango/mountain region",
            "Campus sustainability expectations from students",
            "Environmental science as program strength",
            "Outdoor recreation economy dependency on climate",
            "Wildfire risk to campus and community",
        ],
        "opportunities": [
            "Position as leader in sustainability education",
            "Climate resilience research opportunities",
            "Green campus initiatives for recruitment",
        ],
    },
}

# ============================================================================
# PORTER'S FIVE FORCES  [METHODOLOGY: Internet - Applied to FLC context]
# ============================================================================

PORTERS_DATA = {
    "Competitive Rivalry": {
        "rating": "High",
        "score": 4.5,
        "color": "#e74c3c",
        "description": "Intense competition from CU system, CSU, Western Colorado, and online programs",
        "indicators": [
            {"name": "Number of competing institutions in CO", "value": "30+", "trend": "Increasing"},
            {"name": "FLC market share of CO HS graduates", "value": "~2%", "trend": "Stable"},
            {"name": "Enrollment change vs peers", "value": "-2.5%", "trend": "Declining"},
            {"name": "Tuition discount rate pressure", "value": "High", "trend": "Increasing"},
            {"name": "Online program competition", "value": "Significant", "trend": "Increasing"},
        ],
    },
    "Threat of New Entrants": {
        "rating": "Medium-High",
        "score": 3.5,
        "color": "#e67e22",
        "description": "Online programs and micro-credentials lowering traditional barriers to entry",
        "indicators": [
            {"name": "Accreditation barriers", "value": "High", "trend": "Stable"},
            {"name": "Online program launches (competing)", "value": "Growing", "trend": "Increasing"},
            {"name": "Boot camp / certificate programs", "value": "Moderate", "trend": "Increasing"},
            {"name": "Community college expansion", "value": "Active", "trend": "Increasing"},
            {"name": "Capital requirements barrier", "value": "Moderate", "trend": "Decreasing"},
        ],
    },
    "Bargaining Power of Students": {
        "rating": "High",
        "score": 4.0,
        "color": "#e74c3c",
        "description": "Students have many choices; FLC must compete on value, experience, and outcomes",
        "indicators": [
            {"name": "Yield rate (confirmed to enrolled)", "value": "~87%", "trend": "Improving"},
            {"name": "Summer melt rate (FY)", "value": "12.9%", "trend": "Improving"},
            {"name": "Transfer-out competition", "value": "Moderate", "trend": "Stable"},
            {"name": "Price sensitivity", "value": "High", "trend": "Increasing"},
            {"name": "Information transparency", "value": "High", "trend": "Increasing"},
        ],
    },
    "Bargaining Power of Suppliers": {
        "rating": "Medium-High",
        "score": 3.5,
        "color": "#e67e22",
        "description": "Faculty recruitment challenging due to remote location and salary competition",
        "indicators": [
            {"name": "Faculty with terminal degrees", "value": "98%", "trend": "Stable"},
            {"name": "Durango cost of living", "value": "High", "trend": "Increasing"},
            {"name": "Specialized faculty scarcity", "value": "Moderate", "trend": "Increasing"},
            {"name": "Technology vendor dependency", "value": "Moderate", "trend": "Stable"},
            {"name": "Salary competitiveness vs peers", "value": "Below avg", "trend": "Stable"},
        ],
    },
    "Threat of Substitutes": {
        "rating": "Medium-High",
        "score": 3.5,
        "color": "#e67e22",
        "description": "Online degrees, certificates, and workforce programs offer alternatives to 4-year degree",
        "indicators": [
            {"name": "Online degree program growth", "value": "Rapid", "trend": "Increasing"},
            {"name": "Micro-credential adoption", "value": "Growing", "trend": "Increasing"},
            {"name": "Community college pathways", "value": "Strong", "trend": "Increasing"},
            {"name": "Employer credential acceptance", "value": "Expanding", "trend": "Increasing"},
            {"name": "FLC experiential differentiation", "value": "Strong", "trend": "Stable"},
        ],
    },
}

PORTERS_INSIGHTS = [
    "Overall competitive intensity is HIGH - FLC operates in a challenging market requiring clear differentiation.",
    "FLC's strongest defensive positions: unique Native American mission, outdoor recreation lifestyle, and small liberal arts experience.",
    "Greatest threats: online competition eroding geographic advantage, student price sensitivity, and faculty recruitment in Durango.",
    "Strategic imperative: Leverage unique mission and location as competitive moats while expanding program relevance.",
]

# ============================================================================
# GRAY ASSOCIATES PORTFOLIO ANALYSIS  [METHODOLOGY: Internet - Applied to FLC data]
# ============================================================================
# Gray Associates uses: Market Score (student demand + employment + competition)
# vs. Program Economics (contribution margin). Applied using FLC enrollment and
# BCG data as proxies.

GRAY_ASSOCIATES_DATA = pd.DataFrame({
    "Program": [
        "Business Administration", "Psychology", "Engineering",
        "Exercise Physiology", "Environmental Conservation & Mgmt",
        "Criminology & Justice Studies", "Environmental Science",
        "Health Sciences", "Adventure Education", "Computer Information Systems",
        "Biology", "English", "Mathematics", "Sociology",
        "Art & Design", "Chemistry", "Teacher Education",
        "Economics", "Political Science", "Accounting",
        "History", "Philosophy", "Anthropology",
    ],
    "Enrollment": [
        298, 227, 207, 163, 133, 103, 87, 86, 78, 77,
        70, 65, 55, 50, 40, 45, 48, 30, 25, 35,
        32, 20, 22,
    ],
    "Student_Demand_Score": [
        90, 85, 88, 75, 70, 72, 68, 78, 65, 82,
        65, 45, 40, 55, 35, 42, 60, 38, 30, 70,
        35, 20, 28,
    ],
    "Employment_Score": [
        85, 70, 92, 80, 65, 68, 72, 85, 55, 90,
        60, 40, 55, 50, 30, 55, 65, 45, 35, 75,
        30, 25, 30,
    ],
    "Competition_Score": [
        40, 45, 55, 60, 75, 55, 70, 65, 85, 50,
        45, 40, 50, 50, 60, 55, 45, 50, 55, 55,
        60, 70, 65,
    ],
    "Market_Score": [
        74, 68, 79, 72, 68, 65, 68, 76, 65, 75,
        57, 42, 47, 52, 39, 50, 58, 43, 38, 67,
        39, 34, 38,
    ],
    "Economics_Score": [
        78, 72, 60, 65, 55, 70, 50, 62, 45, 68,
        55, 80, 82, 70, 40, 48, 55, 60, 58, 72,
        65, 75, 50,
    ],
    "Mission_Alignment": [
        "Medium", "High", "High", "High", "High",
        "Medium", "High", "High", "High", "High",
        "High", "Medium", "Medium", "Medium", "Medium",
        "Medium", "High", "Low", "Low", "Medium",
        "Medium", "Low", "Medium",
    ],
    "GA_Recommendation": [
        "Grow", "Grow", "Grow", "Grow", "Sustain",
        "Sustain", "Sustain", "Grow", "Sustain", "Grow",
        "Sustain", "Transform", "Transform", "Sustain", "Sunset Review",
        "Sustain", "Sustain", "Evaluate", "Sunset Review", "Grow",
        "Evaluate", "Evaluate", "Evaluate",
    ],
})

GA_RECOMMENDATION_COLORS = {
    "Grow": "#2ecc71",
    "Sustain": "#3498db",
    "Transform": "#f39c12",
    "Evaluate": "#e67e22",
    "Sunset Review": "#e74c3c",
}

GA_INSIGHTS = [
    "GROW programs (high market + strong economics): Business Admin, Psychology, Engineering, Health Sciences, Computer Info Systems, Exercise Physiology, Accounting.",
    "SUSTAIN programs (solid market, needs efficiency): Environmental programs, Criminology, Biology, Sociology, Teacher Education.",
    "TRANSFORM programs (weak market, strong economics): English and Mathematics generate revenue but face enrollment pressure - innovate delivery.",
    "EVALUATE/SUNSET programs (weak market + economics): Political Science, Philosophy, and Art & Design need strategic review for restructuring or phase-out.",
]

# ============================================================================
# PHASE 2: IMPLEMENTATION TRACKING
# ============================================================================

STRATEGIC_INITIATIVES = pd.DataFrame({
    "ID": [f"SI-{i:03d}" for i in range(1, 16)],
    "Initiative": [
        "Expand Business Administration online offerings",
        "Launch Health Sciences graduate certificate",
        "Restructure Environmental Science/Studies programs",
        "Develop AI Institute partnerships and curriculum",
        "Create Engineering co-op/internship pipeline",
        "Implement retention intervention for First-Gen students",
        "Review and consolidate low-enrollment humanities programs",
        "Expand dual enrollment feeder pipeline",
        "Develop outdoor recreation industry partnerships",
        "Create data-driven advising system",
        "Launch transfer-friendly marketing campaign",
        "Pilot competency-based credentials in IT",
        "Strengthen Native American student support services",
        "Develop faculty recruitment incentive package",
        "Establish program-level KPI dashboard",
    ],
    "Phase": [
        "Phase 2", "Phase 2", "Phase 2", "Phase 2", "Phase 2",
        "Phase 2", "Phase 2", "Phase 2", "Phase 2", "Phase 2",
        "Phase 2", "Phase 2", "Phase 2", "Phase 2", "Phase 2",
    ],
    "Source_Framework": [
        "BCG / Gray Associates", "Gray Associates", "BCG / Gray Associates",
        "PESTLE", "Porter's / Gray Associates", "PESTLE / Institutional Data",
        "BCG / Gray Associates", "Institutional Data", "Porter's Five Forces",
        "PESTLE / Institutional Data", "Porter's Five Forces",
        "Gray Associates / Porter's", "PESTLE / Mission", "Porter's Five Forces",
        "All Frameworks",
    ],
    "Department": [
        "Business", "Health Sciences", "Environment & Sustainability",
        "Computer Science / AI", "Engineering", "Student Affairs",
        "Arts & Humanities", "Enrollment Management", "Adventure Education",
        "Academic Affairs", "Enrollment Management", "Computer Info Systems",
        "Student Affairs", "Academic Affairs", "Provost Office",
    ],
    "Priority": [
        "High", "High", "High", "High", "Medium",
        "High", "Medium", "Medium", "Medium", "High",
        "Medium", "Low", "High", "Medium", "High",
    ],
    "Status": [
        "In Progress", "Planning", "In Progress", "Planning", "Not Started",
        "In Progress", "Planning", "In Progress", "Not Started", "Planning",
        "In Progress", "Not Started", "In Progress", "Not Started", "Planning",
    ],
    "Completion_Pct": [
        35, 10, 25, 15, 0,
        40, 10, 30, 0, 20,
        45, 0, 50, 0, 15,
    ],
    "Start_Date": [
        "2025-09-01", "2025-11-01", "2025-09-01", "2025-10-01", "2026-03-01",
        "2025-08-01", "2025-11-01", "2025-08-01", "2026-01-01", "2025-10-01",
        "2025-09-01", "2026-06-01", "2025-08-01", "2026-01-01", "2025-10-01",
    ],
    "Target_Date": [
        "2026-08-01", "2026-08-01", "2026-12-01", "2027-01-01", "2026-12-01",
        "2026-05-01", "2026-12-01", "2026-08-01", "2026-08-01", "2026-05-01",
        "2026-03-01", "2027-01-01", "2026-08-01", "2026-08-01", "2026-03-01",
    ],
    "Owner": [
        "Dean of Business", "Dean of Health Sciences", "Provost",
        "AI Institute Director", "Dean of Engineering", "VP Student Affairs",
        "Provost", "VP Enrollment", "Dean of Education",
        "Provost", "VP Enrollment", "Dept Chair CIS",
        "VP Student Affairs", "VP Academic Affairs", "Provost",
    ],
})

MILESTONES = pd.DataFrame({
    "ID": ["MS-001", "MS-002", "MS-003", "MS-004", "MS-005", "MS-006",
           "MS-007", "MS-008", "MS-009", "MS-010"],
    "Milestone": [
        "Phase 1 Analysis Complete",
        "Board Presentation of Findings",
        "Implementation Plans Approved",
        "Q1 Progress Review",
        "Retention Intervention Pilot Launch",
        "Online Program Proposal Submitted",
        "Dual Enrollment Partnerships Signed",
        "Mid-Year Progress Report",
        "Program Restructuring Plans Finalized",
        "Year 1 Implementation Review",
    ],
    "Target_Date": [
        "2025-12-15", "2026-01-20", "2026-02-15", "2026-03-31",
        "2026-02-01", "2026-03-01", "2026-04-15", "2026-06-30",
        "2026-05-01", "2026-08-15",
    ],
    "Status": [
        "Complete", "Complete", "In Progress", "Upcoming",
        "In Progress", "In Progress", "Not Started", "Upcoming",
        "Not Started", "Upcoming",
    ],
    "Notes": [
        "BCG, PESTLE, Porter's, Gray Associates analyses delivered",
        "Presented to Board of Trustees at January retreat",
        "Department chairs reviewing implementation details",
        "Scheduled for March 31",
        "First-gen student cohort identified, mentors assigned",
        "Business Admin online MBA proposal in review",
        "Target: San Juan College, Pueblo CC, Red Rocks CC",
        "All initiatives report progress mid-year",
        "Environmental programs and humanities consolidation",
        "Full assessment of Year 1 outcomes",
    ],
})

KPIS = pd.DataFrame({
    "KPI": [
        "Total Enrollment", "Retention Rate (FTFT)",
        "Graduate Enrollment", "First-Year Class Size",
        "Degrees Awarded", "Student-Faculty Ratio",
        "Online Course Offerings", "Dual Enrollment Students",
        "Transfer Students", "Program Completion Rate",
        "First-Gen Retention Gap", "Native American Retention",
    ],
    "Category": [
        "Enrollment", "Retention", "Enrollment", "Enrollment",
        "Outcomes", "Efficiency", "Growth", "Growth",
        "Enrollment", "Outcomes", "Equity", "Equity",
    ],
    "Current_Value": [
        3457, 66.1, 160, 777, 489, 15, 25, 235, 190, 42, 5.2, 61,
    ],
    "Target_Value": [
        3700, 72.0, 200, 850, 525, 14, 60, 300, 220, 48, 2.0, 68,
    ],
    "Unit": [
        "students", "%", "students", "students",
        "degrees", ":1", "courses", "students",
        "students", "%", "pp", "%",
    ],
    "Timeframe": [
        "Fall 2027", "Fall 2027", "Fall 2027", "Fall 2027",
        "FY 2026-27", "Fall 2027", "Fall 2027", "Fall 2027",
        "Fall 2027", "FY 2026-27", "Fall 2027", "Fall 2027",
    ],
})

RESOURCE_ALLOCATION = pd.DataFrame({
    "Category": [
        "Faculty Hiring (STEM/Health)", "Online Program Development",
        "Student Support Services", "Technology Infrastructure",
        "Marketing & Recruitment", "Dual Enrollment Expansion",
        "Faculty Development", "Data Analytics Platform",
    ],
    "Allocated_Budget": [
        450000, 300000, 250000, 200000, 350000, 150000, 100000, 125000,
    ],
    "Spent_to_Date": [
        120000, 85000, 150000, 60000, 200000, 45000, 30000, 50000,
    ],
    "Department": [
        "Academic Affairs", "Academic Affairs", "Student Affairs",
        "IT", "Enrollment Management", "Academic Affairs",
        "Provost Office", "Institutional Research",
    ],
})

# ============================================================================
# DATA SOURCE ATTRIBUTION
# ============================================================================

DATA_SOURCES = {
    "PESTLE Analysis": {
        "source": "Internal FLC Documents",
        "files": ["PESTLE_Report_FLC.docx", "External Forces Shaping FLC.pptx"],
        "icon": "check-circle",
        "badge_color": "#2ecc71",
    },
    "BCG Growth-Share Matrix": {
        "source": "Internal FLC Documents",
        "files": ["BCG Presentation.pptx", "BCG-growthMatrixDepts.png"],
        "icon": "check-circle",
        "badge_color": "#2ecc71",
    },
    "Porter's Five Forces": {
        "source": "Internet Methodology + FLC Data",
        "files": ["Framework from published research; metrics from FLC institutional data"],
        "icon": "info-circle",
        "badge_color": "#3498db",
    },
    "Gray Associates Portfolio": {
        "source": "Internet Methodology + FLC Data",
        "files": ["Gray Associates PES methodology; applied to FLC enrollment & BCG data"],
        "icon": "info-circle",
        "badge_color": "#3498db",
    },
    "SWOT Analysis": {
        "source": "Synthesized from All Phase 1 Frameworks",
        "files": ["PESTLE, Porter's, Gray Associates, BCG analyses + FLC institutional data"],
        "icon": "check-circle",
        "badge_color": "#8e44ad",
    },
    "Zone to Win": {
        "source": "Synthesized from SWOT + All Frameworks",
        "files": ["Zone to Win methodology (Geoffrey Moore); applied to FLC strategic context"],
        "icon": "info-circle",
        "badge_color": "#8e44ad",
    },
    "Strategic Roadmap": {
        "source": "Synthesized from All Phases",
        "files": ["Implementation plan derived from Zone to Win scenarios + Phase 1 analyses"],
        "icon": "info-circle",
        "badge_color": "#8e44ad",
    },
}

# ============================================================================
# FRAMEWORK DESCRIPTIONS (2-3 sentences for top of each Phase 1 tab)
# ============================================================================

FRAMEWORK_DESCRIPTIONS = {
    "PESTLE": (
        "PESTLE Analysis examines six categories of external macro-environmental factors "
        "that influence Fort Lewis College's strategic positioning: Political, Economic, Social, "
        "Technological, Legal, and Environmental. This framework identifies forces beyond the "
        "institution's direct control that create both risks and opportunities for Academic Affairs. "
        "Each factor is assessed for impact severity and directional trend to prioritize strategic responses."
    ),
    "Porters": (
        "Porter's Five Forces framework assesses the competitive intensity and attractiveness "
        "of the higher education market in which Fort Lewis College operates. By analyzing the "
        "threat of new entrants, bargaining power of suppliers (faculty/vendors) and buyers (students), "
        "threat of substitutes, and rivalry among existing competitors, this model reveals FLC's "
        "competitive position and informs differentiation strategy."
    ),
    "Gray": (
        "Gray Associates Portfolio Analysis evaluates academic programs using a data-driven "
        "methodology that plots Market Score (student demand, employment outlook, and competitive "
        "positioning) against Program Economics (revenue efficiency and contribution margin). "
        "This framework classifies programs into actionable categories\u2014Grow, Sustain, Transform, "
        "Evaluate, or Sunset Review\u2014to guide investment and restructuring decisions."
    ),
    "BCG": (
        "The BCG Growth-Share Matrix, adapted from the Boston Consulting Group framework, "
        "categorizes FLC's academic departments based on two dimensions: relative market share "
        "(measured by percentage of total Student Credit Hours) and growth rate (2-year enrollment "
        "change). Programs are classified as Stars, Cash Cows, Question Marks, or Concerns to "
        "guide resource allocation priorities."
    ),
}

# ============================================================================
# PHASE 2: SWOT ANALYSIS  [Synthesized from all Phase 1 frameworks]
# ============================================================================

SWOT_DATA = {
    "Strengths": {
        "color": "#2ecc71",
        "icon": "S",
        "items": [
            {
                "title": "Unique Native American Mission",
                "detail": "Federal obligation to serve Native American students with tuition waiver creates a distinctive institutional identity serving 166 tribes. 37% of students receive the Native American Tuition Waiver.",
                "source": "PESTLE (Social/Political), Institutional Data",
            },
            {
                "title": "Strong Star Programs",
                "detail": "Business Administration (298 enrolled, +4% growth) and Psychology (227 enrolled, +3% growth) demonstrate both high market share and positive trajectory.",
                "source": "BCG Matrix, Gray Associates",
            },
            {
                "title": "High-SCH Cash Cow Programs",
                "detail": "Nine departments (English, Math, Biology, HHP, etc.) generate the bulk of student credit hours, providing a stable revenue foundation despite enrollment softness.",
                "source": "BCG Matrix",
            },
            {
                "title": "Outdoor Recreation & Location Differentiator",
                "detail": "Durango's mountain setting and outdoor lifestyle create a powerful recruitment differentiator that online competitors cannot replicate. Adventure Education is uniquely positioned.",
                "source": "Porter's Five Forces, PESTLE (Social/Environmental)",
            },
            {
                "title": "Growing Graduate Programs",
                "detail": "Graduate enrollment has grown from 10 (2016) to 160 (Fall 2025), a 16x increase demonstrating capacity to launch and scale new credential levels.",
                "source": "Institutional Data, Gray Associates",
            },
            {
                "title": "Small Class Sizes & Faculty Quality",
                "detail": "Average class size of 19, 100% of classes under 50 students, 98% of tenure-track faculty hold terminal degrees. 15:1 student-faculty ratio.",
                "source": "Institutional Data",
            },
            {
                "title": "Strong Employment-Aligned Programs",
                "detail": "Engineering (92), Computer Info Systems (90), Business Admin (85), and Health Sciences (85) score highest on employment outlook in Gray Associates analysis.",
                "source": "Gray Associates",
            },
        ],
    },
    "Weaknesses": {
        "color": "#e74c3c",
        "icon": "W",
        "items": [
            {
                "title": "Declining Undergraduate Enrollment",
                "detail": "UG degree-seeking enrollment fell from 3,498 (2016) to 3,021 (2025), a -13.6% decline over 10 years. Total headcount down -2.5% YoY.",
                "source": "Institutional Data, Enrollment Overview",
            },
            {
                "title": "Multiple Concern-Quadrant Programs",
                "detail": "Nine departments in BCG Concern quadrant: Political Science (-26%), Economics (-24%), Art & Design (-18%), Geosciences (-15.5%) face both low market share and steep enrollment declines.",
                "source": "BCG Matrix",
            },
            {
                "title": "Retention Below National Average",
                "detail": "66.1% FTFT retention rate is below the national average for public 4-year institutions (~73%). Equity gaps persist: First-Gen (60.9%), Pell (61.7%), Students of Color (62.6%).",
                "source": "Institutional Data, PESTLE (Social)",
            },
            {
                "title": "Remote Location Faculty Recruitment",
                "detail": "Durango's high cost of living and geographic isolation create persistent challenges in attracting and retaining specialized faculty. Salary competitiveness is below average.",
                "source": "Porter's Five Forces (Supplier Power)",
            },
            {
                "title": "Tuition Waiver Revenue Impact",
                "detail": "The Native American tuition waiver, while mission-critical, affects revenue generation with ~37% of students receiving the waiver, creating dependency on state funding.",
                "source": "PESTLE (Economic/Political)",
            },
            {
                "title": "Limited Online Program Offerings",
                "detail": "Only 25 online course offerings currently, significantly limiting reach and competitiveness against institutions with robust online portfolios.",
                "source": "Porter's Five Forces, Gray Associates",
            },
        ],
    },
    "Opportunities": {
        "color": "#3498db",
        "icon": "O",
        "items": [
            {
                "title": "Expand High-Demand Programs",
                "detail": "Gray Associates identifies 7 programs for GROW status: Business Admin, Psychology, Engineering, Health Sciences, Computer Info Systems, Exercise Physiology, Accounting. These align with strong employment markets.",
                "source": "Gray Associates, BCG Matrix",
            },
            {
                "title": "Graduate Program Expansion",
                "detail": "16x growth in graduate enrollment (10 to 160) over 9 years demonstrates untapped capacity. Health Sciences graduate certificate and online MBA are immediate opportunities.",
                "source": "Institutional Data, Gray Associates",
            },
            {
                "title": "AI Institute Development",
                "detail": "FLC's AI Institute represents an emerging strength in a high-demand field. Partnerships and curriculum integration can attract new student segments and research funding.",
                "source": "PESTLE (Technological), Institutional Data",
            },
            {
                "title": "Dual Enrollment Pipeline Growth",
                "detail": "Dual enrollment grew from 52 (2016) to 235 (2025), 4.5x increase. 27 prior dual-enrollment students converted to degree-seeking in Fall 2025. Expansion partnerships with San Juan College, Pueblo CC, and Red Rocks CC are viable.",
                "source": "Institutional Data, Enrollment Overview",
            },
            {
                "title": "Sustainability & Environmental Leadership",
                "detail": "FLC's Environmental Conservation & Management (133 enrolled) and Environmental Science (87 enrolled) programs, combined with Durango's setting, position the institution for climate/sustainability leadership.",
                "source": "PESTLE (Environmental), Gray Associates",
            },
            {
                "title": "Indigenous Education National Leadership",
                "detail": "Serving 166 tribes with 26.5% Native American enrollment creates opportunity to become the premier Indigenous higher education institution nationally, attracting federal grants and partnerships.",
                "source": "PESTLE (Social/Political), Institutional Data",
            },
        ],
    },
    "Threats": {
        "color": "#e67e22",
        "icon": "T",
        "items": [
            {
                "title": "Intensifying Online Competition",
                "detail": "Online programs from large universities erode FLC's geographic advantage. Porter's rates Competitive Rivalry at 4.5/5 and Threat of Substitutes at 3.5/5.",
                "source": "Porter's Five Forces",
            },
            {
                "title": "State Funding Volatility",
                "detail": "Colorado state appropriations per student are declining. Performance-based funding models create additional uncertainty for smaller institutions.",
                "source": "PESTLE (Political/Economic)",
            },
            {
                "title": "Declining College-Going Rates",
                "detail": "National college-going rates are declining, particularly affecting small public liberal arts institutions. Colorado first-year student pipeline down -7.6% for FLC in 2025.",
                "source": "PESTLE (Social), Enrollment Overview",
            },
            {
                "title": "Student Price Sensitivity",
                "detail": "Porter's rates Bargaining Power of Students at 4.0/5 (High). Rising tuition sensitivity, student debt concerns, and increasing discount rate pressure threaten net revenue.",
                "source": "Porter's Five Forces, PESTLE (Economic)",
            },
            {
                "title": "Alternative Credential Growth",
                "detail": "Micro-credentials, boot camps, and certificate programs are growing rapidly. Employers increasingly accept alternative credentials, threatening traditional 4-year degree demand.",
                "source": "Porter's Five Forces (Substitutes)",
            },
            {
                "title": "Political Pressure on DEI & Public Higher Ed",
                "detail": "Political landscape creates uncertainty for diversity programs, public institution funding, and federal financial aid policy (Pell Grant, Title IV).",
                "source": "PESTLE (Political)",
            },
        ],
    },
}

# ============================================================================
# PHASE 3: ZONE TO WIN  [Geoffrey Moore framework applied to FLC]
# ============================================================================

ZONE_TO_WIN_DATA = {
    "Performance Zone": {
        "color": "#2ecc71",
        "description": "Revenue maintenance and growth in existing strong programs. Focus on scaling proven programs that drive enrollment and tuition revenue.",
        "programs": [
            {"name": "Business Administration", "action": "Expand online offerings, add MBA track", "investment": "High"},
            {"name": "Psychology", "action": "Grow applied psychology tracks, add graduate options", "investment": "High"},
            {"name": "Engineering", "action": "Strengthen co-op/internship pipeline, industry partnerships", "investment": "High"},
            {"name": "Health Sciences", "action": "Launch graduate certificate, expand clinical partnerships", "investment": "High"},
            {"name": "Computer Information Systems", "action": "Align with AI Institute, add cybersecurity track", "investment": "Medium"},
            {"name": "Exercise Physiology", "action": "Develop sports medicine specializations", "investment": "Medium"},
        ],
    },
    "Productivity Zone": {
        "color": "#3498db",
        "description": "Enabling investments for operational efficiency and effectiveness across academic and administrative support functions.",
        "programs": [
            {"name": "Advising System Overhaul", "action": "Implement data-driven predictive advising platform", "investment": "High"},
            {"name": "Retention Programs", "action": "First-Gen/Pell student intervention programs, Compass expansion", "investment": "High"},
            {"name": "IT Infrastructure", "action": "LMS upgrade, data analytics platform, classroom technology", "investment": "Medium"},
            {"name": "Faculty Recruitment Package", "action": "Housing assistance, salary competitiveness, remote work options", "investment": "Medium"},
            {"name": "Transfer Pathway Optimization", "action": "Streamline articulation agreements with top feeder schools", "investment": "Low"},
            {"name": "Marketing & Communications", "action": "Targeted digital recruitment, brand positioning refresh", "investment": "Medium"},
        ],
    },
    "Incubation Zone": {
        "color": "#f39c12",
        "description": "Disciplined experimentation with emerging academic, administrative, community, or external opportunities that could become future revenue streams.",
        "programs": [
            {"name": "AI Institute Expansion", "action": "Corporate partnerships, research grants, certificate programs", "investment": "Medium"},
            {"name": "Online Degree Programs", "action": "Pilot 2-3 fully online bachelor's completions (Business, Psychology)", "investment": "Medium"},
            {"name": "Micro-Credentials & Badges", "action": "Stackable certificates in IT, sustainability, outdoor leadership", "investment": "Low"},
            {"name": "Dual Enrollment Expansion", "action": "New partnerships with regional high schools and community colleges", "investment": "Low"},
            {"name": "Workforce Development Partnerships", "action": "Employer-sponsored programs in healthcare, technology, outdoor industry", "investment": "Low"},
        ],
    },
    "Transformation Zone": {
        "color": "#9b59b6",
        "description": "Strategic bets on future-defining innovations and new markets that ensure a thriving Academic Affairs at FLC over the long term.",
        "programs": [
            {"name": "Indigenous Education Hub", "action": "National center for Indigenous higher education research and practice", "investment": "High"},
            {"name": "Sustainability & Climate Institute", "action": "Leverage location and programs for interdisciplinary climate research center", "investment": "Medium"},
            {"name": "Experiential Learning Model", "action": "Rebrand FLC as the premier outdoor experiential learning institution nationally", "investment": "Medium"},
            {"name": "Program Portfolio Restructuring", "action": "Consolidate/transform Concern-quadrant humanities into interdisciplinary programs", "investment": "Medium"},
        ],
    },
}

# Three Strategic Scenarios
SCENARIOS = {
    "Optimistic": {
        "description": "Favorable market conditions, successful execution of all strategic initiatives, and strong institutional and state support.",
        "color": "#2ecc71",
        "enrollment_target": 3800,
        "retention_target": 75.0,
        "graduate_target": 250,
        "online_courses": 80,
        "new_programs": 5,
        "assumptions": [
            "State funding increases 3-5% annually",
            "Successful launch of 3+ online degree programs",
            "AI Institute secures major grant funding",
            "Retention interventions close equity gaps by 50%",
            "Dual enrollment pipeline exceeds 350 students",
        ],
        "zone_allocation": {"Performance": 40, "Productivity": 25, "Incubation": 20, "Transformation": 15},
    },
    "Most Likely": {
        "description": "Realistic constraints with incremental progress on initiatives, moderate state support, and steady but competitive market conditions.",
        "color": "#f39c12",
        "enrollment_target": 3550,
        "retention_target": 70.0,
        "graduate_target": 200,
        "online_courses": 50,
        "new_programs": 3,
        "assumptions": [
            "State funding flat or slight increase (0-2%)",
            "1-2 online programs launched successfully",
            "AI Institute grows but external funding uncertain",
            "Retention improves incrementally (2-3 pp)",
            "Dual enrollment grows to 280-300 students",
        ],
        "zone_allocation": {"Performance": 45, "Productivity": 30, "Incubation": 15, "Transformation": 10},
    },
    "Conservative": {
        "description": "Challenging conditions including potential funding cuts, continued enrollment pressure, and increased competition, while maintaining core institutional strengths.",
        "color": "#e74c3c",
        "enrollment_target": 3300,
        "retention_target": 67.0,
        "graduate_target": 170,
        "online_courses": 35,
        "new_programs": 1,
        "assumptions": [
            "State funding decreases 2-5%",
            "Online competition intensifies significantly",
            "Limited resources for new program launches",
            "Focus on protecting core programs and retention",
            "Accelerate program sunset reviews for cost savings",
        ],
        "zone_allocation": {"Performance": 50, "Productivity": 35, "Incubation": 10, "Transformation": 5},
    },
}

# ============================================================================
# PHASE 3: STRATEGIC ROADMAP
# ============================================================================

ROADMAP_MILESTONES = pd.DataFrame({
    "ID": [f"RM-{i:03d}" for i in range(1, 21)],
    "Milestone": [
        "Phase 1 Framework Analyses Complete",
        "SWOT Synthesis Delivered to Provost",
        "Zone to Win Scenarios Presented to Board",
        "Program Sunset Review Initiated (Concern programs)",
        "Retention Intervention Pilot Launched",
        "Online Program Task Force Established",
        "AI Institute Partnership MOU Signed",
        "Dual Enrollment Expansion Agreements (3 schools)",
        "Faculty Recruitment Incentive Package Approved",
        "Business Admin Online MBA Proposal Submitted",
        "Q1 KPI Review & Course Correction",
        "Indigenous Education Hub Feasibility Study",
        "Sustainability Institute Concept Paper",
        "Mid-Year Strategic Progress Report",
        "Program Restructuring Plans Finalized",
        "Year 1 Online Program Enrollment Results",
        "Budget Reallocation Based on Zone Performance",
        "Year 1 Comprehensive Implementation Review",
        "Year 2 Strategic Plan Refinement",
        "Board Presentation: 2-Year Progress & Future Direction",
    ],
    "Phase": [
        "Phase 1", "Phase 2", "Phase 3", "Phase 3", "Phase 3",
        "Phase 3", "Phase 3", "Phase 3", "Phase 3", "Phase 3",
        "Phase 3", "Phase 3", "Phase 3", "Phase 3", "Phase 3",
        "Phase 3", "Phase 3", "Phase 3", "Phase 3", "Phase 3",
    ],
    "Start_Date": [
        "2025-09-01", "2025-12-01", "2026-01-15", "2026-02-01", "2026-02-01",
        "2026-02-15", "2026-03-01", "2026-03-15", "2026-03-01", "2026-04-01",
        "2026-04-01", "2026-04-15", "2026-05-01", "2026-06-01", "2026-06-15",
        "2026-09-15", "2026-10-01", "2026-12-01", "2027-01-15", "2027-06-01",
    ],
    "Target_Date": [
        "2025-12-15", "2026-01-15", "2026-02-01", "2026-05-01", "2026-03-01",
        "2026-03-15", "2026-06-01", "2026-05-15", "2026-05-01", "2026-06-01",
        "2026-04-30", "2026-08-01", "2026-08-01", "2026-07-01", "2026-08-15",
        "2026-11-01", "2026-11-15", "2027-01-15", "2027-03-01", "2027-07-01",
    ],
    "Status": [
        "Complete", "Complete", "In Progress", "In Progress", "In Progress",
        "In Progress", "Not Started", "Not Started", "Not Started", "Not Started",
        "Upcoming", "Not Started", "Not Started", "Upcoming", "Not Started",
        "Upcoming", "Upcoming", "Upcoming", "Upcoming", "Upcoming",
    ],
    "Zone": [
        "All", "All", "All", "Performance", "Productivity",
        "Incubation", "Incubation", "Incubation", "Productivity", "Performance",
        "All", "Transformation", "Transformation", "All", "Performance",
        "Incubation", "All", "All", "All", "All",
    ],
    "Owner": [
        "Consulting Team", "Provost", "Provost/President", "Provost", "VP Student Affairs",
        "VP Academic Affairs", "AI Institute Director", "VP Enrollment", "VP Academic Affairs", "Dean of Business",
        "Provost", "Provost", "Dean of Sciences", "Provost", "Provost",
        "VP Academic Affairs", "CFO/Provost", "President/Provost", "Provost", "President",
    ],
})

ROADMAP_KPIS = pd.DataFrame({
    "KPI": [
        "Total Enrollment", "FTFT Retention Rate", "Graduate Enrollment",
        "First-Year Class Size", "Degrees Awarded", "Online Course Offerings",
        "Dual Enrollment Students", "Transfer Students",
        "First-Gen Retention Gap", "Native American Retention Rate",
        "Programs in Grow/Sustain Status", "Program Completion Rate",
    ],
    "Category": [
        "Enrollment", "Retention", "Enrollment", "Enrollment",
        "Outcomes", "Growth", "Growth", "Enrollment",
        "Equity", "Equity", "Portfolio Health", "Outcomes",
    ],
    "Baseline_Value": [3457, 66.1, 160, 777, 489, 25, 235, 190, 5.2, 61.0, 14, 42],
    "Year1_Target": [3500, 68.0, 180, 800, 500, 40, 270, 200, 4.0, 63.0, 16, 44],
    "Year2_Target": [3600, 70.0, 200, 830, 520, 60, 300, 215, 3.0, 66.0, 17, 47],
    "Stretch_Target": [3800, 75.0, 250, 870, 550, 80, 350, 240, 2.0, 70.0, 19, 50],
    "Unit": ["students", "%", "students", "students", "degrees", "courses",
             "students", "students", "pp", "%", "programs", "%"],
    "Measurement": [
        "Fall census", "Fall-to-Fall FTFT", "Fall census", "Fall census",
        "Annual", "Fall semester", "Fall census", "Fall census",
        "Total pop minus First-Gen", "Fall-to-Fall AIAN", "Gray Associates assessment", "6-year rate",
    ],
})

RISK_MITIGATION = pd.DataFrame({
    "Risk": [
        "State funding cut exceeds 5%",
        "Online program launch delays",
        "Key faculty departures",
        "Enrollment falls below 3,200",
        "Retention drops below 62%",
        "AI Institute funding not secured",
        "Political pressure on DEI programs",
        "Community college competition intensifies",
    ],
    "Probability": ["Medium", "Medium", "High", "Low", "Low", "Medium", "High", "Medium"],
    "Impact": ["High", "Medium", "Medium", "High", "High", "Medium", "High", "Medium"],
    "Mitigation_Strategy": [
        "Diversify revenue: grow graduate programs, online offerings, and auxiliary revenue. Build 6-month operating reserve.",
        "Phase launches incrementally; start with hybrid delivery before full online. Maintain parallel in-person tracks.",
        "Implement faculty retention package (housing, salary). Build succession plans for critical positions.",
        "Activate emergency recruitment campaign. Accelerate dual enrollment and transfer pipelines.",
        "Scale Compass program. Deploy early-alert system. Increase advisor-to-student ratio for at-risk populations.",
        "Pursue alternative funding (NSF, private sector). Scale down to pilot-size if grants not secured.",
        "Frame programs under student success and institutional mission. Diversify terminology while maintaining commitment.",
        "Differentiate on residential experience, outdoor lifestyle, and 4-year degree completion. Strengthen transfer articulation.",
    ],
    "Owner": [
        "CFO", "VP Academic Affairs", "VP Academic Affairs", "VP Enrollment",
        "VP Student Affairs", "AI Institute Director", "President/General Counsel", "VP Enrollment",
    ],
})
