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
# BCG GROWTH-SHARE MATRIX  [INTERNAL - Dataset_Majors.xlsx percentChange tab]
# ============================================================================
# 48 majors with 2022 vs 2024 enrollment (Select Majors Combined).
# Quadrants computed at module load: X = Enrollment_2024, Y = Pct_Change.
# Small_Base flags programs with < 20 students in 2022 (% change unreliable).

_bcg_raw = pd.DataFrame({
    "Major": [
        "Accounting", "Adventure Education", "Anthropology",
        "Art K-12 Education", "Biochemistry",
        "Biology/Cellular & Molecular Biology", "Borders & Languages",
        "Business Administration", "Chemistry",
        "Communication Design", "Computer Engineering",
        "Computer Information Systems", "Criminology & Justice Studies",
        "Early Childhood Education", "Economics",
        "Educational Studies", "Elementary Education",
        "Engineering", "English", "English Secondary Education",
        "Entrepreneurship & Small Business", "Environmental Conservation and Management",
        "Environmental Science", "Environmental Studies",
        "Exercise & Health Promotion", "Exercise Physiology",
        "Exercise Science K-12 Education", "Gender & Sexuality Studies",
        "Geology", "Health Sciences", "History",
        "Journalism & Multimedia Studies", "Marketing",
        "Mathematics", "Music",
        "Native American & Indigenous Studies", "Nutrition",
        "Philosophy", "Physics",
        "Political Science", "Pre-Major Accounting",
        "Psychology", "Public Health",
        "Sociology and Human Services", "Sport Administration",
        "Studio Art", "Theatre", "Writing",
    ],
    "Enrollment_2022": [
        63, 69, 24, 15, 49, 193, 6,
        295, 27, 71, 65, 68, 101,
        7, 38, 7, 18,
        196, 21, 19,
        49, 82, 100, 54,
        35, 137, 5, 7,
        51, 51, 37,
        45, 51, 16, 12,
        26, 20, 9, 14,
        36, 149, 272, 86,
        65, 46, 44, 27, 21,
    ],
    "Enrollment_2024": [
        60, 77, 53, 7, 66, 111, 3,
        325, 27, 48, 73, 77, 103,
        10, 9, 4, 21,
        204, 26, 10,
        40, 133, 86, 56,
        27, 160, 6, 4,
        59, 86, 37,
        34, 68, 21, 27,
        19, 23, 9, 10,
        31, 102, 224, 37,
        65, 40, 57, 15, 20,
    ],
    "Pct_Change": [
        -4.76, 11.59, 120.83, -53.33, 34.69, -42.49, -50.00,
        10.17, 0.00, -32.39, 12.31, 13.24, 1.98,
        42.86, -76.32, -42.86, 16.67,
        4.08, 23.81, -47.37,
        -18.37, 62.20, -14.00, 3.70,
        -22.86, 16.79, 20.00, -42.86,
        15.69, 68.63, 0.00,
        -24.44, 33.33, 31.25, 125.00,
        -26.92, 15.00, 0.00, -28.57,
        -13.89, -31.54, -17.65, -56.98,
        0.00, -13.04, 29.55, -44.44, -4.76,
    ],
})

# Derived columns
_bcg_raw["Abs_Change"] = _bcg_raw["Enrollment_2024"] - _bcg_raw["Enrollment_2022"]
_bcg_raw["Small_Base"] = _bcg_raw["Enrollment_2022"] < 20

# Quartile assignment (matches Excel cutoffs: Q1 <= -28.6%, Q2 <= 0%, Q3 <= 16.7%)
def _assign_quartile(pct):
    if pct <= -28.57:
        return 1
    elif pct <= 0.0:
        return 2
    elif pct <= 16.67:
        return 3
    else:
        return 4

_bcg_raw["Quartile"] = _bcg_raw["Pct_Change"].apply(_assign_quartile)

# BCG Quadrant assignment: X = Enrollment_2024, Y = Pct_Change
_median_enrollment = _bcg_raw["Enrollment_2024"].median()

def _assign_quadrant(row):
    large = row["Enrollment_2024"] >= _median_enrollment
    growing = row["Pct_Change"] > 0
    if large and growing:
        return "Star"
    elif large and not growing:
        return "Cash Cow"
    elif not large and growing:
        return "Question Mark"
    else:
        return "Concern"

_bcg_raw["Quadrant"] = _bcg_raw.apply(_assign_quadrant, axis=1)

BCG_DATA = _bcg_raw.copy()

# Department-level BCG (SCH-based, 22 departments) — original FLC analysis
BCG_DEPT_DATA = pd.DataFrame({
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

BCG_DEPT_INSIGHTS = [
    "Stars (High SCH Share, Growing): Business Administration and Psychology are the only departments with "
    "both large SCH share (>4%) and positive 2-year growth \u2014 invest to sustain momentum.",
    "Cash Cows (High SCH Share, Declining): English (10%), Mathematics (8.5%), Biology (7%), HHP (7.5%), "
    "and Sociology (6%) generate the bulk of institutional SCH but all show 2-year declines \u2014 optimize "
    "efficiency and protect enrollment in these revenue-critical departments.",
    "Question Marks (Low SCH Share, Growing): Accounting (+12%) and History (+3%) show positive trends from "
    "small bases \u2014 evaluate whether growth justifies increased investment.",
    "Concerns (Low SCH Share, Declining): Political Science (\u221226%), Economics (\u221224%), Art & Design "
    "(\u221218%), and Geosciences (\u221215.5%) face both small SCH share and steep declines \u2014 candidates "
    "for structured program review. Note: some serve general education or mission-critical roles.",
]

BCG_QUADRANT_COLORS = {
    "Star": "#2ecc71",
    "Cash Cow": "#3498db",
    "Question Mark": "#f1c40f",
    "Concern": "#e74c3c",
}

BCG_INSIGHTS = [
    "Stars (Large & Growing): Business Administration (+10%, 325 enrolled), Exercise Physiology (+17%, 160), "
    "and Environmental Conservation & Mgmt (+62%, 133) lead FLC's portfolio with strong enrollment and positive growth.",
    "Cash Cows (Large & Declining): Psychology (-18%, 224), Biology/CMB (-42%, 111), and Pre-Major Accounting "
    "(-32%, 102) maintain significant enrollment but face declining trends requiring efficiency optimization.",
    "Question Marks (Small & Growing): Anthropology (+121%, 53) and Music (+125%, 27) show dramatic growth "
    "but from small bases \u2014 evaluate whether growth is sustainable before major investment.",
    "Concerns (Small & Declining): Economics (-76%, 9), MND: Art & Design (-98%, 3), and Public Health "
    "(-57%, 37) face both small enrollment and steep declines \u2014 candidates for structured program review.",
    "Small-base caution: 12 programs had fewer than 20 students in 2022. Their percentage changes can be "
    "misleading (e.g., Music: 12\u219227 = +125% from just 15 additional students).",
]

# ============================================================================
# PESTLE ANALYSIS  [INTERNAL - PESTLE_Report_FLC.docx & External Forces deck]
# ============================================================================

PESTLE_DATA = {
    "Political": {
        "impact": "High",
        "impact_score": 5,
        "trend": "Negative",
        "factors": [
            "Trump administration (2025\u20132029) reducing federal HE funding; 120 TRIO programs terminated",
            "DEI programs under HIGH scrutiny \u2014 executive order targeting DEI in accreditation (Apr 2025)",
            "Tribal education funding VOLATILE: 109% increase Sept 2025, but FY2026 proposes 24% cuts",
            "Colorado FY 2025\u201326: $38.4M increase (far less than $95M requested); 3.5% tuition cap",
            "Native American Tuition Waiver at risk of misclassification as DEI (waiver is statutory, not DEI)",
            "HLC providing flexibility on diversity standards, but federal pressure on accreditors continues",
        ],
        "opportunities": [
            "Reframe Indigenous programs through statutory obligations (CRS 23-52-105) and cultural preservation (legally safe)",
            "Use 'first-generation support' and 'inclusive excellence' framing (avoids identity-based language)",
            "Advocate for rural institution support in state legislature",
        ],
    },
    "Economic": {
        "impact": "High",
        "impact_score": 5,
        "trend": "Negative",
        "factors": [
            "Colorado shifts costs to students via tuition rather than state appropriations",
            "Rising tuition sensitivity; students increasingly price-conscious and comparison-shopping",
            "Durango housing crisis \u2014 major hidden barrier for student attendance AND faculty recruitment",
            "Native American tuition waiver revenue impact (~37% of students at zero tuition)",
            "Regional economy tourism-dependent (seasonal, variable); limited large employers",
            "Skills-based hiring growing \u2014 degrees less of an automatic hiring requirement",
        ],
        "opportunities": [
            "Healthcare/nursing programs (strong regional employer demand)",
            "Expand dual enrollment pipeline (Pueblo CC, San Juan College feeders)",
            "Develop workforce-aligned certificates and micro-credentials",
            "Position as affordable rural alternative to cost-climbing urban institutions",
        ],
    },
    "Social": {
        "impact": "Medium-High",
        "impact_score": 4,
        "trend": "Mixed",
        "factors": [
            "Declining college-going rates nationally and in Colorado",
            "Career outcome expectations dominant ('What job will I get?')",
            "Indigenous education opportunity IS REAL (166 tribes, 37% waiver, underserved nationally)",
            "First-generation students (43%) need targeted support systems",
            "Growing skepticism about ROI of higher education; trade/vocational paths gaining acceptance",
            "Strong outdoor/recreation culture aligns with FLC place-based brand",
        ],
        "opportunities": [
            "Indigenous education leadership \u2014 reframe through statutory obligations (CRS 23-52-105), not DEI",
            "First-generation student success programs (safe framing, encompasses many Indigenous students)",
            "Place-based brand leveraging Durango outdoor lifestyle as recruitment differentiator",
            "Career outcome emphasis across all programs",
        ],
    },
    "Technological": {
        "impact": "High",
        "impact_score": 4,
        "trend": "Rapidly Changing",
        "factors": [
            "AI disruption transforming pedagogy, assessment, and student expectations",
            "Online graduate market SATURATED \u2014 ASU, SNHU, Western Governors dominate ($50M+ marketing)",
            "FLC has NO online brand nationally; ~25 online courses (~10% of offerings)",
            "Passive video lectures becoming obsolete; AI-enabled adaptive learning replacing them",
            "AI Institute at FLC as emerging institutional strength",
            "Online program development requires 1\u20132+ years governance + substantial investment",
        ],
        "opportunities": [
            "AI Institute partnerships and curriculum integration",
            "AI-enabled advising, early alerts, and retention prediction tools",
            "AI literacy across all disciplines as differentiator",
        ],
    },
    "Legal": {
        "impact": "High",
        "impact_score": 4,
        "trend": "Deteriorating",
        "factors": [
            "Title VI scrutiny \u2014 50+ universities under investigation for race-conscious programs",
            "Native American Tuition Waiver has DISTINCT legal basis (CRS 23-52-105, since 1911)",
            "HLC accreditation: federal pressure on DEI standards, but HLC offers flexibility",
            "Trump administration revising Title IX regulations (definitions, due process in flux)",
            "FERPA compliance critical for AI tools processing student data",
            "Programs framed as 'equity-focused' are primary federal targets",
        ],
        "opportunities": [
            "NATW defensible under Title VI (statutory basis per CRS 23-52-105, not voluntary DEI)",
            "Government-to-government tribal partnerships (sovereignty framing, not race-based)",
            "HLC flexibility allows alternative methods to meet diversity-related standards",
        ],
    },
    "Environmental": {
        "impact": "Medium",
        "impact_score": 3,
        "trend": "Negative",
        "factors": [
            "Southwest Colorado wildfire risk increasing \u2014 smoke impacts air quality and outdoor activities",
            "Colorado River basin under long-term drought stress; water rights contentious",
            "Snowpack variability affects regional economy (ski, rafting, outdoor recreation)",
            "Outdoor recreation brand is FLC strength but CLIMATE-VULNERABLE",
            "Sustainability compliance is baseline, not differentiator",
        ],
        "opportunities": [
            "Proactive sustainability initiatives to build brand beyond compliance",
            "Emergency preparedness planning as operational strength",
            "Environmental science/conservation programs align with regional needs",
        ],
    },
}

# ============================================================================
# PORTER'S FIVE FORCES  [METHODOLOGY: Internet - Applied to FLC context]
# ============================================================================

PORTERS_DATA = {
    "Competitive Rivalry": {
        "rating": "High",
        "score": 4.0,
        "color": "#e74c3c",
        "description": "Intense competition from CU system, CSU, Western Colorado; online threat assumed but unverified for FLC",
        "indicators": [
            {"name": "Number of competing institutions in CO", "value": "30+", "trend": "Increasing"},
            {"name": "FLC market share of CO HS graduates", "value": "~2%", "trend": "Stable"},
            {"name": "Enrollment change vs peers", "value": "-2.5%", "trend": "Declining"},
            {"name": "Tuition discount rate pressure", "value": "High", "trend": "Increasing"},
            {"name": "Online competition (unverified for FLC)", "value": "Assumed", "trend": "Unknown"},
        ],
    },
    "Threat of New Entrants": {
        "rating": "Medium",
        "score": 3.0,
        "color": "#f39c12",
        "description": "Accreditation remains HIGH barrier for degree-granting; certificate/non-degree entrants are real threat",
        "indicators": [
            {"name": "Accreditation barriers (degree)", "value": "High", "trend": "Stable"},
            {"name": "Certificate/non-degree barriers", "value": "Low", "trend": "Decreasing"},
            {"name": "Boot camp / micro-credential programs", "value": "Growing", "trend": "Increasing"},
            {"name": "Community college expansion", "value": "Active", "trend": "Increasing"},
            {"name": "Capital requirements for online", "value": "High", "trend": "Stable"},
        ],
    },
    "Bargaining Power of Students": {
        "rating": "High",
        "score": 4.0,
        "color": "#e74c3c",
        "description": "Students have many choices; price sensitivity high; FLC must compete on value and outcomes",
        "indicators": [
            {"name": "Yield rate (needs verification)", "value": "Unverified", "trend": "Unknown"},
            {"name": "Summer melt rate (FY)", "value": "12.9%", "trend": "Improving"},
            {"name": "Transfer-out competition", "value": "Moderate", "trend": "Stable"},
            {"name": "Price sensitivity", "value": "High", "trend": "Increasing"},
            {"name": "Career outcome expectations", "value": "High", "trend": "Increasing"},
        ],
    },
    "Bargaining Power of Suppliers": {
        "rating": "Medium",
        "score": 3.0,
        "color": "#f39c12",
        "description": "National faculty supply HIGH in most fields; real issue is Durango recruitment (cost of living + salary)",
        "indicators": [
            {"name": "National faculty supply", "value": "High", "trend": "Stable"},
            {"name": "Durango cost of living barrier", "value": "High", "trend": "Increasing"},
            {"name": "Durango recruitment competitiveness", "value": "Below avg", "trend": "Worsening"},
            {"name": "High-demand fields (nursing, CS, engr)", "value": "Tight", "trend": "Increasing"},
            {"name": "Salary competitiveness vs peers", "value": "Below avg", "trend": "Stable"},
        ],
    },
    "Threat of Substitutes": {
        "rating": "Medium",
        "score": 3.0,
        "color": "#f39c12",
        "description": "Online/certificates growing nationally, but FLC's place-based brand serves experience-preferring students",
        "indicators": [
            {"name": "Online degree program growth (national)", "value": "Rapid", "trend": "Increasing"},
            {"name": "Micro-credential adoption", "value": "Growing", "trend": "Increasing"},
            {"name": "Community college pathways", "value": "Strong", "trend": "Increasing"},
            {"name": "FLC place-based differentiation", "value": "Strong", "trend": "Stable"},
            {"name": "Are FLC students choosing online? (unverified)", "value": "Unknown", "trend": "Unknown"},
        ],
    },
}

PORTERS_INSIGHTS = [
    "Overall competitive intensity is HIGH, but FLC's place-based, experiential value proposition serves a distinct market segment.",
    "FLC's strongest defensive positions: statutory Native American mission (CRS 23-52-105, federal-state contract), outdoor recreation lifestyle, and small class sizes.",
    "Online competition assumed based on national trends but UNVERIFIED for FLC \u2014 admitted-but-not-enrolled survey data needed.",
    "Faculty recruitment is a Durango problem (cost of living + salary), not a national supply problem \u2014 except in nursing, CS, and engineering.",
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
    "GROW programs (high market + strong economics): Business Admin, Psychology, Engineering, Health Sciences, CIS, Exercise Physiology, Accounting show strongest investment case.",
    "SUSTAIN programs (solid market, needs efficiency): Environmental programs, Criminology, Biology, Sociology, Teacher Education maintain enrollment but need optimization.",
    "TRANSFORM programs (weak market, strong economics): English and Math generate significant SCH as foundational/service courses \u2014 low Market Score reflects major enrollment, not institutional value.",
    "EVALUATE/SUNSET programs: Political Science, Art & Design require strategic review. Note: NAIS is mission-critical and must not be evaluated on enrollment metrics alone.",
    "Data source disclaimer: FLC does not have a Gray Associates subscription. Scores are proxy estimates based on FLC data, not official Gray Associates output.",
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
        "Porter's Five Forces framework assesses the competitive intensity of the higher education "
        "market in which Fort Lewis College operates. This analysis corrects common AI assumptions: "
        "online competition is unverified for FLC (place-based students may not be choosing online), "
        "faculty 'scarcity' is a Durango recruitment issue (national supply is HIGH in most fields), "
        "and FLC's experiential value proposition serves a distinct market segment."
    ),
    "Gray": (
        "Gray Associates Portfolio Analysis evaluates academic programs by plotting Market Score "
        "(student demand 40% + employment 40% + competition 20%) against Program Economics "
        "(SCH efficiency + cost structure). Programs are classified as Grow, Sustain, Transform, "
        "Evaluate, or Sunset Review. Important: FLC does not have a Gray Associates subscription; "
        "scores are proxy estimates based on FLC institutional data, not official Gray output."
    ),
    "BCG": (
        "The BCG Growth-Share Matrix, adapted from the Boston Consulting Group framework, analyzes "
        "FLC's academic portfolio at two levels. The department-level view maps 22 departments by "
        "SCH share (% of total Student Credit Hours) vs. 2-year enrollment change, identifying which "
        "departments generate institutional revenue. The major-level view maps 48 individual majors by "
        "2024 enrollment headcount vs. 2022\u20132024 percentage change, with bubble size representing "
        "absolute enrollment change. Programs with fewer than 20 students in 2022 are flagged as "
        "\u2018small base\u2019 since their percentage changes can be misleading."
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
                "title": "Statutory Native American Mission",
                "detail": "Federal-state-tribal obligation (CRS 23-52-105, since 1911) to serve Native American students with tuition waiver. Serves 166 tribes; 37% of students receive waiver. This is a values-driven commitment, not a market instrument.",
                "source": "PESTLE (Political/Legal), Institutional Data",
            },
            {
                "title": "Strong Star Programs (Major-Level BCG)",
                "detail": "Business Administration (+10%, 325 enrolled), Exercise Physiology (+17%, 160), Environmental Conservation & Mgmt (+62%, 133), and Biochemistry (+35%, 66) lead with both size and growth.",
                "source": "BCG Matrix (48-major analysis), Gray Associates",
            },
            {
                "title": "Place-Based Brand & Outdoor Differentiation",
                "detail": "Durango's mountain setting and outdoor lifestyle create a recruitment differentiator that online competitors cannot replicate. FLC serves experience-preferring students, a distinct market segment.",
                "source": "Porter's Five Forces, PESTLE (Social/Environmental)",
            },
            {
                "title": "Small Class Sizes & Teaching Focus",
                "detail": "Average class size of 19, 15:1 student-faculty ratio. 98% of tenure-track faculty hold terminal degrees (note: terminal degree % is a common proxy but does not directly measure teaching effectiveness).",
                "source": "Institutional Data",
            },
            {
                "title": "Growing Graduate Program",
                "detail": "Graduate enrollment grew from 10 (2016) to 160 (Fall 2025). However, FLC has only ONE graduate program \u2014 further expansion requires 1\u20132+ years shared governance, accreditation, and significant startup investment.",
                "source": "Institutional Data, Budget Constraints",
            },
            {
                "title": "Strong Employment-Aligned Programs",
                "detail": "Engineering (92), CIS (90), Business Admin (85), and Health Sciences (85) score highest on Gray Associates employment outlook. Healthcare and STEM fields show strongest regional job demand.",
                "source": "Gray Associates, PESTLE (Economic)",
            },
            {
                "title": "NATW Legal Foundation",
                "detail": "The Native American Tuition Waiver has a distinct legal basis (CRS 23-52-105, 1911 federal-state contract) separate from voluntary DEI programs. This is defensible under current Title VI scrutiny.",
                "source": "PESTLE (Legal), Context Files",
            },
        ],
    },
    "Weaknesses": {
        "color": "#e74c3c",
        "icon": "W",
        "items": [
            {
                "title": "Declining Undergraduate Enrollment",
                "detail": "UG degree-seeking enrollment fell from 3,498 (2016) to 3,021 (2025), -13.6% over 10 years. Major-level data shows overall -3.1% decline (2,899\u21922,810) from 2022\u20132024.",
                "source": "Institutional Data, BCG Matrix (48-major analysis)",
            },
            {
                "title": "17 Concern-Quadrant Majors",
                "detail": "BCG analysis shows 17 of 48 majors in the Concern quadrant (small & declining). Economics (-76%, 9 enrolled), MND Art & Design (-98%, 3), and Public Health (-57%, 37) face critical enrollment declines.",
                "source": "BCG Matrix (48-major analysis)",
            },
            {
                "title": "Retention Below National Average",
                "detail": "66.1% FTFT retention vs. ~73% national average. Equity gaps persist: First-Gen 60.9%, Pell 61.7%, Students of Color 62.6%.",
                "source": "Institutional Data, PESTLE (Social)",
            },
            {
                "title": "Durango Housing Crisis & Faculty Recruitment",
                "detail": "Durango cost of living is a hidden barrier for both students (attendance) and faculty (recruitment). National faculty supply is HIGH in most fields \u2014 the real issue is FLC's location + salary competitiveness.",
                "source": "Porter's Five Forces, PESTLE (Economic)",
            },
            {
                "title": "Faculty Footprint Disproportionate to Enrollment",
                "detail": "Number of programs and faculty positions are disproportionately large relative to student enrollment. Faculty governance resistance expected for any consolidation.",
                "source": "Institutional Priorities, SWOT Context",
            },
            {
                "title": "No Online Brand or Infrastructure",
                "detail": "Only ~25 online courses (~10% of offerings). Online market is SATURATED (ASU, SNHU, WGU spend $50M+ on marketing). FLC has no online brand nationally and cannot compete on price with scale players.",
                "source": "PESTLE (Technological), Porter's Five Forces",
            },
            {
                "title": "Tuition Waiver Revenue Impact",
                "detail": "~37% of students at zero tuition via NATW creates dependency on state appropriations. With state funding declining, this structural gap widens.",
                "source": "PESTLE (Economic/Political), Budget Constraints",
            },
        ],
    },
    "Opportunities": {
        "color": "#3498db",
        "icon": "O",
        "items": [
            {
                "title": "Invest in Star Programs",
                "detail": "BCG Star programs (Business Admin, Exercise Physiology, Env Conservation, Biochemistry, Engineering, etc.) show enrollment growth. Gray Associates classifies 7 programs for GROW status with strong employment alignment.",
                "source": "BCG Matrix, Gray Associates",
            },
            {
                "title": "Indigenous Education Leadership (Statutorily Grounded)",
                "detail": "Serving 166 tribes with 37% waiver enrollment is a genuine national distinction. Must be framed through statutory obligations (CRS 23-52-105), cultural preservation, and sovereign agreements \u2014 not DEI language \u2014 to remain viable in current political climate. Reconciles with DEI threat below.",
                "source": "PESTLE (Social/Legal), SWOT Context",
            },
            {
                "title": "Dual Enrollment Pipeline Growth",
                "detail": "Dual enrollment grew 4.5x (52\u2192235) since 2016. 27 converted to degree-seeking in Fall 2025. Partnerships with Pueblo CC, San Juan College, Red Rocks CC are viable low-risk growth channels.",
                "source": "Institutional Data, Enrollment Overview",
            },
            {
                "title": "AI Institute & Technology Integration",
                "detail": "FLC's AI Institute is an emerging strength. AI-enabled advising, retention prediction, and curriculum integration can attract new students. Requires realistic investment assessment.",
                "source": "PESTLE (Technological), Institutional Data",
            },
            {
                "title": "Healthcare & Workforce Programs",
                "detail": "Healthcare sector showing strongest regional job growth. Health Sciences (+69%, 86 enrolled) is a BCG Star. Nursing, allied health, and behavioral health have strong employer demand in SW Colorado.",
                "source": "PESTLE (Economic), Gray Associates, BCG Matrix",
            },
            {
                "title": "First-Generation Student Success",
                "detail": "43% first-gen population is a safe, non-identity-based framing that encompasses many Indigenous students. First-gen support programs are politically viable and address a real retention gap (60.9% vs 66.1%).",
                "source": "PESTLE (Social/Political), Institutional Data",
            },
            {
                "title": "Graduate Certificate Development (Long-Term)",
                "detail": "Graduate enrollment growth (10\u2192160) shows capacity, but expansion requires defensible niche (e.g., tribal governance, outdoor education leadership). Generic MBA/MEd markets are saturated. Timeline: 2\u20133+ years.",
                "source": "Budget Constraints, SWOT Context, PESTLE (Technological)",
            },
        ],
    },
    "Threats": {
        "color": "#e67e22",
        "icon": "T",
        "items": [
            {
                "title": "DEI & Federal Scrutiny of Public Higher Ed",
                "detail": "120 TRIO programs terminated; 50+ universities under Title VI investigation. Programs framed as 'equity-focused' are primary targets. FLC's NATW is legally defensible (statutory basis) but could be misclassified as DEI. Reconcile: Indigenous programs must use statutory/sovereign framing.",
                "source": "PESTLE (Political/Legal)",
            },
            {
                "title": "Tribal Education Funding Volatility",
                "detail": "Federal tribal education funding is VOLATILE: 109% increase Sept 2025, but FY2026 budget proposes 24% cuts. State appropriations also falling short ($38.4M vs $95M requested).",
                "source": "PESTLE (Political/Economic)",
            },
            {
                "title": "Declining College-Going Rates",
                "detail": "National and Colorado college-going rates declining. Growing skepticism about ROI of degrees. FLC first-year pipeline down -7.6% in 2025. Small public liberal arts institutions most affected.",
                "source": "PESTLE (Social), Enrollment Overview",
            },
            {
                "title": "Durango Housing Crisis",
                "detail": "Dramatic cost increases affect student attendance and faculty recruitment. This is a hidden barrier that compounds enrollment decline and makes salary offers less competitive.",
                "source": "PESTLE (Economic), Porter's Five Forces",
            },
            {
                "title": "Alternative Credentials Eroding Degree Value",
                "detail": "Micro-credentials, boot camps, and certificates growing rapidly. Skills-based hiring means degrees are less of an automatic requirement. FLC's liberal arts value proposition harder to sell without career framing.",
                "source": "Porter's Five Forces (Substitutes), PESTLE (Economic)",
            },
            {
                "title": "Climate Vulnerability of Outdoor Brand",
                "detail": "Southwest Colorado wildfire risk increasing, drought stressing Colorado River basin, snowpack variability affecting ski/rafting economy. FLC's outdoor brand is a strength but climate-vulnerable.",
                "source": "PESTLE (Environmental)",
            },
            {
                "title": "Shared Governance Constraints on Speed",
                "detail": "Faculty governance takes 1\u20132+ years for program changes. Conservative Board risk tolerance. Any restructuring or new program requires significant political capital and patience.",
                "source": "SWOT Context, Institutional Priorities",
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
        "description": "Revenue maintenance and growth in proven programs. Focus on BCG Stars and Gray Associates GROW programs that drive enrollment.",
        "programs": [
            {"name": "Business Administration", "action": "Protect capacity, expand pathways (325 enrolled, +10%)", "investment": "High"},
            {"name": "Exercise Physiology", "action": "Develop sports medicine and wellness specializations (160, +17%)", "investment": "Medium"},
            {"name": "Environmental Conservation & Mgmt", "action": "Leverage regional economy and sustainability demand (133, +62%)", "investment": "Medium"},
            {"name": "Engineering", "action": "Strengthen co-op/internship pipeline, industry partnerships (51, +4%)", "investment": "High"},
            {"name": "Health Sciences", "action": "Expand clinical partnerships; strongest regional job demand (86, +69%)", "investment": "High"},
            {"name": "Computer Information Systems", "action": "Align with AI Institute; cybersecurity and data science tracks (77, +13%)", "investment": "Medium"},
        ],
    },
    "Productivity Zone": {
        "color": "#3498db",
        "description": "Enabling investments for retention, efficiency, and operational excellence. Highest-priority revenue stream: retention improvement.",
        "programs": [
            {"name": "Retention Programs", "action": "First-gen support (safe framing), Compass expansion, early alerts; close 66.1% \u2192 70% gap", "investment": "High"},
            {"name": "Advising System Overhaul", "action": "AI-enabled predictive advising; retention risk identification", "investment": "High"},
            {"name": "Faculty Recruitment Package", "action": "Housing assistance + salary competitiveness for Durango (national supply is HIGH; issue is location)", "investment": "Medium"},
            {"name": "Transfer Pathway Optimization", "action": "Streamline articulation with Pueblo CC, San Juan College, Red Rocks CC", "investment": "Low"},
            {"name": "Marketing & Communications", "action": "Place-based brand positioning; statutorily grounded Indigenous recruitment (not DEI language)", "investment": "Medium"},
            {"name": "Program Review Process", "action": "Structured review of 17 Concern-quadrant majors; protect mission-critical programs (NAIS)", "investment": "Low"},
        ],
    },
    "Incubation Zone": {
        "color": "#f39c12",
        "description": "Disciplined experimentation with emerging opportunities. Online programs heavily constrained by saturated market and governance timelines.",
        "programs": [
            {"name": "AI Institute Expansion", "action": "Corporate partnerships, research grants, AI literacy certificates", "investment": "Medium"},
            {"name": "Dual Enrollment Expansion", "action": "New partnerships with regional high schools and community colleges (grew 4.5x since 2016)", "investment": "Low"},
            {"name": "Workforce Certificates", "action": "Stackable micro-credentials in healthcare, IT, sustainability (must complement not cannibalize degrees)", "investment": "Low"},
            {"name": "Healthcare Pipeline", "action": "Nursing and allied health programs aligned with regional employer demand", "investment": "Medium"},
            {"name": "Online Program Pilot", "action": "CAUTION: Market saturated (ASU/SNHU/WGU). Requires 1\u20132yr governance + $50K+ marketing. Start with Indigenous-niche hybrid only.", "investment": "Medium"},
        ],
    },
    "Transformation Zone": {
        "color": "#9b59b6",
        "description": "Strategic bets that change institutional trajectory. Must be framed through statutory obligations (CRS 23-52-105) and sovereign agreements, not DEI. Governance timeline: 2\u20135 years.",
        "programs": [
            {"name": "Indigenous Education Hub", "action": "National center for Indigenous higher education \u2014 framed through statutory obligations and cultural preservation (not DEI)", "investment": "High"},
            {"name": "Experiential Learning Model", "action": "Position FLC as premier outdoor experiential institution; differentiate from online competitors", "investment": "Medium"},
            {"name": "Graduate Certificate Development", "action": "Defensible niche only (tribal governance, outdoor ed leadership). Generic MBA/MEd is a losing strategy. Timeline: 2\u20133+ years.", "investment": "Medium"},
            {"name": "Program Portfolio Restructuring", "action": "Consolidate small declining majors into interdisciplinary programs; requires faculty governance buy-in (1\u20132+ years)", "investment": "Medium"},
        ],
    },
}

# Cross-references linking Zone to Win programs back to Phase 1 & Phase 2 findings
ZONE_CROSS_REFERENCES = {
    # ── Performance Zone ──
    "Business Administration": {
        "supporting": [
            {"text": "BCG Star: 325 enrolled, +10% growth, +30 students (largest absolute gain)", "source": "BCG Matrix (48-major)"},
            {"text": "Gray Associates GROW \u2014 highest market score (74), strong employment (85)", "source": "Gray Associates"},
        ],
        "risks": [
            {"text": "Student price sensitivity HIGH; tuition cap at 3.5%", "source": "PESTLE Economic / Porter's"},
            {"text": "Online competition assumed but unverified for FLC students", "source": "Porter's Five Forces"},
        ],
    },
    "Exercise Physiology": {
        "supporting": [
            {"text": "BCG Star: 160 enrolled, +17% growth, +23 students", "source": "BCG Matrix (48-major)"},
            {"text": "Gray Associates GROW \u2014 employment score 80, strong career alignment", "source": "Gray Associates"},
            {"text": "Place-based brand and outdoor differentiation enhance program appeal", "source": "SWOT Strength"},
        ],
        "risks": [
            {"text": "Overall enrollment declining -3.1%; pressure on all programs", "source": "BCG Matrix / Institutional Data"},
        ],
    },
    "Environmental Conservation & Mgmt": {
        "supporting": [
            {"text": "BCG Star: 133 enrolled, +62% growth, +51 students (second-largest absolute gain)", "source": "BCG Matrix (48-major)"},
            {"text": "Aligns with regional economy and sustainability demand", "source": "PESTLE Economic/Environmental"},
        ],
        "risks": [
            {"text": "Climate vulnerability \u2014 wildfire/drought may erode outdoor brand assets", "source": "PESTLE Environmental"},
        ],
    },
    "Engineering": {
        "supporting": [
            {"text": "Gray Associates GROW \u2014 highest employment score (92/100)", "source": "Gray Associates"},
            {"text": "Healthcare and STEM show strongest regional job demand", "source": "PESTLE Economic"},
        ],
        "risks": [
            {"text": "Durango recruitment challenge for specialized engineering faculty (tight labor market)", "source": "Porter's / PESTLE Economic"},
        ],
    },
    "Health Sciences": {
        "supporting": [
            {"text": "BCG Star: 86 enrolled, +69% growth, +35 students", "source": "BCG Matrix (48-major)"},
            {"text": "Healthcare sector strongest regional employer demand", "source": "PESTLE Economic"},
            {"text": "Gray Associates GROW \u2014 employment score 85, market score 76", "source": "Gray Associates"},
        ],
        "risks": [
            {"text": "Clinical program accreditation adds compliance complexity", "source": "PESTLE Legal"},
        ],
    },
    "Computer Information Systems": {
        "supporting": [
            {"text": "BCG Star: 77 enrolled, +13% growth", "source": "BCG Matrix (48-major)"},
            {"text": "AI Institute alignment creates curriculum synergy", "source": "SWOT Opportunity"},
            {"text": "Gray Associates GROW \u2014 employment score 90", "source": "Gray Associates"},
        ],
        "risks": [
            {"text": "Alternative credentials (boot camps) compete in tech space", "source": "SWOT Threat / Porter's"},
        ],
    },
    # ── Productivity Zone ──
    "Retention Programs": {
        "supporting": [
            {"text": "66.1% retention vs ~73% national avg; First-Gen 60.9%, Pell 61.7%", "source": "SWOT Weakness / Institutional Data"},
            {"text": "Declining college-going rates make retention more critical than recruitment", "source": "PESTLE Social"},
            {"text": "First-gen support is politically safe framing (not identity-based)", "source": "PESTLE Political / SWOT Opportunity"},
        ],
        "risks": [
            {"text": "DEI scrutiny may affect equity-framed retention programs; use first-gen framing", "source": "SWOT Threat / PESTLE Political"},
        ],
    },
    "Advising System Overhaul": {
        "supporting": [
            {"text": "AI-enabled advising and early alerts identified as viable technology investment", "source": "PESTLE Technological"},
            {"text": "Retention gap directly addressable through proactive advising", "source": "SWOT Weakness"},
        ],
        "risks": [
            {"text": "FERPA compliance critical for AI tools processing student data", "source": "PESTLE Legal"},
            {"text": "State funding volatility may constrain technology investment", "source": "SWOT Threat"},
        ],
    },
    "Faculty Recruitment Package": {
        "supporting": [
            {"text": "National faculty supply HIGH in most fields \u2014 issue is Durango, not supply", "source": "Porter's Five Forces (corrected)"},
            {"text": "Small class sizes and teaching focus are core differentiators worth protecting", "source": "SWOT Strength"},
        ],
        "risks": [
            {"text": "Durango housing crisis + below-avg salary = persistent recruitment barrier", "source": "SWOT Weakness / PESTLE Economic"},
            {"text": "Nursing, CS, engineering have genuinely tight labor markets", "source": "Porter's Five Forces"},
        ],
    },
    "Transfer Pathway Optimization": {
        "supporting": [
            {"text": "Dual enrollment grew 4.5x (52\u2192235); 27 converted to degree-seeking FY25", "source": "SWOT Opportunity / Institutional Data"},
            {"text": "Community college pathways are a low-risk growth channel", "source": "Porter's Five Forces"},
        ],
        "risks": [
            {"text": "CU/CSU also competing for transfer students", "source": "Porter's Five Forces"},
        ],
    },
    "Marketing & Communications": {
        "supporting": [
            {"text": "Place-based brand is FLC's strongest competitive moat", "source": "SWOT Strength / Porter's"},
            {"text": "Must use statutory/sovereign framing for Indigenous recruitment, not DEI language", "source": "PESTLE Political/Legal"},
        ],
        "risks": [
            {"text": "Declining college-going rates = shrinking prospect pool", "source": "SWOT Threat / PESTLE Social"},
        ],
    },
    "Program Review Process": {
        "supporting": [
            {"text": "17 of 48 majors in Concern quadrant (small & declining)", "source": "BCG Matrix (48-major)"},
            {"text": "Gray Associates identifies EVALUATE/SUNSET programs for structured review", "source": "Gray Associates"},
        ],
        "risks": [
            {"text": "Faculty governance resistance expected; process takes 1\u20132+ years", "source": "SWOT Threat / Institutional Priorities"},
            {"text": "NAIS is mission-critical \u2014 must NEVER be evaluated on enrollment metrics alone", "source": "SWOT Context"},
        ],
    },
    # ── Incubation Zone ──
    "AI Institute Expansion": {
        "supporting": [
            {"text": "AI Institute is emerging institutional strength in high-demand field", "source": "SWOT Opportunity / PESTLE Technological"},
            {"text": "AI literacy across all disciplines is a viable differentiator", "source": "PESTLE Technological"},
        ],
        "risks": [
            {"text": "Specialized AI faculty face tight labor market + Durango recruitment barrier", "source": "Porter's / PESTLE Economic"},
            {"text": "Tech boot camps and online AI certificates compete in this space", "source": "SWOT Threat"},
        ],
    },
    "Dual Enrollment Expansion": {
        "supporting": [
            {"text": "Proven 4.5x growth track record; low-risk pipeline", "source": "SWOT Opportunity / Institutional Data"},
            {"text": "Partnerships with Pueblo CC, San Juan College viable", "source": "PESTLE Economic"},
        ],
        "risks": [
            {"text": "Community college expansion could compete for same students", "source": "Porter's Five Forces"},
        ],
    },
    "Workforce Certificates": {
        "supporting": [
            {"text": "Healthcare and outdoor recreation sectors have regional employer demand", "source": "PESTLE Economic"},
            {"text": "Micro-credentials can complement (not replace) degree programs", "source": "SWOT Opportunity"},
        ],
        "risks": [
            {"text": "Must complement not cannibalize existing degree programs", "source": "Zone to Win Context"},
            {"text": "Credential market increasingly crowded", "source": "Porter's (Substitutes)"},
        ],
    },
    "Healthcare Pipeline": {
        "supporting": [
            {"text": "Healthcare sector showing strongest regional job growth", "source": "PESTLE Economic"},
            {"text": "Health Sciences is a BCG Star (+69%, 86 enrolled)", "source": "BCG Matrix (48-major)"},
        ],
        "risks": [
            {"text": "Clinical program accreditation and faculty recruitment in nursing are barriers", "source": "PESTLE Legal / Porter's"},
        ],
    },
    "Online Program Pilot": {
        "supporting": [
            {"text": "Indigenous-niche online (tuition waiver moat) is only defensible online strategy", "source": "Zone to Win Context"},
        ],
        "risks": [
            {"text": "SATURATED market: ASU/SNHU/WGU spend $50M+ on marketing", "source": "PESTLE Technological"},
            {"text": "FLC has NO online brand; ~25 courses (~10%); 1\u20132yr governance to launch", "source": "SWOT Weakness / PESTLE Technological"},
            {"text": "Building 'traditional' online programs invests in yesterday's model", "source": "PESTLE Technological Context"},
        ],
    },
    # ── Transformation Zone ──
    "Indigenous Education Hub": {
        "supporting": [
            {"text": "Statutory mission: CRS 23-52-105 (since 1911), 166 tribes, 37% waiver", "source": "SWOT Strength / PESTLE Legal"},
            {"text": "Indigenous education opportunity IS REAL and demographically sound", "source": "PESTLE Social"},
            {"text": "NATW has distinct legal basis (CRS 23-52-105, federal-state contract) separate from DEI", "source": "PESTLE Legal"},
        ],
        "risks": [
            {"text": "Must use statutory/cultural preservation framing, NOT DEI language", "source": "PESTLE Political / SWOT Context"},
            {"text": "Tribal education funding VOLATILE: +109% then proposed -24%", "source": "PESTLE Political"},
            {"text": "NATW could be misclassified as DEI by federal investigators", "source": "SWOT Threat / PESTLE Legal"},
        ],
    },
    "Experiential Learning Model": {
        "supporting": [
            {"text": "Place-based brand is strongest competitive moat vs online competitors", "source": "SWOT Strength / Porter's"},
            {"text": "Adventure Education (+12%, 77 enrolled) is a BCG Star", "source": "BCG Matrix (48-major)"},
        ],
        "risks": [
            {"text": "Climate vulnerability: wildfire smoke, drought, variable snowpack", "source": "PESTLE Environmental"},
        ],
    },
    "Graduate Certificate Development": {
        "supporting": [
            {"text": "Graduate enrollment grew 10\u2192160 (16x), demonstrating scaling capacity", "source": "SWOT Strength / Institutional Data"},
        ],
        "risks": [
            {"text": "FLC has ONE grad program; new programs need 1\u20132+ years governance", "source": "Budget Constraints"},
            {"text": "Generic MBA/MEd markets saturated; only defensible niches viable", "source": "SWOT Context / PESTLE Technological"},
            {"text": "Online grad market dominated by ASU/SNHU/WGU with massive scale", "source": "PESTLE Technological"},
        ],
    },
    "Program Portfolio Restructuring": {
        "supporting": [
            {"text": "17 Concern-quadrant majors; Economics (-76%), MND Art (-98%), Public Health (-57%)", "source": "BCG Matrix (48-major)"},
            {"text": "TRANSFORM: English and Math have strong economics (service courses) despite low Market Score", "source": "Gray Associates"},
        ],
        "risks": [
            {"text": "Faculty governance takes 1\u20132+ years; resistance is default assumption", "source": "SWOT Threat / SWOT Context"},
            {"text": "NAIS is mission-critical and MUST NOT be recommended for reduction", "source": "SWOT Context"},
        ],
    },
}

# Three Strategic Scenarios
SCENARIOS = {
    "Incremental": {
        "description": "Low-risk strategy that builds on current strengths. Focuses on stabilizing enrollment through retention gains, protecting high-performing programs, and improving operational efficiency. No major new ventures.",
        "color": "#2ecc71",
        "enrollment_target": 3450,
        "retention_target": 68.0,
        "graduate_target": 170,
        "online_courses": 30,
        "new_programs": 0,
        "assumptions": [
            "State funding flat (0-1% change); FY2026 appropriations likely down",
            "No new degree programs launched — focus on strengthening existing portfolio",
            "Retention interventions (Compass, advising redesign) yield 2 pp gain",
            "Dual enrollment grows modestly to 260-270 students",
            "Faculty governance limits prevent structural changes in Year 1",
        ],
        "zone_allocation": {"Performance": 50, "Productivity": 30, "Incubation": 10, "Transformation": 10},
        "strategic_bet": "Retention and operational efficiency — stabilize before expanding",
        "risk_level": "Low",
        "success_probability": "High (70-80%)",
        "investment_needs": "Minimal new investment; reallocation of existing resources",
        "zone_recommendations": {
            "Performance": "Invest in BCG Stars (Business Admin, Exercise Physiology, Psychology). Protect mission-critical programs (NAIS). Begin structured review of 17 Concern-quadrant majors.",
            "Productivity": "Redesign advising model per NACADA review. Optimize course scheduling. Address Durango housing recruitment barrier with signing incentives.",
            "Incubation": "AI Institute continues at current scale. Dual enrollment expands with existing partners. No new online programs.",
            "Transformation": "Feasibility study only — Indigenous Education Hub concept paper. No large capital commitments.",
        },
    },
    "Moderate-Adaptive": {
        "description": "Balanced strategy with selective investment in differentiated strengths. Invests in Indigenous education (statutorily grounded), experiential learning brand, and targeted workforce alignment while protecting core programs.",
        "color": "#f39c12",
        "enrollment_target": 3550,
        "retention_target": 70.0,
        "graduate_target": 180,
        "online_courses": 35,
        "new_programs": 1,
        "assumptions": [
            "State funding flat to slight decrease; supplemented by targeted grant revenue",
            "Indigenous education initiatives framed through statutory obligations attract federal/foundation support",
            "One new graduate certificate launched in existing program area (leverages current faculty)",
            "Retention improvements of 3-4 pp through scaled Compass program and early-alert systems",
            "Dual enrollment pipeline reaches 290-310 students through 3+ school partnerships",
        ],
        "zone_allocation": {"Performance": 45, "Productivity": 25, "Incubation": 15, "Transformation": 15},
        "strategic_bet": "Indigenous education differentiation + experiential learning brand",
        "risk_level": "Medium",
        "success_probability": "Moderate (50-65%)",
        "investment_needs": "Moderate — marketing investment for Indigenous programs, one new certificate development, advising technology",
        "zone_recommendations": {
            "Performance": "Invest in Stars, sunset-review lowest-enrollment Concern programs, grow Healthcare and Engineering pipelines. Begin faculty realignment discussions.",
            "Productivity": "Scale Compass retention program. Implement early-alert system. Launch Durango faculty housing partnership. Streamline program review process.",
            "Incubation": "AI Institute pursues NSF/foundation grants. Dual enrollment expands to 3+ high schools. Pilot one workforce certificate aligned with regional demand. Indigenous online course pilot (small scale, NATW niche only).",
            "Transformation": "Launch Indigenous Education Hub initiative (statutory/sovereign framing, not DEI). Develop experiential learning as core brand differentiator. Begin graduate certificate feasibility in defensible niche.",
        },
    },
    "Disruptive": {
        "description": "Bold repositioning strategy that restructures FLC's academic portfolio and identity. Pursues Indigenous online niche nationally, workforce credentials, and significant program consolidation. Highest potential reward but requires substantial political capital and investment.",
        "color": "#e74c3c",
        "enrollment_target": 3700,
        "retention_target": 72.0,
        "graduate_target": 200,
        "online_courses": 45,
        "new_programs": 3,
        "assumptions": [
            "State funding declines but offset by new revenue streams (grants, workforce partnerships, tuition from national Indigenous student recruitment)",
            "Significant marketing investment ($200K+) for national Indigenous online student recruitment",
            "Faculty governance navigated through president's political capital and shared governance engagement",
            "3-5 low-enrollment programs consolidated or restructured (requires 12-18 months governance process)",
            "Workforce credential partnerships secured with regional employers (healthcare, energy, outdoor industry)",
        ],
        "zone_allocation": {"Performance": 35, "Productivity": 20, "Incubation": 20, "Transformation": 25},
        "strategic_bet": "Institutional repositioning — Indigenous education leader + workforce alignment + portfolio restructuring",
        "risk_level": "High",
        "success_probability": "Lower (30-45%) — depends on governance buy-in, marketing effectiveness, and external funding",
        "investment_needs": "Significant — marketing ($200K+), new program development, faculty buyouts/realignment, technology infrastructure",
        "zone_recommendations": {
            "Performance": "Aggressively invest in Stars. Consolidate or restructure 5+ Concern programs. Reallocate faculty lines from declining to growing programs. Protect NAIS regardless of enrollment.",
            "Productivity": "Full advising redesign. Implement predictive analytics for retention. Faculty recruitment overhaul with Durango housing solutions. Administrative consolidation across small departments.",
            "Incubation": "AI Institute scales with external funding. Launch 2-3 workforce certificates (healthcare, outdoor industry, AI/tech). Indigenous online programs expand beyond pilot. Dual enrollment targets 350+ students.",
            "Transformation": "Indigenous Education Hub becomes national brand (statutory NATW mission as moat). Launch sub-baccalaureate workforce credentials. Graduate certificates in 2 defensible niches. Complete program portfolio restructuring aligned with labor market demand.",
        },
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
        "Program Sunset Review Initiated (17 Concern-quadrant majors)",
        "Retention Intervention Pilot Launched (Compass expansion)",
        "Dual Enrollment Expansion Agreements (3+ high schools)",
        "AI Institute Partnership MOU Signed",
        "Indigenous Education Hub Feasibility Study Complete",
        "Faculty Recruitment Incentive Package Approved (Durango housing)",
        "Advising Redesign Implementation (NACADA recommendations)",
        "Q1 KPI Review & Course Correction",
        "Workforce Certificate Feasibility (regional demand analysis)",
        "Graduate Certificate Proposal (existing program area)",
        "Mid-Year Strategic Progress Report",
        "Program Restructuring Plans Through Faculty Governance",
        "Early-Alert Retention System Deployed",
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
        "2026-02-15", "2026-03-01", "2026-03-01", "2026-03-01", "2026-04-01",
        "2026-04-01", "2026-05-01", "2026-05-01", "2026-06-01", "2026-06-01",
        "2026-08-01", "2026-10-01", "2026-12-01", "2027-01-15", "2027-06-01",
    ],
    "Target_Date": [
        "2025-12-15", "2026-01-15", "2026-02-01", "2026-06-01", "2026-04-01",
        "2026-05-15", "2026-06-01", "2026-08-01", "2026-06-01", "2026-08-01",
        "2026-04-30", "2026-09-01", "2026-10-01", "2026-07-01", "2027-02-01",
        "2026-10-15", "2026-11-15", "2027-01-15", "2027-03-01", "2027-07-01",
    ],
    "Status": [
        "Complete", "Complete", "In Progress", "In Progress", "In Progress",
        "Not Started", "Not Started", "Not Started", "Not Started", "Not Started",
        "Upcoming", "Not Started", "Not Started", "Upcoming", "Not Started",
        "Not Started", "Upcoming", "Upcoming", "Upcoming", "Upcoming",
    ],
    "Zone": [
        "All", "All", "All", "Performance", "Productivity",
        "Incubation", "Incubation", "Transformation", "Productivity", "Productivity",
        "All", "Incubation", "Performance", "All", "Performance",
        "Productivity", "All", "All", "All", "All",
    ],
    "Owner": [
        "Consulting Team", "Provost", "Provost/President", "Provost", "VP Student Affairs",
        "VP Enrollment", "AI Institute Director", "Provost", "VP Academic Affairs", "VP Student Affairs",
        "Provost", "VP Academic Affairs", "Dean/Provost", "Provost", "Provost/Faculty Senate",
        "VP Student Affairs", "CFO/Provost", "President/Provost", "Provost", "President",
    ],
})

ROADMAP_KPIS = pd.DataFrame({
    "KPI": [
        "Total Enrollment", "FTFT Retention Rate", "Graduate Enrollment",
        "First-Year Class Size", "Dual Enrollment Students", "Transfer Students",
        "First-Gen Retention Gap", "Native American Retention Rate",
        "Programs in Grow/Sustain Status", "Concern Programs Under Review",
        "Degrees Awarded", "Program Completion Rate",
    ],
    "Category": [
        "Enrollment", "Retention", "Enrollment", "Enrollment",
        "Growth Pipeline", "Growth Pipeline",
        "Equity", "Equity",
        "Portfolio Health", "Portfolio Health",
        "Outcomes", "Outcomes",
    ],
    "Baseline_Value": [3457, 66.1, 160, 777, 235, 190, 5.2, 61.0, 14, 0, 489, 42],
    "Year1_Target": [3450, 68.0, 165, 780, 270, 200, 4.5, 63.0, 15, 8, 490, 43],
    "Year2_Target": [3550, 70.0, 180, 800, 310, 215, 3.5, 65.0, 17, 15, 510, 45],
    "Stretch_Target": [3700, 72.0, 200, 830, 350, 235, 2.5, 68.0, 19, 19, 530, 48],
    "Unit": ["students", "%", "students", "students",
             "students", "students", "pp", "%",
             "programs", "programs", "degrees", "%"],
    "Measurement": [
        "Fall census", "Fall-to-Fall FTFT", "Fall census (1 existing program)", "Fall census",
        "Fall census", "Fall census",
        "Total pop minus First-Gen", "Fall-to-Fall AIAN",
        "Gray Associates assessment", "Programs with sunset/restructure review initiated", "Annual", "6-year rate",
    ],
})

RISK_MITIGATION = pd.DataFrame({
    "Risk": [
        "Native American tuition waiver misclassified as DEI and defunded",
        "Federal DEI policy disrupts TRIO/diversity programs",
        "State funding cut exceeds 3%",
        "Durango housing crisis worsens faculty/staff recruitment",
        "Enrollment falls below 3,200",
        "Retention drops below 62%",
        "Tribal education funding volatility",
        "Faculty governance blocks program restructuring",
        "AI Institute external funding not secured",
        "Online investment exceeds return (saturated market)",
        "Community college competition intensifies",
        "Climate events (wildfire/drought) disrupt operations or brand",
    ],
    "Probability": ["Medium", "High", "Medium", "High", "Low", "Low", "Medium", "High", "Medium", "High", "Medium", "Medium"],
    "Impact": ["Critical", "High", "High", "High", "High", "High", "High", "High", "Medium", "Medium", "Medium", "Medium"],
    "Mitigation_Strategy": [
        "Proactively document NATW legal basis (CRS 23-52-105, 1911 federal-state contract). Frame through statutory obligations and state law, not DEI. Engage state legislators and tribal partners as advocates.",
        "Reframe Indigenous education and first-gen programs through statutory/sovereign and student success language. Maintain commitment to outcomes while adapting terminology. Document Title VI compliance.",
        "Diversify revenue through retention gains, dual enrollment growth, and workforce partnerships. Build contingency reserve. Prioritize revenue-generating Stars programs.",
        "Develop faculty housing partnership with City of Durango. Implement signing incentives and salary adjustments for hard-to-fill positions. Expand remote/hybrid work where possible.",
        "Activate emergency recruitment campaign. Accelerate dual enrollment and transfer pipelines. Increase Durango-area marketing. Consider strategic tuition adjustments.",
        "Scale Compass program college-wide. Deploy early-alert system. Increase advisor capacity for at-risk populations. Address first-gen and Native American retention gaps specifically.",
        "Diversify Indigenous education funding sources (federal, foundation, state, tribal). Build relationships with multiple tribal nations. Document statutory obligation to maintain state support.",
        "Engage faculty senate early in restructuring discussions. Use data-driven program review (BCG + Gray Associates) to build evidence base. Accept 12-18 month governance timeline; do not bypass shared governance.",
        "Pursue alternative funding (NSF, private sector, foundation grants). Scale down to pilot-size if grants not secured. Maintain AI integration in curriculum regardless of institute scale.",
        "Start with Indigenous online niche only (defensible NATW moat). Avoid generic online degrees. Pilot small before investing in marketing. Maintain parallel in-person delivery.",
        "Differentiate on residential experience, outdoor lifestyle, and place-based education. Strengthen transfer articulation agreements. Emphasize 4-year degree completion value proposition.",
        "Integrate climate resilience into campus planning. Develop emergency response protocols. Position sustainability programs as responsive to environmental reality, not just academic interest.",
    ],
    "Owner": [
        "President/General Counsel", "President/General Counsel", "CFO", "VP Academic Affairs/CFO",
        "VP Enrollment", "VP Student Affairs", "Provost/President", "Provost/Faculty Senate",
        "AI Institute Director", "VP Academic Affairs", "VP Enrollment", "VP Operations",
    ],
})
