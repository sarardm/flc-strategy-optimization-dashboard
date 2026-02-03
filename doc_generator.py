"""
Document Generator for FLC Portfolio Optimization Dashboard
============================================================
Generates executive summary .docx and slide deck .pptx files
for each Phase 1 framework analysis.
"""

import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pptx import Presentation
from pptx.util import Inches as PptxInches, Pt as PptxPt
from pptx.dml.color import RGBColor as PptxRGB
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Emu

from data import (
    INSTITUTION, PESTLE_DATA, PORTERS_DATA, PORTERS_INSIGHTS,
    BCG_DATA, BCG_INSIGHTS, BCG_QUADRANT_COLORS,
    GRAY_ASSOCIATES_DATA, GA_INSIGHTS, FRAMEWORK_DESCRIPTIONS,
    SWOT_DATA,
)

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "generated_docs")
os.makedirs(OUTPUT_DIR, exist_ok=True)

FLC_NAVY = RGBColor(0x00, 0x30, 0x57)
FLC_GOLD = RGBColor(0xC8, 0xA4, 0x15)


# ============================================================================
# SHARED HELPERS
# ============================================================================

def _add_heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = FLC_NAVY
    return h


def _add_para(doc, text, bold=False, size=11):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.bold = bold
    return p


def _doc_header(doc, title, subtitle):
    doc.add_paragraph()
    h = doc.add_heading(title, level=0)
    for run in h.runs:
        run.font.color.rgb = FLC_NAVY
    p = doc.add_paragraph()
    run = p.add_run(subtitle)
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    p = doc.add_paragraph()
    run = p.add_run(f"Fort Lewis College | {INSTITUTION['location']} | {INSTITUTION['type']}")
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
    doc.add_paragraph("_" * 60)
    return doc


def _pptx_title_slide(prs, title, subtitle):
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = title
    slide.placeholders[1].text = subtitle
    return slide


def _pptx_content_slide(prs, title, bullets):
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = title
    body = slide.placeholders[1]
    tf = body.text_frame
    tf.clear()
    for i, bullet in enumerate(bullets):
        if i == 0:
            tf.text = bullet
        else:
            p = tf.add_paragraph()
            p.text = bullet
            p.level = 0
    return slide


# ============================================================================
# PESTLE DOCUMENTS
# ============================================================================

def generate_pestle_docx():
    doc = Document()
    _doc_header(doc, "PESTLE Analysis: Executive Summary",
                "External Environmental Scan for Fort Lewis College Academic Affairs")

    _add_heading(doc, "1. Introduction & Methodology", 2)
    _add_para(doc, FRAMEWORK_DESCRIPTIONS["PESTLE"])
    _add_para(doc, (
        "This analysis draws on FLC's internal PESTLE report, the External Forces Shaping "
        "Fort Lewis College presentation, institutional enrollment data, and the 2025-2030 "
        "Strategic Plan. Each of the six PESTLE dimensions was assessed for impact severity "
        "(1-5 scale) and directional trend (Positive, Mixed, Negative, Stable, Opportunity)."
    ))

    _add_heading(doc, "2. Key Findings by Category", 2)
    for category, d in PESTLE_DATA.items():
        _add_heading(doc, f"{category} (Impact: {d['impact']}, Trend: {d['trend']})", 3)
        for factor in d["factors"]:
            doc.add_paragraph(factor, style="List Bullet")
        _add_para(doc, "Opportunities:", bold=True, size=10)
        for opp in d["opportunities"]:
            doc.add_paragraph(opp, style="List Bullet 2")

    _add_heading(doc, "3. Impact Assessment Summary", 2)
    _add_para(doc, (
        "The highest-impact categories are Economic (5/5) and Political/Social (4/5 each). "
        "Economic factors\u2014particularly state funding volatility, tuition sensitivity, and the "
        "revenue impact of the Native American tuition waiver\u2014represent the most significant "
        "external pressures on FLC's financial model. Political factors including performance-based "
        "funding and pressure on DEI programs add further uncertainty. Social factors, notably "
        "declining college-going rates and shifting student expectations toward career outcomes, "
        "require programmatic adaptation."
    ))
    _add_para(doc, (
        "Technological, Legal, and Environmental factors are rated Medium impact (3/5) but "
        "present notable opportunities. The AI Institute, online program expansion, and sustainability "
        "leadership represent areas where FLC can differentiate and grow."
    ))

    _add_heading(doc, "4. Strategic Implications for Academic Affairs", 2)
    _add_para(doc, (
        "The PESTLE analysis reveals that FLC operates in an environment where external economic "
        "and political pressures demand proactive financial diversification. The institution cannot "
        "rely solely on traditional state funding and in-person tuition revenue. Key strategic "
        "imperatives include: (1) diversifying revenue through graduate programs and online offerings, "
        "(2) strengthening the dual enrollment pipeline as a hedge against declining first-year "
        "enrollment, (3) leveraging the Native American mission as a unique competitive advantage "
        "rather than viewing it solely as a cost center, and (4) investing in technology and AI "
        "capabilities as both a program differentiator and operational efficiency driver."
    ))

    _add_heading(doc, "5. Recommendations", 2)
    recommendations = [
        "Prioritize revenue diversification through graduate program expansion and online course development",
        "Strengthen advocacy for state funding while reducing dependency through alternative revenue streams",
        "Position Indigenous education mission as a national leadership opportunity attracting federal grants",
        "Invest in AI Institute and sustainability programs as emerging institutional differentiators",
        "Develop workforce-aligned certificates and micro-credentials to address student career expectations",
        "Build data-driven retention programs targeting equity gaps in First-Gen and Pell-eligible populations",
    ]
    for rec in recommendations:
        doc.add_paragraph(rec, style="List Bullet")

    _add_para(doc, "\nSource: FLC PESTLE Report, External Forces Shaping FLC presentation, "
              "FLC Institutional Data, 2025-2030 Strategic Plan", size=9)

    path = os.path.join(OUTPUT_DIR, "PESTLE_Executive_Summary.docx")
    doc.save(path)
    return path


def generate_pestle_pptx():
    prs = Presentation()
    _pptx_title_slide(prs, "PESTLE Analysis", "External Environmental Scan\nFort Lewis College Academic Affairs")

    _pptx_content_slide(prs, "Methodology", [
        FRAMEWORK_DESCRIPTIONS["PESTLE"],
        "Sources: PESTLE_Report_FLC.docx, External Forces Shaping FLC.pptx, Institutional Data",
        "Six dimensions assessed on 1-5 impact scale with directional trends",
    ])

    for category, d in PESTLE_DATA.items():
        bullets = [f"Impact: {d['impact']} ({d['impact_score']}/5) | Trend: {d['trend']}"]
        bullets.extend(d["factors"][:4])
        bullets.append(f"Opportunity: {d['opportunities'][0]}")
        _pptx_content_slide(prs, f"{category} Factors", bullets)

    _pptx_content_slide(prs, "Impact Assessment Summary", [
        "Highest impact: Economic (5/5), Political & Social (4/5 each)",
        "Medium impact: Technological, Legal, Environmental (3/5 each)",
        "Key risk: State funding volatility and tuition waiver revenue impact",
        "Key opportunity: AI Institute, sustainability leadership, Indigenous education hub",
    ])

    _pptx_content_slide(prs, "Strategic Recommendations", [
        "Diversify revenue through graduate programs and online offerings",
        "Strengthen advocacy while reducing state funding dependency",
        "Position Indigenous mission as national leadership opportunity",
        "Invest in AI Institute and sustainability as differentiators",
        "Build data-driven retention programs for equity gap closure",
    ])

    path = os.path.join(OUTPUT_DIR, "PESTLE_Slide_Deck.pptx")
    prs.save(path)
    return path


# ============================================================================
# PORTER'S FIVE FORCES DOCUMENTS
# ============================================================================

def generate_porters_docx():
    doc = Document()
    _doc_header(doc, "Porter's Five Forces: Executive Summary",
                "Competitive Analysis of FLC's Higher Education Market")

    _add_heading(doc, "1. Introduction & Methodology", 2)
    _add_para(doc, FRAMEWORK_DESCRIPTIONS["Porters"])
    _add_para(doc, (
        "This analysis applies Porter's framework using published higher education competitive "
        "research methodologies combined with FLC institutional data, enrollment trends, and "
        "regional market intelligence. Each force is rated on a 1-5 scale (Low to High competitive "
        "intensity) with supporting indicators and trend analysis."
    ))

    _add_heading(doc, "2. Analysis of Five Forces", 2)
    for force, d in PORTERS_DATA.items():
        _add_heading(doc, f"{force}: {d['rating']} ({d['score']}/5)", 3)
        _add_para(doc, d["description"])
        _add_para(doc, "Key Indicators:", bold=True, size=10)
        for ind in d["indicators"]:
            doc.add_paragraph(
                f"{ind['name']}: {ind['value']} (Trend: {ind['trend']})",
                style="List Bullet"
            )

    _add_heading(doc, "3. Overall Competitive Assessment", 2)
    _add_para(doc, (
        "The aggregate competitive intensity facing Fort Lewis College is HIGH, with an average "
        "force score of 3.8/5 across all five dimensions. Competitive Rivalry (4.5/5) and Bargaining "
        "Power of Students (4.0/5) represent the most intense pressures. The higher education market "
        "in Colorado features 30+ competing institutions, robust online program offerings from larger "
        "universities, and increasing student price sensitivity."
    ))
    _add_para(doc, (
        "FLC's strongest defensive positions are its unique Native American mission and tuition "
        "waiver program (which creates a competitive moat for serving Indigenous students), the "
        "Durango outdoor recreation lifestyle (which cannot be replicated by online competitors), "
        "and the intimate small-college experience with 19-student average class sizes."
    ))

    _add_heading(doc, "4. Strategic Implications", 2)
    for insight in PORTERS_INSIGHTS:
        doc.add_paragraph(insight, style="List Bullet")

    _add_heading(doc, "5. Competitive Positioning Recommendations", 2)
    recommendations = [
        "Double down on experiential/outdoor differentiators that online competitors cannot replicate",
        "Develop strategic pricing and financial aid models to address student price sensitivity",
        "Build transfer-friendly pathways with community colleges rather than competing against them",
        "Invest in unique program offerings (Indigenous education, sustainability, AI) with limited regional competition",
        "Address faculty recruitment challenges through creative compensation and quality-of-life incentives",
        "Expand online offerings strategically to meet student flexibility expectations while preserving residential core",
    ]
    for rec in recommendations:
        doc.add_paragraph(rec, style="List Bullet")

    _add_para(doc, "\nSource: Porter's Five Forces framework (published research); "
              "Applied to FLC institutional data and regional market intelligence", size=9)

    path = os.path.join(OUTPUT_DIR, "Porters_Executive_Summary.docx")
    doc.save(path)
    return path


def generate_porters_pptx():
    prs = Presentation()
    _pptx_title_slide(prs, "Porter's Five Forces Analysis",
                       "Competitive Market Assessment\nFort Lewis College Academic Affairs")

    _pptx_content_slide(prs, "Methodology", [
        FRAMEWORK_DESCRIPTIONS["Porters"],
        "Source: Published higher education research + FLC institutional data",
        "Each force rated 1-5 (Low-High intensity) with indicators and trends",
    ])

    for force, d in PORTERS_DATA.items():
        bullets = [f"Rating: {d['rating']} ({d['score']}/5)"]
        bullets.append(d["description"])
        for ind in d["indicators"][:3]:
            bullets.append(f"{ind['name']}: {ind['value']} ({ind['trend']})")
        _pptx_content_slide(prs, force, bullets)

    _pptx_content_slide(prs, "Overall Assessment", [
        "Average competitive intensity: 3.8/5 (HIGH)",
        "Strongest pressures: Competitive Rivalry (4.5), Student Power (4.0)",
        "FLC defensive moats: Native American mission, outdoor lifestyle, small-college experience",
        "Greatest vulnerabilities: online competition, price sensitivity, faculty recruitment",
    ])

    _pptx_content_slide(prs, "Strategic Recommendations", [
        "Leverage experiential/outdoor differentiators against online competitors",
        "Develop strategic pricing models for student price sensitivity",
        "Build transfer pathways with community colleges as partners, not rivals",
        "Invest in unique programs with limited regional competition",
        "Address faculty recruitment through creative compensation",
    ])

    path = os.path.join(OUTPUT_DIR, "Porters_Slide_Deck.pptx")
    prs.save(path)
    return path


# ============================================================================
# GRAY ASSOCIATES DOCUMENTS
# ============================================================================

def generate_gray_docx():
    doc = Document()
    _doc_header(doc, "Gray Associates Portfolio Analysis: Executive Summary",
                "Academic Program Evaluation for Fort Lewis College")

    _add_heading(doc, "1. Introduction & Methodology", 2)
    _add_para(doc, FRAMEWORK_DESCRIPTIONS["Gray"])
    _add_para(doc, (
        "This analysis applies the Gray Associates Program Evaluation System (PES) methodology "
        "to FLC's academic programs using institutional enrollment data, BCG market share analysis, "
        "and regional employment projections. Market Score is a weighted composite of Student "
        "Demand (40%), Employment Outlook (40%), and Competitive Position (20%). Economics Score "
        "is derived from SCH generation efficiency and program cost structures."
    ))
    _add_para(doc, (
        "Note: This analysis uses the Gray Associates framework methodology sourced from published "
        "research and applied to FLC's internal institutional data. FLC does not have a Gray Associates "
        "subscription; scores represent best estimates based on available data."
    ), size=10)

    _add_heading(doc, "2. Program Portfolio Classification", 2)
    categories = {"Grow": [], "Sustain": [], "Transform": [], "Evaluate": [], "Sunset Review": []}
    for _, row in GRAY_ASSOCIATES_DATA.iterrows():
        categories[row["GA_Recommendation"]].append(row)

    for cat, programs in categories.items():
        if programs:
            _add_heading(doc, f"{cat} ({len(programs)} programs)", 3)
            for p in programs:
                doc.add_paragraph(
                    f"{p['Program']}: Market Score {p['Market_Score']}, Economics {p['Economics_Score']}, "
                    f"Enrollment {p['Enrollment']}, Mission: {p['Mission_Alignment']}",
                    style="List Bullet"
                )

    _add_heading(doc, "3. Key Findings", 2)
    for insight in GA_INSIGHTS:
        doc.add_paragraph(insight, style="List Bullet")

    _add_heading(doc, "4. Market Demand Analysis", 2)
    _add_para(doc, (
        "Programs with the highest Market Scores reflect strong alignment between student interest "
        "and labor market demand. Engineering (79), Health Sciences (76), Computer Information "
        "Systems (75), and Business Administration (74) lead on market positioning. These programs "
        "benefit from robust regional and national employment demand in STEM, healthcare, and "
        "business fields."
    ))
    _add_para(doc, (
        "Programs with the lowest Market Scores\u2014Philosophy (34), Political Science (38), "
        "Anthropology (38), Art & Design (39), and History (39)\u2014face challenging demand dynamics. "
        "However, some of these programs play important roles in general education requirements and "
        "institutional mission fulfillment that pure market metrics do not capture."
    ))

    _add_heading(doc, "5. Recommendations", 2)
    recommendations = [
        "Invest aggressively in Grow programs: add capacity, online options, and graduate tracks",
        "Improve efficiency of Sustain programs through shared resources and cross-listing",
        "Innovate delivery for Transform programs: English and Math could adopt co-requisite and applied models",
        "Conduct deep-dive reviews of Evaluate and Sunset Review programs with faculty input",
        "Consider interdisciplinary consolidation for low-enrollment programs to preserve breadth while reducing cost",
        "Weight mission alignment alongside market/economics in final program decisions",
    ]
    for rec in recommendations:
        doc.add_paragraph(rec, style="List Bullet")

    _add_para(doc, "\nSource: Gray Associates PES methodology (internet research); "
              "Applied to FLC enrollment data, BCG analysis, and regional employment data", size=9)

    path = os.path.join(OUTPUT_DIR, "Gray_Executive_Summary.docx")
    doc.save(path)
    return path


def generate_gray_pptx():
    prs = Presentation()
    _pptx_title_slide(prs, "Gray Associates Portfolio Analysis",
                       "Academic Program Evaluation\nFort Lewis College")

    _pptx_content_slide(prs, "Methodology", [
        FRAMEWORK_DESCRIPTIONS["Gray"],
        "Market Score = Student Demand (40%) + Employment (40%) + Competition (20%)",
        "Economics Score = SCH generation efficiency + program cost structure",
        "Data Source: Gray Associates PES methodology applied to FLC institutional data",
    ])

    for cat in ["Grow", "Sustain", "Transform", "Evaluate", "Sunset Review"]:
        progs = GRAY_ASSOCIATES_DATA[GRAY_ASSOCIATES_DATA["GA_Recommendation"] == cat]
        if not progs.empty:
            bullets = [f"{len(progs)} programs classified as {cat}"]
            for _, p in progs.iterrows():
                bullets.append(f"{p['Program']} (Market: {p['Market_Score']}, Econ: {p['Economics_Score']})")
            _pptx_content_slide(prs, f"{cat} Programs", bullets[:6])

    _pptx_content_slide(prs, "Key Findings", GA_INSIGHTS)

    _pptx_content_slide(prs, "Recommendations", [
        "Invest in Grow programs: capacity, online, graduate tracks",
        "Improve efficiency of Sustain programs via shared resources",
        "Innovate delivery for Transform programs (English, Math)",
        "Deep-dive reviews for Evaluate/Sunset programs with faculty",
        "Weight mission alignment in final decisions",
    ])

    path = os.path.join(OUTPUT_DIR, "Gray_Slide_Deck.pptx")
    prs.save(path)
    return path


# ============================================================================
# BCG MATRIX DOCUMENTS
# ============================================================================

def generate_bcg_docx():
    doc = Document()
    _doc_header(doc, "BCG Growth-Share Matrix: Executive Summary",
                "Department Portfolio Analysis for Fort Lewis College")

    _add_heading(doc, "1. Introduction & Methodology", 2)
    _add_para(doc, FRAMEWORK_DESCRIPTIONS["BCG"])
    _add_para(doc, (
        "This analysis uses FLC's internal BCG presentation data, which maps 22 academic departments "
        "on two axes: (1) the department's share of total Student Credit Hours (SCH) in 2023-24, "
        "representing relative market share, and (2) the 2-year percentage change in SCH, representing "
        "growth rate. The intersection of these dimensions creates four strategic quadrants."
    ))
    _add_para(doc, (
        "Important terminology note: The traditional BCG framework labels the low-growth/low-share "
        "quadrant as 'Dogs.' Fort Lewis College uses the term 'Concerns' to describe these programs, "
        "reflecting the need for careful review rather than automatic elimination."
    ))

    _add_heading(doc, "2. Quadrant Analysis", 2)
    quadrants = {
        "Stars (High Share, Growing)": "Star",
        "Cash Cows (High Share, Declining)": "Cash Cow",
        "Question Marks (Low Share, Growing)": "Question Mark",
        "Concerns (Low Share, Declining)": "Concern",
    }
    for label, q in quadrants.items():
        depts = BCG_DATA[BCG_DATA["Quadrant"] == q]
        _add_heading(doc, f"{label}: {len(depts)} departments", 3)
        for _, row in depts.iterrows():
            doc.add_paragraph(
                f"{row['Department']}: {row['SCH_Pct']}% of SCH, {row['Two_Year_Change']:+.1f}% 2-year change",
                style="List Bullet"
            )

    _add_heading(doc, "3. Key Findings", 2)
    for insight in BCG_INSIGHTS:
        doc.add_paragraph(insight, style="List Bullet")

    _add_heading(doc, "4. Portfolio Health Assessment", 2)
    _add_para(doc, (
        "FLC's academic portfolio shows a concerning imbalance: only 2 of 22 departments (9%) are "
        "Stars, while 9 departments (41%) fall in the Concern quadrant. The 9 Cash Cow departments "
        "represent the institutional backbone, generating the majority of SCH, but their declining "
        "trajectory requires attention. The 2 Question Marks (Accounting and History) present "
        "intriguing growth signals that merit further investigation."
    ))
    _add_para(doc, (
        "The heavy concentration in the Concern quadrant is particularly notable. However, several "
        "Concern programs (Anthropology, Philosophy, Environment & Sustainability) contribute to "
        "FLC's liberal arts mission and general education requirements. Strategic review should "
        "balance market metrics with mission contribution and interdisciplinary value."
    ))

    _add_heading(doc, "5. Investment Recommendations", 2)
    recommendations = [
        "Stars: Invest to accelerate growth. Add capacity, online options, and graduate-level offerings in Business Admin and Psychology.",
        "Cash Cows: Optimize for efficiency. Protect SCH generation while reducing per-student costs through class size optimization and cross-listing.",
        "Question Marks: Evaluate selectively. Accounting's 12% growth warrants investment; History's 3% growth needs demand validation.",
        "Concerns: Conduct structured program review. Differentiate between mission-critical programs needing restructuring vs. programs for potential phase-out.",
        "Cross-cutting: Explore interdisciplinary combinations that can revitalize Concern programs by linking them to Star or Cash Cow departments.",
    ]
    for rec in recommendations:
        doc.add_paragraph(rec, style="List Bullet")

    _add_para(doc, "\nSource: FLC BCG Presentation, BCG-growthMatrixDepts.png, "
              "FLC Institutional Data (23-24 SCH)", size=9)

    path = os.path.join(OUTPUT_DIR, "BCG_Executive_Summary.docx")
    doc.save(path)
    return path


def generate_bcg_pptx():
    prs = Presentation()
    _pptx_title_slide(prs, "BCG Growth-Share Matrix",
                       "Department Portfolio Analysis\nFort Lewis College")

    _pptx_content_slide(prs, "Methodology", [
        FRAMEWORK_DESCRIPTIONS["BCG"],
        "X-axis: % of Total SCH (23-24) = Market Share",
        "Y-axis: 2-Year Change % = Growth Rate",
        "Note: 'Concerns' used instead of 'Dogs' per FLC preference",
        "Source: BCG Presentation.pptx, BCG-growthMatrixDepts.png",
    ])

    for label, q in [("Stars", "Star"), ("Cash Cows", "Cash Cow"),
                      ("Question Marks", "Question Mark"), ("Concerns", "Concern")]:
        depts = BCG_DATA[BCG_DATA["Quadrant"] == q]
        bullets = [f"{len(depts)} departments in this quadrant"]
        for _, row in depts.iterrows():
            bullets.append(f"{row['Department']}: {row['SCH_Pct']}% SCH, {row['Two_Year_Change']:+.1f}%")
        _pptx_content_slide(prs, label, bullets[:7])

    _pptx_content_slide(prs, "Portfolio Health", [
        "2 Stars (9%), 9 Cash Cows (41%), 2 Question Marks (9%), 9 Concerns (41%)",
        "Concerning imbalance: 41% of departments in Concern quadrant",
        "Cash Cows provide stable SCH but face declining trajectories",
        "Mission-critical Concern programs need restructuring, not elimination",
    ])

    _pptx_content_slide(prs, "Investment Recommendations", [
        "Stars: Invest to accelerate (online, graduate tracks)",
        "Cash Cows: Optimize efficiency, protect SCH generation",
        "Question Marks: Evaluate selectively (Accounting promising)",
        "Concerns: Structured review; differentiate mission-critical vs. phase-out",
        "Explore interdisciplinary combinations for Concern programs",
    ])

    path = os.path.join(OUTPUT_DIR, "BCG_Slide_Deck.pptx")
    prs.save(path)
    return path


# ============================================================================
# SWOT MATRIX SLIDE
# ============================================================================

def generate_swot_pptx():
    """Generate a single landscape SWOT matrix slide with 2x2 grid."""
    prs = Presentation()
    prs.slide_width = PptxInches(13.333)
    prs.slide_height = PptxInches(7.5)

    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout

    # --- Title bar at top ---
    title_box = slide.shapes.add_textbox(
        PptxInches(0.4), PptxInches(0.25), PptxInches(12.5), PptxInches(0.6),
    )
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = "SWOT Analysis"
    run.font.size = PptxPt(28)
    run.font.color.rgb = PptxRGB(0x00, 0x30, 0x57)
    run.font.bold = True
    # Subtitle
    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.LEFT
    run2 = p2.add_run()
    run2.text = (
        f"Fort Lewis College | {INSTITUTION['location']} | "
        "Cross-Framework Strategic Synthesis"
    )
    run2.font.size = PptxPt(11)
    run2.font.color.rgb = PptxRGB(0x66, 0x66, 0x66)

    # --- Quadrant layout ---
    quadrant_config = [
        ("Strengths",     PptxRGB(0x2E, 0xCC, 0x71), 0.4,   1.15),  # top-left
        ("Weaknesses",    PptxRGB(0xE7, 0x4C, 0x3C), 6.9,   1.15),  # top-right
        ("Opportunities", PptxRGB(0x34, 0x98, 0xDB), 0.4,   4.3),   # bottom-left
        ("Threats",       PptxRGB(0xE6, 0x7E, 0x22), 6.9,   4.3),   # bottom-right
    ]
    box_width = PptxInches(6.1)
    box_height = PptxInches(2.95)

    for label, color, left, top in quadrant_config:
        data = SWOT_DATA[label]

        # Colored header bar
        header = slide.shapes.add_textbox(
            PptxInches(left), PptxInches(top), box_width, PptxInches(0.38),
        )
        header.fill.solid()
        header.fill.fore_color.rgb = color
        htf = header.text_frame
        htf.word_wrap = True
        htf.margin_top = PptxPt(3)
        htf.margin_bottom = PptxPt(3)
        htf.margin_left = PptxPt(8)
        hp = htf.paragraphs[0]
        hp.alignment = PP_ALIGN.LEFT
        hr = hp.add_run()
        hr.text = f"{label}  ({len(data['items'])} items)"
        hr.font.size = PptxPt(13)
        hr.font.bold = True
        hr.font.color.rgb = PptxRGB(0xFF, 0xFF, 0xFF)

        # Content body
        body = slide.shapes.add_textbox(
            PptxInches(left), PptxInches(top + 0.38),
            box_width, PptxInches(box_height.inches - 0.38),
        )
        body.fill.solid()
        body.fill.fore_color.rgb = PptxRGB(0xFA, 0xFA, 0xFA)
        body.line.color.rgb = PptxRGB(0xE0, 0xE0, 0xE0)
        body.line.width = PptxPt(0.5)
        btf = body.text_frame
        btf.word_wrap = True
        btf.margin_top = PptxPt(6)
        btf.margin_left = PptxPt(8)
        btf.margin_right = PptxPt(6)
        btf.vertical_anchor = MSO_ANCHOR.TOP

        for idx, item in enumerate(data["items"]):
            # Title line (bold)
            if idx == 0:
                p_title = btf.paragraphs[0]
            else:
                p_title = btf.add_paragraph()
            p_title.space_before = PptxPt(4) if idx > 0 else PptxPt(0)
            p_title.space_after = PptxPt(1)
            r_bullet = p_title.add_run()
            r_bullet.text = f"\u2022 {item['title']}"
            r_bullet.font.size = PptxPt(9)
            r_bullet.font.bold = True
            r_bullet.font.color.rgb = PptxRGB(0x00, 0x30, 0x57)

            # Detail line
            p_detail = btf.add_paragraph()
            p_detail.space_before = PptxPt(0)
            p_detail.space_after = PptxPt(1)
            r_detail = p_detail.add_run()
            r_detail.text = f"   {item['detail']}"
            r_detail.font.size = PptxPt(7.5)
            r_detail.font.color.rgb = PptxRGB(0x33, 0x33, 0x33)

            # Source line (italic)
            p_src = btf.add_paragraph()
            p_src.space_before = PptxPt(0)
            p_src.space_after = PptxPt(2)
            r_src = p_src.add_run()
            r_src.text = f"   Source: {item['source']}"
            r_src.font.size = PptxPt(6.5)
            r_src.font.italic = True
            r_src.font.color.rgb = PptxRGB(0x99, 0x99, 0x99)

    path = os.path.join(OUTPUT_DIR, "SWOT_Matrix.pptx")
    prs.save(path)
    return path


# ============================================================================
# PUBLIC API
# ============================================================================

GENERATORS = {
    "pestle_docx": generate_pestle_docx,
    "pestle_pptx": generate_pestle_pptx,
    "porters_docx": generate_porters_docx,
    "porters_pptx": generate_porters_pptx,
    "gray_docx": generate_gray_docx,
    "gray_pptx": generate_gray_pptx,
    "bcg_docx": generate_bcg_docx,
    "bcg_pptx": generate_bcg_pptx,
    "swot_pptx": generate_swot_pptx,
}


def generate_all():
    """Generate all documents. Returns dict of {name: filepath}."""
    results = {}
    for name, fn in GENERATORS.items():
        results[name] = fn()
        print(f"  Generated: {results[name]}")
    return results


if __name__ == "__main__":
    print("\nGenerating all executive summaries and slide decks...\n")
    generate_all()
    print(f"\nAll documents saved to: {OUTPUT_DIR}\n")
