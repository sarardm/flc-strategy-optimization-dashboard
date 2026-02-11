"""
Fort Lewis College Portfolio Optimization Dashboard
=====================================================
A comprehensive Dash application for tracking FLC's strategic portfolio
analysis across three phases:
  Phase 1: Environmental Scanning (PESTLE, Porter's, Gray Associates, BCG)
  Phase 2: Strategic Synthesis (SWOT)
  Phase 3: Strategic Direction (Zone to Win, Strategic Roadmap)

Run:  python app.py
Then open http://127.0.0.1:8050 in your browser.
"""

import dash_auth

# Password protection
VALID_USERNAME_PASSWORD_PAIRS = {
    'mtzstrategy': 'mtz2026flc'
}

import os
import dash
from dash import dcc, html, dash_table, callback_context
from dash.dependencies import Input, Output
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import pandas as pd
import numpy as np
from datetime import datetime

from data import (
    INSTITUTION, ENROLLMENT_HISTORY, GRADUATE_ENROLLMENT,
    RETENTION_HISTORY, RETENTION_BY_DEMO,
    TOP_MAJORS_ENROLLMENT, DEGREES_AWARDED,
    BCG_DATA, BCG_DEPT_DATA, BCG_DEPT_INSIGHTS, BCG_QUADRANT_COLORS, BCG_INSIGHTS,
    PESTLE_DATA,
    PORTERS_DATA, PORTERS_INSIGHTS,
    GRAY_ASSOCIATES_DATA, GA_RECOMMENDATION_COLORS, GA_INSIGHTS,
    STRATEGIC_INITIATIVES, MILESTONES, KPIS, RESOURCE_ALLOCATION,
    DATA_SOURCES, FRAMEWORK_DESCRIPTIONS,
    SWOT_DATA, ZONE_TO_WIN_DATA, ZONE_CROSS_REFERENCES, SCENARIOS,
    ROADMAP_MILESTONES, ROADMAP_KPIS, RISK_MITIGATION,
)
from doc_generator import (
    generate_swot_pptx,
    generate_exec_summary_docx,
    generate_exec_summary_pptx,
    generate_final_report_docx,
    generate_final_report_pptx,
)

# ============================================================================
# APP SETUP
# ============================================================================

GENERATED_DOCS_DIR = os.path.join(os.path.dirname(__file__), "generated_docs")

app = dash.Dash(
    __name__,
    suppress_callback_exceptions=True,
    title="FLC Portfolio Optimization Dashboard",
)

auth = dash_auth.BasicAuth(
    app,
    VALID_USERNAME_PASSWORD_PAIRS
)

# FLC brand colors (matched to official FLC templates)
FLC_NAVY = "#003057"
FLC_BLUE = "#0066b3"
FLC_BLUE_LIGHT = "#2a8fd4"
FLC_GOLD = "#c8a415"
FLC_LIGHT = "#f5f8fb"
FLC_BLUE_PALE = "#d6e8f7"
FLC_BLUE_WASH = "#eaf2fa"
BG_WHITE = "#ffffff"
FLC_BORDER = "#c8daea"

# Shared style constants
CARD_STYLE = {
    "backgroundColor": BG_WHITE,
    "borderRadius": "10px",
    "padding": "22px",
    "marginBottom": "16px",
    "boxShadow": "0 1px 3px rgba(0,48,87,0.06), 0 1px 2px rgba(0,48,87,0.04)",
    "border": f"1px solid {FLC_BORDER}",
}

SECTION_TITLE = {
    "color": FLC_BLUE,
    "fontSize": "20px",
    "fontWeight": "700",
    "marginBottom": "12px",
    "borderBottom": f"2px solid {FLC_BLUE}",
    "paddingBottom": "8px",
    "letterSpacing": "0.3px",
}

TAB_STYLE = {
    "fontWeight": "600",
    "fontSize": "13px",
    "padding": "12px 18px",
    "border": "none",
    "borderBottom": "3px solid transparent",
    "color": "#4a6070",
}
TAB_SELECTED = {
    "fontWeight": "700",
    "fontSize": "13px",
    "padding": "12px 18px",
    "borderTop": "none",
    "borderBottom": f"3px solid {FLC_BLUE}",
    "color": FLC_NAVY,
}

# FLC-branded Plotly chart template applied to ALL visualizations
FLC_COLORWAY = [FLC_NAVY, FLC_BLUE, FLC_BLUE_LIGHT, "#5ba3d9", "#8cc0e8", "#b8d8f0"]
FLC_CHART_TEMPLATE = go.layout.Template(
    layout=go.Layout(
        font=dict(family="Segoe UI, Tahoma, Geneva, Verdana, sans-serif", color=FLC_NAVY),
        title=dict(font=dict(color=FLC_NAVY, size=16, family="Segoe UI, sans-serif"), x=0, xanchor="left"),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        colorway=FLC_COLORWAY,
        xaxis=dict(gridcolor=FLC_BLUE_PALE, linecolor=FLC_BORDER, zerolinecolor=FLC_BORDER,
                   title_font=dict(color=FLC_NAVY, size=12), tickfont=dict(color="#4a6070", size=11)),
        yaxis=dict(gridcolor=FLC_BLUE_PALE, linecolor=FLC_BORDER, zerolinecolor=FLC_BORDER,
                   title_font=dict(color=FLC_NAVY, size=12), tickfont=dict(color="#4a6070", size=11)),
        legend=dict(font=dict(size=11, color=FLC_NAVY)),
        hoverlabel=dict(bgcolor=FLC_NAVY, font_size=12, font_color="white"),
    )
)

# Semantic colors (muted, professional versions for indicators)
CLR_HIGH = "#c53030"       # deep red
CLR_MEDIUM = "#d69e2e"     # warm amber
CLR_LOW = "#276749"        # forest green
CLR_POSITIVE = "#2b6cb0"   # positive blue
CLR_NEUTRAL = "#718096"    # neutral gray-blue

# Phase accent colors (blue family)
PHASE_COLORS = [FLC_BLUE_LIGHT, FLC_BLUE, FLC_NAVY]

# Status colors (blue-toned)
STATUS_COLORS = {
    "Complete": "#2b6cb0",     # blue
    "In Progress": FLC_BLUE_LIGHT,
    "Not Started": "#a0aec0",  # gray-blue
    "Upcoming": "#4299e1",     # lighter blue
}

# Badge helper
def _badge(text, bg_color):
    return html.Span(text, style={
        "backgroundColor": bg_color, "color": "white", "padding": "2px 9px",
        "borderRadius": "10px", "fontSize": "10px", "fontWeight": "600",
        "marginLeft": "8px", "whiteSpace": "nowrap",
    })

# Shared DataTable style dict (Slide 7 style: light blue banding, thin borders)
TABLE_HEADER_STYLE = {
    "backgroundColor": BG_WHITE, "color": FLC_NAVY, "fontWeight": "700",
    "fontSize": "11px", "borderBottom": f"2px solid {FLC_BLUE}",
    "textTransform": "uppercase", "letterSpacing": "0.5px",
    "padding": "10px 8px",
}
TABLE_CELL_STYLE = {
    "textAlign": "center", "padding": "8px", "fontSize": "12px",
    "color": FLC_NAVY, "border": f"1px solid {FLC_BLUE_PALE}",
    "fontFamily": "Segoe UI, Tahoma, sans-serif",
}
TABLE_ROW_BANDING = [
    {"if": {"row_index": "odd"}, "backgroundColor": FLC_BLUE_WASH},
    {"if": {"row_index": "even"}, "backgroundColor": BG_WHITE},
]

# ============================================================================
# HELPERS
# ============================================================================

def data_source_badge(framework_name):
    """Render a data-source attribution badge."""
    info = DATA_SOURCES.get(framework_name, {})
    source = info.get("source", "Unknown")
    return html.Div([
        html.Span("DATA SOURCE: ", style={"fontWeight": "bold", "fontSize": "11px", "color": FLC_NAVY}),
        html.Span(source, style={
            "backgroundColor": FLC_BLUE, "color": "white", "padding": "2px 8px",
            "borderRadius": "10px", "fontSize": "11px", "fontWeight": "600",
        }),
        html.Span(
            f"  ({', '.join(info.get('files', []))})",
            style={"fontSize": "11px", "color": "#6b8299", "marginLeft": "8px"},
        ),
    ], style={"marginBottom": "12px"})


def framework_description_block(key):
    """Render a 2-3 sentence framework description at the top of a Phase 1 tab."""
    text = FRAMEWORK_DESCRIPTIONS.get(key, "")
    return html.Div([
        html.P(text, style={
            "fontSize": "13px", "color": "#3a5068", "lineHeight": "1.7",
            "backgroundColor": FLC_BLUE_WASH, "padding": "14px 18px",
            "borderLeft": f"4px solid {FLC_BLUE}", "borderRadius": "4px",
            "marginBottom": "16px",
        }),
    ])


def download_buttons(framework_label):
    """Render download buttons for .docx and .pptx for a Phase 1 framework."""
    file_map = {
        "PESTLE": ("PESTLE_Executive_Summary.docx", "PESTLE_Slide_Deck.pptx"),
        "Porters": ("Porters_Executive_Summary.docx", "Porters_Slide_Deck.pptx"),
        "Gray": ("Gray_Executive_Summary.docx", "Gray_Slide_Deck.pptx"),
        "BCG": ("BCG_Executive_Summary.docx", "BCG_Slide_Deck.pptx"),
    }
    docx_file, pptx_file = file_map.get(framework_label, ("", ""))
    docx_path = os.path.join(GENERATED_DOCS_DIR, docx_file)
    pptx_path = os.path.join(GENERATED_DOCS_DIR, pptx_file)
    docx_exists = os.path.exists(docx_path)
    pptx_exists = os.path.exists(pptx_path)

    btn_style = {
        "backgroundColor": FLC_BLUE, "color": "white", "border": "none",
        "padding": "9px 18px", "borderRadius": "6px", "cursor": "pointer",
        "fontSize": "12px", "fontWeight": "600", "marginRight": "8px",
        "transition": "background-color 0.2s ease",
        "letterSpacing": "0.2px",
    }
    btn_disabled_style = {**btn_style, "backgroundColor": FLC_BORDER, "cursor": "not-allowed", "color": "#8a9baa"}

    buttons = []
    if docx_exists:
        buttons.append(html.Button(
            "Download Executive Summary (.docx)",
            id={"type": "dl-btn", "name": f"{framework_label}-docx"},
            style=btn_style,
        ))
        buttons.append(dcc.Download(id={"type": "dl-target", "name": f"{framework_label}-docx"}))
    else:
        buttons.append(html.Button("Executive Summary (.docx) - not generated",
                                   disabled=True, style=btn_disabled_style))

    if pptx_exists:
        buttons.append(html.Button(
            "Download Slide Deck (.pptx)",
            id={"type": "dl-btn", "name": f"{framework_label}-pptx"},
            style=btn_style,
        ))
        buttons.append(dcc.Download(id={"type": "dl-target", "name": f"{framework_label}-pptx"}))
    else:
        buttons.append(html.Button("Slide Deck (.pptx) - not generated",
                                   disabled=True, style=btn_disabled_style))

    return html.Div(buttons, style={"marginBottom": "16px"})


def source_annotation(text):
    """Small italic source citation below a chart or table."""
    return html.Div(text, style={
        "fontSize": "10px", "color": "#8a9bb0", "fontStyle": "italic",
        "textAlign": "right", "marginTop": "-8px", "marginBottom": "8px",
    })


# ============================================================================
# TAB BUILDERS
# ============================================================================

def build_summary_page():
    """Executive summary page with 3-phase overview."""
    # KPI cards
    kpi_cards = []
    quick_stats = [
        ("Total Enrollment", f"{INSTITUTION['total_enrollment_f25']:,}", "Fall 2025", "-2.5% YoY"),
        ("Retention Rate", f"{INSTITUTION['retention_rate_f24']}%", "FTFT Students", "Recovering"),
        ("Graduate Students", "160", "Fall 2025", "+5.3% YoY"),
        ("Programs Analyzed", "23", "All Frameworks", "4 Frameworks"),
    ]
    kpi_accents = [FLC_NAVY, FLC_BLUE, FLC_BLUE_LIGHT, "#5ba3d9"]
    for i, (title, value, sub, trend) in enumerate(quick_stats):
        accent = kpi_accents[i]
        kpi_cards.append(html.Div([
            html.Div(title, style={"fontSize": "11px", "color": FLC_BLUE, "textTransform": "uppercase", "fontWeight": "700", "letterSpacing": "1px"}),
            html.Div(value, style={"fontSize": "36px", "fontWeight": "800", "color": FLC_NAVY, "marginTop": "6px", "lineHeight": "1"}),
            html.Div(sub, style={"fontSize": "12px", "color": "#6b8299", "marginTop": "6px"}),
            html.Div(trend, style={"fontSize": "12px", "color": accent, "fontWeight": "700", "marginTop": "4px"}),
        ], style={**CARD_STYLE, "textAlign": "center", "flex": "1", "minWidth": "180px",
                  "borderTop": f"4px solid {accent}", "padding": "24px 16px"}))

    # Enrollment trend mini chart
    fig_enroll = go.Figure()
    fig_enroll.add_trace(go.Scatter(
        x=ENROLLMENT_HISTORY["Year"], y=ENROLLMENT_HISTORY["Total_Headcount"],
        mode="lines+markers", line=dict(color=FLC_BLUE, width=3),
        marker=dict(size=8, color=FLC_NAVY, line=dict(width=2, color=FLC_BLUE)),
        name="Total Headcount",
        fill="tozeroy", fillcolor="rgba(0,102,179,0.07)",
    ))
    fig_enroll.update_layout(
        template=FLC_CHART_TEMPLATE,
        title=dict(text="10-Year Enrollment Trend"),
        height=300, margin=dict(l=40, r=20, t=50, b=30), showlegend=False,
    )

    # 3-phase overview cards (blue family accents)
    phase_cards = []
    phases = [
        ("Phase 1: Environmental Scanning", FLC_BLUE_LIGHT,
         "PESTLE, Porter's Five Forces, Gray Associates Portfolio Analysis, and BCG Growth-Share Matrix provide comprehensive external and internal scanning.",
         "4 frameworks completed"),
        ("Phase 2: Strategic Synthesis", FLC_BLUE,
         "SWOT Analysis synthesizes all Phase 1 findings into actionable Strengths, Weaknesses, Opportunities, and Threats with source attribution.",
         "Cross-framework synthesis"),
        ("Phase 3: Strategic Direction", FLC_NAVY,
         "Zone to Win framework with 3 strategic scenarios. Strategic Roadmap with risk assessment, implementation timeline, and resource allocation based on the Moderate-Adaptive scenario.",
         "Implementation planning"),
    ]
    for name, color, desc, badge_text in phases:
        phase_cards.append(html.Div([
            html.Div([
                html.Strong(name, style={"fontSize": "15px", "color": FLC_NAVY}),
                _badge(badge_text, color),
            ]),
            html.P(desc, style={"fontSize": "13px", "color": "#4a6070", "marginTop": "6px", "marginBottom": "0", "lineHeight": "1.6"}),
        ], style={**CARD_STYLE, "padding": "16px", "borderLeft": f"4px solid {color}"}))

    # Framework highlight summaries
    framework_summaries = []
    _bcg_q = BCG_DATA["Quadrant"].value_counts()
    _dept_q = BCG_DEPT_DATA["Quadrant"].value_counts()
    fw_data = [
        ("PESTLE Analysis", "Internal FLC Documents",
         "Political and Economic factors rated highest impact. Key risks: federal DEI policy disruption, tribal waiver vulnerability, state funding decline. Key opportunity: Indigenous education (statutorily grounded), AI Institute."),
        ("BCG Analysis", "22 Depts + 48 Majors",
         f"Departments (SCH): {_dept_q.get('Star', 0)} Stars, {_dept_q.get('Cash Cow', 0)} Cash Cows, {_dept_q.get('Concern', 0)} Concerns. "
         f"Majors (enrollment): {_bcg_q.get('Star', 0)} Stars, {_bcg_q.get('Cash Cow', 0)} Cash Cows, {_bcg_q.get('Concern', 0)} Concerns. "
         f"{int(BCG_DATA['Small_Base'].sum())} small-base programs flagged."),
        ("Porter's Analysis", "Internet Methodology + FLC Data",
         "Overall competitive intensity: HIGH. Strongest defense: statutory Native American mission and outdoor lifestyle. Online competition unverified for FLC specifically. Durango housing constrains faculty recruitment."),
        ("Gray Analysis", "Internet Methodology + FLC Data",
         "7 programs recommended to GROW, 7 to SUSTAIN, 2 to TRANSFORM, 4 to EVALUATE, 2 for SUNSET REVIEW. Note: Gray methodology applied with proxy data; scores are directional, not definitive."),
    ]
    fw_badge_colors = [FLC_NAVY, FLC_NAVY, FLC_BLUE, FLC_BLUE]
    for i, (name, source, summary) in enumerate(fw_data):
        framework_summaries.append(html.Div([
            html.Div([
                html.Strong(name, style={"fontSize": "14px", "color": FLC_NAVY}),
                _badge(source, fw_badge_colors[i]),
            ]),
            html.P(summary, style={"fontSize": "13px", "color": "#4a6070", "marginTop": "6px", "marginBottom": "0", "lineHeight": "1.6"}),
        ], style={**CARD_STYLE, "padding": "16px"}))

    # Retention trend mini chart
    fig_retention = go.Figure()
    fig_retention.add_trace(go.Scatter(
        x=RETENTION_HISTORY["Year"], y=RETENTION_HISTORY["Retention_Rate"],
        mode="lines+markers", line=dict(color=FLC_NAVY, width=3),
        marker=dict(size=8, color=FLC_BLUE, line=dict(width=2, color=FLC_NAVY)),
        name="Retention Rate",
        fill="tozeroy", fillcolor="rgba(0,48,87,0.05)",
    ))
    fig_retention.add_hline(y=73, line_dash="dash", line_color=CLR_HIGH, line_width=1,
                            annotation_text="National Avg (73%)", annotation_position="right")
    fig_retention.update_layout(
        template=FLC_CHART_TEMPLATE,
        title=dict(text="FTFT Retention Rate Trend"),
        height=300, margin=dict(l=40, r=20, t=50, b=30), showlegend=False,
        yaxis_range=[50, 80],
    )

    # Download buttons for project deliverables
    dl_btn_style = {
        "backgroundColor": FLC_BLUE, "color": "white", "border": "none",
        "padding": "9px 18px", "borderRadius": "6px", "cursor": "pointer",
        "fontSize": "12px", "fontWeight": "600", "marginRight": "8px",
        "transition": "background-color 0.2s ease", "letterSpacing": "0.2px",
    }
    deliverables_block = html.Div([
        html.Div("Project Deliverables", style={
            "fontSize": "13px", "color": FLC_NAVY, "fontWeight": "700",
            "textTransform": "uppercase", "letterSpacing": "1px",
            "marginBottom": "10px",
        }),
        html.Div([
            html.Button("Download Executive Summary (.docx)",
                        id="dl-exec-summary-docx-btn", style=dl_btn_style),
            dcc.Download(id="dl-exec-summary-docx"),
            html.Button("Download Executive Summary Deck (.pptx)",
                        id="dl-exec-summary-pptx-btn", style=dl_btn_style),
            dcc.Download(id="dl-exec-summary-pptx"),
        ], style={"marginBottom": "8px"}),
        html.Div([
            html.Button("Download Final Report (.docx)",
                        id="dl-final-report-docx-btn", style=dl_btn_style),
            dcc.Download(id="dl-final-report-docx"),
            html.Button("Download Final Presentation (.pptx)",
                        id="dl-final-report-pptx-btn", style=dl_btn_style),
            dcc.Download(id="dl-final-report-pptx"),
        ]),
    ], style={**CARD_STYLE, "borderLeft": f"4px solid {FLC_GOLD}",
              "marginBottom": "20px"})

    return html.Div([
        html.H2("Executive Summary", style={**SECTION_TITLE, "fontSize": "24px"}),
        html.P(
            f"Fort Lewis College Portfolio Optimization Project | Updated {datetime.now().strftime('%B %d, %Y')}",
            style={"color": "#6b8299", "marginBottom": "16px", "fontSize": "13px"},
        ),
        deliverables_block,

        # KPI row
        html.Div(kpi_cards, style={"display": "flex", "gap": "14px", "flexWrap": "wrap", "marginBottom": "20px"}),

        # Two-column: enrollment + retention trends
        html.Div([
            html.Div([
                dcc.Graph(figure=fig_enroll, config={"displayModeBar": False}),
                source_annotation("Source: FLC Enrollment Overview PDF, Fall census data"),
            ], style={**CARD_STYLE, "flex": "1"}),
            html.Div([
                dcc.Graph(figure=fig_retention, config={"displayModeBar": False}),
                source_annotation("Source: FLC Institutional Data, FTFT cohort tracking"),
            ], style={**CARD_STYLE, "flex": "1"}),
        ], style={"display": "flex", "gap": "16px"}),

        # 3-Phase overview
        html.H3("Three-Phase Strategic Framework", style={**SECTION_TITLE, "fontSize": "18px"}),
        html.Div(phase_cards),

        # Framework summaries
        html.H3("Phase 1 Framework Highlights", style={**SECTION_TITLE, "fontSize": "18px"}),
        html.Div(framework_summaries),
    ])


def build_pestle_tab():
    """PESTLE Analysis tab with radar chart, bar chart, and factor details."""
    categories = list(PESTLE_DATA.keys())
    scores = [PESTLE_DATA[c]["impact_score"] for c in categories]

    fig_radar = go.Figure(data=go.Scatterpolar(
        r=scores + [scores[0]],
        theta=categories + [categories[0]],
        fill="toself",
        fillcolor="rgba(0,102,179,0.12)",
        line=dict(color=FLC_BLUE, width=2),
        marker=dict(size=8, color=FLC_NAVY),
    ))
    fig_radar.update_layout(
        template=FLC_CHART_TEMPLATE,
        polar=dict(
            radialaxis=dict(visible=True, range=[0, 5], tickvals=[1, 2, 3, 4, 5],
                            gridcolor=FLC_BLUE_PALE, linecolor=FLC_BORDER),
            angularaxis=dict(gridcolor=FLC_BLUE_PALE, linecolor=FLC_BORDER),
            bgcolor="rgba(0,0,0,0)",
        ),
        title=dict(text="PESTLE Impact Assessment (1-5 scale)"),
        height=400, margin=dict(l=60, r=60, t=50, b=30),
    )

    impact_colors = {"High": CLR_HIGH, "Medium": CLR_MEDIUM, "Low": CLR_LOW}
    # Use a blue gradient for the bars, with text showing impact level
    bar_blues = [FLC_NAVY, FLC_BLUE, FLC_BLUE_LIGHT, "#5ba3d9", "#8cc0e8", "#b8d8f0"]
    fig_bar = go.Figure(data=[go.Bar(
        x=categories, y=scores,
        marker_color=bar_blues[:len(categories)],
        text=[PESTLE_DATA[c]["impact"] for c in categories],
        textposition="outside", textfont=dict(color=FLC_NAVY, size=11),
    )])
    fig_bar.update_layout(
        template=FLC_CHART_TEMPLATE,
        title=dict(text="Impact Level by PESTLE Category"),
        yaxis_title="Impact Score", yaxis_range=[0, 6],
        height=350, margin=dict(l=40, r=20, t=50, b=30),
    )

    detail_cards = []
    trend_colors = {"Negative": CLR_HIGH, "Mixed": CLR_MEDIUM,
                    "Stable": FLC_BLUE, "Opportunity": CLR_POSITIVE}
    for cat in categories:
        d = PESTLE_DATA[cat]
        detail_cards.append(html.Div([
            html.Div([
                html.Strong(cat, style={"fontSize": "16px", "color": FLC_NAVY}),
                _badge(f"Impact: {d['impact']}", impact_colors.get(d["impact"], CLR_NEUTRAL)),
                _badge(f"Trend: {d['trend']}", trend_colors.get(d["trend"], CLR_NEUTRAL)),
            ]),
            html.Div([
                html.Strong("Key Factors:", style={"fontSize": "12px", "color": FLC_NAVY}),
                html.Ul([html.Li(f, style={"fontSize": "12px", "color": "#4a6070"}) for f in d["factors"]],
                        style={"marginTop": "4px", "marginBottom": "4px"}),
            ], style={"marginTop": "8px"}),
            html.Div([
                html.Strong("Opportunities:", style={"fontSize": "12px", "color": FLC_BLUE}),
                html.Ul([html.Li(o, style={"fontSize": "12px", "color": "#4a6070"}) for o in d["opportunities"]],
                        style={"marginTop": "4px"}),
            ]),
        ], style={**CARD_STYLE, "padding": "16px", "borderLeft": f"3px solid {FLC_BLUE}"}))

    return html.Div([
        html.H2("PESTLE Analysis", style=SECTION_TITLE),
        framework_description_block("PESTLE"),
        data_source_badge("PESTLE Analysis"),
        download_buttons("PESTLE"),
        html.Div([
            html.Div([
                dcc.Graph(figure=fig_radar, config={"displayModeBar": False}),
                source_annotation("Source: PESTLE_Report_FLC.docx, External Forces Shaping FLC.pptx"),
            ], style={**CARD_STYLE, "flex": "1"}),
            html.Div([
                dcc.Graph(figure=fig_bar, config={"displayModeBar": False}),
                source_annotation("Source: PESTLE_Report_FLC.docx"),
            ], style={**CARD_STYLE, "flex": "1"}),
        ], style={"display": "flex", "gap": "16px"}),
        html.H3("Detailed Factor Analysis", style={**SECTION_TITLE, "fontSize": "16px"}),
        html.Div(detail_cards),
    ])


def build_porters_tab():
    """Porter's Analysis tab with radar and detail cards."""
    forces = list(PORTERS_DATA.keys())
    scores = [PORTERS_DATA[f]["score"] for f in forces]

    fig_radar = go.Figure(data=go.Scatterpolar(
        r=scores + [scores[0]],
        theta=forces + [forces[0]],
        fill="toself",
        fillcolor="rgba(0,48,87,0.10)",
        line=dict(color=FLC_NAVY, width=2),
        marker=dict(size=8, color=FLC_BLUE),
    ))
    fig_radar.update_layout(
        template=FLC_CHART_TEMPLATE,
        polar=dict(
            radialaxis=dict(visible=True, range=[0, 5], tickvals=[1, 2, 3, 4, 5],
                            ticktext=["1-Low", "2", "3-Med", "4", "5-High"],
                            gridcolor=FLC_BLUE_PALE, linecolor=FLC_BORDER),
            angularaxis=dict(gridcolor=FLC_BLUE_PALE, linecolor=FLC_BORDER),
            bgcolor="rgba(0,0,0,0)",
        ),
        title=dict(text="Porter's Five Forces - Competitive Intensity"),
        height=420, margin=dict(l=80, r=80, t=50, b=30),
    )

    porter_blues = [FLC_NAVY, FLC_BLUE, FLC_BLUE_LIGHT, "#5ba3d9", "#8cc0e8"]
    fig_bar = go.Figure(data=[go.Bar(
        y=forces, x=scores, orientation="h",
        marker_color=porter_blues[:len(forces)],
        text=[PORTERS_DATA[f]["rating"] for f in forces],
        textposition="outside", textfont=dict(color=FLC_NAVY, size=11),
    )])
    fig_bar.update_layout(
        template=FLC_CHART_TEMPLATE,
        title=dict(text="Force Intensity Ratings"),
        xaxis_title="Intensity (1=Low, 5=High)", xaxis_range=[0, 5.5],
        height=350, margin=dict(l=180, r=40, t=50, b=30),
    )

    trend_colors_map = {"Increasing": CLR_HIGH, "Decreasing": CLR_LOW,
                        "Stable": FLC_BLUE, "Improving": CLR_POSITIVE}
    force_cards = []
    for fi, force in enumerate(forces):
        d = PORTERS_DATA[force]
        accent = porter_blues[fi % len(porter_blues)]
        ind_rows = []
        for ri, ind in enumerate(d["indicators"]):
            trend_icon = {"Increasing": "^", "Decreasing": "v", "Stable": "-", "Improving": "^"}.get(ind["trend"], "?")
            row_bg = FLC_BLUE_WASH if ri % 2 == 0 else BG_WHITE
            ind_rows.append(html.Tr([
                html.Td(ind["name"], style={"fontSize": "12px", "padding": "6px 10px", "color": FLC_NAVY, "backgroundColor": row_bg}),
                html.Td(ind["value"], style={"fontSize": "12px", "padding": "6px 10px", "fontWeight": "600", "color": FLC_NAVY, "backgroundColor": row_bg}),
                html.Td(f"{trend_icon} {ind['trend']}", style={
                    "fontSize": "12px", "padding": "6px 10px", "color": trend_colors_map.get(ind["trend"], CLR_NEUTRAL),
                    "fontWeight": "600", "backgroundColor": row_bg,
                }),
            ]))
        force_cards.append(html.Div([
            html.Div([
                html.Strong(force, style={"fontSize": "15px", "color": FLC_NAVY}),
                _badge(d["rating"], accent),
            ]),
            html.P(d["description"], style={"fontSize": "12px", "color": "#4a6070", "margin": "6px 0", "lineHeight": "1.6"}),
            html.Table([
                html.Thead(html.Tr([
                    html.Th("Indicator", style={"fontSize": "11px", "padding": "8px 10px", "color": FLC_NAVY, "fontWeight": "700",
                                                "borderBottom": f"2px solid {FLC_BLUE}", "textTransform": "uppercase", "letterSpacing": "0.5px", "backgroundColor": BG_WHITE}),
                    html.Th("Value", style={"fontSize": "11px", "padding": "8px 10px", "color": FLC_NAVY, "fontWeight": "700",
                                            "borderBottom": f"2px solid {FLC_BLUE}", "textTransform": "uppercase", "letterSpacing": "0.5px", "backgroundColor": BG_WHITE}),
                    html.Th("Trend", style={"fontSize": "11px", "padding": "8px 10px", "color": FLC_NAVY, "fontWeight": "700",
                                            "borderBottom": f"2px solid {FLC_BLUE}", "textTransform": "uppercase", "letterSpacing": "0.5px", "backgroundColor": BG_WHITE}),
                ])),
                html.Tbody(ind_rows),
            ], style={"width": "100%", "borderCollapse": "collapse", "marginTop": "10px", "border": f"1px solid {FLC_BLUE_PALE}"}),
        ], style={**CARD_STYLE, "padding": "16px", "borderLeft": f"3px solid {accent}"}))

    insight_box = html.Div([
        html.H3("Strategic Implications", style={**SECTION_TITLE, "fontSize": "16px"}),
        html.Ul([html.Li(i, style={"marginBottom": "8px", "fontSize": "13px", "color": "#4a6070", "lineHeight": "1.6"}) for i in PORTERS_INSIGHTS]),
    ], style=CARD_STYLE)

    return html.Div([
        html.H2("Porter's Analysis", style=SECTION_TITLE),
        framework_description_block("Porters"),
        data_source_badge("Porter's Five Forces"),
        download_buttons("Porters"),
        html.Div([
            html.Div([
                dcc.Graph(figure=fig_radar, config={"displayModeBar": False}),
                source_annotation("Source: Porter's Five Forces methodology applied to FLC institutional data"),
            ], style={**CARD_STYLE, "flex": "1"}),
            html.Div([
                dcc.Graph(figure=fig_bar, config={"displayModeBar": False}),
                source_annotation("Source: Porter's Five Forces methodology applied to FLC institutional data"),
            ], style={**CARD_STYLE, "flex": "1"}),
        ], style={"display": "flex", "gap": "16px"}),
        html.H3("Force Analysis Details", style={**SECTION_TITLE, "fontSize": "16px"}),
        html.Div(force_cards),
        insight_box,
    ])


def build_gray_tab():
    """Gray Analysis tab - preserving the bubble chart exactly."""
    df = GRAY_ASSOCIATES_DATA.copy()

    # Bubble chart: Market Score vs Economics Score, size=Enrollment
    fig = go.Figure()
    for rec in df["GA_Recommendation"].unique():
        df_r = df[df["GA_Recommendation"] == rec]
        fig.add_trace(go.Scatter(
            x=df_r["Economics_Score"], y=df_r["Market_Score"],
            mode="markers+text", name=rec,
            marker=dict(
                size=df_r["Enrollment"] / 5 + 8,
                color=GA_RECOMMENDATION_COLORS.get(rec, "#999"),
                opacity=0.8, line=dict(width=1, color="white"),
            ),
            text=df_r["Program"],
            textposition="top center",
            textfont=dict(size=8),
        ))

    fig.add_hline(y=55, line_dash="dash", line_color="#aaa", line_width=1)
    fig.add_vline(x=55, line_dash="dash", line_color="#aaa", line_width=1)

    fig.update_layout(
        template=FLC_CHART_TEMPLATE,
        title=dict(text="Gray Associates Portfolio Matrix: Market Score vs. Program Economics"),
        xaxis_title="Program Economics Score (Revenue Efficiency)",
        yaxis_title="Market Score (Student Demand + Employment + Competition)",
        height=600,
        margin=dict(l=50, r=30, t=50, b=50),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
        annotations=[
            dict(x=30, y=80, text="SUSTAIN", showarrow=False,
                 font=dict(size=14, color=FLC_BLUE_LIGHT), opacity=0.5),
            dict(x=80, y=80, text="GROW", showarrow=False,
                 font=dict(size=14, color=FLC_NAVY), opacity=0.5),
            dict(x=30, y=30, text="SUNSET REVIEW", showarrow=False,
                 font=dict(size=14, color=CLR_HIGH), opacity=0.5),
            dict(x=80, y=30, text="TRANSFORM", showarrow=False,
                 font=dict(size=14, color=CLR_MEDIUM), opacity=0.5),
        ],
    )

    # Recommendation summary bar
    rec_counts = df["GA_Recommendation"].value_counts()
    ga_blue_map = {"Grow": FLC_NAVY, "Sustain": FLC_BLUE, "Transform": FLC_BLUE_LIGHT,
                   "Evaluate": "#5ba3d9", "Sunset Review": "#8cc0e8"}
    fig_bar = go.Figure(data=[go.Bar(
        x=rec_counts.index, y=rec_counts.values,
        marker_color=[ga_blue_map.get(r, "#b8d8f0") for r in rec_counts.index],
        text=rec_counts.values, textposition="outside",
        textfont=dict(color=FLC_NAVY, size=12, family="Segoe UI"),
    )])
    fig_bar.update_layout(
        template=FLC_CHART_TEMPLATE,
        title=dict(text="Programs by Recommendation"), height=300,
        margin=dict(l=40, r=20, t=50, b=30),
        yaxis_title="Number of Programs",
    )

    table_df = df[["Program", "Enrollment", "Student_Demand_Score", "Employment_Score",
                   "Competition_Score", "Market_Score", "Economics_Score",
                   "Mission_Alignment", "GA_Recommendation"]].sort_values("Market_Score", ascending=False)

    # Blue-toned recommendation colors for table rows
    rec_color_map = {
        "Grow": FLC_BLUE_WASH, "Sustain": FLC_BLUE_PALE,
        "Transform": "#e8f0f8", "Evaluate": "#f0f4f8", "Sunset Review": "#f5f0f0",
    }
    style_conditions = TABLE_ROW_BANDING + [
        {"if": {"filter_query": f'{{GA_Recommendation}} = "{rec}"'},
         "backgroundColor": color}
        for rec, color in rec_color_map.items()
    ]

    return html.Div([
        html.H2("Gray Analysis", style=SECTION_TITLE),
        framework_description_block("Gray"),
        data_source_badge("Gray Associates Portfolio"),
        download_buttons("Gray"),
        html.Div([dcc.Graph(figure=fig, config={"displayModeBar": False})], style=CARD_STYLE),
        source_annotation("Source: Gray Associates PES methodology applied to FLC enrollment & BCG data"),
        html.Div([
            html.Div([
                dcc.Graph(figure=fig_bar, config={"displayModeBar": False}),
                source_annotation("Source: Gray Associates classification of 23 FLC programs"),
            ], style={**CARD_STYLE, "flex": "1"}),
            html.Div([
                html.H3("Key Insights", style={**SECTION_TITLE, "fontSize": "16px"}),
                html.Ul([html.Li(i, style={"marginBottom": "8px", "fontSize": "13px", "color": "#4a6070", "lineHeight": "1.6"}) for i in GA_INSIGHTS]),
            ], style={**CARD_STYLE, "flex": "1"}),
        ], style={"display": "flex", "gap": "16px"}),

        # --- Methodology Explanation ---
        html.H3("Methodology", style={**SECTION_TITLE, "fontSize": "16px", "marginTop": "24px"}),
        html.Div([
            html.P("The Gray Associates Program Evaluation System (PES) plots each academic program on two axes to identify "
                   "investment priorities. FLC does not hold a Gray Associates subscription \u2014 scores below are estimated by applying "
                   "the published Gray methodology to FLC institutional data.",
                   style={"fontSize": "13px", "color": "#4a6070", "lineHeight": "1.7", "marginBottom": "14px"}),

            html.H4("Market Score (Y-Axis)", style={"color": FLC_NAVY, "fontSize": "14px", "fontWeight": "700", "marginBottom": "6px"}),
            html.P("Measures how attractive a program is in the external market.",
                   style={"fontSize": "13px", "color": "#4a6070", "lineHeight": "1.6", "marginBottom": "6px"}),
            html.Div("Market Score = (Student Demand \u00d7 0.40) + (Employment \u00d7 0.40) + (Competition \u00d7 0.20)",
                     style={"fontFamily": "Consolas, monospace", "fontSize": "13px", "color": FLC_NAVY,
                            "backgroundColor": FLC_BLUE_WASH, "padding": "10px 14px", "borderRadius": "6px",
                            "border": f"1px solid {FLC_BLUE_PALE}", "marginBottom": "10px"}),
            html.Ul([
                html.Li([html.Strong("Student Demand (40%): "), "Current enrollment and enrollment trend \u2014 sourced from FLC enrollment data and BCG market-share analysis."],
                        style={"fontSize": "12px", "color": "#4a6070", "marginBottom": "4px", "lineHeight": "1.6"}),
                html.Li([html.Strong("Employment (40%): "), "Career prospects and regional job-market strength for graduates \u2014 sourced from regional employment projections."],
                        style={"fontSize": "12px", "color": "#4a6070", "marginBottom": "4px", "lineHeight": "1.6"}),
                html.Li([html.Strong("Competition (20%): "), "Market saturation (lower competition = higher score) \u2014 sourced from BCG competitive positioning data."],
                        style={"fontSize": "12px", "color": "#4a6070", "marginBottom": "4px", "lineHeight": "1.6"}),
            ], style={"paddingLeft": "20px", "marginBottom": "16px"}),

            html.H4("Economics Score (X-Axis)", style={"color": FLC_NAVY, "fontSize": "14px", "fontWeight": "700", "marginBottom": "6px"}),
            html.P("Measures how efficiently a program generates revenue relative to its costs.",
                   style={"fontSize": "13px", "color": "#4a6070", "lineHeight": "1.6", "marginBottom": "6px"}),
            html.Ul([
                html.Li([html.Strong("SCH Generation Efficiency: "), "Percentage of total institutional Student Credit Hours generated by the program."],
                        style={"fontSize": "12px", "color": "#4a6070", "marginBottom": "4px", "lineHeight": "1.6"}),
                html.Li([html.Strong("Program Cost Structure: "), "Revenue efficiency relative to operating costs (faculty, labs, facilities, student support)."],
                        style={"fontSize": "12px", "color": "#4a6070", "marginBottom": "4px", "lineHeight": "1.6"}),
            ], style={"paddingLeft": "20px", "marginBottom": "16px"}),

            html.P([html.Strong("Interpretation: "), "Scores above 65 indicate strength; 50\u201365 is solid; below 50 indicates weakness on that axis."],
                   style={"fontSize": "12px", "color": "#4a6070", "lineHeight": "1.6", "marginBottom": "4px"}),
        ], style=CARD_STYLE),

        # --- Decision Rules ---
        html.H3("Program Scorecard \u2014 Decision Rules", style={**SECTION_TITLE, "fontSize": "16px"}),
        html.Div([
            html.P("Each program is assigned to one of five categories based on where it falls on the Market Score and Economics Score axes:",
                   style={"fontSize": "13px", "color": "#4a6070", "lineHeight": "1.7", "marginBottom": "12px"}),
            html.Table([
                html.Thead(html.Tr([
                    html.Th("Recommendation", style={**TABLE_HEADER_STYLE, "textAlign": "left", "width": "18%"}),
                    html.Th("Market Score", style={**TABLE_HEADER_STYLE, "textAlign": "center", "width": "14%"}),
                    html.Th("Economics Score", style={**TABLE_HEADER_STYLE, "textAlign": "center", "width": "16%"}),
                    html.Th("Action", style={**TABLE_HEADER_STYLE, "textAlign": "left", "width": "52%"}),
                ])),
                html.Tbody([
                    html.Tr([html.Td("Grow", style={"fontWeight": "600"}), html.Td("> 65", style={"textAlign": "center"}), html.Td("> 65", style={"textAlign": "center"}), html.Td("Invest aggressively \u2014 scale enrollment, add capacity")],
                            style={"backgroundColor": FLC_BLUE_WASH}),
                    html.Tr([html.Td("Sustain", style={"fontWeight": "600"}), html.Td("50\u201365", style={"textAlign": "center"}), html.Td("Any", style={"textAlign": "center"}), html.Td("Maintain current enrollment, improve efficiency")],
                            style={"backgroundColor": FLC_BLUE_PALE}),
                    html.Tr([html.Td("Transform", style={"fontWeight": "600"}), html.Td("< 50", style={"textAlign": "center"}), html.Td("> 60", style={"textAlign": "center"}), html.Td("Weak market but strong economics \u2014 innovate delivery model")],
                            style={"backgroundColor": "#e8f0f8"}),
                    html.Tr([html.Td("Evaluate", style={"fontWeight": "600"}), html.Td("< 50", style={"textAlign": "center"}), html.Td("< 60", style={"textAlign": "center"}), html.Td("Requires deep-dive review before deciding direction")],
                            style={"backgroundColor": "#f0f4f8"}),
                    html.Tr([html.Td("Sunset Review", style={"fontWeight": "600"}), html.Td("< 40", style={"textAlign": "center"}), html.Td("< 50", style={"textAlign": "center"}), html.Td("Weak on both axes \u2014 consider phase-out or major restructuring")],
                            style={"backgroundColor": "#f5f0f0"}),
                ], style={"fontSize": "12px", "color": FLC_NAVY, "lineHeight": "1.6"}),
            ], style={"width": "100%", "borderCollapse": "collapse", "border": f"1px solid {FLC_BLUE_PALE}"}),
            html.P([
                html.Strong("Note: "),
                "Mission alignment is noted in the scorecard but not weighted into the quantitative score. "
                "Programs serving foundational/service roles (e.g., English, Mathematics) or tribal obligations "
                "should not be evaluated solely on these metrics."
            ], style={"fontSize": "12px", "color": "#4a6070", "lineHeight": "1.6", "marginTop": "12px", "fontStyle": "italic"}),
        ], style=CARD_STYLE),

        # --- Program Scorecard Data Table ---
        html.H3("Program Scorecard", style={**SECTION_TITLE, "fontSize": "16px"}),
        html.Div([dash_table.DataTable(
            data=table_df.to_dict("records"),
            columns=[
                {"name": "Program", "id": "Program"},
                {"name": "Enrollment", "id": "Enrollment"},
                {"name": "Student Demand", "id": "Student_Demand_Score"},
                {"name": "Employment", "id": "Employment_Score"},
                {"name": "Competition", "id": "Competition_Score"},
                {"name": "Market Score", "id": "Market_Score"},
                {"name": "Economics", "id": "Economics_Score"},
                {"name": "Mission", "id": "Mission_Alignment"},
                {"name": "Recommendation", "id": "GA_Recommendation"},
            ],
            style_cell=TABLE_CELL_STYLE,
            style_header=TABLE_HEADER_STYLE,
            style_data_conditional=style_conditions,
            sort_action="native",
            filter_action="native",
            page_size=25,
        )], style=CARD_STYLE),
        source_annotation("Source: Gray Associates PES methodology applied to FLC institutional data"),
    ])


def build_bcg_tab():
    """BCG Analysis tab \u2014 department-level (SCH) + major-level (enrollment) views."""

    # ═══════════════════════════════════════════════════════════════════════
    # DEPARTMENT-LEVEL BCG (22 departments, SCH-based)
    # ═══════════════════════════════════════════════════════════════════════
    dept_fig = go.Figure()
    for quadrant in ["Star", "Cash Cow", "Question Mark", "Concern"]:
        df_q = BCG_DEPT_DATA[BCG_DEPT_DATA["Quadrant"] == quadrant]
        dept_fig.add_trace(go.Scatter(
            x=df_q["SCH_Pct"], y=df_q["Two_Year_Change"],
            mode="markers+text", name=quadrant,
            marker=dict(
                size=df_q["SCH_Pct"] * 4 + 12,
                color=BCG_QUADRANT_COLORS[quadrant],
                opacity=0.85,
                line=dict(width=1, color="white"),
            ),
            text=df_q["Department"],
            textposition="top center",
            textfont=dict(size=9, color=FLC_NAVY),
            hovertemplate=(
                "<b>%{text}</b><br>"
                "SCH Share: %{x:.1f}%<br>"
                "2-Year Change: %{y:+.1f}%<br>"
                "<extra>%{fullData.name}</extra>"
            ),
        ))
    dept_fig.add_hline(y=0, line_dash="dash", line_color="#aaa", line_width=1)
    dept_fig.add_vline(x=4.0, line_dash="dash", line_color="#aaa", line_width=1)
    dept_annotations = [
        dict(x=1.5, y=14, text="Question Marks", showarrow=False,
             font=dict(size=13, color="#5ba3d9"), opacity=0.5),
        dict(x=8, y=14, text="Stars", showarrow=False,
             font=dict(size=13, color=FLC_NAVY), opacity=0.5),
        dict(x=1.5, y=-28, text="Concerns", showarrow=False,
             font=dict(size=13, color="#8cc0e8"), opacity=0.5),
        dict(x=8, y=-28, text="Cash Cows", showarrow=False,
             font=dict(size=13, color=FLC_BLUE), opacity=0.5),
    ]
    dept_fig.update_layout(
        template=FLC_CHART_TEMPLATE,
        title=dict(text="BCG Growth-Share Matrix (Departments)"),
        yaxis_title="2-Year Change % (Growth Rate)",
        xaxis_title="% of Total SCH (Market Share)",
        height=600,
        annotations=dept_annotations,
        margin=dict(l=50, r=30, t=50, b=50),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
    )

    dept_summary = BCG_DEPT_DATA.groupby("Quadrant").agg(
        Count=("Department", "count"),
        Avg_SCH_Pct=("SCH_Pct", "mean"),
        Avg_Change=("Two_Year_Change", "mean"),
    ).reindex(["Star", "Cash Cow", "Question Mark", "Concern"]).reset_index().round(1)

    dept_insight_list = html.Div([
        html.H3("Department-Level Insights", style={**SECTION_TITLE, "fontSize": "16px"}),
        html.Ul([html.Li(i, style={"marginBottom": "8px", "fontSize": "13px",
                                    "color": "#4a6070", "lineHeight": "1.6"})
                  for i in BCG_DEPT_INSIGHTS]),
    ], style=CARD_STYLE)

    # ═══════════════════════════════════════════════════════════════════════
    # MAJOR-LEVEL BCG (48 majors, enrollment-based)
    # ═══════════════════════════════════════════════════════════════════════
    df = BCG_DATA.copy()
    median_enroll = df["Enrollment_2024"].median()

    # --- Bubble chart ---
    fig = go.Figure()

    for quadrant in ["Star", "Cash Cow", "Question Mark", "Concern"]:
        for is_small in [False, True]:
            subset = df[(df["Quadrant"] == quadrant) & (df["Small_Base"] == is_small)]
            if subset.empty:
                continue
            # Bubble size proportional to |absolute change|; minimum size 8
            sizes = subset["Abs_Change"].abs().clip(lower=2) * 1.1 + 8
            fig.add_trace(go.Scatter(
                x=subset["Enrollment_2024"],
                y=subset["Pct_Change"],
                mode="markers",
                name=f"{quadrant}{' (small base)' if is_small else ''}",
                legendgroup=quadrant,
                showlegend=(not is_small),
                marker=dict(
                    size=sizes,
                    color=BCG_QUADRANT_COLORS[quadrant] if not is_small else "rgba(255,255,255,0.3)",
                    opacity=0.85 if not is_small else 0.55,
                    line=dict(
                        width=2 if is_small else 1,
                        color=BCG_QUADRANT_COLORS[quadrant],
                    ),
                ),
                text=subset["Major"],
                hovertemplate=(
                    "<b>%{text}</b><br>"
                    "2024 Enrollment: %{x}<br>"
                    "% Change (2022\u20132024): %{y:+.1f}%<br>"
                    "Absolute Change: %{customdata[0]:+d}<br>"
                    "2022 Enrollment: %{customdata[1]}<br>"
                    "Quartile: %{customdata[2]}<br>"
                    "<extra>%{customdata[3]}</extra>"
                ),
                customdata=np.column_stack([
                    subset["Abs_Change"].values,
                    subset["Enrollment_2022"].values,
                    subset["Quartile"].values,
                    [f"{quadrant} \u2022 {'Small Base' if is_small else ''}" for _ in range(len(subset))],
                ]),
            ))

    # Smart labels: top 12 by enrollment + extreme % changes
    label_top = df.nlargest(12, "Enrollment_2024")
    label_extreme = df[df["Pct_Change"].abs() > 60]
    labels = pd.concat([label_top, label_extreme]).drop_duplicates(subset="Major")
    # Truncate long names for readability
    label_texts = labels["Major"].str[:22]

    fig.add_trace(go.Scatter(
        x=labels["Enrollment_2024"],
        y=labels["Pct_Change"],
        mode="text",
        text=label_texts,
        textposition="top center",
        textfont=dict(size=8, color=FLC_NAVY),
        showlegend=False,
        hoverinfo="skip",
    ))

    # Quadrant dividers
    fig.add_hline(y=0, line_dash="dash", line_color="#aaa", line_width=1)
    fig.add_vline(x=median_enroll, line_dash="dash", line_color="#aaa", line_width=1)

    x_max = df["Enrollment_2024"].max()
    y_max = df["Pct_Change"].max()
    y_min = df["Pct_Change"].min()
    annotations = [
        dict(x=median_enroll / 2, y=y_max * 0.85, text="Question Marks",
             showarrow=False, font=dict(size=12, color="#5ba3d9"), opacity=0.5),
        dict(x=median_enroll + (x_max - median_enroll) / 2, y=y_max * 0.85, text="Stars",
             showarrow=False, font=dict(size=12, color=FLC_NAVY), opacity=0.5),
        dict(x=median_enroll / 2, y=y_min * 0.85, text="Concerns",
             showarrow=False, font=dict(size=12, color="#8cc0e8"), opacity=0.5),
        dict(x=median_enroll + (x_max - median_enroll) / 2, y=y_min * 0.85, text="Cash Cows",
             showarrow=False, font=dict(size=12, color=FLC_BLUE), opacity=0.5),
    ]
    fig.update_layout(
        template=FLC_CHART_TEMPLATE,
        title=dict(text="BCG Growth-Share Matrix (48 Majors, 2022\u20132024)"),
        xaxis_title="2024 Enrollment (Institutional Market Share Proxy)",
        yaxis_title="% Change 2022\u20132024 (Growth Rate)",
        height=650,
        annotations=annotations,
        margin=dict(l=50, r=30, t=50, b=50),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
    )

    # --- How to Read This Chart card ---
    reading_guide = html.Div([
        html.H3("How to Read This Chart", style={**SECTION_TITLE, "fontSize": "15px"}),
        html.Ul([
            html.Li([html.Strong("Bubble position: "), "X = 2024 enrollment size, Y = % change since 2022"],
                     style={"fontSize": "12px", "marginBottom": "4px", "color": "#4a6070"}),
            html.Li([html.Strong("Bubble size: "), "Proportional to absolute enrollment change (students gained/lost)"],
                     style={"fontSize": "12px", "marginBottom": "4px", "color": "#4a6070"}),
            html.Li([html.Strong("Hollow bubbles: "),
                      "Programs with fewer than 20 students in 2022 \u2014 their % changes can be misleading"],
                     style={"fontSize": "12px", "marginBottom": "4px", "color": "#4a6070"}),
            html.Li([html.Strong("Dashed lines: "),
                      f"Vertical = median enrollment ({int(median_enroll)}), Horizontal = 0% growth"],
                     style={"fontSize": "12px", "marginBottom": "4px", "color": "#4a6070"}),
        ], style={"paddingLeft": "16px", "margin": "8px 0"}),
    ], style={**CARD_STYLE, "backgroundColor": "#f8fafb", "borderLeft": f"4px solid {FLC_BLUE}"})

    # --- Quadrant summary ---
    summary = df.groupby("Quadrant").agg(
        Count=("Major", "count"),
        Avg_Enrollment=("Enrollment_2024", "mean"),
        Avg_Change=("Pct_Change", "mean"),
        Total_Abs_Change=("Abs_Change", "sum"),
    ).reindex(["Star", "Cash Cow", "Question Mark", "Concern"]).reset_index().round(1)

    insight_list = html.Div([
        html.H3("Key Insights", style={**SECTION_TITLE, "fontSize": "16px"}),
        html.Ul([html.Li(i, style={"marginBottom": "8px", "fontSize": "13px",
                                    "color": "#4a6070", "lineHeight": "1.6"})
                  for i in BCG_INSIGHTS]),
    ], style=CARD_STYLE)

    # --- Full program detail table ---
    detail_df = df[["Major", "Enrollment_2022", "Enrollment_2024", "Abs_Change",
                     "Pct_Change", "Quartile", "Quadrant", "Small_Base"]].copy()
    detail_df["Pct_Change"] = detail_df["Pct_Change"].round(1)
    # Append asterisk to major name for small-base programs so it's visible on every page
    detail_df["Major"] = detail_df.apply(
        lambda r: f"{r['Major']} *" if r["Small_Base"] else r["Major"], axis=1)
    detail_df = detail_df.drop(columns=["Small_Base"])
    detail_df = detail_df.sort_values("Enrollment_2024", ascending=False)

    return html.Div([
        html.H2("BCG Analysis", style=SECTION_TITLE),
        framework_description_block("BCG"),
        data_source_badge("BCG Growth-Share Matrix"),
        download_buttons("BCG"),

        # ── Department-Level View ──
        html.H3("Department-Level Analysis (22 Departments \u2014 SCH-Based)", style={**SECTION_TITLE, "fontSize": "18px"}),
        html.Div([dcc.Graph(figure=dept_fig, config={"displayModeBar": False})], style=CARD_STYLE),
        source_annotation("Source: BCG Presentation.pptx, BCG-growthMatrixDepts.png (FLC Internal)"),
        html.Div([
            html.Div([
                html.H3("Department Quadrant Summary", style={**SECTION_TITLE, "fontSize": "16px"}),
                dash_table.DataTable(
                    data=dept_summary.to_dict("records"),
                    columns=[
                        {"name": "Quadrant", "id": "Quadrant"},
                        {"name": "# Departments", "id": "Count"},
                        {"name": "Avg SCH %", "id": "Avg_SCH_Pct"},
                        {"name": "Avg 2-Yr Change", "id": "Avg_Change"},
                    ],
                    style_cell=TABLE_CELL_STYLE,
                    style_header=TABLE_HEADER_STYLE,
                    style_data_conditional=TABLE_ROW_BANDING + [
                        {"if": {"filter_query": '{Quadrant} = "Star"'}, "backgroundColor": FLC_BLUE_WASH},
                        {"if": {"filter_query": '{Quadrant} = "Concern"'}, "backgroundColor": "#f5f0f0"},
                    ],
                ),
            ], style={**CARD_STYLE, "flex": "1"}),
            html.Div([dept_insight_list], style={"flex": "1"}),
        ], style={"display": "flex", "gap": "16px"}),

        # ── Major-Level View ──
        html.H3("Major-Level Analysis (48 Majors \u2014 Enrollment-Based)", style={**SECTION_TITLE, "fontSize": "18px", "marginTop": "32px"}),
        html.Div([dcc.Graph(figure=fig, config={"displayModeBar": False})], style=CARD_STYLE),
        reading_guide,
        source_annotation("Source: Dataset_Majors.xlsx (FLC Institutional Data, 2022\u20132024)"),
        html.Div([
            html.Div([
                html.H3("Quadrant Summary", style={**SECTION_TITLE, "fontSize": "16px"}),
                dash_table.DataTable(
                    data=summary.to_dict("records"),
                    columns=[
                        {"name": "Quadrant", "id": "Quadrant"},
                        {"name": "# Majors", "id": "Count"},
                        {"name": "Avg Enrollment", "id": "Avg_Enrollment"},
                        {"name": "Avg % Change", "id": "Avg_Change"},
                        {"name": "Total \u0394 Students", "id": "Total_Abs_Change"},
                    ],
                    style_cell=TABLE_CELL_STYLE,
                    style_header=TABLE_HEADER_STYLE,
                    style_data_conditional=TABLE_ROW_BANDING + [
                        {"if": {"filter_query": '{Quadrant} = "Star"'}, "backgroundColor": FLC_BLUE_WASH},
                        {"if": {"filter_query": '{Quadrant} = "Concern"'}, "backgroundColor": "#f5f0f0"},
                    ],
                ),
            ], style={**CARD_STYLE, "flex": "1"}),
            html.Div([insight_list], style={"flex": "1"}),
        ], style={"display": "flex", "gap": "16px"}),
        html.Div([
            html.H3("Program Detail (All 48 Majors)", style={**SECTION_TITLE, "fontSize": "16px"}),
            dash_table.DataTable(
                data=detail_df.to_dict("records"),
                columns=[
                    {"name": "Major", "id": "Major"},
                    {"name": "2022", "id": "Enrollment_2022"},
                    {"name": "2024", "id": "Enrollment_2024"},
                    {"name": "\u0394 Students", "id": "Abs_Change"},
                    {"name": "% Change", "id": "Pct_Change"},
                    {"name": "Quartile", "id": "Quartile"},
                    {"name": "Quadrant", "id": "Quadrant"},
                ],
                style_cell=TABLE_CELL_STYLE,
                style_header=TABLE_HEADER_STYLE,
                style_data_conditional=TABLE_ROW_BANDING + [
                    {"if": {"filter_query": '{Quadrant} = "Star"'}, "backgroundColor": FLC_BLUE_WASH},
                    {"if": {"filter_query": '{Quadrant} = "Concern"'}, "backgroundColor": "#f5f0f0"},
                    {"if": {"filter_query": '{Major} contains "*"'}, "fontStyle": "italic", "color": "#888"},
                ],
                sort_action="native",
                filter_action="native",
                page_size=25,
            ),
            html.P("* Small base: fewer than 20 students in 2022 \u2014 percentage changes may be misleading.",
                    style={"fontSize": "11px", "color": "#888", "fontStyle": "italic", "marginTop": "6px"}),
        ], style=CARD_STYLE),
    ])


def build_swot_tab():
    """Phase 2: SWOT Analysis synthesizing all Phase 1 frameworks."""
    quadrants = []
    for label in ["Strengths", "Weaknesses", "Opportunities", "Threats"]:
        data = SWOT_DATA[label]
        items = []
        for item in data["items"]:
            items.append(html.Div([
                html.Div([
                    html.Strong(item["title"], style={"fontSize": "13px", "color": FLC_NAVY}),
                ]),
                html.P(item["detail"], style={"fontSize": "12px", "color": "#444", "margin": "4px 0"}),
                html.Div(f"Source: {item['source']}", style={
                    "fontSize": "10px", "color": "#999", "fontStyle": "italic",
                }),
            ], style={
                "padding": "10px", "marginBottom": "8px",
                "backgroundColor": FLC_BLUE_WASH, "borderRadius": "6px",
                "borderLeft": f"3px solid {data['color']}",
            }))

        quadrants.append(html.Div([
            html.Div([
                html.Span(data["icon"], style={
                    "display": "inline-block", "width": "28px", "height": "28px",
                    "lineHeight": "28px", "textAlign": "center", "borderRadius": "50%",
                    "backgroundColor": data["color"], "color": "white",
                    "fontWeight": "bold", "fontSize": "14px", "marginRight": "8px",
                }),
                html.Strong(label, style={"fontSize": "18px", "color": FLC_NAVY}),
                html.Span(f"  {len(data['items'])} items", style={
                    "fontSize": "11px", "color": "#888", "marginLeft": "8px",
                }),
            ], style={"marginBottom": "12px"}),
            html.Div(items),
        ], style={**CARD_STYLE, "flex": "1", "minWidth": "420px"}))

    return html.Div([
        html.H2("SWOT Analysis", style=SECTION_TITLE),
        html.P(
            "Phase 2 synthesizes findings from all four Phase 1 frameworks (PESTLE, Porter's Five Forces, "
            "Gray Associates, BCG Matrix) into a unified Strengths-Weaknesses-Opportunities-Threats analysis. "
            "Each item includes source attribution to its originating framework(s).",
            style={
                "fontSize": "13px", "color": "#444", "lineHeight": "1.6",
                "backgroundColor": FLC_BLUE_WASH, "padding": "14px 18px",
                "borderLeft": f"4px solid {FLC_BLUE}", "borderRadius": "4px",
                "marginBottom": "16px",
            },
        ),
        data_source_badge("SWOT Analysis"),
        html.Div([
            html.Button(
                "Download SWOT Matrix (.pptx)",
                id="dl-swot-pptx-btn",
                style={
                    "backgroundColor": FLC_BLUE, "color": "white", "border": "none",
                    "padding": "9px 18px", "borderRadius": "6px", "cursor": "pointer",
                    "fontSize": "12px", "fontWeight": "600", "marginRight": "8px",
                    "transition": "background-color 0.2s ease", "letterSpacing": "0.2px",
                },
            ),
            dcc.Download(id="dl-swot-pptx"),
        ], style={"marginBottom": "16px"}),

        # 2x2 SWOT grid
        html.Div([quadrants[0], quadrants[1]], style={"display": "flex", "gap": "16px", "flexWrap": "wrap"}),
        html.Div([quadrants[2], quadrants[3]], style={"display": "flex", "gap": "16px", "flexWrap": "wrap"}),
    ])


def _build_zone_section(zone_name, zone_data, recommendation_text):
    """Build a zone sub-section with recommendation text and programs table."""
    programs = zone_data["programs"]
    inv_colors = {"High": CLR_HIGH, "Medium": CLR_MEDIUM, "Low": FLC_BLUE}
    program_rows = []
    for pi, p in enumerate(programs):
        row_bg = FLC_BLUE_WASH if pi % 2 == 0 else BG_WHITE
        program_rows.append(html.Tr([
            html.Td(p["name"], style={"fontSize": "12px", "padding": "8px 10px", "fontWeight": "600", "color": FLC_NAVY, "backgroundColor": row_bg}),
            html.Td(p["action"], style={"fontSize": "12px", "padding": "8px 10px", "color": "#4a6070", "backgroundColor": row_bg}),
            html.Td(p["investment"], style={
                "fontSize": "12px", "padding": "8px 10px", "textAlign": "center",
                "color": inv_colors.get(p["investment"], CLR_NEUTRAL), "fontWeight": "700", "backgroundColor": row_bg,
            }),
        ]))
        # Cross-reference commentary row
        xref = ZONE_CROSS_REFERENCES.get(p["name"])
        if xref:
            xref_children = []
            if xref.get("supporting"):
                items = [f'"{f["text"]}" ({f["source"]})' for f in xref["supporting"]]
                xref_children.append(html.Div([
                    html.Span("\u2713 Supporting: ", style={
                        "color": "#276749", "fontWeight": "700", "fontSize": "10px",
                    }),
                    html.Span("; ".join(items), style={
                        "color": "#4a6070", "fontSize": "10px",
                    }),
                ], style={"marginBottom": "3px"}))
            if xref.get("risks"):
                items = [f'"{f["text"]}" ({f["source"]})' for f in xref["risks"]]
                xref_children.append(html.Div([
                    html.Span("\u26A0 Risks: ", style={
                        "color": "#c53030", "fontWeight": "700", "fontSize": "10px",
                    }),
                    html.Span("; ".join(items), style={
                        "color": "#4a6070", "fontSize": "10px",
                    }),
                ]))
            program_rows.append(html.Tr([
                html.Td(xref_children, colSpan=3, style={
                    "padding": "6px 10px 10px 20px", "backgroundColor": "#f8fafc",
                    "borderBottom": f"1px solid {FLC_BLUE_PALE}",
                }),
            ]))

    return html.Div([
        # Zone heading with color dot
        html.Div([
            html.Div(style={
                "display": "inline-block", "width": "10px", "height": "10px",
                "borderRadius": "50%", "backgroundColor": zone_data["color"],
                "marginRight": "8px", "verticalAlign": "middle",
            }),
            html.Strong(zone_name, style={"fontSize": "14px", "color": FLC_NAVY}),
            html.Span(f"  {len(programs)} initiatives", style={
                "fontSize": "11px", "color": "#888", "marginLeft": "8px",
            }),
        ], style={"marginBottom": "6px"}),
        # Scenario-specific recommendation
        html.P(recommendation_text, style={
            "fontSize": "12px", "color": "#444", "lineHeight": "1.5",
            "backgroundColor": "#f8fafc", "padding": "8px 12px",
            "borderLeft": f"3px solid {zone_data['color']}", "borderRadius": "3px",
            "marginBottom": "8px",
        }),
        # Programs table
        html.Table([
            html.Thead(html.Tr([
                html.Th("Program/Initiative", style={"fontSize": "11px", "padding": "8px 10px", "backgroundColor": BG_WHITE, "color": FLC_NAVY,
                                                     "fontWeight": "700", "borderBottom": f"2px solid {FLC_BLUE}", "textTransform": "uppercase", "letterSpacing": "0.5px"}),
                html.Th("Strategic Action", style={"fontSize": "11px", "padding": "8px 10px", "backgroundColor": BG_WHITE, "color": FLC_NAVY,
                                                   "fontWeight": "700", "borderBottom": f"2px solid {FLC_BLUE}", "textTransform": "uppercase", "letterSpacing": "0.5px"}),
                html.Th("Investment", style={"fontSize": "11px", "padding": "8px 10px", "backgroundColor": BG_WHITE, "color": FLC_NAVY,
                                             "fontWeight": "700", "borderBottom": f"2px solid {FLC_BLUE}", "textTransform": "uppercase", "letterSpacing": "0.5px", "textAlign": "center"}),
            ])),
            html.Tbody(program_rows),
        ], style={"width": "100%", "borderCollapse": "collapse", "border": f"1px solid {FLC_BLUE_PALE}"}),
    ], style={
        "borderLeft": f"4px solid {zone_data['color']}",
        "paddingLeft": "12px", "marginBottom": "16px",
    })


def build_zone_to_win_tab():
    """Phase 3: Zone to Win framework — scenarios as top-level organizer, each containing 4 zones."""
    # Zone key: "Performance", "Productivity", etc. (without " Zone" suffix)
    zone_keys = ["Performance", "Productivity", "Incubation", "Transformation"]

    # Build scenario cards, each containing pie chart + 4 zone sub-sections
    scenario_cards = []
    for scenario_name, s_data in SCENARIOS.items():
        # Pie chart for zone allocation
        alloc = s_data["zone_allocation"]
        fig_pie = go.Figure(data=[go.Pie(
            labels=list(alloc.keys()),
            values=list(alloc.values()),
            marker=dict(colors=[
                ZONE_TO_WIN_DATA[f"{z} Zone"]["color"]
                for z in alloc.keys()
            ]),
            hole=0.4, textinfo="label+percent",
        )])
        fig_pie.update_layout(
            template=FLC_CHART_TEMPLATE,
            title=dict(text="Zone Allocation"), height=280,
            margin=dict(l=20, r=20, t=40, b=20), showlegend=False,
        )

        # Scenario metadata badges
        meta_items = [
            ("Strategic Bet", s_data["strategic_bet"]),
            ("Risk Level", s_data["risk_level"]),
            ("Success Probability", s_data["success_probability"]),
            ("Investment Needs", s_data["investment_needs"]),
        ]
        meta_badges = html.Div([
            html.Div([
                html.Span(label + ": ", style={"fontWeight": "700", "fontSize": "11px", "color": FLC_NAVY}),
                html.Span(value, style={"fontSize": "11px", "color": "#4a6070"}),
            ], style={"marginBottom": "4px"})
            for label, value in meta_items
        ], style={"marginTop": "8px"})

        # Zone sub-sections within this scenario
        zone_sections = []
        for zk in zone_keys:
            zone_full = f"{zk} Zone"
            zone_data = ZONE_TO_WIN_DATA[zone_full]
            rec_text = s_data["zone_recommendations"].get(zk, "")
            zone_sections.append(_build_zone_section(zone_full, zone_data, rec_text))

        scenario_cards.append(html.Div([
            # Header row: scenario info (left) + pie chart (right)
            html.Div([
                html.Div([
                    html.Strong(f"{scenario_name} Scenario", style={"fontSize": "17px", "color": s_data["color"]}),
                    html.P(s_data["description"], style={"fontSize": "12px", "color": "#555", "margin": "6px 0"}),
                    html.Strong("Key Assumptions:", style={"fontSize": "12px"}),
                    html.Ul([html.Li(a, style={"fontSize": "11px"}) for a in s_data["assumptions"]],
                            style={"marginTop": "4px", "marginBottom": "4px"}),
                    meta_badges,
                ], style={"flex": "1"}),
                html.Div([
                    dcc.Graph(figure=fig_pie, config={"displayModeBar": False}),
                ], style={"flex": "1", "minWidth": "300px"}),
            ], style={"display": "flex", "gap": "16px", "marginBottom": "16px"}),
            # Zone sub-sections
            html.Hr(style={"border": "none", "borderTop": f"1px solid {FLC_BLUE_PALE}", "margin": "8px 0 16px 0"}),
            html.Div(zone_sections),
        ], style={**CARD_STYLE, "borderLeft": f"4px solid {s_data['color']}"}))

    # Scenario comparison bar chart
    fig_compare = go.Figure()
    metrics = ["enrollment_target", "retention_target", "graduate_target", "online_courses"]
    metric_labels = ["Enrollment", "Retention %", "Graduate Enroll.", "Online Courses"]
    x_labels = list(SCENARIOS.keys())
    for i, (metric, label) in enumerate(zip(metrics, metric_labels)):
        vals = [SCENARIOS[s][metric] for s in x_labels]
        fig_compare.add_trace(go.Bar(
            name=label, x=x_labels, y=vals,
            text=[f"{v:,.0f}" if v > 100 else f"{v}" for v in vals],
            textposition="outside",
        ))
    fig_compare.update_layout(
        template=FLC_CHART_TEMPLATE,
        title=dict(text="Scenario Target Comparison"),
        barmode="group", height=380,
        margin=dict(l=40, r=20, t=50, b=30),
    )

    return html.Div([
        html.H2("Zone to Win", style=SECTION_TITLE),
        html.P(
            "Geoffrey Moore's Zone to Win framework organizes FLC's strategic initiatives into four zones: "
            "Performance (revenue growth), Productivity (operational efficiency), Incubation (emerging opportunities), "
            "and Transformation (future-defining bets). Three scenarios model different resource allocation strategies.",
            style={
                "fontSize": "13px", "color": "#444", "lineHeight": "1.6",
                "backgroundColor": FLC_BLUE_WASH, "padding": "14px 18px",
                "borderLeft": f"4px solid {FLC_BLUE}", "borderRadius": "4px",
                "marginBottom": "16px",
            },
        ),
        data_source_badge("Zone to Win"),

        # Scenario cards (top-level organizer)
        html.H3("Strategic Scenarios", style={**SECTION_TITLE, "fontSize": "18px"}),
        html.Div(scenario_cards),

        # Comparison chart
        html.Div([
            dcc.Graph(figure=fig_compare, config={"displayModeBar": False}),
            source_annotation("Source: Zone to Win methodology (Geoffrey Moore) applied to FLC strategic context"),
        ], style=CARD_STYLE),
    ])


def build_roadmap_tab():
    """Phase 3: Strategic Roadmap — simplified view with risk assessment first, then implementation overview."""
    scenario = SCENARIOS["Moderate-Adaptive"]

    # ── Risk Assessment Matrix (kept & improved) ──
    risk_df = RISK_MITIGATION.copy()
    prob_map = {"Low": 1, "Medium": 2, "High": 3}
    impact_map = {"Low": 1, "Medium": 2, "High": 3, "Critical": 4}
    risk_df["Prob_Num"] = risk_df["Probability"].map(prob_map)
    risk_df["Impact_Num"] = risk_df["Impact"].map(impact_map)
    risk_df["Risk_Score"] = risk_df["Prob_Num"] * risk_df["Impact_Num"]

    # Jitter overlapping points slightly for readability
    seen = {}
    jitter_x, jitter_y = [], []
    for _, row in risk_df.iterrows():
        key = (row["Prob_Num"], row["Impact_Num"])
        offset = seen.get(key, 0)
        seen[key] = offset + 1
        jitter_x.append(row["Prob_Num"] + offset * 0.12)
        jitter_y.append(row["Impact_Num"] + offset * 0.08)
    risk_df["Jitter_X"] = jitter_x
    risk_df["Jitter_Y"] = jitter_y

    fig_risk = go.Figure(data=go.Scatter(
        x=risk_df["Jitter_X"], y=risk_df["Jitter_Y"],
        mode="markers+text",
        marker=dict(
            size=risk_df["Risk_Score"] * 8 + 10,
            color=risk_df["Risk_Score"],
            colorscale=[[0, "#8cc0e8"], [0.5, FLC_BLUE], [1, FLC_NAVY]],
            showscale=True, colorbar=dict(title="Risk Score", tickfont=dict(color=FLC_NAVY)),
            opacity=0.85,
            line=dict(width=1, color="white"),
        ),
        text=risk_df["Risk"].str[:25],
        textposition=["top center" if i % 2 == 0 else "bottom center" for i in range(len(risk_df))],
        textfont=dict(size=7, color=FLC_NAVY),
        hovertext=risk_df.apply(
            lambda r: f"<b>{r['Risk']}</b><br>Probability: {r['Probability']} | Impact: {r['Impact']}"
                      f"<br><br>Mitigation: {r['Mitigation_Strategy']}",
            axis=1,
        ),
        hoverinfo="text",
    ))
    fig_risk.update_layout(
        template=FLC_CHART_TEMPLATE,
        title=dict(text="Risk Assessment Matrix"),
        xaxis=dict(title="Probability", tickvals=[1, 2, 3], ticktext=["Low", "Medium", "High"], range=[0.5, 3.5]),
        yaxis=dict(title="Impact", tickvals=[1, 2, 3, 4], ticktext=["Low", "Medium", "High", "Critical"], range=[0.5, 4.5]),
        height=450,
        margin=dict(l=60, r=30, t=50, b=50),
    )
    # Background shading: low-risk zone (bottom-left), high-risk zone (top-right), critical zone (top rows)
    fig_risk.add_shape(type="rect", x0=0.5, y0=0.5, x1=1.5, y1=2.5,
                       fillcolor="rgba(140,192,232,0.08)", line_width=0)
    fig_risk.add_shape(type="rect", x0=1.5, y0=2.5, x1=3.5, y1=3.5,
                       fillcolor="rgba(0,48,87,0.08)", line_width=0)
    fig_risk.add_shape(type="rect", x0=0.5, y0=3.5, x1=3.5, y1=4.5,
                       fillcolor="rgba(0,48,87,0.12)", line_width=0)

    # ── Scenario 2 summary metrics ──
    target_style = {"textAlign": "center", "flex": "1", "padding": "12px 8px"}
    scenario_targets = html.Div([
        html.Div([
            html.Div("Enrollment", style={"fontSize": "10px", "color": FLC_BLUE, "textTransform": "uppercase", "fontWeight": "700", "letterSpacing": "1px"}),
            html.Div(f"{scenario['enrollment_target']:,}", style={"fontSize": "28px", "fontWeight": "800", "color": FLC_NAVY}),
            html.Div("students", style={"fontSize": "11px", "color": "#6b8299"}),
        ], style=target_style),
        html.Div([
            html.Div("Retention", style={"fontSize": "10px", "color": FLC_BLUE, "textTransform": "uppercase", "fontWeight": "700", "letterSpacing": "1px"}),
            html.Div(f"{scenario['retention_target']}%", style={"fontSize": "28px", "fontWeight": "800", "color": FLC_NAVY}),
            html.Div("FTFT", style={"fontSize": "11px", "color": "#6b8299"}),
        ], style=target_style),
        html.Div([
            html.Div("Graduate", style={"fontSize": "10px", "color": FLC_BLUE, "textTransform": "uppercase", "fontWeight": "700", "letterSpacing": "1px"}),
            html.Div(f"{scenario['graduate_target']}", style={"fontSize": "28px", "fontWeight": "800", "color": FLC_NAVY}),
            html.Div("students", style={"fontSize": "11px", "color": "#6b8299"}),
        ], style=target_style),
        html.Div([
            html.Div("Online Courses", style={"fontSize": "10px", "color": FLC_BLUE, "textTransform": "uppercase", "fontWeight": "700", "letterSpacing": "1px"}),
            html.Div(f"{scenario['online_courses']}", style={"fontSize": "28px", "fontWeight": "800", "color": FLC_NAVY}),
            html.Div("courses", style={"fontSize": "11px", "color": "#6b8299"}),
        ], style=target_style),
        html.Div([
            html.Div("New Programs", style={"fontSize": "10px", "color": FLC_BLUE, "textTransform": "uppercase", "fontWeight": "700", "letterSpacing": "1px"}),
            html.Div(f"{scenario['new_programs']}", style={"fontSize": "28px", "fontWeight": "800", "color": FLC_NAVY}),
            html.Div("programs", style={"fontSize": "11px", "color": "#6b8299"}),
        ], style=target_style),
    ], style={"display": "flex", "gap": "8px"})

    # ── High-level implementation timeline (plain HTML table) ──
    tl_cell = {"padding": "10px 12px", "fontSize": "12px", "color": FLC_NAVY,
               "borderBottom": f"1px solid {FLC_BLUE_PALE}", "lineHeight": "1.6", "verticalAlign": "top"}
    tl_hdr = {**TABLE_HEADER_STYLE, "textAlign": "left", "padding": "10px 12px"}

    timeline_table = html.Table([
        html.Thead(html.Tr([
            html.Th("Timeframe", style={**tl_hdr, "width": "14%"}),
            html.Th("Focus", style={**tl_hdr, "width": "16%"}),
            html.Th("Key Actions", style={**tl_hdr, "width": "50%"}),
            html.Th("Primary Zones", style={**tl_hdr, "width": "20%"}),
        ])),
        html.Tbody([
            html.Tr([
                html.Td("Year 1 (2026)", style={**tl_cell, "fontWeight": "700"}),
                html.Td("Foundation", style=tl_cell),
                html.Td([
                    "Program sunset reviews for 17 Concern-quadrant majors. ",
                    "Retention pilot expansion (Compass, early-alert system). ",
                    "Advising redesign per NACADA recommendations. ",
                    "Faculty recruitment package (Durango housing incentive). ",
                    "Dual enrollment expansion (3+ high schools).",
                ], style=tl_cell),
                html.Td("Performance, Productivity", style=tl_cell),
            ], style={"backgroundColor": FLC_BLUE_WASH}),
            html.Tr([
                html.Td("Year 2 (2027)", style={**tl_cell, "fontWeight": "700"}),
                html.Td("Selective Investment", style=tl_cell),
                html.Td([
                    "Indigenous Education Hub feasibility complete; launch decision. ",
                    "AI Institute partnership MOUs and grant applications. ",
                    "Workforce certificate feasibility (regional demand analysis). ",
                    "Program restructuring through faculty governance (12\u201318 mo process). ",
                    "Budget reallocation based on Year 1 Zone performance.",
                ], style=tl_cell),
                html.Td("Incubation, Transformation", style=tl_cell),
            ]),
            html.Tr([
                html.Td("Year 3 (2028)", style={**tl_cell, "fontWeight": "700"}),
                html.Td("Differentiation", style=tl_cell),
                html.Td([
                    "Indigenous Education Hub operational (statutory/sovereign framing). ",
                    "First graduate certificate enrollment (existing program area). ",
                    "Experiential learning brand formalized. ",
                    "Portfolio rebalanced: faculty lines aligned to growing programs.",
                ], style=tl_cell),
                html.Td("Transformation, Performance", style=tl_cell),
            ], style={"backgroundColor": FLC_BLUE_WASH}),
            html.Tr([
                html.Td("Years 4\u20135 (2029\u201330)", style={**tl_cell, "fontWeight": "700"}),
                html.Td("Maturation", style=tl_cell),
                html.Td([
                    "Indigenous online niche scaled (if pilot successful + marketing investment secured). ",
                    "Workforce credentials aligned with regional employers. ",
                    "Full portfolio optimized around Grow/Sustain programs. ",
                    "Place-based experiential learning model nationally recognized.",
                ], style=tl_cell),
                html.Td("Transformation, Performance", style=tl_cell),
            ]),
        ]),
    ], style={"width": "100%", "borderCollapse": "collapse", "border": f"1px solid {FLC_BLUE_PALE}"})

    # ── Zone allocation visual ──
    za = scenario["zone_allocation"]
    zone_colors = {"Performance": FLC_NAVY, "Productivity": FLC_BLUE,
                   "Incubation": FLC_BLUE_LIGHT, "Transformation": "#5ba3d9"}
    zone_bar = html.Div([
        html.Div(f"{zone} {pct}%", style={
            "flex": str(pct), "backgroundColor": zone_colors.get(zone, "#999"),
            "color": "white", "textAlign": "center", "padding": "8px 4px",
            "fontSize": "11px", "fontWeight": "600",
        }) for zone, pct in za.items()
    ], style={"display": "flex", "borderRadius": "6px", "overflow": "hidden", "marginTop": "8px"})

    return html.Div([
        html.H2("Strategic Roadmap", style=SECTION_TITLE),
        html.P(
            "Implementation plan based on the Moderate-Adaptive scenario \u2014 selective investment in "
            "differentiated strengths while protecting core programs. Risk assessment synthesized "
            "from PESTLE, Porter's, BCG (48 majors), Gray Associates, and SWOT analyses.",
            style={
                "fontSize": "13px", "color": "#444", "lineHeight": "1.6",
                "backgroundColor": FLC_BLUE_WASH, "padding": "14px 18px",
                "borderLeft": f"4px solid {FLC_BLUE}", "borderRadius": "4px",
                "marginBottom": "16px",
            },
        ),

        # ── RISK ASSESSMENT (top) ──
        html.H3("Risk Assessment & Mitigation", style={**SECTION_TITLE, "fontSize": "16px"}),
        html.Div([
            dcc.Graph(figure=fig_risk, config={"displayModeBar": False}),
            source_annotation("Source: Risk analysis synthesized from all Phase 1 and Phase 2 framework analyses"),
        ], style=CARD_STYLE),

        html.Div([dash_table.DataTable(
            data=risk_df[["Risk", "Probability", "Impact", "Mitigation_Strategy", "Owner"]].to_dict("records"),
            columns=[
                {"name": "Risk", "id": "Risk"},
                {"name": "Probability", "id": "Probability"},
                {"name": "Impact", "id": "Impact"},
                {"name": "Mitigation Strategy", "id": "Mitigation_Strategy"},
                {"name": "Owner", "id": "Owner"},
            ],
            style_cell={**TABLE_CELL_STYLE, "textAlign": "left", "whiteSpace": "normal", "height": "auto"},
            style_header=TABLE_HEADER_STYLE,
            style_data_conditional=TABLE_ROW_BANDING + [
                {"if": {"filter_query": '{Impact} = "Critical"', "column_id": "Impact"},
                 "color": "#8b0000", "fontWeight": "bold", "backgroundColor": "#fde8e8"},
                {"if": {"filter_query": '{Impact} = "High"', "column_id": "Impact"},
                 "color": FLC_NAVY, "fontWeight": "bold", "backgroundColor": "#e8f0f8"},
                {"if": {"filter_query": '{Probability} = "High"', "column_id": "Probability"},
                 "color": FLC_NAVY, "fontWeight": "bold", "backgroundColor": "#e8f0f8"},
            ],
            sort_action="native",
        )], style=CARD_STYLE),

        # ── IMPLEMENTATION OVERVIEW ──
        html.H3("Implementation Overview \u2014 Moderate-Adaptive Scenario", style={**SECTION_TITLE, "fontSize": "16px", "marginTop": "24px"}),
        html.Div([
            html.P(scenario["description"],
                   style={"fontSize": "13px", "color": "#4a6070", "lineHeight": "1.7", "marginBottom": "12px"}),
            scenario_targets,
            html.Div([
                html.Div("Investment Allocation by Zone", style={"fontSize": "11px", "color": FLC_BLUE,
                          "fontWeight": "700", "textTransform": "uppercase", "letterSpacing": "1px", "marginTop": "16px"}),
                zone_bar,
            ]),
        ], style=CARD_STYLE),

        html.Div([
            html.P("Key Assumptions:", style={"fontSize": "13px", "color": FLC_NAVY, "fontWeight": "700", "marginBottom": "6px"}),
            html.Ul([html.Li(a, style={"fontSize": "12px", "color": "#4a6070", "marginBottom": "4px", "lineHeight": "1.6"})
                     for a in scenario["assumptions"]]),
        ], style={**CARD_STYLE, "padding": "16px"}),

        # ── HIGH-LEVEL TIMELINE ──
        html.H3("Implementation Timeline (2026\u20132030)", style={**SECTION_TITLE, "fontSize": "16px"}),
        html.Div([timeline_table], style=CARD_STYLE),
        source_annotation("Source: Implementation plan derived from Zone to Win Moderate-Adaptive scenario + all Phase 1\u20132 analyses"),
    ])


# ============================================================================
# MAIN LAYOUT
# ============================================================================

app.layout = html.Div([
    # Header with mountain silhouette gradient
    html.Div([
        html.Div([
            html.H1("Fort Lewis College", style={
                "color": "white", "margin": "0", "fontSize": "26px", "fontWeight": "800",
                "letterSpacing": "0.5px", "textShadow": "0 1px 3px rgba(0,0,0,0.2)",
            }),
            html.Div("Portfolio Optimization Dashboard", style={
                "color": FLC_BLUE_PALE, "fontSize": "14px", "fontWeight": "500",
                "marginTop": "2px", "letterSpacing": "0.3px",
            }),
        ], style={"flex": "1", "position": "relative", "zIndex": "10"}),
        html.Div([
            html.Div("Phase 1: Environmental Scanning", style={
                "color": "rgba(255,255,255,0.7)", "fontSize": "11px",
            }),
            html.Div("Phase 2: Strategic Synthesis", style={
                "color": "rgba(255,255,255,0.7)", "fontSize": "11px",
            }),
            html.Div("Phase 3: Strategic Direction", style={
                "color": "rgba(255,255,255,0.7)", "fontSize": "11px",
            }),
        ], style={"textAlign": "right", "position": "relative", "zIndex": "10"}),
    ], className="flc-header", style={
        "padding": "20px 28px 36px 28px",
        "display": "flex", "alignItems": "center", "justifyContent": "space-between",
    }),

    # Tab navigation
    dcc.Tabs(id="main-tabs", value="summary", children=[
        dcc.Tab(label="Executive Summary", value="summary", style=TAB_STYLE, selected_style=TAB_SELECTED),
        dcc.Tab(label="PESTLE Analysis", value="pestle", style=TAB_STYLE, selected_style=TAB_SELECTED),
        dcc.Tab(label="Porter's Analysis", value="porters", style=TAB_STYLE, selected_style=TAB_SELECTED),
        dcc.Tab(label="Gray Analysis", value="gray", style=TAB_STYLE, selected_style=TAB_SELECTED),
        dcc.Tab(label="BCG Analysis", value="bcg", style=TAB_STYLE, selected_style=TAB_SELECTED),
        dcc.Tab(label="SWOT Analysis", value="swot", style=TAB_STYLE, selected_style=TAB_SELECTED),
        dcc.Tab(label="Zone to Win", value="zonetowin", style=TAB_STYLE, selected_style=TAB_SELECTED),
        dcc.Tab(label="Strategic Roadmap", value="roadmap", style=TAB_STYLE, selected_style=TAB_SELECTED),
    ], style={"marginBottom": "0", "backgroundColor": BG_WHITE, "borderBottom": f"1px solid {FLC_BORDER}"}),

    # Tab content
    html.Div(id="tab-content", style={
        "padding": "24px 28px",
        "backgroundColor": FLC_LIGHT,
        "minHeight": "calc(100vh - 180px)",
        "position": "relative",
        "zIndex": "1",
    }),

    # Footer with mountain silhouette
    html.Div([
        html.Div([
            html.Span("Fort Lewis College Portfolio Optimization Project", style={
                "fontWeight": "600", "color": "rgba(255,255,255,0.9)",
            }),
            html.Span(" | ", style={"color": "rgba(255,255,255,0.4)"}),
            html.Span("PESTLE  |  Porter's  |  Gray Associates  |  BCG  |  SWOT  |  Zone to Win  |  Roadmap", style={
                "color": "rgba(255,255,255,0.6)", "fontSize": "10px",
            }),
        ], style={"position": "relative", "zIndex": "10"}),
    ], className="flc-footer", style={
        "backgroundColor": FLC_NAVY, "color": "#aaa", "padding": "20px 28px 14px 28px",
        "fontSize": "11px", "textAlign": "center",
    }),
], className="flc-topo-bg", style={
    "fontFamily": "'Segoe UI', -apple-system, BlinkMacSystemFont, Tahoma, Geneva, Verdana, sans-serif",
})


# ============================================================================
# CALLBACKS
# ============================================================================

@app.callback(
    Output("tab-content", "children"),
    Input("main-tabs", "value"),
)
def render_tab(tab):
    if tab == "summary":
        return build_summary_page()
    elif tab == "pestle":
        return build_pestle_tab()
    elif tab == "porters":
        return build_porters_tab()
    elif tab == "gray":
        return build_gray_tab()
    elif tab == "bcg":
        return build_bcg_tab()
    elif tab == "swot":
        return build_swot_tab()
    elif tab == "zonetowin":
        return build_zone_to_win_tab()
    elif tab == "roadmap":
        return build_roadmap_tab()
    return html.Div("Select a tab")


# Download callbacks for Phase 1 document generation
@app.callback(
    Output({"type": "dl-target", "name": "PESTLE-docx"}, "data"),
    Input({"type": "dl-btn", "name": "PESTLE-docx"}, "n_clicks"),
    prevent_initial_call=True,
)
def dl_pestle_docx(n):
    return dcc.send_file(os.path.join(GENERATED_DOCS_DIR, "PESTLE_Executive_Summary.docx"))

@app.callback(
    Output({"type": "dl-target", "name": "PESTLE-pptx"}, "data"),
    Input({"type": "dl-btn", "name": "PESTLE-pptx"}, "n_clicks"),
    prevent_initial_call=True,
)
def dl_pestle_pptx(n):
    return dcc.send_file(os.path.join(GENERATED_DOCS_DIR, "PESTLE_Slide_Deck.pptx"))

@app.callback(
    Output({"type": "dl-target", "name": "Porters-docx"}, "data"),
    Input({"type": "dl-btn", "name": "Porters-docx"}, "n_clicks"),
    prevent_initial_call=True,
)
def dl_porters_docx(n):
    return dcc.send_file(os.path.join(GENERATED_DOCS_DIR, "Porters_Executive_Summary.docx"))

@app.callback(
    Output({"type": "dl-target", "name": "Porters-pptx"}, "data"),
    Input({"type": "dl-btn", "name": "Porters-pptx"}, "n_clicks"),
    prevent_initial_call=True,
)
def dl_porters_pptx(n):
    return dcc.send_file(os.path.join(GENERATED_DOCS_DIR, "Porters_Slide_Deck.pptx"))

@app.callback(
    Output({"type": "dl-target", "name": "Gray-docx"}, "data"),
    Input({"type": "dl-btn", "name": "Gray-docx"}, "n_clicks"),
    prevent_initial_call=True,
)
def dl_gray_docx(n):
    return dcc.send_file(os.path.join(GENERATED_DOCS_DIR, "Gray_Executive_Summary.docx"))

@app.callback(
    Output({"type": "dl-target", "name": "Gray-pptx"}, "data"),
    Input({"type": "dl-btn", "name": "Gray-pptx"}, "n_clicks"),
    prevent_initial_call=True,
)
def dl_gray_pptx(n):
    return dcc.send_file(os.path.join(GENERATED_DOCS_DIR, "Gray_Slide_Deck.pptx"))

@app.callback(
    Output({"type": "dl-target", "name": "BCG-docx"}, "data"),
    Input({"type": "dl-btn", "name": "BCG-docx"}, "n_clicks"),
    prevent_initial_call=True,
)
def dl_bcg_docx(n):
    return dcc.send_file(os.path.join(GENERATED_DOCS_DIR, "BCG_Executive_Summary.docx"))

@app.callback(
    Output({"type": "dl-target", "name": "BCG-pptx"}, "data"),
    Input({"type": "dl-btn", "name": "BCG-pptx"}, "n_clicks"),
    prevent_initial_call=True,
)
def dl_bcg_pptx(n):
    return dcc.send_file(os.path.join(GENERATED_DOCS_DIR, "BCG_Slide_Deck.pptx"))


@app.callback(
    Output("dl-swot-pptx", "data"),
    Input("dl-swot-pptx-btn", "n_clicks"),
    prevent_initial_call=True,
)
def dl_swot_pptx(n):
    path = generate_swot_pptx()
    return dcc.send_file(path)


# --- Project-level deliverable downloads (generated on-demand) ---

@app.callback(
    Output("dl-exec-summary-docx", "data"),
    Input("dl-exec-summary-docx-btn", "n_clicks"),
    prevent_initial_call=True,
)
def dl_exec_summary_docx(n):
    return dcc.send_file(generate_exec_summary_docx())


@app.callback(
    Output("dl-exec-summary-pptx", "data"),
    Input("dl-exec-summary-pptx-btn", "n_clicks"),
    prevent_initial_call=True,
)
def dl_exec_summary_pptx(n):
    return dcc.send_file(generate_exec_summary_pptx())


@app.callback(
    Output("dl-final-report-docx", "data"),
    Input("dl-final-report-docx-btn", "n_clicks"),
    prevent_initial_call=True,
)
def dl_final_report_docx(n):
    return dcc.send_file(generate_final_report_docx())


@app.callback(
    Output("dl-final-report-pptx", "data"),
    Input("dl-final-report-pptx-btn", "n_clicks"),
    prevent_initial_call=True,
)
def dl_final_report_pptx(n):
    return dcc.send_file(generate_final_report_pptx())


# ============================================================================
# RUN
# ============================================================================

def _kill_stale_process(port):
    """Kill any existing process listening on the given port (Windows only)."""
    import subprocess, re
    try:
        result = subprocess.run(["netstat", "-ano"], capture_output=True, text=True)
        for line in result.stdout.splitlines():
            if f"0.0.0.0:{port}" in line and "LISTENING" in line:
                pid = line.strip().split()[-1]
                if pid.isdigit() and int(pid) != os.getpid():
                    subprocess.run(["taskkill", "/F", "/PID", pid],
                                   capture_output=True)
                    print(f"  Killed stale process on port {port} (PID {pid})")
    except Exception:
        pass  # Non-Windows or netstat unavailable — skip silently


if __name__ == "__main__":
    PORT = 8080
    _kill_stale_process(PORT)
    print("\n" + "=" * 60)
    print("  FLC Portfolio Optimization Dashboard")
    print("  Phase 1: Environmental Scanning")
    print("  Phase 2: Strategic Synthesis (SWOT)")
    print("  Phase 3: Strategic Direction (Zone to Win + Roadmap)")
    print(f"  Open http://127.0.0.1:{PORT} in your browser")
    print("=" * 60 + "\n")
    app.run(host='0.0.0.0', port=PORT, debug=False)
