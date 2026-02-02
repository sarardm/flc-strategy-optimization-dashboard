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
    BCG_DATA, BCG_QUADRANT_COLORS, BCG_INSIGHTS,
    PESTLE_DATA,
    PORTERS_DATA, PORTERS_INSIGHTS,
    GRAY_ASSOCIATES_DATA, GA_RECOMMENDATION_COLORS, GA_INSIGHTS,
    STRATEGIC_INITIATIVES, MILESTONES, KPIS, RESOURCE_ALLOCATION,
    DATA_SOURCES, FRAMEWORK_DESCRIPTIONS,
    SWOT_DATA, ZONE_TO_WIN_DATA, SCENARIOS,
    ROADMAP_MILESTONES, ROADMAP_KPIS, RISK_MITIGATION,
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

# FLC brand colors
FLC_NAVY = "#003057"
FLC_BLUE = "#0066b3"
FLC_GOLD = "#c8a415"
FLC_LIGHT = "#f0f4f8"
BG_WHITE = "#ffffff"

# Shared style constants
CARD_STYLE = {
    "backgroundColor": BG_WHITE,
    "borderRadius": "8px",
    "padding": "20px",
    "marginBottom": "16px",
    "boxShadow": "0 2px 4px rgba(0,0,0,0.08)",
    "border": "1px solid #e2e8f0",
}

SECTION_TITLE = {
    "color": FLC_NAVY,
    "fontSize": "20px",
    "fontWeight": "bold",
    "marginBottom": "12px",
    "borderBottom": f"3px solid {FLC_GOLD}",
    "paddingBottom": "8px",
}

TAB_STYLE = {"fontWeight": "600"}
TAB_SELECTED = {"fontWeight": "700", "borderTop": f"3px solid {FLC_GOLD}"}

# ============================================================================
# HELPERS
# ============================================================================

def data_source_badge(framework_name):
    """Render a data-source attribution badge."""
    info = DATA_SOURCES.get(framework_name, {})
    source = info.get("source", "Unknown")
    color = info.get("badge_color", "#999")
    return html.Div([
        html.Span("DATA SOURCE: ", style={"fontWeight": "bold", "fontSize": "11px"}),
        html.Span(source, style={
            "backgroundColor": color, "color": "white", "padding": "2px 8px",
            "borderRadius": "10px", "fontSize": "11px", "fontWeight": "600",
        }),
        html.Span(
            f"  ({', '.join(info.get('files', []))})",
            style={"fontSize": "11px", "color": "#666", "marginLeft": "8px"},
        ),
    ], style={"marginBottom": "12px"})


def framework_description_block(key):
    """Render a 2-3 sentence framework description at the top of a Phase 1 tab."""
    text = FRAMEWORK_DESCRIPTIONS.get(key, "")
    return html.Div([
        html.P(text, style={
            "fontSize": "13px", "color": "#444", "lineHeight": "1.6",
            "backgroundColor": "#f8fafc", "padding": "12px 16px",
            "borderLeft": f"4px solid {FLC_GOLD}", "borderRadius": "4px",
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
        "backgroundColor": FLC_NAVY, "color": "white", "border": "none",
        "padding": "8px 16px", "borderRadius": "6px", "cursor": "pointer",
        "fontSize": "12px", "fontWeight": "600", "marginRight": "8px",
    }
    btn_disabled_style = {**btn_style, "backgroundColor": "#ccc", "cursor": "not-allowed"}

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
        "fontSize": "10px", "color": "#999", "fontStyle": "italic",
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
    for title, value, sub, trend in quick_stats:
        kpi_cards.append(html.Div([
            html.Div(title, style={"fontSize": "12px", "color": "#666", "textTransform": "uppercase"}),
            html.Div(value, style={"fontSize": "32px", "fontWeight": "bold", "color": FLC_NAVY}),
            html.Div(sub, style={"fontSize": "12px", "color": "#888"}),
            html.Div(trend, style={"fontSize": "12px", "color": FLC_BLUE, "fontWeight": "600"}),
        ], style={**CARD_STYLE, "textAlign": "center", "flex": "1", "minWidth": "180px"}))

    # Enrollment trend mini chart
    fig_enroll = go.Figure()
    fig_enroll.add_trace(go.Scatter(
        x=ENROLLMENT_HISTORY["Year"], y=ENROLLMENT_HISTORY["Total_Headcount"],
        mode="lines+markers", line=dict(color=FLC_NAVY, width=3),
        marker=dict(size=8), name="Total Headcount",
    ))
    fig_enroll.update_layout(
        title="10-Year Enrollment Trend", height=300,
        margin=dict(l=40, r=20, t=40, b=30),
        template="plotly_white", showlegend=False,
    )

    # 3-phase overview cards
    phase_cards = []
    phases = [
        ("Phase 1: Environmental Scanning", "#2ecc71",
         "PESTLE, Porter's Five Forces, Gray Associates Portfolio Analysis, and BCG Growth-Share Matrix provide comprehensive external and internal scanning.",
         "4 frameworks completed"),
        ("Phase 2: Strategic Synthesis", "#3498db",
         "SWOT Analysis synthesizes all Phase 1 findings into actionable Strengths, Weaknesses, Opportunities, and Threats with source attribution.",
         "Cross-framework synthesis"),
        ("Phase 3: Strategic Direction", "#8e44ad",
         "Zone to Win framework with 3 strategic scenarios and a detailed Strategic Roadmap with KPIs, milestones, and risk mitigation.",
         "Implementation planning"),
    ]
    for name, color, desc, badge in phases:
        phase_cards.append(html.Div([
            html.Div([
                html.Strong(name, style={"fontSize": "15px", "color": FLC_NAVY}),
                html.Span(f"  {badge}", style={
                    "backgroundColor": color, "color": "white",
                    "padding": "2px 8px", "borderRadius": "10px",
                    "fontSize": "10px", "marginLeft": "8px",
                }),
            ]),
            html.P(desc, style={"fontSize": "13px", "color": "#444", "marginTop": "6px", "marginBottom": "0"}),
        ], style={**CARD_STYLE, "padding": "14px", "borderLeft": f"4px solid {color}"}))

    # Framework highlight summaries
    framework_summaries = []
    fw_data = [
        ("PESTLE Analysis", "Internal FLC Documents",
         "Economic and Social factors rated highest impact. Key risk: state funding volatility. Key opportunity: AI Institute and sustainability leadership."),
        ("BCG Analysis", "Internal FLC Documents",
         "2 Stars (Business Admin, Psychology), 9 Cash Cows generating bulk SCH, 2 Question Marks with potential, 9 Concern programs needing action."),
        ("Porter's Analysis", "Internet Methodology + FLC Data",
         "Overall competitive intensity: HIGH. Strongest defense: unique Native American mission and outdoor lifestyle. Greatest threat: online competition and price sensitivity."),
        ("Gray Analysis", "Internet Methodology + FLC Data",
         "7 programs recommended to GROW, 7 to SUSTAIN, 2 to TRANSFORM, 4 to EVALUATE, 2 for SUNSET REVIEW."),
    ]
    for name, source, summary in fw_data:
        badge_color = "#2ecc71" if "Internal" in source else "#3498db"
        framework_summaries.append(html.Div([
            html.Div([
                html.Strong(name, style={"fontSize": "14px", "color": FLC_NAVY}),
                html.Span(f"  {source}", style={
                    "backgroundColor": badge_color, "color": "white",
                    "padding": "1px 6px", "borderRadius": "8px",
                    "fontSize": "10px", "marginLeft": "8px",
                }),
            ]),
            html.P(summary, style={"fontSize": "13px", "color": "#444", "marginTop": "6px", "marginBottom": "0"}),
        ], style={**CARD_STYLE, "padding": "14px"}))

    # Retention trend mini chart
    fig_retention = go.Figure()
    fig_retention.add_trace(go.Scatter(
        x=RETENTION_HISTORY["Year"], y=RETENTION_HISTORY["Retention_Rate"],
        mode="lines+markers", line=dict(color=FLC_GOLD, width=3),
        marker=dict(size=8), name="Retention Rate",
    ))
    fig_retention.add_hline(y=73, line_dash="dash", line_color="#e74c3c", line_width=1,
                            annotation_text="National Avg (73%)", annotation_position="right")
    fig_retention.update_layout(
        title="FTFT Retention Rate Trend", height=300,
        margin=dict(l=40, r=20, t=40, b=30),
        template="plotly_white", showlegend=False,
        yaxis_range=[50, 80],
    )

    return html.Div([
        html.H2("Executive Summary", style={**SECTION_TITLE, "fontSize": "24px"}),
        html.P(
            f"Fort Lewis College Portfolio Optimization Project | Updated {datetime.now().strftime('%B %d, %Y')}",
            style={"color": "#666", "marginBottom": "16px"},
        ),

        # KPI row
        html.Div(kpi_cards, style={"display": "flex", "gap": "12px", "flexWrap": "wrap", "marginBottom": "16px"}),

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
        fillcolor="rgba(0,48,87,0.15)",
        line=dict(color=FLC_NAVY, width=2),
        marker=dict(size=8, color=FLC_GOLD),
    ))
    fig_radar.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0, 5], tickvals=[1, 2, 3, 4, 5])),
        title="PESTLE Impact Assessment (1-5 scale)",
        height=400, margin=dict(l=60, r=60, t=50, b=30),
    )

    impact_colors = {"High": "#e74c3c", "Medium": "#f39c12", "Low": "#2ecc71"}
    fig_bar = go.Figure(data=[go.Bar(
        x=categories, y=scores,
        marker_color=[impact_colors.get(PESTLE_DATA[c]["impact"], "#999") for c in categories],
        text=[PESTLE_DATA[c]["impact"] for c in categories],
        textposition="outside",
    )])
    fig_bar.update_layout(
        title="Impact Level by PESTLE Category",
        yaxis_title="Impact Score", yaxis_range=[0, 6],
        height=350, margin=dict(l=40, r=20, t=50, b=30),
        template="plotly_white",
    )

    detail_cards = []
    for cat in categories:
        d = PESTLE_DATA[cat]
        trend_color = {"Negative": "#e74c3c", "Mixed": "#f39c12",
                       "Stable": "#3498db", "Opportunity": "#2ecc71"}.get(d["trend"], "#999")
        detail_cards.append(html.Div([
            html.Div([
                html.Strong(cat, style={"fontSize": "16px", "color": FLC_NAVY}),
                html.Span(f"  Impact: {d['impact']}", style={
                    "backgroundColor": impact_colors.get(d["impact"], "#999"),
                    "color": "white", "padding": "2px 8px", "borderRadius": "10px",
                    "fontSize": "11px", "marginLeft": "8px",
                }),
                html.Span(f"  Trend: {d['trend']}", style={
                    "backgroundColor": trend_color,
                    "color": "white", "padding": "2px 8px", "borderRadius": "10px",
                    "fontSize": "11px", "marginLeft": "4px",
                }),
            ]),
            html.Div([
                html.Strong("Key Factors:", style={"fontSize": "12px"}),
                html.Ul([html.Li(f, style={"fontSize": "12px"}) for f in d["factors"]],
                        style={"marginTop": "4px", "marginBottom": "4px"}),
            ], style={"marginTop": "8px"}),
            html.Div([
                html.Strong("Opportunities:", style={"fontSize": "12px", "color": "#2ecc71"}),
                html.Ul([html.Li(o, style={"fontSize": "12px"}) for o in d["opportunities"]],
                        style={"marginTop": "4px"}),
            ]),
        ], style={**CARD_STYLE, "padding": "14px"}))

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
        fillcolor="rgba(0,102,179,0.15)",
        line=dict(color=FLC_BLUE, width=2),
        marker=dict(size=8, color=FLC_GOLD),
    ))
    fig_radar.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0, 5], tickvals=[1, 2, 3, 4, 5],
                                    ticktext=["1-Low", "2", "3-Med", "4", "5-High"])),
        title="Porter's Five Forces - Competitive Intensity",
        height=420, margin=dict(l=80, r=80, t=50, b=30),
    )

    fig_bar = go.Figure(data=[go.Bar(
        y=forces, x=scores, orientation="h",
        marker_color=[PORTERS_DATA[f]["color"] for f in forces],
        text=[PORTERS_DATA[f]["rating"] for f in forces],
        textposition="outside",
    )])
    fig_bar.update_layout(
        title="Force Intensity Ratings",
        xaxis_title="Intensity (1=Low, 5=High)", xaxis_range=[0, 5.5],
        height=350, margin=dict(l=180, r=40, t=50, b=30),
        template="plotly_white",
    )

    force_cards = []
    for force in forces:
        d = PORTERS_DATA[force]
        ind_rows = []
        for ind in d["indicators"]:
            trend_icon = {"Increasing": "^", "Decreasing": "v", "Stable": "-", "Improving": "^"}.get(ind["trend"], "?")
            trend_color = {"Increasing": "#e74c3c", "Decreasing": "#2ecc71",
                           "Stable": "#3498db", "Improving": "#2ecc71"}.get(ind["trend"], "#999")
            ind_rows.append(html.Tr([
                html.Td(ind["name"], style={"fontSize": "12px", "padding": "4px 8px"}),
                html.Td(ind["value"], style={"fontSize": "12px", "padding": "4px 8px", "fontWeight": "600"}),
                html.Td(f"{trend_icon} {ind['trend']}", style={
                    "fontSize": "12px", "padding": "4px 8px", "color": trend_color, "fontWeight": "600",
                }),
            ]))
        force_cards.append(html.Div([
            html.Div([
                html.Strong(force, style={"fontSize": "15px", "color": FLC_NAVY}),
                html.Span(f"  {d['rating']}", style={
                    "backgroundColor": d["color"], "color": "white",
                    "padding": "2px 10px", "borderRadius": "10px",
                    "fontSize": "12px", "marginLeft": "8px",
                }),
            ]),
            html.P(d["description"], style={"fontSize": "12px", "color": "#555", "margin": "6px 0"}),
            html.Table([
                html.Thead(html.Tr([
                    html.Th("Indicator", style={"fontSize": "11px", "padding": "4px 8px", "backgroundColor": "#f0f4f8"}),
                    html.Th("Value", style={"fontSize": "11px", "padding": "4px 8px", "backgroundColor": "#f0f4f8"}),
                    html.Th("Trend", style={"fontSize": "11px", "padding": "4px 8px", "backgroundColor": "#f0f4f8"}),
                ])),
                html.Tbody(ind_rows),
            ], style={"width": "100%", "borderCollapse": "collapse", "marginTop": "8px"}),
        ], style={**CARD_STYLE, "padding": "14px"}))

    insight_box = html.Div([
        html.H3("Strategic Implications", style={**SECTION_TITLE, "fontSize": "16px"}),
        html.Ul([html.Li(i, style={"marginBottom": "8px", "fontSize": "13px"}) for i in PORTERS_INSIGHTS]),
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
        title="Gray Associates Portfolio Matrix: Market Score vs. Program Economics",
        xaxis_title="Program Economics Score (Revenue Efficiency)",
        yaxis_title="Market Score (Student Demand + Employment + Competition)",
        height=600, template="plotly_white",
        margin=dict(l=50, r=30, t=50, b=50),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
        annotations=[
            dict(x=30, y=80, text="SUSTAIN", showarrow=False,
                 font=dict(size=14, color="#3498db"), opacity=0.4),
            dict(x=80, y=80, text="GROW", showarrow=False,
                 font=dict(size=14, color="#2ecc71"), opacity=0.4),
            dict(x=30, y=30, text="SUNSET REVIEW", showarrow=False,
                 font=dict(size=14, color="#e74c3c"), opacity=0.4),
            dict(x=80, y=30, text="TRANSFORM", showarrow=False,
                 font=dict(size=14, color="#f39c12"), opacity=0.4),
        ],
    )

    # Recommendation summary bar
    rec_counts = df["GA_Recommendation"].value_counts()
    fig_bar = go.Figure(data=[go.Bar(
        x=rec_counts.index, y=rec_counts.values,
        marker_color=[GA_RECOMMENDATION_COLORS.get(r, "#999") for r in rec_counts.index],
        text=rec_counts.values, textposition="outside",
    )])
    fig_bar.update_layout(
        title="Programs by Recommendation", height=300,
        template="plotly_white", margin=dict(l=40, r=20, t=50, b=30),
        yaxis_title="Number of Programs",
    )

    table_df = df[["Program", "Enrollment", "Student_Demand_Score", "Employment_Score",
                   "Competition_Score", "Market_Score", "Economics_Score",
                   "Mission_Alignment", "GA_Recommendation"]].sort_values("Market_Score", ascending=False)

    rec_color_map = {
        "Grow": "#e8f5e9", "Sustain": "#e3f2fd", "Transform": "#fff3e0",
        "Evaluate": "#fce4ec", "Sunset Review": "#ffebee",
    }
    style_conditions = [
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
                html.Ul([html.Li(i, style={"marginBottom": "8px", "fontSize": "13px"}) for i in GA_INSIGHTS]),
            ], style={**CARD_STYLE, "flex": "1"}),
        ], style={"display": "flex", "gap": "16px"}),
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
            style_cell={"textAlign": "center", "padding": "6px", "fontSize": "12px"},
            style_header={"backgroundColor": FLC_NAVY, "color": "white", "fontWeight": "bold", "fontSize": "11px"},
            style_data_conditional=style_conditions,
            sort_action="native",
            filter_action="native",
            page_size=25,
        )], style=CARD_STYLE),
        source_annotation("Source: Gray Associates PES methodology; Market Score = Student Demand (40%) + Employment (40%) + Competition (20%)"),
    ])


def build_bcg_tab():
    """BCG Analysis tab with scatter plot - uses 'Concern' not 'Dog'."""
    fig = go.Figure()

    for quadrant in ["Star", "Cash Cow", "Question Mark", "Concern"]:
        df_q = BCG_DATA[BCG_DATA["Quadrant"] == quadrant]
        fig.add_trace(go.Scatter(
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
            textfont=dict(size=9),
        ))

    fig.add_hline(y=0, line_dash="dash", line_color="#aaa", line_width=1)
    fig.add_vline(x=4.0, line_dash="dash", line_color="#aaa", line_width=1)

    annotations = [
        dict(x=1.5, y=14, text="Question Marks", showarrow=False,
             font=dict(size=13, color="#f1c40f", family="Arial Black"), opacity=0.5),
        dict(x=8, y=14, text="Stars", showarrow=False,
             font=dict(size=13, color="#2ecc71", family="Arial Black"), opacity=0.5),
        dict(x=1.5, y=-28, text="Concerns", showarrow=False,
             font=dict(size=13, color="#e74c3c", family="Arial Black"), opacity=0.5),
        dict(x=8, y=-28, text="Cash Cows", showarrow=False,
             font=dict(size=13, color="#3498db", family="Arial Black"), opacity=0.5),
    ]
    fig.update_layout(
        title="BCG Growth-Share Matrix (Departments)",
        xaxis_title="23-24 % of Total SCH (Market Share)",
        yaxis_title="2-Year Change % (Growth Rate)",
        height=600, template="plotly_white",
        annotations=annotations,
        margin=dict(l=50, r=30, t=50, b=50),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
    )

    summary = BCG_DATA.groupby("Quadrant").agg(
        Count=("Department", "count"),
        Avg_SCH_Pct=("SCH_Pct", "mean"),
        Avg_Change=("Two_Year_Change", "mean"),
    ).reset_index().round(1)

    insight_list = html.Div([
        html.H3("Key Insights", style={**SECTION_TITLE, "fontSize": "16px"}),
        html.Ul([html.Li(i, style={"marginBottom": "8px", "fontSize": "13px"}) for i in BCG_INSIGHTS]),
    ], style=CARD_STYLE)

    return html.Div([
        html.H2("BCG Analysis", style=SECTION_TITLE),
        framework_description_block("BCG"),
        data_source_badge("BCG Growth-Share Matrix"),
        download_buttons("BCG"),
        html.Div([dcc.Graph(figure=fig, config={"displayModeBar": False})], style=CARD_STYLE),
        source_annotation("Source: BCG Presentation.pptx, BCG-growthMatrixDepts.png (FLC Internal)"),
        html.Div([
            html.Div([
                html.H3("Quadrant Summary", style={**SECTION_TITLE, "fontSize": "16px"}),
                dash_table.DataTable(
                    data=summary.to_dict("records"),
                    columns=[
                        {"name": "Quadrant", "id": "Quadrant"},
                        {"name": "# Departments", "id": "Count"},
                        {"name": "Avg SCH %", "id": "Avg_SCH_Pct"},
                        {"name": "Avg 2-Yr Change", "id": "Avg_Change"},
                    ],
                    style_cell={"textAlign": "center", "padding": "8px", "fontSize": "13px"},
                    style_header={"backgroundColor": FLC_NAVY, "color": "white", "fontWeight": "bold"},
                    style_data_conditional=[
                        {"if": {"filter_query": '{Quadrant} = "Star"'}, "backgroundColor": "#e8f5e9"},
                        {"if": {"filter_query": '{Quadrant} = "Concern"'}, "backgroundColor": "#fce4ec"},
                    ],
                ),
            ], style={**CARD_STYLE, "flex": "1"}),
            html.Div([insight_list], style={"flex": "1"}),
        ], style={"display": "flex", "gap": "16px"}),
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
                "backgroundColor": "#fafafa", "borderRadius": "6px",
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

    # SWOT summary counts for visualization
    fig_counts = go.Figure(data=[go.Bar(
        x=["Strengths", "Weaknesses", "Opportunities", "Threats"],
        y=[len(SWOT_DATA[k]["items"]) for k in ["Strengths", "Weaknesses", "Opportunities", "Threats"]],
        marker_color=[SWOT_DATA[k]["color"] for k in ["Strengths", "Weaknesses", "Opportunities", "Threats"]],
        text=[len(SWOT_DATA[k]["items"]) for k in ["Strengths", "Weaknesses", "Opportunities", "Threats"]],
        textposition="outside",
    )])
    fig_counts.update_layout(
        title="SWOT Factor Count by Category", height=280,
        template="plotly_white", margin=dict(l=40, r=20, t=50, b=30),
        yaxis_title="Number of Factors",
    )

    return html.Div([
        html.H2("SWOT Analysis", style=SECTION_TITLE),
        html.P(
            "Phase 2 synthesizes findings from all four Phase 1 frameworks (PESTLE, Porter's Five Forces, "
            "Gray Associates, BCG Matrix) into a unified Strengths-Weaknesses-Opportunities-Threats analysis. "
            "Each item includes source attribution to its originating framework(s).",
            style={
                "fontSize": "13px", "color": "#444", "lineHeight": "1.6",
                "backgroundColor": "#f8fafc", "padding": "12px 16px",
                "borderLeft": f"4px solid {FLC_GOLD}", "borderRadius": "4px",
                "marginBottom": "16px",
            },
        ),
        data_source_badge("SWOT Analysis"),

        html.Div([
            dcc.Graph(figure=fig_counts, config={"displayModeBar": False}),
            source_annotation("Source: Cross-framework synthesis of PESTLE, Porter's, Gray Associates, BCG analyses"),
        ], style=CARD_STYLE),

        # 2x2 SWOT grid
        html.Div([quadrants[0], quadrants[1]], style={"display": "flex", "gap": "16px", "flexWrap": "wrap"}),
        html.Div([quadrants[2], quadrants[3]], style={"display": "flex", "gap": "16px", "flexWrap": "wrap"}),
    ])


def build_zone_to_win_tab():
    """Phase 3: Zone to Win framework with 4 zones and 3 scenarios."""
    # Zone overview cards
    zone_cards = []
    for zone_name, zone_data in ZONE_TO_WIN_DATA.items():
        programs = zone_data["programs"]
        inv_colors = {"High": "#e74c3c", "Medium": "#f39c12", "Low": "#3498db"}
        program_rows = []
        for p in programs:
            program_rows.append(html.Tr([
                html.Td(p["name"], style={"fontSize": "12px", "padding": "6px 8px", "fontWeight": "600"}),
                html.Td(p["action"], style={"fontSize": "12px", "padding": "6px 8px"}),
                html.Td(p["investment"], style={
                    "fontSize": "12px", "padding": "6px 8px", "textAlign": "center",
                    "color": inv_colors.get(p["investment"], "#999"), "fontWeight": "600",
                }),
            ]))

        zone_cards.append(html.Div([
            html.Div([
                html.Div(style={
                    "display": "inline-block", "width": "12px", "height": "12px",
                    "borderRadius": "50%", "backgroundColor": zone_data["color"],
                    "marginRight": "8px", "verticalAlign": "middle",
                }),
                html.Strong(zone_name, style={"fontSize": "16px", "color": FLC_NAVY}),
                html.Span(f"  {len(programs)} initiatives", style={
                    "fontSize": "11px", "color": "#888", "marginLeft": "8px",
                }),
            ], style={"marginBottom": "8px"}),
            html.P(zone_data["description"], style={"fontSize": "12px", "color": "#555", "marginBottom": "8px"}),
            html.Table([
                html.Thead(html.Tr([
                    html.Th("Program/Initiative", style={"fontSize": "11px", "padding": "6px 8px", "backgroundColor": "#f0f4f8"}),
                    html.Th("Strategic Action", style={"fontSize": "11px", "padding": "6px 8px", "backgroundColor": "#f0f4f8"}),
                    html.Th("Investment", style={"fontSize": "11px", "padding": "6px 8px", "backgroundColor": "#f0f4f8", "textAlign": "center"}),
                ])),
                html.Tbody(program_rows),
            ], style={"width": "100%", "borderCollapse": "collapse"}),
        ], style={**CARD_STYLE, "borderLeft": f"4px solid {zone_data['color']}"}))

    # Zone allocation pie chart for each scenario
    scenario_figs = []
    for scenario_name, s_data in SCENARIOS.items():
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
            title=f"{scenario_name} Scenario", height=280,
            margin=dict(l=20, r=20, t=40, b=20), showlegend=False,
        )
        scenario_figs.append((scenario_name, s_data, fig_pie))

    # Scenario comparison bar chart
    fig_compare = go.Figure()
    metrics = ["enrollment_target", "retention_target", "graduate_target", "online_courses"]
    metric_labels = ["Enrollment", "Retention %", "Graduate Enroll.", "Online Courses"]
    current_vals = [3457, 66.1, 160, 25]
    x_labels = list(SCENARIOS.keys())
    for i, (metric, label) in enumerate(zip(metrics, metric_labels)):
        vals = [SCENARIOS[s][metric] for s in x_labels]
        fig_compare.add_trace(go.Bar(
            name=label, x=x_labels, y=vals,
            text=[f"{v:,.0f}" if v > 100 else f"{v}" for v in vals],
            textposition="outside",
        ))
    fig_compare.update_layout(
        title="Scenario Target Comparison",
        barmode="group", height=380,
        template="plotly_white",
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
                "backgroundColor": "#f8fafc", "padding": "12px 16px",
                "borderLeft": f"4px solid {FLC_GOLD}", "borderRadius": "4px",
                "marginBottom": "16px",
            },
        ),
        data_source_badge("Zone to Win"),

        # Four zone cards
        html.H3("Strategic Zones", style={**SECTION_TITLE, "fontSize": "18px"}),
        html.Div(zone_cards),

        # Scenarios section
        html.H3("Strategic Scenarios", style={**SECTION_TITLE, "fontSize": "18px"}),

        # Scenario details + pie charts
        html.Div([
            html.Div([
                html.Div([
                    html.Strong(name, style={"fontSize": "15px", "color": s_data["color"]}),
                    html.P(s_data["description"], style={"fontSize": "12px", "color": "#555", "margin": "6px 0"}),
                    html.Strong("Key Assumptions:", style={"fontSize": "12px"}),
                    html.Ul([html.Li(a, style={"fontSize": "11px"}) for a in s_data["assumptions"]],
                            style={"marginTop": "4px"}),
                ], style={"flex": "1"}),
                html.Div([
                    dcc.Graph(figure=fig_pie, config={"displayModeBar": False}),
                ], style={"flex": "1"}),
            ], style={**CARD_STYLE, "display": "flex", "gap": "16px",
                      "borderLeft": f"4px solid {s_data['color']}"})
            for name, s_data, fig_pie in scenario_figs
        ]),

        # Comparison chart
        html.Div([
            dcc.Graph(figure=fig_compare, config={"displayModeBar": False}),
            source_annotation("Source: Zone to Win methodology (Geoffrey Moore) applied to FLC strategic context"),
        ], style=CARD_STYLE),
    ])


def build_roadmap_tab():
    """Phase 3: Strategic Roadmap with timeline, KPIs, Gantt, milestones, and risks."""
    rm = ROADMAP_MILESTONES.copy()

    # Gantt-style timeline
    status_colors = {
        "Complete": "#2ecc71", "In Progress": "#3498db",
        "Not Started": "#95a5a6", "Upcoming": "#f39c12",
    }
    fig_gantt = go.Figure()
    for _, row in rm.iterrows():
        color = status_colors.get(row["Status"], "#999")
        label = row["Milestone"][:50] + "..." if len(row["Milestone"]) > 50 else row["Milestone"]
        fig_gantt.add_trace(go.Bar(
            y=[label],
            x=[pd.Timestamp(row["Target_Date"]) - pd.Timestamp(row["Start_Date"])],
            base=[row["Start_Date"]],
            orientation="h",
            marker=dict(color=color, opacity=0.85),
            showlegend=False,
            hovertext=f"{row['Milestone']}<br>Phase: {row['Phase']}<br>Zone: {row['Zone']}<br>Status: {row['Status']}<br>Owner: {row['Owner']}",
            hoverinfo="text",
        ))
    fig_gantt.update_layout(
        title="Strategic Implementation Timeline (2025-2027)",
        height=700, template="plotly_white",
        margin=dict(l=340, r=30, t=50, b=50),
        xaxis_title="Timeline",
        barmode="overlay",
    )
    # Add legend manually
    for status, color in status_colors.items():
        fig_gantt.add_trace(go.Bar(
            y=[None], x=[None], orientation="h",
            marker=dict(color=color), name=status, showlegend=True,
        ))
    fig_gantt.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="center", x=0.5))

    # Milestone status summary
    status_counts = rm["Status"].value_counts()
    fig_status = go.Figure(data=[go.Pie(
        labels=status_counts.index, values=status_counts.values,
        marker=dict(colors=[status_colors.get(s, "#999") for s in status_counts.index]),
        hole=0.5, textinfo="label+value",
    )])
    fig_status.update_layout(
        title="Milestone Status", height=280,
        margin=dict(l=20, r=20, t=40, b=20), showlegend=False,
    )

    # Zone distribution
    zone_counts = rm["Zone"].value_counts()
    fig_zone = go.Figure(data=[go.Bar(
        x=zone_counts.index, y=zone_counts.values,
        marker_color=[{
            "Performance": "#2ecc71", "Productivity": "#3498db",
            "Incubation": "#f39c12", "Transformation": "#9b59b6", "All": FLC_NAVY,
        }.get(z, "#999") for z in zone_counts.index],
        text=zone_counts.values, textposition="outside",
    )])
    fig_zone.update_layout(
        title="Milestones by Zone", height=280,
        template="plotly_white", margin=dict(l=40, r=20, t=50, b=30),
    )

    # KPIs table with progress visualization
    kpis = ROADMAP_KPIS.copy()
    kpi_fig = go.Figure()
    for _, row in kpis.iterrows():
        baseline = row["Baseline_Value"]
        y1 = row["Year1_Target"]
        y2 = row["Year2_Target"]
        stretch = row["Stretch_Target"]
        kpi_fig.add_trace(go.Bar(
            y=[row["KPI"]], x=[y1 - baseline], base=[baseline],
            orientation="h", name="Year 1" if _ == 0 else None,
            marker=dict(color="#3498db", opacity=0.7),
            showlegend=(_ == 0),
            hovertext=f"{row['KPI']}: Baseline={baseline}, Y1={y1}, Y2={y2}, Stretch={stretch}",
            hoverinfo="text",
        ))
        kpi_fig.add_trace(go.Bar(
            y=[row["KPI"]], x=[y2 - y1], base=[y1],
            orientation="h", name="Year 2" if _ == 0 else None,
            marker=dict(color="#2ecc71", opacity=0.7),
            showlegend=(_ == 0),
            hoverinfo="skip",
        ))
        kpi_fig.add_trace(go.Bar(
            y=[row["KPI"]], x=[stretch - y2], base=[y2],
            orientation="h", name="Stretch" if _ == 0 else None,
            marker=dict(color="#f39c12", opacity=0.5),
            showlegend=(_ == 0),
            hoverinfo="skip",
        ))
    kpi_fig.update_layout(
        title="KPI Targets: Baseline to Stretch",
        barmode="stack", height=500,
        template="plotly_white",
        margin=dict(l=200, r=40, t=50, b=30),
        legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="center", x=0.5),
    )

    # Risk matrix
    risk_df = RISK_MITIGATION.copy()
    prob_map = {"Low": 1, "Medium": 2, "High": 3}
    impact_map = {"Low": 1, "Medium": 2, "High": 3}
    risk_df["Prob_Num"] = risk_df["Probability"].map(prob_map)
    risk_df["Impact_Num"] = risk_df["Impact"].map(impact_map)
    risk_df["Risk_Score"] = risk_df["Prob_Num"] * risk_df["Impact_Num"]

    fig_risk = go.Figure(data=go.Scatter(
        x=risk_df["Prob_Num"], y=risk_df["Impact_Num"],
        mode="markers+text",
        marker=dict(
            size=risk_df["Risk_Score"] * 8 + 10,
            color=risk_df["Risk_Score"],
            colorscale=[[0, "#2ecc71"], [0.5, "#f39c12"], [1, "#e74c3c"]],
            showscale=True, colorbar=dict(title="Risk Score"),
            opacity=0.8,
        ),
        text=risk_df["Risk"].str[:30],
        textposition="top center",
        textfont=dict(size=9),
        hovertext=risk_df.apply(
            lambda r: f"{r['Risk']}<br>Prob: {r['Probability']}, Impact: {r['Impact']}<br>Mitigation: {r['Mitigation_Strategy'][:100]}...",
            axis=1,
        ),
        hoverinfo="text",
    ))
    fig_risk.update_layout(
        title="Risk Assessment Matrix",
        xaxis=dict(title="Probability", tickvals=[1, 2, 3], ticktext=["Low", "Medium", "High"], range=[0.5, 3.5]),
        yaxis=dict(title="Impact", tickvals=[1, 2, 3], ticktext=["Low", "Medium", "High"], range=[0.5, 3.5]),
        height=450, template="plotly_white",
        margin=dict(l=60, r=30, t=50, b=50),
    )
    # Add risk zones
    fig_risk.add_shape(type="rect", x0=0.5, y0=2.5, x1=1.5, y1=3.5,
                       fillcolor="rgba(243,156,18,0.1)", line_width=0)
    fig_risk.add_shape(type="rect", x0=1.5, y0=1.5, x1=3.5, y1=3.5,
                       fillcolor="rgba(231,76,60,0.1)", line_width=0)
    fig_risk.add_shape(type="rect", x0=0.5, y0=0.5, x1=1.5, y1=1.5,
                       fillcolor="rgba(46,204,113,0.1)", line_width=0)

    return html.Div([
        html.H2("Strategic Roadmap", style=SECTION_TITLE),
        html.P(
            "The Strategic Roadmap translates Zone to Win scenarios into a detailed implementation plan with "
            "milestones, KPI targets, and risk mitigation strategies across the 2025-2027 planning horizon.",
            style={
                "fontSize": "13px", "color": "#444", "lineHeight": "1.6",
                "backgroundColor": "#f8fafc", "padding": "12px 16px",
                "borderLeft": f"4px solid {FLC_GOLD}", "borderRadius": "4px",
                "marginBottom": "16px",
            },
        ),
        data_source_badge("Strategic Roadmap"),

        # Status summary row
        html.Div([
            html.Div([
                dcc.Graph(figure=fig_status, config={"displayModeBar": False}),
            ], style={**CARD_STYLE, "flex": "1"}),
            html.Div([
                dcc.Graph(figure=fig_zone, config={"displayModeBar": False}),
            ], style={**CARD_STYLE, "flex": "1"}),
        ], style={"display": "flex", "gap": "16px"}),

        # Gantt timeline
        html.Div([
            dcc.Graph(figure=fig_gantt, config={"displayModeBar": False}),
            source_annotation("Source: Implementation plan derived from Zone to Win scenarios + Phase 1 analyses"),
        ], style=CARD_STYLE),

        # Milestone detail table
        html.H3("Milestone Tracker", style={**SECTION_TITLE, "fontSize": "16px"}),
        html.Div([dash_table.DataTable(
            data=rm.to_dict("records"),
            columns=[
                {"name": "ID", "id": "ID"},
                {"name": "Milestone", "id": "Milestone"},
                {"name": "Phase", "id": "Phase"},
                {"name": "Start", "id": "Start_Date"},
                {"name": "Target", "id": "Target_Date"},
                {"name": "Status", "id": "Status"},
                {"name": "Zone", "id": "Zone"},
                {"name": "Owner", "id": "Owner"},
            ],
            style_cell={"textAlign": "left", "padding": "6px", "fontSize": "12px",
                         "whiteSpace": "normal", "height": "auto"},
            style_header={"backgroundColor": FLC_NAVY, "color": "white", "fontWeight": "bold", "fontSize": "11px"},
            style_data_conditional=[
                {"if": {"filter_query": '{Status} = "Complete"'}, "backgroundColor": "#e8f5e9"},
                {"if": {"filter_query": '{Status} = "In Progress"'}, "backgroundColor": "#e3f2fd"},
                {"if": {"filter_query": '{Status} = "Upcoming"'}, "backgroundColor": "#fff8e1"},
            ],
            sort_action="native",
            filter_action="native",
            page_size=20,
        )], style=CARD_STYLE),

        # KPI targets
        html.H3("KPI Targets & Tracking", style={**SECTION_TITLE, "fontSize": "16px"}),
        html.Div([
            dcc.Graph(figure=kpi_fig, config={"displayModeBar": False}),
            source_annotation("Source: Baseline from Fall 2025 census; targets aligned to Most Likely scenario"),
        ], style=CARD_STYLE),

        html.Div([dash_table.DataTable(
            data=kpis.to_dict("records"),
            columns=[
                {"name": "KPI", "id": "KPI"},
                {"name": "Category", "id": "Category"},
                {"name": "Baseline", "id": "Baseline_Value"},
                {"name": "Year 1", "id": "Year1_Target"},
                {"name": "Year 2", "id": "Year2_Target"},
                {"name": "Stretch", "id": "Stretch_Target"},
                {"name": "Unit", "id": "Unit"},
                {"name": "Measurement", "id": "Measurement"},
            ],
            style_cell={"textAlign": "center", "padding": "6px", "fontSize": "12px"},
            style_header={"backgroundColor": FLC_NAVY, "color": "white", "fontWeight": "bold", "fontSize": "11px"},
            sort_action="native",
        )], style=CARD_STYLE),

        # Risk management
        html.H3("Risk Assessment & Mitigation", style={**SECTION_TITLE, "fontSize": "16px"}),
        html.Div([
            dcc.Graph(figure=fig_risk, config={"displayModeBar": False}),
            source_annotation("Source: Risk analysis synthesized from all Phase 1 frameworks"),
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
            style_cell={"textAlign": "left", "padding": "8px", "fontSize": "12px",
                         "whiteSpace": "normal", "height": "auto"},
            style_header={"backgroundColor": FLC_NAVY, "color": "white", "fontWeight": "bold", "fontSize": "11px"},
            style_data_conditional=[
                {"if": {"filter_query": '{Impact} = "High"', "column_id": "Impact"},
                 "color": "#e74c3c", "fontWeight": "bold"},
                {"if": {"filter_query": '{Probability} = "High"', "column_id": "Probability"},
                 "color": "#e74c3c", "fontWeight": "bold"},
            ],
            sort_action="native",
        )], style=CARD_STYLE),
    ])


# ============================================================================
# MAIN LAYOUT
# ============================================================================

app.layout = html.Div([
    # Header
    html.Div([
        html.Div([
            html.H1("Fort Lewis College", style={
                "color": "white", "margin": "0", "fontSize": "24px", "fontWeight": "bold",
            }),
            html.Div("Portfolio Optimization Dashboard", style={
                "color": FLC_GOLD, "fontSize": "14px", "fontWeight": "600",
            }),
        ], style={"flex": "1"}),
        html.Div([
            html.Div("Phase 1: Environmental Scanning | Phase 2: Strategic Synthesis | Phase 3: Strategic Direction", style={
                "color": "#aaa", "fontSize": "11px",
            }),
        ]),
    ], style={
        "backgroundColor": FLC_NAVY, "padding": "16px 24px",
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
    ], style={"marginBottom": "0"}),

    # Tab content
    html.Div(id="tab-content", style={
        "padding": "20px 24px",
        "backgroundColor": FLC_LIGHT,
        "minHeight": "calc(100vh - 140px)",
    }),

    # Footer
    html.Div([
        html.Span("Fort Lewis College Portfolio Optimization Project | "),
        html.Span("Phase 1: PESTLE, Porter's, Gray Associates, BCG | "),
        html.Span("Phase 2: SWOT Synthesis | Phase 3: Zone to Win & Roadmap"),
    ], style={
        "backgroundColor": FLC_NAVY, "color": "#aaa", "padding": "10px 24px",
        "fontSize": "11px", "textAlign": "center",
    }),
], style={"fontFamily": "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif"})


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


# ============================================================================
# RUN
# ============================================================================

if __name__ == "__main__":
    print("\n" + "=" * 60)
    print("  FLC Portfolio Optimization Dashboard")
    print("  Phase 1: Environmental Scanning")
    print("  Phase 2: Strategic Synthesis (SWOT)")
    print("  Phase 3: Strategic Direction (Zone to Win + Roadmap)")
    print("  Open http://127.0.0.1:8050 in your browser")
    print("=" * 60 + "\n")
    app.run(host='0.0.0.0', port=8080, debug=False)
