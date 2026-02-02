"""
Generate Excel templates for FLC Portfolio Optimization Dashboard data input.
Run this script to create editable Excel files in the dashboard/templates/ folder.
Users can fill in these templates and the dashboard will read from them.

Usage:  python create_excel_templates.py
"""

import os
import pandas as pd
from data import (
    BCG_DATA, PESTLE_DATA, PORTERS_DATA,
    GRAY_ASSOCIATES_DATA, STRATEGIC_INITIATIVES,
    MILESTONES, KPIS, RESOURCE_ALLOCATION,
    ENROLLMENT_HISTORY, RETENTION_HISTORY,
    TOP_MAJORS_ENROLLMENT, DEGREES_AWARDED,
)

TEMPLATE_DIR = os.path.join(os.path.dirname(__file__), "templates")
os.makedirs(TEMPLATE_DIR, exist_ok=True)


def create_bcg_template():
    """BCG Growth-Share Matrix data template."""
    path = os.path.join(TEMPLATE_DIR, "bcg_data.xlsx")
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        BCG_DATA.to_excel(writer, sheet_name="BCG_Matrix", index=False)
        pd.DataFrame({
            "Quadrant": ["Star", "Cash Cow", "Question Mark", "Concern"],
            "Description": [
                "High market share + Growing (invest to maintain)",
                "High market share + Declining (optimize efficiency)",
                "Low market share + Growing (evaluate investment)",
                "Low market share + Declining (restructure or sunset)",
            ],
        }).to_excel(writer, sheet_name="Quadrant_Definitions", index=False)
    print(f"  Created: {path}")


def create_pestle_template():
    """PESTLE Analysis data template."""
    path = os.path.join(TEMPLATE_DIR, "pestle_data.xlsx")
    rows = []
    for category, d in PESTLE_DATA.items():
        for i, factor in enumerate(d["factors"]):
            rows.append({
                "Category": category,
                "Impact": d["impact"],
                "Impact_Score": d["impact_score"],
                "Trend": d["trend"],
                "Factor": factor,
                "Type": "Factor",
            })
        for opp in d["opportunities"]:
            rows.append({
                "Category": category,
                "Impact": d["impact"],
                "Impact_Score": d["impact_score"],
                "Trend": d["trend"],
                "Factor": opp,
                "Type": "Opportunity",
            })
    df = pd.DataFrame(rows)
    df.to_excel(path, index=False, engine="openpyxl")
    print(f"  Created: {path}")


def create_porters_template():
    """Porter's Five Forces data template."""
    path = os.path.join(TEMPLATE_DIR, "porters_five_forces.xlsx")
    rows = []
    for force, d in PORTERS_DATA.items():
        for ind in d["indicators"]:
            rows.append({
                "Force": force,
                "Rating": d["rating"],
                "Score": d["score"],
                "Description": d["description"],
                "Indicator_Name": ind["name"],
                "Indicator_Value": ind["value"],
                "Indicator_Trend": ind["trend"],
            })
    df = pd.DataFrame(rows)
    df.to_excel(path, index=False, engine="openpyxl")
    print(f"  Created: {path}")


def create_gray_template():
    """Gray Associates Portfolio Analysis template."""
    path = os.path.join(TEMPLATE_DIR, "gray_associates_data.xlsx")
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        GRAY_ASSOCIATES_DATA.to_excel(writer, sheet_name="Program_Scores", index=False)
        pd.DataFrame({
            "Recommendation": ["Grow", "Sustain", "Transform", "Evaluate", "Sunset Review"],
            "Description": [
                "High market + strong economics: prioritize investment",
                "Solid market, needs efficiency improvements",
                "Weak market but strong economics: innovate delivery",
                "Needs deeper analysis before deciding direction",
                "Weak market + economics: consider phase-out",
            ],
            "Market_Score_Range": [">65", "50-65", "<50", "<50", "<40"],
            "Economics_Score_Range": [">65", "Any", ">60", "<60", "<50"],
        }).to_excel(writer, sheet_name="Scoring_Guide", index=False)
    print(f"  Created: {path}")


def create_implementation_template():
    """Phase 2 implementation tracking template."""
    path = os.path.join(TEMPLATE_DIR, "implementation_tracking.xlsx")
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        STRATEGIC_INITIATIVES.to_excel(writer, sheet_name="Initiatives", index=False)
        MILESTONES.to_excel(writer, sheet_name="Milestones", index=False)
        KPIS.to_excel(writer, sheet_name="KPIs", index=False)
        RESOURCE_ALLOCATION.to_excel(writer, sheet_name="Resources", index=False)
    print(f"  Created: {path}")


def create_enrollment_template():
    """Enrollment and institutional data template."""
    path = os.path.join(TEMPLATE_DIR, "enrollment_data.xlsx")
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        ENROLLMENT_HISTORY.to_excel(writer, sheet_name="Enrollment_History", index=False)
        RETENTION_HISTORY.to_excel(writer, sheet_name="Retention_History", index=False)
        TOP_MAJORS_ENROLLMENT.to_excel(writer, sheet_name="Top_Majors", index=False)
        DEGREES_AWARDED.to_excel(writer, sheet_name="Degrees_Awarded", index=False)
    print(f"  Created: {path}")


if __name__ == "__main__":
    print("\nCreating Excel templates in dashboard/templates/\n")
    create_bcg_template()
    create_pestle_template()
    create_porters_template()
    create_gray_template()
    create_implementation_template()
    create_enrollment_template()
    print(f"\nAll templates created in: {TEMPLATE_DIR}")
    print("Edit these files and update data.py to reflect changes.\n")
