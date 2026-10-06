"""ADK workflow definition for staged dashboard comparison.

The existing FastAPI engine can adopt this root agent incrementally. Each stage
writes a compact JSON artifact to session state instead of forwarding a full
-dashboard extraction between agents.
"""

from google.adk.agents import LlmAgent, SequentialAgent
from google.genai import types

MODEL = "gemini-2.5-flash"

INVENTORY_INSTRUCTION = """You are the Dashboard Inventory Agent.
Inspect the supplied dashboard image and return only compact JSON inventory data:
dashboard title, page name, visible filters, visual count, visual IDs, titles,
optional descriptive types, normalized bounding boxes, and reading order.
Do not extract chart points, table rows, KPI values, or treemap tile values.
Do not return explanations or duplicate visuals. Store the result in inventory."""

LAYOUT_INSTRUCTION = """You are the Layout Comparison Agent.
Use the two inventory objects in session state: {inventory_a} and {inventory_b}.
Return only JSON with matched_visuals, missing_visuals, added_visuals, moved_visuals,
and changed_visuals. Match by title, semantic identity, relative position, and
bounding box. Do not extract metric values. Store the result in layout_changes."""

DETAIL_INSTRUCTION = """You are the Visual Detail Agent.
Use layout_changes to identify changed visual IDs. Extract details only for those
visual crops supplied in the request. Return compact JSON with visible labels,
values, units, truncation flags, bounding boxes, and confidence. Never infer
hidden values and never return unchanged visual data. Store the result in details."""

REPORT_INSTRUCTION = """You are the Dashboard Report Agent.
Use inventory_a, inventory_b, layout_changes, and details from session state.
Return only the final comparison JSON with overall_similarity, status, matched,
missing, added, changed counts, and evidence-backed anomalies. Mark incomplete
extraction as Uncertain; never treat identical null or N/A placeholders as a match.
Store the result in comparison_report."""


def build_dashboard_workflow() -> SequentialAgent:
    """Build the staged ADK workflow without making any model calls."""
    inventory_a = LlmAgent(
        name="dashboard_inventory_agent_a",
        description="Build a compact inventory for the baseline dashboard image.",
        model=MODEL,
        instruction=INVENTORY_INSTRUCTION,
        output_key="inventory_a",
        generate_content_config=types.GenerateContentConfig(
            temperature=0,
            max_output_tokens=4096,
            response_mime_type="application/json",
        ),
    )
    inventory_b = LlmAgent(
        name="dashboard_inventory_agent_b",
        description="Build a compact inventory for the current dashboard image.",
        model=MODEL,
        instruction=INVENTORY_INSTRUCTION,
        output_key="inventory_b",
        generate_content_config=types.GenerateContentConfig(
            temperature=0,
            max_output_tokens=4096,
            response_mime_type="application/json",
        ),
    )
    layout = LlmAgent(
        name="layout_comparison_agent",
        description="Compare compact dashboard inventories before deep extraction.",
        model=MODEL,
        instruction=LAYOUT_INSTRUCTION,
        output_key="layout_changes",
        generate_content_config=types.GenerateContentConfig(
            temperature=0,
            max_output_tokens=4096,
            response_mime_type="application/json",
        ),
    )
    detail = LlmAgent(
        name="visual_detail_agent",
        description="Extract only changed visual details from supplied crops.",
        model=MODEL,
        instruction=DETAIL_INSTRUCTION,
        output_key="details",
        generate_content_config=types.GenerateContentConfig(
            temperature=0,
            max_output_tokens=8192,
            response_mime_type="application/json",
        ),
    )
    report = LlmAgent(
        name="dashboard_report_agent",
        description="Produce the evidence-based comparison report.",
        model=MODEL,
        instruction=REPORT_INSTRUCTION,
        output_key="comparison_report",
        generate_content_config=types.GenerateContentConfig(
            temperature=0,
            max_output_tokens=4096,
            response_mime_type="application/json",
        ),
    )
    return SequentialAgent(
        name="dashboard_comparison_workflow",
        description="Inventory, compare layout, extract changed visuals, and report.",
        sub_agents=[inventory_a, inventory_b, layout, detail, report],
    )


root_agent = build_dashboard_workflow()
