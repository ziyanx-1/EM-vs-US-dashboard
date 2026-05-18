import json
import re
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import boto3
from botocore.config import Config
from pathlib import Path

st.set_page_config(page_title="MAS Analytics", layout="wide")

# ── Paths ──────────────────────────────────────────────────────────────────────
BASE_DIR     = Path(__file__).parent
REPORTS_DIR  = BASE_DIR / "saved"
TEMPLATE_DIR = BASE_DIR / "template"
REPORTS_DIR.mkdir(exist_ok=True)

EM_US_EXCEL    = BASE_DIR / "External Research Summary - EM vs US V2 (Filtered).xlsx"
EM_US_TEMPLATE = TEMPLATE_DIR / "EM_vs_US_Dashboard_V3_Final.html"
OPUS_MODEL     = "us.anthropic.claude-opus-4-6-v1"
SONNET_MODEL   = "us.anthropic.claude-sonnet-4-20250514-v1:0"

# ── Sidebar navigation ─────────────────────────────────────────────────────────
st.sidebar.title("MAS Analytics")

SECTIONS  = ["PM Analysis", "External Research Insights", "Models"]
ERI_TOPICS = ["EQ/FI", "VALUE/GROWTH", "SMALL/LARGE", "US/EAFE/EM", "FI/DURATION", "FI/CREDIT"]

section = st.sidebar.radio("", SECTIONS, label_visibility="collapsed")

# ── Bedrock helpers (from app.py) ──────────────────────────────────────────────

@st.cache_resource
def get_bedrock_client():
    return boto3.client(
        "bedrock-runtime",
        region_name="us-east-1",
        config=Config(retries={"max_attempts": 3}),
    )


def excel_to_text(path: Path) -> str:
    xl = pd.ExcelFile(path)
    parts = []
    for sheet in xl.sheet_names:
        df = xl.parse(sheet, header=None, dtype=str).fillna("")
        parts.append(f"=== Sheet: {sheet} ===\n{df.to_string(index=False)}")
    return "\n\n".join(parts)


def strip_code_fence(text: str) -> str:
    text = text.strip()
    text = re.sub(r"^```[a-zA-Z]*\n?", "", text)
    text = re.sub(r"\n?```$", "", text)
    return text.strip()


def bedrock_stream(prompt: str, max_tokens: int, budget_tokens: int, model: str = OPUS_MODEL):
    body = json.dumps({
        "anthropic_version": "bedrock-2023-05-31",
        "max_tokens": max_tokens,
        "thinking": {"type": "enabled", "budget_tokens": budget_tokens},
        "messages": [{"role": "user", "content": [{"type": "text", "text": prompt}]}],
    })
    resp = get_bedrock_client().invoke_model_with_response_stream(
        modelId=model, body=body,
        contentType="application/json", accept="application/json",
    )
    current_type = None
    for event in resp["body"]:
        chunk = json.loads(event["chunk"]["bytes"])
        evt = chunk.get("type")
        if evt == "content_block_start":
            current_type = chunk.get("content_block", {}).get("type")
        elif evt == "content_block_delta":
            delta = chunk.get("delta", {})
            if delta.get("type") == "thinking_delta":
                yield "thinking", delta.get("thinking", "")
            elif delta.get("type") == "text_delta":
                yield "text", delta.get("text", "")


ANALYSIS_PROMPT = """\
You are a senior investment research analyst. Your task is to carefully read a set of \
investment research articles (provided as raw Excel data) and extract structured facts \
that will be used to populate a financial dashboard.

CONTEXT — what this dashboard is:
The dashboard summarises external sell-side research articles that each express a view on \
whether to Overweight EM (Emerging Markets) equities vs US equities, Underweight EM vs US, \
take a Selective / country-specific stance, or remain Neutral. The dashboard surfaces the \
consensus picture across all articles: how many hold each stance, what arguments each camp \
makes, what risks are cited most, what catalysts could change views, and how valuations compare.

─── EXCEL DATA ──────────────────────────────────────────────────
{excel_text}
─────────────────────────────────────────────────────────────────

Extract the following with exact counts and zero rounding errors. \
Count every data row carefully — do not guess or estimate.

1. OVERVIEW
   - Total number of research reports (= data rows, excluding header)
   - Institution names and date range covered

2. STANCE SPLIT — exact count of rows for each classification:
   - Overweight EM (or "OW EM")
   - Underweight EM (or "UW EM")
   - Selective / Selective EM
   - Neutral
   (Sum must equal total reports — double-check this.)

3. CONVICTION BY STANCE — for each stance, summarise the typical Conviction Strength values \
   (e.g. Weak, Moderate, Strong) to derive "average" conviction range per stance.

4. KEY DRIVERS — for each stance (OW / UW / Selective), list the primary and supporting \
   drivers grouped by category (Macro, Fundamentals, Valuation, Sentiment, Technicals). \
   Deduplicate and summarise to 2–4 concise bullet points per category.

5. TOP RISKS — list the most frequently cited risks across ALL reports, with a count of \
   how many reports cite each, and classify each as: Macro / Geo / Tech / Flow / Other.

6. CATALYSTS — separate into:
   - Bull (EM-positive catalysts)
   - Bear (US-positive catalysts)
   - Bilateral (outcome-dependent)
   Include a citation count where determinable.

7. VALUATION SNAPSHOT — any forward P/E, discount, or valuation metrics mentioned. \
   Include the index (S&P 500, MSCI EM, specific countries) and the reported figure.

8. PRICED IN vs NOT PRICED IN — list items explicitly flagged as already reflected in \
   prices vs items the market has not yet priced in.

Output a clearly structured plain-text report with section headers. \
Be numerically precise — every count must be verifiable from the data above.
"""

GENERATION_PROMPT = """\
You are an expert frontend developer specialising in financial research dashboards.

CONTEXT — what this dashboard is:
The dashboard is a one-page visual summary for portfolio managers, showing the consensus \
and dispersion of views across sell-side investment research articles on the topic of \
EM (Emerging Markets) equities vs US equities. It must be accurate, professional, and \
instantly scannable.

You are given three inputs:
1. A reference HTML template — copy its EXACT CSS design system, component structure, \
   and JavaScript. Do not change a single CSS variable or class name.
2. A pre-verified structured analysis of the research data — use these numbers and labels \
   as the authoritative source of truth for every value in the dashboard.
3. The raw Excel data — use only as a secondary reference if you need extra detail.

─── TEMPLATE HTML ───────────────────────────────────────────────
{template_html}
─────────────────────────────────────────────────────────────────

─── PRE-VERIFIED DATA ANALYSIS ──────────────────────────────────
{analysis}
─────────────────────────────────────────────────────────────────

─── RAW EXCEL DATA (secondary reference) ────────────────────────
{excel_text}
─────────────────────────────────────────────────────────────────

Instructions:
- Reproduce the complete HTML document with the same structure and CSS.
- Replace ALL hardcoded values (numbers, labels, pill text, lollipop data array, \
  header subtitle, institution list, date range) with values from the analysis above.
- The stance counts in the banner MUST match the analysis exactly.
- The lollipop JS data array MUST use the risk counts from the analysis.
- Keep all tab-switching and lollipop-rendering JavaScript intact.
- Output ONLY raw HTML. Start with <!DOCTYPE html> and end with </html>.
- No markdown, no code fences, no commentary outside the HTML.
"""


def run_step1(excel_text: str, status_widget) -> str:
    prompt = ANALYSIS_PROMPT.format(excel_text=excel_text)
    thinking_buf, text_buf = "", ""
    thinking_box = status_widget.empty()
    output_box   = status_widget.empty()
    for block_type, chunk in bedrock_stream(prompt, max_tokens=12000, budget_tokens=9000):
        if block_type == "thinking":
            thinking_buf += chunk
            thinking_box.caption(f"Thinking… ({len(thinking_buf):,} chars)")
        else:
            text_buf += chunk
            output_box.markdown(f"```\n…{text_buf[-500:]}\n```")
    thinking_box.empty()
    output_box.empty()
    return text_buf.strip()


def run_step2(excel_text: str, analysis: str, template_html: str, status_widget) -> str:
    prompt = GENERATION_PROMPT.format(
        template_html=template_html, analysis=analysis, excel_text=excel_text,
    )
    thinking_buf, html_buf = "", ""
    thinking_box = status_widget.empty()
    counter_box  = status_widget.empty()
    preview_box  = status_widget.empty()
    for block_type, chunk in bedrock_stream(prompt, max_tokens=20000, budget_tokens=6000, model=SONNET_MODEL):
        if block_type == "thinking":
            thinking_buf += chunk
            thinking_box.caption(f"Thinking… ({len(thinking_buf):,} chars)")
        else:
            html_buf += chunk
            thinking_box.empty()
            counter_box.caption(f"Writing HTML… {len(html_buf):,} characters so far")
            tail = html_buf[-300:].replace("<", "&lt;").replace(">", "&gt;")
            preview_box.markdown(f"```html\n…{tail}\n```")
    thinking_box.empty()
    counter_box.empty()
    preview_box.empty()
    return strip_code_fence(html_buf)

# ── Report persistence ─────────────────────────────────────────────────────────

def report_path(topic: str) -> Path:
    return REPORTS_DIR / f"{topic.replace('/', '_')}.html"


def load_report(topic: str) -> str | None:
    p = report_path(topic)
    return p.read_text(encoding="utf-8") if p.exists() else None


def save_report(topic: str, html: str) -> None:
    report_path(topic).write_text(html, encoding="utf-8")

# ── Section renderers ──────────────────────────────────────────────────────────

def render_pm_analysis():
    st.header("PM Analysis")
    st.info("PM Analysis content coming soon.")


def render_models():
    st.header("Models")
    st.info("Models content coming soon.")


def render_us_eafe_em(is_admin: bool):
    st.subheader("US / EAFE / EM")

    report_html = load_report("US/EAFE/EM")

    if is_admin:
        st.markdown("**Admin — generate / update report**")

        # Show data preview
        if EM_US_EXCEL.exists():
            with st.expander(f"Data source: {EM_US_EXCEL.name}", expanded=False):
                df_preview = pd.read_excel(EM_US_EXCEL)
                st.dataframe(df_preview.head(20), use_container_width=True)
                st.caption(f"{len(df_preview):,} rows · {len(df_preview.columns)} columns")
        else:
            st.error(f"Excel file not found: {EM_US_EXCEL.name}")
            return

        if not EM_US_TEMPLATE.exists():
            st.error(f"Template not found: {EM_US_TEMPLATE.name}")
            return

        if st.button("Run report", type="primary", key="run_us_eafe_em"):
            excel_text    = excel_to_text(EM_US_EXCEL)
            template_html = EM_US_TEMPLATE.read_text(encoding="utf-8")

            try:
                with st.status("Step 1 / 2 — Analysing research data (Opus 4.6)…", expanded=True) as s1:
                    analysis = run_step1(excel_text, s1)
                    s1.update(label="Step 1 / 2 — Analysis complete ✓", state="complete", expanded=False)

                with st.expander("View extracted analysis", expanded=False):
                    st.text(analysis)

                with st.status("Step 2 / 2 — Generating dashboard HTML (Sonnet 4.6)…", expanded=True) as s2:
                    html = run_step2(excel_text, analysis, template_html, s2)
                    s2.update(label="Step 2 / 2 — Dashboard ready ✓", state="complete", expanded=False)

                save_report("US/EAFE/EM", html)
                st.success("Report saved and now visible to all users.")
                report_html = html

            except Exception as e:
                st.error(f"Generation failed: {e}")

        if report_html:
            st.divider()

    if report_html:
        col_dl, _ = st.columns([1, 5])
        with col_dl:
            st.download_button(
                "Download HTML",
                data=report_html,
                file_name="US_EAFE_EM_report.html",
                mime="text/html",
            )
        components.html(report_html, height=950, scrolling=True)
    else:
        st.info("No report available yet. Ask an admin to run it.")


def render_eri_topic_placeholder(topic: str):
    st.subheader(topic)
    st.info(f"**{topic}** report coming soon.")


def render_external_research(is_admin: bool):
    st.header("External Research Insights")

    topic = st.sidebar.radio("Select topic", ERI_TOPICS, key="eri_topic")

    if topic == "US/EAFE/EM":
        render_us_eafe_em(is_admin)
    else:
        render_eri_topic_placeholder(topic)


# ── Admin toggle (sidebar) ─────────────────────────────────────────────────────
is_admin = st.sidebar.checkbox("Admin mode", key="global_admin")

# ── Main router ────────────────────────────────────────────────────────────────
if section == "PM Analysis":
    render_pm_analysis()
elif section == "External Research Insights":
    render_external_research(is_admin)
elif section == "Models":
    render_models()
