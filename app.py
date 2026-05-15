import json
import re
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import boto3
from botocore.config import Config
from pathlib import Path

st.set_page_config(page_title="Research Dashboard Builder", layout="wide")

TEMPLATE_DIR = Path(__file__).parent / "template"
OPUS_MODEL = "us.anthropic.claude-opus-4-6-v1"


@st.cache_resource
def get_bedrock_client():
    return boto3.client(
        "bedrock-runtime",
        region_name="us-east-1",
        config=Config(retries={"max_attempts": 3}),
    )


def list_templates() -> list[str]:
    return [f.name for f in sorted(TEMPLATE_DIR.glob("*.html"))]


def excel_to_text(file) -> str:
    xl = pd.ExcelFile(file)
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
    """
    Generator yielding (block_type, text) tuples.
    block_type is 'thinking' or 'text'.
    Extended thinking is enabled with the given budget.
    """
    body = json.dumps({
        "anthropic_version": "bedrock-2023-05-31",
        "max_tokens": max_tokens,
        "thinking": {"type": "enabled", "budget_tokens": budget_tokens},
        "messages": [
            {"role": "user", "content": [{"type": "text", "text": prompt}]}
        ],
    })

    resp = get_bedrock_client().invoke_model_with_response_stream(
        modelId=model,
        body=body,
        contentType="application/json",
        accept="application/json",
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
    """Analyse the Excel and return the structured text summary."""
    prompt = ANALYSIS_PROMPT.format(excel_text=excel_text)
    thinking_buf, text_buf = "", ""

    thinking_box = status_widget.empty()
    output_box   = status_widget.empty()

    for block_type, chunk in bedrock_stream(prompt, max_tokens=12000, budget_tokens=9000, model=OPUS_MODEL):
        if block_type == "thinking":
            thinking_buf += chunk
            thinking_box.caption(f"Thinking… ({len(thinking_buf):,} chars)")
        else:
            text_buf += chunk
            tail = text_buf[-500:]
            output_box.markdown(f"```\n…{tail}\n```")

    thinking_box.empty()
    output_box.empty()
    return text_buf.strip()


def run_step2(excel_text: str, analysis: str, template_html: str, status_widget) -> str:
    """Generate the final HTML dashboard."""
    prompt = GENERATION_PROMPT.format(
        template_html=template_html,
        analysis=analysis,
        excel_text=excel_text,
    )
    thinking_buf, html_buf = "", ""

    thinking_box = status_widget.empty()
    counter_box  = status_widget.empty()
    preview_box  = status_widget.empty()

    for block_type, chunk in bedrock_stream(prompt, max_tokens=20000, budget_tokens=6000, model=OPUS_MODEL):
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


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("Settings")

    uploaded = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

    templates = list_templates()
    if templates:
        template = st.selectbox("Select template", templates)
    else:
        st.warning("No HTML templates found in /template/")
        template = None

    generate = st.button(
        "Generate Dashboard",
        type="primary",
        disabled=not (uploaded and template),
    )

# ── Main ──────────────────────────────────────────────────────────────────────
st.title("Research Dashboard Builder")

if not uploaded:
    st.info("Upload an Excel file and select a template in the sidebar to get started.")
    st.stop()

df = pd.read_excel(uploaded)

with st.expander(
    f"Data preview — {uploaded.name}  ({len(df):,} rows · {len(df.columns)} columns)",
    expanded=False,
):
    st.dataframe(df.head(20), use_container_width=True)

if generate:
    excel_text    = excel_to_text(uploaded)
    template_html = (TEMPLATE_DIR / template).read_text(encoding="utf-8")

    try:
        # ── Step 1 ────────────────────────────────────────────────────────────
        with st.status("Step 1 / 2 — Analysing research data (Opus 4.6)…", expanded=True) as s1:
            analysis = run_step1(excel_text, s1)
            s1.update(label="Step 1 / 2 — Analysis complete ✓", state="complete", expanded=False)

        with st.expander("View extracted analysis", expanded=False):
            st.text(analysis)

        # ── Step 2 ────────────────────────────────────────────────────────────
        with st.status("Step 2 / 2 — Generating dashboard HTML (Opus 4.6)…", expanded=True) as s2:
            html = run_step2(excel_text, analysis, template_html, s2)
            s2.update(label="Step 2 / 2 — Dashboard ready ✓", state="complete", expanded=False)

        st.session_state["dashboard_html"]     = html
        st.session_state["dashboard_label"]    = template
        st.session_state["dashboard_analysis"] = analysis

    except Exception as e:
        st.error(f"Generation failed: {e}")

if "dashboard_html" in st.session_state:
    html  = st.session_state["dashboard_html"]
    label = st.session_state["dashboard_label"]

    col_info, col_save, col_dl = st.columns([5, 1, 1])
    with col_info:
        st.success(f"Template: **{label}**")
    with col_save:
        if st.button("💾 Save", use_container_width=True):
            from datetime import datetime
            stem = Path(label).stem
            fname = f"{stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
            save_path = Path(__file__).parent / "saved" / fname
            save_path.write_text(html, encoding="utf-8")
            st.toast(f"Saved → saved/{fname}", icon="✅")
    with col_dl:
        st.download_button(
            "⬇ Download",
            data=html,
            file_name=f"{Path(label).stem}.html",
            mime="text/html",
            use_container_width=True,
        )

    components.html(html, height=950, scrolling=True)
