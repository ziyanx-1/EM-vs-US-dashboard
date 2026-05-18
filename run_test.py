"""
Standalone pipeline test — runs Step 1 (Opus) + Step 2 (Sonnet) and saves the report.
Run with: python3 run_test.py
"""
import json
import re
import sys
import traceback
import pandas as pd
import boto3
from botocore.config import Config
from pathlib import Path

BASE_DIR       = Path(__file__).parent
EM_US_EXCEL    = BASE_DIR / "External Research Summary - EM vs US V2 (Filtered).xlsx"
EM_US_TEMPLATE = BASE_DIR / "template" / "EM_vs_US_Dashboard_V3_Final.html"
REPORTS_DIR    = BASE_DIR / "saved"
OPUS_MODEL     = "us.anthropic.claude-opus-4-6-v1"
SONNET_MODEL   = "us.anthropic.claude-sonnet-4-20250514-v1:0"

client = boto3.client(
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


def bedrock_stream(prompt: str, max_tokens: int, budget_tokens: int, model: str):
    body = json.dumps({
        "anthropic_version": "bedrock-2023-05-31",
        "max_tokens": max_tokens,
        "thinking": {"type": "enabled", "budget_tokens": budget_tokens},
        "messages": [{"role": "user", "content": [{"type": "text", "text": prompt}]}],
    })
    resp = client.invoke_model_with_response_stream(
        modelId=model, body=body,
        contentType="application/json", accept="application/json",
    )
    for event in resp["body"]:
        chunk = json.loads(event["chunk"]["bytes"])
        if chunk.get("type") == "content_block_delta":
            delta = chunk["delta"]
            if delta.get("type") == "thinking_delta":
                yield "thinking", delta.get("thinking", "")
            elif delta.get("type") == "text_delta":
                yield "text", delta.get("text", "")


# Pull prompts from mas_analytics.py
src = (BASE_DIR / "mas_analytics.py").read_text()
ANALYSIS_PROMPT   = re.search(r'ANALYSIS_PROMPT\s*=\s*"""\\\n(.*?)\n"""', src, re.DOTALL).group(1)
GENERATION_PROMPT = re.search(r'GENERATION_PROMPT\s*=\s*"""\\\n(.*?)\n"""', src, re.DOTALL).group(1)

print(f"ANALYSIS_PROMPT:   {len(ANALYSIS_PROMPT):,} chars")
print(f"GENERATION_PROMPT: {len(GENERATION_PROMPT):,} chars")
sys.stdout.flush()

# Load data
print("\nReading Excel...")
sys.stdout.flush()
excel_text = excel_to_text(EM_US_EXCEL)
print(f"Excel text: {len(excel_text):,} chars")
sys.stdout.flush()

# Step 1 — Opus analysis
print("\n=== STEP 1: Analysis (Opus 4.6) ===")
sys.stdout.flush()
prompt1 = ANALYSIS_PROMPT.format(excel_text=excel_text)
thinking_chars, text_buf = 0, ""
for block_type, chunk in bedrock_stream(prompt1, max_tokens=12000, budget_tokens=9000, model=OPUS_MODEL):
    if block_type == "thinking":
        thinking_chars += len(chunk)
    else:
        text_buf += chunk
        if len(text_buf) % 500 == 0:
            print(f"  Analysis: {len(text_buf):,} chars written...")
            sys.stdout.flush()
analysis = text_buf.strip()
print(f"Step 1 complete — {len(analysis):,} chars, thinking used {thinking_chars:,} chars")
print("\n--- Analysis (first 1000 chars) ---")
print(analysis[:1000])
print("---")
sys.stdout.flush()

# Step 2 — Sonnet HTML generation
print("\n=== STEP 2: HTML generation (Sonnet 4.6) ===")
sys.stdout.flush()
template_html = EM_US_TEMPLATE.read_text(encoding="utf-8")
prompt2 = GENERATION_PROMPT.format(
    template_html=template_html, analysis=analysis, excel_text=excel_text,
)
html_buf = ""
for block_type, chunk in bedrock_stream(prompt2, max_tokens=20000, budget_tokens=6000, model=SONNET_MODEL):
    if block_type != "thinking":
        html_buf += chunk
        if len(html_buf) % 2000 == 0:
            print(f"  HTML: {len(html_buf):,} chars written...")
            sys.stdout.flush()
html = strip_code_fence(html_buf)
print(f"Step 2 complete — {len(html):,} chars")
sys.stdout.flush()

# Save
REPORTS_DIR.mkdir(exist_ok=True)
out = REPORTS_DIR / "US_EAFE_EM.html"
out.write_text(html, encoding="utf-8")
print(f"\nSaved → {out}")
print(f"Valid DOCTYPE: {html.strip().startswith('<!DOCTYPE')}")
print(f"Ends </html>:  {html.strip().endswith('</html>')}")
