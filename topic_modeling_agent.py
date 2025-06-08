#!/usr/bin/env python3
import os
import sys
import json
import argparse
import pdfplumber
import pandas as pd
from openai import OpenAI, OpenAIError

# Paths for external config specific to topic modeling
METHODS_FILE = "topic_methods.json"

# Load or initialize known methods from external JSON
if os.path.exists(METHODS_FILE):
    with open(METHODS_FILE, 'r', encoding='utf-8') as f:
        enum_methods = json.load(f)
else:
    enum_methods = {
        "AnalyzedFields": [],
        "TermPreprocessing": [],
        "Clustering": [],
        "ClusterAnalysis": []
    }

# Few-shot examples for prompting
few_shot_examples = [
    {
        "Document": "Topic modeling of seawater desalination research",
        "AnalyzedFields": "Title; Abstract; Keywords",
        "TermPreprocessing": "Lemmatization; Stop-word removal",
        "Clustering": "Latent Dirichlet Allocation (LDA); Dynamic Topic Modeling (DTM)",
        "ClusterAnalysis": "Topic identification; Trend visualization; Most cited articles",
        "PaperCount": "11,942 (2000–2024)",
        "TermCount": "500 terms"
    },
    {
        "Document": "Exploring themes in circulating tumor cell research",
        "AnalyzedFields": "Keywords",
        "TermPreprocessing": "Not specified",
        "Clustering": "BERTopic; Correlation Explanation (CorEx)",
        "ClusterAnalysis": "Topic quality evaluation; Connection analysis",
        "PaperCount": "1,747 (2008–2023)",
        "TermCount": "45 terms"
    }
]

# Initialize OpenAI client
api_key = os.getenv("OPENAI_API_KEY")
if not api_key:
    sys.exit("Error: OPENAI_API_KEY environment variable not set")
client = OpenAI(api_key=api_key)

# In-memory normalization map starts from loaded methods
normalize_map = {k: set(v) for k, v in enum_methods.items()}

def build_prompt(full_text: str) -> str:
    km_json = json.dumps(enum_methods, ensure_ascii=False)
    examples_txt = "\n".join(json.dumps(ex, ensure_ascii=False) for ex in few_shot_examples)
    return f"""
You are given the full text of a research article. Using KNOWN_METHODS:
{km_json}
And observing the examples:
{examples_txt}

Extract exactly this JSON structure:
{{
  "Document": "...",
  "AnalyzedFields": "...",
  "TermPreprocessing": "...",
  "Clustering": "...",
  "ClusterAnalysis": "...",
  "PaperCount": "...",
  "TermCount": "..."
}}
- Separate multiple items with semicolons (;).
- Match values exactly from KNOWN_METHODS lists when possible.
- Use "Not specified" only if the section truly has no data.
- If the analysis includes keywords, specify the type of keywords used (e.g. Author Keywords, WoS Keywords, Keywords Plus).
Article text:
{full_text}
"""

def process_pdf(pdf_path: str) -> dict:
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = "\n\n".join(page.extract_text() or "" for page in pdf.pages)
    except Exception as e:
        raise RuntimeError(f"PDF read error: {e}")

    prompt = build_prompt(text)
    try:
        resp = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=1500,
            temperature=0.0
        )
        content = resp.choices[0].message.content.strip()
    except OpenAIError as e:
        raise RuntimeError(f"OpenAI API error: {e}")

    if content.startswith("```"):
        content = "\n".join(line for line in content.splitlines() if not line.startswith("```"))

    try:
        record = json.loads(content)
    except json.JSONDecodeError:
        raise ValueError(f"Failed to parse JSON:\n{content}")

    for key in enum_methods.keys():
        raw = record.get(key, "Not specified")
        items = [i.strip() for i in raw.split(";") if i.strip()]
        normalized = []
        for it in items:
            if it.lower() == "not specified":
                normalized.append("Not specified")
            else:
                normalized.append(it)
                if it not in enum_methods[key]:
                    normalize_map[key].add(it)
        record[key] = "; ".join(normalized) if normalized else "Not specified"

    return record

def save_methods():
    with open(METHODS_FILE, 'w', encoding='utf-8') as f:
        json.dump({k: sorted(v) for k, v in normalize_map.items()}, f, ensure_ascii=False, indent=2)


def main():
    parser = argparse.ArgumentParser(
        description="Process PDFs for topic modeling and save metadata to Excel, replacing entries for existing files."
    )
    parser.add_argument("input", help="Path to PDF file or directory")
    parser.add_argument(
        "--excel", default="topic_metadata.xlsx",
        help="Output Excel file path (default: topic_metadata.xlsx)"
    )
    args = parser.parse_args()

    if os.path.isdir(args.input):
        pdfs = [os.path.join(args.input, f) for f in os.listdir(args.input) if f.lower().endswith(".pdf")]
    else:
        pdfs = [args.input]

    new_records = []
    for pdf in pdfs:
        try:
            rec = process_pdf(pdf)
            rec["SourceFile"] = os.path.basename(pdf)
            new_records.append(rec)
        except Exception as e:
            print(f"[Error] {pdf}: {e}", file=sys.stderr)

    # Load existing Excel if present
    if os.path.exists(args.excel):
        df_existing = pd.read_excel(args.excel)
    else:
        df_existing = pd.DataFrame()

    # Create DataFrame for new records
    df_new = pd.DataFrame(new_records)

    # Merge: drop existing entries for these files
    if not df_existing.empty:
        df_existing = df_existing[df_existing["SourceFile"].isin(df_new["SourceFile"]) == False]

    # Concatenate and save to Excel
    df_merged = pd.concat([df_existing, df_new], ignore_index=True)
    df_merged.to_excel(args.excel, index=False)
    print(f"Processed {len(new_records)} files. Excel saved to {args.excel}.")

    save_methods()

if __name__ == "__main__":
    main()
