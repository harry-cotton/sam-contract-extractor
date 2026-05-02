# SAM.gov BD Pipeline

A pipeline I built using the Anthropic Claude API to automate the early stages of a government contracting BD workflow — from reading a raw solicitation all the way through to a near-finished proposal draft.

This started as a simple data extractor and kept growing as I realised each output could feed the next step.

## The Problem

Responding to government soliciations is currently very manual and time consuming. Sometimes the solicication docs are long, complicated and contain small inconsistencies. The docs need to be read, cross referenced and then matched to past performance examples, all before writers can start crafting a compelling response.

This pipeline automates that work, from first read to first draft.

## How It Works

Three scripts, each building on the last:

**Step 1 — Triage the opportunity**
- **`extractor.py`** — single document (paste the SAM.gov text into a `.txt` file and run it)
- **`extractor_multi.py`** — full solicitation package (point it at a folder containing PDFs, Word docs, Excel files, or text files — it reads them all and synthesises across the package)

Both produce a structured JSON extraction and a console summary.

**Step 2 — Draft the past performance volume**
- **`pp_writer.py`** — reads the extraction JSON and a folder of past performance PowerPoint decks, selects the most relevant examples, and drafts a standalone past performance volume with citations tied to the evaluation criteria

**Step 3 — Draft the full proposal**
- **`proposal_writer.py`** — reads the extraction JSON and the same PP decks, then writes a complete near-finished proposal response structured around the proposal outline the extractor identified — with past performance woven in as evidence throughout, not bolted on at the end

The output from Step 3 is a markdown file. I drop it into a Claude chat and it converts it to a clean Word doc in one shot.

## What Gets Extracted

The schema is based on a real govcon BD qualification checklist — the minimum a capture manager needs for a go/no-bid call:

- **Opportunity Overview**: Title, solicitation number, agency, notice type, deadlines
- **Contract Details**: NAICS, contract type, set-aside, estimated value, period of performance, place of performance
- **Requirements Summary**: Scope, key technical requirements, clearance needs, key personnel, deliverables
- **Evaluation Criteria**: LPTA vs. Best Value, factors in priority order, submission requirements
- **Competitive Intelligence**: Incumbent, predecessor contract, competitive landscape, notable terms
- **BD Action Items**: Pursuit recommendation, critical dates, risks and red flags
- **Proposal Outline**: Win themes and a section-by-section response outline mapping requirements to proposal volumes — so the JSON output feeds directly into the proposal scripts

## Setup

### Prerequisites
- Python 3.10+
- An Anthropic API key

### Installation

```bash
pip install anthropic pdfplumber python-docx openpyxl python-pptx
```

### Set your API key

On Windows:
```bash
set ANTHROPIC_API_KEY=your-api-key-here
```

On Mac/Linux:
```bash
export ANTHROPIC_API_KEY=your-api-key-here
```

### Configure your firm

Fill in `firm_config.json` once with your firm's details. The proposal scripts use this to write from your firm's perspective throughout:

```json
{
    "firm_name": "Your Firm Name",
    "firm_description": "...",
    "firm_capabilities": "...",
    "certifications": ["SDVOSB"],
    "naics_codes": ["541810"]
}
```

## Usage

**Single document extraction:**
```bash
python extractor.py "Sample Contracts\cms_analytics_solicitation.txt"
```

**Full solicitation package (PDF, Word, Excel, text):**
```bash
python extractor_multi.py "Sample Contracts\US_Secret Service"
```

**Past performance volume:**
```bash
python pp_writer.py "Sample Contracts/US_Secret Service/US_Secret Service_full_extraction.json" "Past Performance"
```

**Full proposal draft:**
```bash
python proposal_writer.py "Sample Contracts/US_Secret Service/US_Secret Service_full_extraction.json" "Past Performance"
```

## Sample Output

See the `Sample Contracts/` folder for example inputs, extracted JSON files, and generated draft outputs — including a US Secret Service recruitment advertising RFQ, a Peace Corps AI chatbot solicitation, and a CMS analytics opportunity.

## Design Notes

The extraction prompt frames Claude as a senior govcon BD analyst because procurement terminology matters. Terms like "LPTA," "set-aside," and "Best Value Tradeoff" have precise meanings that a generic summarizer would miss.

A few choices I made deliberately:

- **"Extract only what is explicitly stated"** — prevents Claude from inferring details that aren't in the document, which matters when the output informs a real bid decision
- **JSON as the handoff format** — the extraction JSON feeds the proposal scripts directly; every field the proposal writer needs is already structured and named
- **Proposal outline in the extraction** — by identifying win themes and section-level guidance at triage time, the extractor does the strategic thinking once, and the proposal scripts just execute against it
- **Firm config as a separate file** — fill it in once, and every proposal draft is written from your firm's perspective without touching the code
- **Banned words list in the prompts** — compliance flags like "ensure" and "ensuring" are blocked at the prompt level so they never appear in output

## Potential Extensions

- **SAM.gov API integration**: Pull opportunities directly from the SAM.gov public API instead of manual copy-paste
- **Bid/no-bid scoring**: Evaluate each extraction against configurable criteria (firm NAICS codes, clearance capabilities, set-aside eligibility) and return a scored recommendation
- **Comparison mode**: Diff a sources sought notice against the subsequent RFP to surface what changed between releases

## Author

Harry Cotton — built as a portfolio project to explore how Claude Code can automate real day-to-day govcon BD work using the Anthropic Claude API.
