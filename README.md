# SAM.gov Contract Data Extractor

A proof-of-concept tool that uses the Anthropic Claude API to extract structured data from unstructured federal contract opportunity text sourced from SAM.gov.

## The Problem

Government contracting Business Development (BD) teams spend significant time manually reading through SAM.gov solicitations, sources sought notices, and RFIs to extract the key data points needed for pursuit decisions. A single solicitation can run 50-100+ pages, and a busy BD team may need to triage dozens of new opportunities per week.

The critical question is always the same: **should we bid on this?** Answering that question requires quickly extracting a standard set of fields — agency, contract type, NAICS code, set-aside status, estimated value, period of performance, clearance requirements, evaluation criteria, and incumbent information — from documents that vary wildly in format and structure.

This tool automates that first-pass extraction, producing structured JSON output that could feed into a CRM, pipeline tracker, or bid/no-bid scoring matrix.

## How It Works

1. Copy the text of a SAM.gov opportunity into a `.txt` file
2. Run the extractor — Claude analyzes the full text using a prompt designed by someone who understands govcon BD workflows
3. The tool outputs both a human-readable summary (printed to console) and a structured JSON file (saved to disk)

## What Gets Extracted

The extraction schema mirrors a standard govcon opportunity qualification checklist:

- **Opportunity Overview**: Title, solicitation number, agency, notice type, deadlines
- **Contract Details**: NAICS, contract type, set-aside, estimated value, period of performance, place of performance
- **Requirements Summary**: Scope overview, key technical requirements, clearance needs, key personnel
- **Evaluation Criteria**: Evaluation method (LPTA vs. Best Value), factors in order of importance
- **Competitive Intelligence**: Incumbent contractor, predecessor contract, competitive landscape indicators
- **BD Action Items**: Pursuit positioning, critical dates, risks and red flags

These fields were chosen because they represent the minimum information a capture manager needs to make an informed go/no-go recommendation. The schema design reflects real-world govcon BD workflows, not a generic document parser.

## Setup

### Prerequisites
- Python 3.10+
- An Anthropic API key ([get one here](https://console.anthropic.com))

### Installation

```bash
pip install anthropic
```

### Set your API key

On Windows (Command Prompt):
```bash
set ANTHROPIC_API_KEY=your-api-key-here
```

On Mac/Linux:
```bash
export ANTHROPIC_API_KEY=your-api-key-here
```

## Usage

```bash
python extractor.py "Sample Contracts\cms_analytics_solicitation.txt"
```

The tool prints a formatted summary to the console and saves the full JSON extraction alongside the original file with an `_extracted.json` suffix.

## Sample Output

See the `Sample Contracts/` folder for example inputs and their extracted JSON outputs.

## Why JSON Output?

Unlike Project A (the Policy Brief Summarizer), which produces human-readable markdown summaries, this tool outputs structured JSON. The reasoning:

- **Machine-readable**: JSON output can be programmatically ingested into a CRM (Salesforce, Pipedrive), a SharePoint tracker, or a custom pipeline dashboard
- **Consistent schema**: Every extraction follows the same structure regardless of how the source document is formatted, enabling apples-to-apples comparison across opportunities
- **Composable**: A downstream system could aggregate extractions to identify trends (e.g., "show me all IDIQ opportunities over $50M with AI/ML requirements posted in the last 90 days")
- **Auditable**: The JSON serves as a structured record that a BD analyst can quickly verify against the source document

## Prompt Design Notes

The system prompt establishes Claude as a senior govcon BD analyst because procurement terminology matters. Terms like "LPTA," "set-aside," "IDIQ ceiling," and "FWA detection" have precise meanings in this domain. A generic summarizer would miss nuances that a BD professional catches immediately — for example, that an "unrestricted" set-aside means large businesses can compete, or that a "Best Value Tradeoff" evaluation means the government can pay more for a stronger technical approach.

Key design choices:
- **"Extract only what is explicitly stated"**: Prevents Claude from inferring or hallucinating details that aren't in the document — critical when the output informs a bid decision
- **Competitive intelligence section**: BD teams always want to know who the incumbent is and what the predecessor contract looked like. This information is often buried in the document
- **BD action items**: Goes beyond extraction to provide actionable next steps, making the output immediately useful rather than just informational
- **JSON format with fallback**: If Claude's response isn't valid JSON (rare but possible), the script catches the error and still saves the raw output rather than crashing

## Potential Extensions

- **SAM.gov API integration**: Pull opportunities directly from the SAM.gov public API instead of manual copy-paste
- **Batch processing**: Extract data from an entire folder of solicitations for pipeline review
- **Bid/no-bid scoring**: Add a scoring module that evaluates the opportunity against configurable criteria (company NAICS codes, clearance capabilities, past performance areas)
- **Comparison mode**: Compare two related solicitations (e.g., a sources sought notice vs. the subsequent RFP) to identify changes

## Author

Harry Cotton — built as a portfolio project demonstrating government-relevant AI applications using the Anthropic Claude API.
