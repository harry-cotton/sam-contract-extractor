 # SAM.gov Contract Data Extractor

  A tool I built using the Anthropic Claude API to extract structured data
  from federal contract solicitations sourced from SAM.gov — turning
  50-100 page documents into a clean JSON record in seconds.

  ## The Problem

  BD teams spend hours manually triaging SAM.gov
  solicitations. A single opportunity can run 50-100+ pages, and a busy
  team may need to review dozens per week. The question is always the same:
  **should we bid on this?**

  Answering that quickly means pulling the same set of fields every time —
  agency, contract type, NAICS code, set-aside, estimated value, period of
  performance, clearance requirements, evaluation criteria, incumbent. From
  documents that vary wildly in format.

  This tool automates that first pass.

  ## How It Works

  Two scripts depending on what you're working with:

  - **`extractor.py`** — single document (paste the SAM.gov text into a
    `.txt` file and run it)
  - **`extractor_multi.py`** — full solicitation package (point it at a
    folder of `.txt` files — synopsis, PWS, Section L/M, amendments — and
    it synthesizes across all of them)

  Both print a formatted summary to the console and save a full JSON
  extraction to disk.

  ## What Gets Extracted

  The schema is based on a real govcon BD qualification checklist — the
  minimum a capture manager needs for a go/no-bid call:

  - **Opportunity Overview**: Title, solicitation number, agency, notice
    type, deadlines
  - **Contract Details**: NAICS, contract type, set-aside, estimated value,
    period of performance, place of performance
  - **Requirements Summary**: Scope, key technical requirements, clearance
    needs, key personnel, deliverables
  - **Evaluation Criteria**: LPTA vs. Best Value, factors in priority order,
    submission requirements
  - **Competitive Intelligence**: Incumbent, predecessor contract,
    competitive landscape, notable terms
  - **BD Action Items**: Pursuit recommendation, critical dates, risks and
    red flags
  - **Proposal Outline**: A short response outline mapping key requirements
    to proposal sections, with win themes — so the output is immediately
    actionable, not just informational

  ## Setup

  ### Prerequisites
  - Python 3.10+
  - An Anthropic API key

  ### Installation

  ```bash
  pip install anthropic

  Set your API key

  On Windows:
  set ANTHROPIC_API_KEY=your-api-key-here

  On Mac/Linux:
  export ANTHROPIC_API_KEY=your-api-key-here

  Usage

  Single document:
  python extractor.py "Sample Contracts\cms_analytics_solicitation.txt"

  Full solicitation package:
  python extractor_multi.py "Sample Contracts\Peacecorps_package"

  Sample Output

  See the Sample Contracts/ folder for example inputs and extracted JSON
  outputs — including a real Peace Corps AI chatbot RFQ and a CMS analytics
  solicitation.

  Design Notes

  The system prompt frames Claude as a senior govcon BD analyst because
  procurement terminology matters. Terms like "LPTA," "set-aside," and
  "Best Value Tradeoff" have precise meanings that a generic summarizer
  would flatten or miss entirely.

  A few choices I made deliberately:

  - "Extract only what is explicitly stated" — prevents Claude from
  inferring details that aren't in the document, which matters when the
  output informs a real bid decision
  - JSON output — every extraction follows the same schema regardless
  of how the source document is formatted, so it can feed directly into a
  CRM, SharePoint tracker, or pipeline dashboard
  - Proposal outline section — goes beyond extraction to give a BD team
  a starting point for structuring their response, with requirements
  mapped to each proposal volume
  - JSON fallback — if the response isn't valid JSON, the script saves
  the raw output rather than crashing

  Potential Extensions

  - SAM.gov API integration: Pull opportunities directly from the
  SAM.gov public API instead of manual copy-paste
  - Bid/no-bid scoring: Evaluate each opportunity against configurable
  criteria (company NAICS codes, clearance capabilities, past performance)
  - Comparison mode: Diff a sources sought notice against the subsequent
  RFP to surface what changed

  Author

  Harry Cotton — portfolio project demonstrating government AI applications
  using the Anthropic Claude API.
