"""
Proposal Writer
===============
Takes the JSON extraction output from extractor_multi.py and a folder of
past performance PowerPoint decks, then uses Claude to produce a near-finished
full proposal response structured around the proposal outline identified in
the extraction JSON.

Unlike pp_writer.py (which drafts the past performance volume in isolation),
this script produces a complete proposal draft -- each volume written as
near-finished prose, with past performance examples woven in as evidence
supporting the technical and management claims. The extraction JSON is the
blueprint; the PP decks are the proof.

Usage:
    python proposal_writer.py <extraction_json> <past_performance_folder>

Example:
    python proposal_writer.py "Sample Contracts/US_Secret Service/US_Secret Service_full_extraction.json" "Past Performance"

Output:
    A markdown file saved alongside the extraction JSON, structured by
    proposal volume and ready for staff review and light editing.

Required libraries:
    pip install anthropic python-pptx

Author: Harry Cotton
"""

import anthropic
import sys
import os
import json
from pathlib import Path

# Shared utilities — see utils.py for implementations
from utils import (
    markdown_to_docx,
    extract_text_from_pptx,
    read_past_performance_decks,
    format_decks_for_prompt,
    check_banned_words,
    warn_on_banned_words,
    compute_word_budget,
)


# ============================================================
# STEP 2: LOAD FIRM CONFIG
# ============================================================

def load_firm_config() -> dict:
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "firm_config.json")

    if not os.path.exists(config_path):
        print("Warning: firm_config.json not found. Using placeholder firm details.")
        return {
            "firm_name": "Our Firm",
            "firm_description": "a government contracting firm",
            "firm_capabilities": "consulting, program management, and technology services",
            "certifications": [],
            "naics_codes": []
        }

    with open(config_path, "r", encoding="utf-8") as f:
        config = json.load(f)

    print(f"Firm: {config.get('firm_name', 'Unknown')}")
    return config


# ============================================================
# STEP 3: BUILD FULL SOLICITATION CONTEXT FROM EXTRACTION JSON
# ============================================================

def build_full_solicitation_context(extraction: dict) -> str:
    """
    Pulls everything from the extraction JSON that informs proposal strategy:
    overview, requirements, eval criteria, competitive intelligence, risks,
    and the full proposal outline with section-level guidance.
    """
    overview = extraction.get("opportunity_overview", {})
    details = extraction.get("contract_details", {})
    requirements = extraction.get("requirements_summary", {})
    section_l = extraction.get("section_l", {})
    section_m = extraction.get("section_m", {})
    intel = extraction.get("competitive_intelligence", {})
    actions = extraction.get("bd_action_items", {})
    outline = extraction.get("proposal_outline", {})

    sub_req = section_l.get("submission_requirements", {})
    factors = section_m.get("factors", [])
    factor_lines = []
    for f in factors:
        if isinstance(f, dict):
            line = f"  - {f.get('name', 'N/A')} [{f.get('weight', 'N/A')}]"
            subs = f.get("subfactors", [])
            if subs:
                line += "\n" + "\n".join(f"      * {s}" for s in subs)
            factor_lines.append(line)
        else:
            factor_lines.append(f"  - {f}")

    context = f"""OPPORTUNITY: {overview.get('title', 'N/A')}
AGENCY: {overview.get('agency', 'N/A')}
SOLICITATION #: {overview.get('solicitation_number', 'N/A')}
RESPONSE DEADLINE: {overview.get('response_deadline', 'N/A')}
CONTRACT TYPE: {details.get('contract_type', 'N/A')}
SET-ASIDE: {details.get('set_aside', 'N/A')}
ESTIMATED VALUE: {details.get('estimated_value', 'N/A')}
PERIOD OF PERFORMANCE: {details.get('period_of_performance', 'N/A')}

SCOPE:
{requirements.get('scope_overview', 'N/A')}

KEY REQUIREMENTS:
{chr(10).join(f'  - {r}' for r in requirements.get('key_requirements', []))}

SECTION M — EVALUATION METHOD: {section_m.get('evaluation_method', 'N/A')}

SECTION M — EVALUATION FACTORS (in order of importance):
{chr(10).join(factor_lines) if factor_lines else '  Not specified'}

SECTION L — SUBMISSION REQUIREMENTS:
  Page limit (total): {sub_req.get('page_limit_total', 'Not specified')}
  Page limits by volume: {sub_req.get('page_limits_by_volume', 'Not specified')}
  Font: {sub_req.get('font', 'Not specified')}
  Line spacing: {sub_req.get('line_spacing', 'Not specified')}
  File format: {sub_req.get('file_format', 'Not specified')}
  Copies: {sub_req.get('copies', 'Not specified')}
  Submission method: {sub_req.get('submission_method', 'Not specified')}

SECTION L — REQUIRED VOLUMES:
{chr(10).join(f'  - {v}' for v in section_l.get('volume_structure', [])) or '  Not specified'}

SECTION L — SIGNING REQUIREMENTS:
{section_l.get('signing_requirements', 'Not specified')}

COMPETITIVE INTELLIGENCE:
  Incumbent: {intel.get('incumbent', 'N/A')}
  Competition: {intel.get('estimated_competition', 'N/A')}
  Notable Terms: {intel.get('notable_terms', 'N/A')}

RISKS AND FLAGS TO ADDRESS IN THE PROPOSAL:
{chr(10).join(f'  - {r}' for r in actions.get('risks_and_flags', []))}

PURSUIT RECOMMENDATION:
{actions.get('pursuit_recommendation', 'N/A')}

WIN THEMES (thread throughout every section):
{chr(10).join(f'  - {t}' for t in outline.get('win_themes', []))}

PROPOSAL OUTLINE — WRITE EACH SECTION BELOW:
"""

    for i, section in enumerate(outline.get("sections", []), 1):
        context += f"""
SECTION {i}: {section.get('title', 'Untitled')}
  Requirements this section must address:
{chr(10).join(f'    - {r}' for r in section.get('requirements_addressed', []))}
  Guidance: {section.get('guidance', 'N/A')}
"""

    return context


# ============================================================
# STEP 4: DRAFT THE FULL PROPOSAL
# ============================================================

def write_full_proposal(extraction: dict, decks: list, firm_config: dict) -> str:
    """
    Sends the full solicitation context and PP decks to Claude.
    Claude writes a complete, near-finished proposal response structured
    around the proposal outline from the extraction JSON, with PP examples
    woven in as evidence throughout.

    Computes an explicit target word budget from the Section L page limit
    and passes it to the model so page discipline is a concrete number, not
    a fuzzy instruction.
    """
    solicitation_context = build_full_solicitation_context(extraction)
    decks_text = format_decks_for_prompt(decks)
    firm_name = firm_config.get("firm_name", "Our Firm")
    firm_description = firm_config.get("firm_description", "")
    firm_capabilities = firm_config.get("firm_capabilities", "")
    certs = ", ".join(firm_config.get("certifications", []))

    # Compute an explicit target word budget from the Section L page limit.
    # This converts "respect the page limit" (fuzzy) into "aim for ~7000 words"
    # (concrete). System-prompt enforcement is unreliable on fuzzy targets.
    sub_req = extraction.get("section_l", {}).get("submission_requirements", {})
    page_limit_total = sub_req.get("page_limit_total")
    word_budget = compute_word_budget(page_limit_total)
    if word_budget:
        budget_instruction = (
            f"TARGET WORD BUDGET: ~{word_budget:,} words across all evaluated content "
            f"(derived from a page limit of {page_limit_total} at ~600 words per page in "
            f"10pt Times New Roman / 1-inch margins / single-spaced). "
            f"Tables and headings count against this budget; calibrate prose accordingly. "
            f"Going significantly over this budget will produce a non-compliant draft."
        )
    else:
        budget_instruction = (
            "TARGET WORD BUDGET: No page limit specified in the extraction. Write at a "
            "reasonable length for federal proposal evaluators - tight, evidence-backed, "
            "no padding."
        )

    client = anthropic.Anthropic()

    system_prompt = f"""You are a senior proposal writer at {firm_name} with 15 years of experience
writing winning federal proposals. You write near-finished, submission-ready prose
that staff can lightly tune -- not templates, not outlines, not placeholders.

ABOUT {firm_name.upper()}:
{firm_name} is {firm_description}
Core capabilities: {firm_capabilities}
Certifications: {certs if certs else 'See firm profile'}
{firm_name} has the depth and breadth of a Big Four consulting firm, with a proven
track record delivering complex, mission-critical programs for federal agencies.

YOUR TASK:
Write a complete proposal response to the solicitation described below. Structure
the response exactly according to the PROPOSAL OUTLINE provided -- one section per
proposal volume. Each section must:

1. Be written as polished, near-finished proposal prose from {firm_name}'s perspective
   using "we," "our team," and "{firm_name}" throughout
2. Directly address every requirement mapped to that section in the proposal outline
3. Weave in specific past performance examples from the provided decks as proof points
   -- cite them inline where they demonstrate a claimed capability, not as a separate
   appendix or afterthought
4. Thread the win themes throughout -- every section should reinforce the same
   overarching differentiators
5. Proactively address any relevant risks or flags identified in the solicitation
   analysis where doing so strengthens the proposal

WRITING STYLE -- PAGE-ECONOMY DISCIPLINE:
Section L page limits are non-negotiable. Every line costs you. Write to the
following discipline:
- Structure each section as: name the requirement, state what {firm_name} will do,
  back the claim with PP evidence -- in that order, in flowing prose. Do not write
  "executive summary" buildups that restate the requirement before answering it.
- Lead with action: "{firm_name} will deliver..." or "We provide..." -- not "It is
  our intention to..." or "The platform supports..."
- A specific TARGET WORD BUDGET is stated at the top of the user prompt. Treat it
  as a hard target, not a suggestion. Distribute the budget across volumes
  proportional to their declared page limits. Drafts exceeding the budget by
  more than ~10% are non-compliant and will need to be rewritten.
- Tables and small frameworks are encouraged where they save space versus prose
  (implementation timelines, KPI tracking, channel mix, compliance matrices,
  staffing matrices). Keep them tight: 4-6 columns max, 5-10 rows max.

PAST PERFORMANCE INTEGRATION:
- Cite PP inline within the prose of the Technical and Management sections, in the
  same sentence or paragraph as the claim it supports. Do NOT create dedicated
  "Relevance to evaluation criteria" sub-paragraphs in the Technical volume. Example:
  "We will maintain a 99.9% uptime SLA -- the same commitment we have delivered to
  SSA at a measured 99.97% over Year 1."
- In the dedicated Past Performance volume (if one exists in the outline), write
  structured narratives in standard govcon format, but keep them tight.
- Every capability claim should be backed by at least one reference to a specific
  past project, metric, or outcome from the PP decks.
- DO NOT invent operational specifics (durations, effort estimates, cost figures,
  performance percentages, headcount, time-to-value claims) that are not in the
  source PP decks. If a specific number is needed and not in the decks, write
  [TO BE CONFIRMED] in its place. Fabricated operational specifics are a
  disqualifying failure mode.

BANNED WORDS -- never use these anywhere in the output:
- "ensure" or "ensuring"
If you would naturally write either word, rewrite the sentence to avoid it
entirely. Do not substitute a hollow synonym.

FORMAT:
- Heading depth: TWO LEVELS MAX. Use "## 1.0 Section Title" and optionally
  "### 1.1 Subsection Title". DO NOT use three-level headings (1.1.1) or deeper.
- Do not use horizontal-rule dividers (---) inside a volume's body. Use them only
  between top-level volumes.
- Match the volume titles from the outline.
- Write in first-person plural ("we / our / {firm_name}").
- Active voice. Specific. Confident.
- Flag missing data as [TO BE CONFIRMED].
- After all proposal sections, include a brief "Proposal Team Checklist" of the
  top 5 items staff must verify or complete before submission."""

    user_prompt = f"""Write a complete, near-finished proposal response for {firm_name} using the
solicitation context and past performance decks below.

{budget_instruction}

Structure the response exactly according to the PROPOSAL OUTLINE sections listed in
the solicitation context. For each section, write polished proposal prose that
addresses the mapped requirements, weaves in relevant PP examples as proof, and
threads the win themes throughout.

This should read like a draft that needs light editing -- not a template.

--- SOLICITATION CONTEXT AND PROPOSAL OUTLINE ---
{solicitation_context}

--- PAST PERFORMANCE DECKS ---
{decks_text}

Write the complete proposal response now, section by section, followed by the
Proposal Team Checklist. Stay within the TARGET WORD BUDGET stated above."""

    # Stream the response. At max_tokens=32000, a generation can exceed the
    # SDK's 10-minute non-streaming timeout window, so streaming is required.
    # Bonus: the draft prints live in the terminal as it's produced, giving
    # visible progress during a 1-3 minute generation.
    print()
    print("=" * 60)
    print("STREAMING DRAFT (live):")
    print("=" * 60)

    chunks = []
    with client.messages.stream(
        model="claude-sonnet-4-6",
        max_tokens=32000,
        system=system_prompt,
        messages=[{"role": "user", "content": user_prompt}],
    ) as stream:
        for text_delta in stream.text_stream:
            print(text_delta, end="", flush=True)
            chunks.append(text_delta)

    print()  # final newline after streamed content
    return "".join(chunks)


# ============================================================
# STEP 5: MAIN
# ============================================================

def main():
    if len(sys.argv) < 3:
        print("Usage: python proposal_writer.py <extraction_json> <past_performance_folder>")
        print()
        print("Example:")
        print('  python proposal_writer.py "Sample Contracts/US_Secret Service/US_Secret Service_full_extraction.json" "Past Performance"')
        sys.exit(1)

    json_path = sys.argv[1]
    pp_folder = sys.argv[2]

    if not os.path.exists(json_path):
        print(f"Error: Extraction JSON not found: {json_path}")
        sys.exit(1)

    if not os.path.isdir(pp_folder):
        print(f"Error: Past performance folder not found: {pp_folder}")
        sys.exit(1)

    # Load firm config and extraction JSON
    firm_config = load_firm_config()

    with open(json_path, "r", encoding="utf-8") as f:
        extraction = json.load(f)

    opportunity_title = extraction.get("opportunity_overview", {}).get("title", "Unknown Opportunity")

    print(f"Opportunity:  {opportunity_title}")
    print(f"Extraction:   {json_path}")
    print(f"PP folder:    {pp_folder}")
    print("-" * 60)

    # Read PP decks
    decks = read_past_performance_decks(pp_folder)
    print()

    # Draft the full proposal
    print("Drafting full proposal response...\n")
    draft = write_full_proposal(extraction, decks, firm_config)

    # Post-generation compliance check - banned words must not appear in
    # federal proposal output. System-prompt enforcement is unreliable so we
    # check deterministically and warn loudly.
    findings = check_banned_words(draft, label="Proposal Draft")
    warn_on_banned_words(findings, label="Proposal Draft")

    json_dir = os.path.dirname(os.path.abspath(json_path))
    stem = Path(json_path).stem.replace("_full_extraction", "")
    output_path = os.path.join(json_dir, f"{stem}_proposal_draft.docx")

    firm_name = firm_config.get("firm_name", "Unknown")
    markdown_to_docx(
        draft_text=draft,
        title="Proposal Draft",
        subtitle_lines=[
            f"Opportunity: {opportunity_title}",
            f"Firm: {firm_name}",
            f"Source: {os.path.basename(json_path)}",
        ],
        output_path=output_path,
    )

    # The draft was already streamed live to the terminal above;
    # no need to reprint it.
    print()
    print("=" * 60)
    print(f"Draft saved to: {output_path}")


if __name__ == "__main__":
    main()
