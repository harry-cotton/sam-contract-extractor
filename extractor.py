"""
SAM.gov Contract Data Extractor
================================
A proof-of-concept tool that uses Claude to extract structured data from
unstructured federal contract opportunity text (solicitations, RFIs, 
sources sought notices, etc.) sourced from SAM.gov.

In government contracting, Business Development teams spend hours manually
reading through solicitations to extract key fields for capture decisions.
This tool automates that first-pass triage, producing a structured JSON
output that could feed into a CRM, pipeline tracker, or bid/no-bid matrix.

Built to demonstrate government-relevant AI use cases using the Anthropic API.

Author: Harry Cotton
"""

import anthropic
import sys
import os
import json


def extract_contract_data(file_path: str) -> str:
    """
    Reads raw solicitation text from a SAM.gov opportunity and extracts
    structured fields using Claude's API.

    Why these specific fields?
    --------------------------
    These are the fields a govcon BD team evaluates first when deciding
    whether to pursue an opportunity:

    - Opportunity Title & Solicitation Number: For tracking and reference
    - Agency / Office: Who is the buyer? Do we have a relationship?
    - Notice Type: Is this an RFI (just intelligence gathering) or a full
      solicitation (requires a proposal)?
    - NAICS Code & Set-Aside: Are we eligible? Is this small business set-aside?
    - Contract Type: FFP, T&M, Cost-Plus? Each carries different risk profiles
    - Period of Performance: Base year + options — what's the total commitment?
    - Place of Performance: On-site, remote, or specific location?
    - Response Deadline: How much time do we have to respond?
    - Estimated Value: Is this worth pursuing given our pipeline?
    - Key Requirements: What capabilities do we need to demonstrate?
    - Evaluation Criteria: How will proposals be scored?
    - Incumbent: Who holds this contract now? Are we displacing someone?
    - Clearance Requirements: Do our people have the right clearances?

    A BD professional at a firm like Deloitte, Booz Allen, or SAIC would
    recognize this as the essential "opportunity qualification" checklist.
    """

    # Read the input document
    with open(file_path, "r", encoding="utf-8") as f:
        solicitation_text = f.read()

    # Initialize the Anthropic client
    client = anthropic.Anthropic()

    # System prompt: Claude acts as an experienced govcon BD analyst
    # who knows what fields matter for capture decisions
    system_prompt = """You are a senior government contracts Business Development 
analyst with 15 years of experience in federal procurement. You specialize in 
rapidly qualifying contract opportunities from SAM.gov by extracting key data 
points that inform bid/no-bid decisions.

You understand FAR/DFARS procurement terminology, NAICS codes, contract types 
(FFP, T&M, CPFF, CPAF, IDIQ, BPA, etc.), set-aside categories (8(a), SDVOSB, 
HUBZone, WOSB, etc.), and evaluation methodologies (LPTA, best value, 
tradeoff analysis).

Guidelines:
- Extract only what is explicitly stated in the document
- If a field is not mentioned or cannot be determined, mark it as "Not specified"
- For NAICS codes, include the code number AND the description if provided
- For dates, use the format stated in the document
- For estimated value, note whether the figure is for the base period, 
  total contract value, or per-task-order ceiling
- Flag any ambiguities or unusual terms that a BD team should investigate
- Be precise — do not infer or assume information that is not in the text"""

    # User prompt: defines the exact JSON structure we want back
    # JSON output was chosen over markdown because structured data is more
    # useful downstream — it could feed into a CRM, database, or dashboard
    user_prompt = f"""Analyze the following federal contract opportunity text from 
SAM.gov and extract the key fields into a structured format.

Return your response as a JSON object with the following structure:

{{
    "opportunity_overview": {{
        "title": "Full opportunity title",
        "solicitation_number": "Solicitation or notice number",
        "notice_type": "e.g., Solicitation, Sources Sought, RFI, Pre-Solicitation, Award Notice",
        "agency": "Contracting agency name",
        "office": "Contracting office",
        "posted_date": "Date posted",
        "response_deadline": "Response due date and time, including timezone"
    }},
    "contract_details": {{
        "naics_code": "NAICS code and description",
        "psc_code": "Product/Service Code if available",
        "contract_type": "e.g., FFP, T&M, CPFF, IDIQ, BPA",
        "set_aside": "e.g., Total Small Business, 8(a), SDVOSB, Unrestricted",
        "estimated_value": "Dollar value and what it covers (base, total, ceiling)",
        "period_of_performance": "Base period and option years",
        "place_of_performance": "Location(s) where work will be performed"
    }},
    "requirements_summary": {{
        "scope_overview": "1 sentence summary of what the government is buying",
        "key_requirements": ["List of the most important technical/functional requirements"],
        "clearance_requirements": "Security clearance level required, if any",
        "key_personnel": "Any specific roles or qualifications required"
    }},
    "evaluation_criteria": {{
        "evaluation_method": "e.g., LPTA, Best Value Tradeoff, Lowest Price",
        "factors": ["List of evaluation factors in order of importance if stated"],
        "notes": "Any additional evaluation details"
    }},
    "competitive_intelligence": {{
        "incumbent": "Current contract holder if mentioned",
        "predecessor_contract": "Previous contract number if referenced",
        "estimated_competition": "Any indicators of competitive landscape",
        "notable_terms": "Unusual clauses, restrictions, or requirements worth flagging"
    }},
    "bd_action_items": {{
        "pursuit_recommendation": "Quick assessment: what type of firm is best positioned",
        "key_dates": ["List all critical dates and deadlines"],
        "risks_and_flags": ["Any red flags or concerns a BD team should investigate"]
    }},
    "proposal_outline": {{
        "sections": [
            {{
                "title": "Proposal section title (e.g., Technical Approach, Management Approach, Past Performance, Price/Cost)",
                "requirements_addressed": ["Key requirements from the solicitation that this section must address, by name or reference"],
                "guidance": "One sentence on what to emphasize in this section to score well"
            }}
        ],
        "win_themes": ["2-3 overarching themes or differentiators to thread throughout the entire proposal"]
    }}
}}

Important: Return ONLY the JSON object, no additional text before or after it.

---

SOLICITATION TEXT:

{solicitation_text}"""

    # Make the API call to Claude
    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        system=system_prompt,
        messages=[
            {"role": "user", "content": user_prompt}
        ]
    )

    return message.content[0].text


def main():
    """
    Entry point: accepts a file path as a command-line argument,
    runs the extractor, and outputs both formatted and raw JSON results.
    """

    if len(sys.argv) < 2:
        print("Usage: python extractor.py <path_to_solicitation.txt>")
        print("Example: python extractor.py \"Sample Contracts\\solicitation.txt\"")
        sys.exit(1)

    file_path = sys.argv[1]

    if not os.path.exists(file_path):
        print(f"Error: File not found: {file_path}")
        sys.exit(1)

    print(f"Reading solicitation: {file_path}")
    print("Extracting contract data...\n")

    raw_response = extract_contract_data(file_path)

    # Try to parse and pretty-print the JSON
    # If Claude returns valid JSON, we format it nicely
    # If not, we still show the raw response
    try:
        # Strip any markdown code fences Claude might add
        cleaned = raw_response.strip()
        if "```json" in cleaned:
            cleaned = cleaned.split("```json")[1]
        if "```" in cleaned:
            cleaned = cleaned.split("```")[0]
        cleaned = cleaned.strip()

        parsed = json.loads(cleaned)

        print("=" * 60)
        print("EXTRACTED CONTRACT DATA")
        print("=" * 60)

        # Print a human-readable summary of the most critical fields
        overview = parsed.get("opportunity_overview", {})
        details = parsed.get("contract_details", {})
        requirements = parsed.get("requirements_summary", {})
        eval_criteria = parsed.get("evaluation_criteria", {})
        intel = parsed.get("competitive_intelligence", {})
        actions = parsed.get("bd_action_items", {})
        outline = parsed.get("proposal_outline", {})

        print(f"\n  OPPORTUNITY: {overview.get('title', 'N/A')}")
        print(f"  SOLICITATION #: {overview.get('solicitation_number', 'N/A')}")
        print(f"  AGENCY: {overview.get('agency', 'N/A')}")
        print(f"  NOTICE TYPE: {overview.get('notice_type', 'N/A')}")
        print(f"  RESPONSE DUE: {overview.get('response_deadline', 'N/A')}")
        print(f"\n  NAICS: {details.get('naics_code', 'N/A')}")
        print(f"  CONTRACT TYPE: {details.get('contract_type', 'N/A')}")
        print(f"  SET-ASIDE: {details.get('set_aside', 'N/A')}")
        print(f"  EST. VALUE: {details.get('estimated_value', 'N/A')}")
        print(f"  PERIOD: {details.get('period_of_performance', 'N/A')}")
        print(f"  LOCATION: {details.get('place_of_performance', 'N/A')}")
        print(f"\n  SCOPE: {requirements.get('scope_overview', 'N/A')}")
        print(f"  CLEARANCE: {requirements.get('clearance_requirements', 'N/A')}")
        print(f"  EVAL METHOD: {eval_criteria.get('evaluation_method', 'N/A')}")
        print(f"  INCUMBENT: {intel.get('incumbent', 'N/A')}")

        # Print key requirements
        key_reqs = requirements.get("key_requirements", [])
        if key_reqs:
            print(f"\n  KEY REQUIREMENTS:")
            for i, req in enumerate(key_reqs, 1):
                print(f"    {i}. {req}")

        # Print evaluation factors
        factors = eval_criteria.get("factors", [])
        if factors:
            print(f"\n  EVALUATION FACTORS:")
            for i, factor in enumerate(factors, 1):
                print(f"    {i}. {factor}")

        # Print risks and flags
        risks = actions.get("risks_and_flags", [])
        if risks:
            print(f"\n  RISKS / FLAGS:")
            for i, risk in enumerate(risks, 1):
                print(f"    {i}. {risk}")

        # Print proposal outline
        win_themes = outline.get("win_themes", [])
        sections = outline.get("sections", [])
        if win_themes or sections:
            print(f"\n  PROPOSAL OUTLINE:")
            if win_themes:
                print(f"    Win Themes:")
                for theme in win_themes:
                    print(f"      - {theme}")
            for section in sections:
                print(f"\n    {section.get('title', 'Section')}:")
                reqs = section.get("requirements_addressed", [])
                if reqs:
                    print(f"      Addresses: {', '.join(reqs)}")
                guidance = section.get("guidance", "")
                if guidance:
                    print(f"      Note: {guidance}")

        print("\n" + "=" * 60)

        # Save the full JSON output
        output_path = file_path.replace(".txt", "_extracted.json")
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(parsed, f, indent=2)
        print(f"\nFull JSON saved to: {output_path}")

    except json.JSONDecodeError:
        # If JSON parsing fails, still save and display the raw response
        print("Note: Response was not valid JSON. Showing raw output.\n")
        print(raw_response)

        output_path = file_path.replace(".txt", "_extracted.txt")
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(raw_response)
        print(f"\nRaw output saved to: {output_path}")


if __name__ == "__main__":
    main()
