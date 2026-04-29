"""
SAM.gov Contract Data Extractor — Multi-Document Version
=========================================================
An enhanced version of the extractor that processes an entire solicitation
package (multiple files in a folder) rather than a single document.

Real-world solicitations on SAM.gov often include multiple attachments:
- The SAM.gov synopsis/notice
- Performance Work Statement (PWS) or Statement of Work (SOW)
- Section L — Instructions to Offerors
- Section M — Evaluation Criteria
- Pricing template or CLIN structure
- Contract Data Requirements List (CDRLs)
- Attachments (org charts, security requirements, past performance forms)

This tool reads all .txt files in a specified folder, labels each one
by filename so Claude can distinguish between documents, and sends the
full package for extraction.

Built to demonstrate government-relevant AI use cases using the Anthropic API.

Author: Harry Cotton
"""

import anthropic
import sys
import os
import json
import glob


def read_all_documents(folder_path: str) -> str:
    """
    Reads all .txt files in a folder and combines them into a single
    labeled string. Each document is clearly marked with its filename
    so Claude can reference specific source documents in its extraction.

    This approach — concatenating documents with clear labels — is simpler
    than RAG for small-to-medium document packages (under ~50 pages total).
    For very large solicitation packages, a chunking/retrieval approach
    would be more appropriate.
    """

    # Find all .txt files in the folder, excluding previously generated output files
    txt_files = sorted(
        f for f in glob.glob(os.path.join(folder_path, "*.txt"))
        if not os.path.basename(f).endswith(("_extraction.txt", "_extracted.txt"))
    )

    if not txt_files:
        print(f"Error: No .txt files found in {folder_path}")
        sys.exit(1)

    print(f"Found {len(txt_files)} document(s):")
    for f in txt_files:
        print(f"  - {os.path.basename(f)}")
    print()

    # Combine all documents with clear separators and labels
    combined_text = ""
    for file_path in txt_files:
        filename = os.path.basename(file_path)
        with open(file_path, "r", encoding="utf-8") as f:
            content = f.read()
        combined_text += f"\n{'='*60}\n"
        combined_text += f"DOCUMENT: {filename}\n"
        combined_text += f"{'='*60}\n\n"
        combined_text += content
        combined_text += "\n\n"

    return combined_text


def extract_contract_data(folder_path: str) -> str:
    """
    Reads all documents in a solicitation package folder and extracts
    structured fields using Claude's API.

    The key difference from the single-document version: the prompt
    instructs Claude to synthesize information across multiple documents,
    noting which document each data point came from when possible. This
    mirrors how a BD analyst would work — cross-referencing the synopsis
    with the PWS, checking Section M against Section L, etc.
    """

    combined_text = read_all_documents(folder_path)

    client = anthropic.Anthropic()

    system_prompt = """You are a senior government contracts Business Development 
analyst with 15 years of experience in federal procurement. You specialize in 
rapidly qualifying contract opportunities from SAM.gov by extracting key data 
points that inform bid/no-bid decisions.

You understand FAR/DFARS procurement terminology, NAICS codes, contract types 
(FFP, T&M, CPFF, CPAF, IDIQ, BPA, etc.), set-aside categories (8(a), SDVOSB, 
HUBZone, WOSB, etc.), and evaluation methodologies (LPTA, best value, 
tradeoff analysis).

You are being given a COMPLETE SOLICITATION PACKAGE consisting of multiple 
documents. These may include the SAM.gov synopsis, a PWS/SOW, evaluation 
criteria, instructions to offerors, pricing templates, and other attachments.

Guidelines:
- Synthesize information across ALL documents in the package
- If the same field appears in multiple documents with different values, 
  flag the discrepancy in capital letters
- Note which source document each key data point comes from when it adds clarity
- Extract only what is explicitly stated — do not infer or assume
- If a field is not found in ANY of the provided documents, mark it as "Not specified"
- Flag any inconsistencies or contradictions between documents
- Be precise with FAR/DFARS terminology, dates, and dollar figures"""

    user_prompt = f"""Analyze the following federal contract opportunity package from 
SAM.gov. This package contains multiple documents from the same solicitation.
Extract the key fields into a structured format by synthesizing information 
across all provided documents.

Return your response as a JSON object with the following structure:

{{
    "opportunity_overview": {{
        "title": "Full opportunity title",
        "solicitation_number": "Solicitation or notice number",
        "notice_type": "e.g., Solicitation, Sources Sought, RFI, Pre-Solicitation, Award Notice",
        "agency": "Contracting agency name",
        "office": "Contracting office",
        "posted_date": "Date posted",
        "response_deadline": "Response due date and time, including timezone",
        "source_documents": "List of documents included in this package"
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
        "key_personnel": "Any specific roles or qualifications required",
        "deliverables": "Key deliverables or CDRLs if specified"
    }},
    "evaluation_criteria": {{
        "evaluation_method": "e.g., LPTA, Best Value Tradeoff, Lowest Price",
        "factors": ["List of evaluation factors in order of importance if stated"],
        "submission_requirements": "Page limits, format requirements, volumes required",
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
        "key_dates": ["List all critical dates and deadlines from across all documents"],
        "risks_and_flags": ["Any red flags, inconsistencies between documents, or concerns a BD team should investigate"],
        "cross_document_discrepancies": ["Any contradictions or inconsistencies found between the provided documents"]
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

SOLICITATION PACKAGE:

{combined_text}"""

    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=8192,
        system=system_prompt,
        messages=[
            {"role": "user", "content": user_prompt}
        ]
    )

    return message.content[0].text


def main():
    """
    Entry point: accepts a folder path containing solicitation documents,
    runs the extractor across all documents, and outputs results.

    Usage:
        python extractor_multi.py "Sample Contracts/CMS_Package"

    The folder should contain .txt files representing the various
    documents in a solicitation package.
    """

    if len(sys.argv) < 2:
        print("Usage: python extractor_multi.py <path_to_solicitation_folder>")
        print("Example: python extractor_multi.py \"Sample Contracts\\CMS_Package\"")
        print("\nThe folder should contain .txt files for each document")
        print("in the solicitation package (synopsis, PWS, Section L/M, etc.)")
        sys.exit(1)

    folder_path = sys.argv[1]

    if not os.path.isdir(folder_path):
        print(f"Error: Folder not found: {folder_path}")
        sys.exit(1)

    print(f"Processing solicitation package: {folder_path}")
    print("-" * 60)

    raw_response = extract_contract_data(folder_path)

    # Parse and display the JSON response
    try:
        cleaned = raw_response.strip()
        if "```json" in cleaned:
            cleaned = cleaned.split("```json")[1]
        if "```" in cleaned:
            cleaned = cleaned.split("```")[0]
        cleaned = cleaned.strip()

        parsed = json.loads(cleaned)

        print("=" * 60)
        print("EXTRACTED CONTRACT DATA (Multi-Document)")
        print("=" * 60)

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
        print(f"  SOURCE DOCS: {overview.get('source_documents', 'N/A')}")

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

        # Key requirements
        key_reqs = requirements.get("key_requirements", [])
        if key_reqs:
            print(f"\n  KEY REQUIREMENTS:")
            for i, req in enumerate(key_reqs, 1):
                print(f"    {i}. {req}")

        # Evaluation factors
        factors = eval_criteria.get("factors", [])
        if factors:
            print(f"\n  EVALUATION FACTORS:")
            for i, factor in enumerate(factors, 1):
                print(f"    {i}. {factor}")

        # Risks and flags
        risks = actions.get("risks_and_flags", [])
        if risks:
            print(f"\n  RISKS / FLAGS:")
            for i, risk in enumerate(risks, 1):
                print(f"    {i}. {risk}")

        # Cross-document discrepancies — unique to multi-doc version
        discrepancies = actions.get("cross_document_discrepancies", [])
        if discrepancies:
            print(f"\n  CROSS-DOCUMENT DISCREPANCIES:")
            for i, d in enumerate(discrepancies, 1):
                print(f"    {i}. {d}")

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
        folder_name = os.path.basename(os.path.normpath(folder_path))
        output_path = os.path.join(folder_path, f"{folder_name}_full_extraction.json")
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(parsed, f, indent=2)
        print(f"\nFull JSON saved to: {output_path}")

    except json.JSONDecodeError:
        print("Note: Response was not valid JSON. Showing raw output.\n")
        print(raw_response)

        folder_name = os.path.basename(os.path.normpath(folder_path))
        output_path = os.path.join(folder_path, f"{folder_name}_raw_extraction.txt")
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(raw_response)
        print(f"\nRaw output saved to: {output_path}")


if __name__ == "__main__":
    main()
