"""
Clause Comparison & LLM Prompt Generation Engine.
Computes unified diffs between specification releases and formats multi-tier
hierarchical context prompts optimized for LLM standards analysis.
"""

import difflib
from typing import Any, Dict, List, Optional, Tuple


FOCUS_PROMPTS = {
    "standards": """### Analysis Directive for LLM
Act as a 3GPP Senior Standards and Telecommunications Architect. Review the diff above in light of the provided procedure context and analyze:
1. **Normative & Behavioral Impact:** What functional behaviors, state transitions, or mandatory requirements change for the UE, (R)AN, and Core Network NFs?
2. **Signalling & Protocol Consequences:** What message call flows, Information Elements (IEs), or parameters are introduced, modified, or deprecated?
3. **Backward Compatibility & Interoperability:** How does a legacy entity behave when communicating with a peer implementing these changes? What are the potential failure modes?
4. **Summary of Consequences:** Provide a concise executive bullet-point summary of the core architectural impact.""",

    "patent": """### Analysis Directive for LLM
Act as a 3GPP Telecommunications Patent Specialist and Standard-Essential Patent (SEP) Analyst. Review the diff above in light of the provided procedure context and analyze:
1. **Novelty & Timing of Introduction:** Identify the exact technical features, call flow steps, or conditional triggers present in the Target Version that were entirely absent in the Base Version.
2. **Claim Limitation Mapping:** Identify specific procedural limitations (e.g., parameter names, conditional checks, timer triggers) that distinguish this release from prior art.
3. **Essentiality Assessment:** Assess whether an implementation of the Target Version strictly requires this novel procedural step, or if alternative standard paths exist.
4. **Prior Art Delta Summary:** Summarize what distinguishes the Target Version from the Base Version regarding patent novelty.""",

    "signalling": """### Analysis Directive for LLM
Act as a 3GPP Protocol Engineering & Signalling Expert. Review the diff above in light of the provided procedure context and analyze:
1. **Message Structure & Encoding:** What exact Information Elements (IEs), bit-fields, or ASN.1 parameters were added or altered?
2. **Call Flow Timing & Ordering:** Are any request/response message sequences, timers, or acknowledgments modified?
3. **Error Handling & Cause Codes:** What error scenarios, reject causes, or fallback procedures are introduced if the new functionality fails?
4. **Implementation Checklist:** Provide a step-by-step checklist of protocol parser changes required for this delta."""
}


def generate_unified_diff(
    base_text: str,
    target_text: str,
    base_label: str = "Base Version",
    target_label: str = "Target Version",
) -> str:
    """Generates a clean unified diff between two text versions."""
    base_lines = base_text.splitlines(keepends=True)
    target_lines = target_text.splitlines(keepends=True)

    diff_lines = list(
        difflib.unified_diff(
            base_lines,
            target_lines,
            fromfile=base_label,
            tofile=target_label,
            lineterm="",
        )
    )

    if not diff_lines:
        return "No text differences detected between these versions."

    return "\n".join(diff_lines)


def build_llm_clause_prompt(
    spec_number: str,
    clause_number: str,
    clause_title: str,
    base_version: str,
    base_date: Optional[str],
    target_version: str,
    target_date: Optional[str],
    diff_text: str,
    tier: int = 2,
    hierarchy: Optional[List[Dict[str, Any]]] = None,
    branch_clauses: Optional[List[Dict[str, Any]]] = None,
    focus_mode: str = "standards",
) -> str:
    """
    Assembles a structured Markdown prompt containing hierarchical context,
    the computed diff, and specialized analytical instructions for an LLM.
    """
    base_date_str = f" ({base_date})" if base_date else ""
    target_date_str = f" ({target_date})" if target_date else ""

    sections = [
        f"# 3GPP Specification Change Analysis: TS {spec_number}",
        f"- **Target Clause:** Clause {clause_number} ({clause_title})",
        f"- **Base Release:** v{base_version}{base_date_str}",
        f"- **Target Release:** v{target_version}{target_date_str}",
        f"- **Context Scope:** Tier {tier} ({'Exact Clause' if tier == 1 else ('Parent Procedure Scope' if tier == 2 else 'Comprehensive Procedure Branch')})",
        "",
        "---",
    ]

    # Add Hierarchical Context for Tier 2 and Tier 3
    if tier >= 2 and hierarchy:
        breadcrumb_parts = [f"Clause {h['clause_number']} ({h['clause_title']})" for h in hierarchy]
        breadcrumb_parts.append(f"**Clause {clause_number} ({clause_title})**")
        breadcrumb_str = " ➔ ".join(breadcrumb_parts)

        sections.append("## 📍 Hierarchical Context & Preconditions (Target Release)")
        sections.append(f"> **Document Structure:** {breadcrumb_str}")
        sections.append("")

        for parent in hierarchy:
            p_num = parent.get("clause_number", "")
            p_title = parent.get("clause_title", "")
            p_content = parent.get("content", "").strip()

            # Truncate overly long parent chapters to keep prompt focused
            if len(p_content) > 1200:
                p_content = p_content[:1200] + "\n... [Remaining parent text omitted for brevity] ..."

            sections.append(f"### Parent Clause {p_num}: {p_title}")
            sections.append(f"```text\n{p_content}\n```")
            sections.append("")

    # Add Sibling Procedure Branch Context for Tier 3
    if tier >= 3 and branch_clauses:
        sections.append("## 🌳 Full Procedure Branch Context (Target Release)")
        for sibling in branch_clauses:
            s_num = sibling.get("clause_number", "")
            s_title = sibling.get("clause_title", "")
            s_content = sibling.get("content", "").strip()

            if s_num != clause_number:
                if len(s_content) > 1000:
                    s_content = s_content[:1000] + "\n... [Truncated for brevity] ..."
                sections.append(f"#### Clause {s_num}: {s_title}")
                sections.append(f"```text\n{s_content}\n```")
                sections.append("")

    # Add the Clause Diff Block
    sections.append("## 📝 Clause Text Diff")
    sections.append(f"```diff\n{diff_text}\n```")
    sections.append("")
    sections.append("---")
    sections.append("")

    # Add Directive
    directive = FOCUS_PROMPTS.get(focus_mode, FOCUS_PROMPTS["standards"])
    sections.append(directive)

    return "\n".join(sections)