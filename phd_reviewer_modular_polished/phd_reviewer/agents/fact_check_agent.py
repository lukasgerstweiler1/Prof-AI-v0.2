from __future__ import annotations

import json
import textwrap
from typing import Any, Dict, List
from openai import OpenAI
from phd_reviewer.pipeline import call_json_model


def run(
    client: OpenAI,
    model: str,
    paragraph: Dict[str, Any],
    cited_evidence: List[Dict[str, Any]],
    document_profile: Dict[str, Any],
    mode: str,
    retries: int,
) -> Dict[str, Any]:
    if not cited_evidence:
        return {"edits": [], "comments": []}

    edit_rule = (
        "Use comments only."
        if mode == "comments_only"
        else "Edits are allowed only for very safe claim-softening changes such as changing absolute language to cautious language."
    )
    system_prompt = textwrap.dedent(
        f"""
        You are the FactCheckAgent in a multi-agent paper-review system.

        Task:
        - Compare the paragraph's explicit scientific claims against the cited-source evidence provided.
        - Use only the supplied internet lookup evidence derived from reference metadata, abstracts, open-access full text, and post-publication status signals when available.
        - Identify overclaiming, citation mismatch, citation-status risk, or claims not clearly supported by the cited evidence.
        - {edit_rule}
        - Do not invent facts.
        - If the evidence is too thin, say so in a comment rather than asserting the claim is false.
        - If the required correction is obvious and low-risk, propose a direct tracked edit and optionally add a brief comment only if human attention is still needed.

        Return strict JSON only:
        {{
          "edits": [
            {{
              "paragraph_id": "p0001",
              "find": "exact risky wording",
              "replace": "more cautious wording",
              "comment": "why softened",
              "kind": "fact_check",
              "confidence": "medium"
            }}
          ],
          "comments": [
            {{
              "paragraph_id": "p0001",
              "anchor_text": "",
              "comment": "specific fact-check note linked to the cited evidence",
              "kind": "fact_check"
            }}
          ]
        }}
        """
    ).strip()
    user_payload = {
        "document_type": document_profile.get("document_type", ""),
        "paragraph": paragraph,
        "cited_evidence": cited_evidence[:8],
    }
    return call_json_model(client, model, system_prompt, json.dumps(user_payload, ensure_ascii=False, indent=2), retries=retries)

# -----------------------------------------------------------------------------
# Deterministic reference verification (Crossref)
# -----------------------------------------------------------------------------
