from __future__ import annotations

import json
import textwrap
from typing import Any, Dict
from openai import OpenAI
from phd_reviewer.pipeline import call_json_model, summarize_document_for_prepass, normalize_ws


def run(
    client: OpenAI,
    model: str,
    payload: Dict[str, Any],
    document_profile: Dict[str, Any],
    retries: int,
) -> Dict[str, Any]:
    system_prompt = textwrap.dedent(
        """
        You are the OverallReviewAgent in a multi-agent paper-review system.

        Task:
        - Review the whole document at a high level.
        - Focus on structure, narrative flow, section ordering, redundancy, logical transitions,
          thesis clarity, and whether conclusions are aligned with the presented material.
        - Keep feedback specific and practical.
        - Use comments only.

        Return strict JSON only:
        {
          "document_summary": "short paragraph",
          "comments": [
            {"paragraph_id": "p0001", "anchor_text": "", "comment": "specific overall comment", "kind": "overall_review"}
          ]
        }
        """
    ).strip()
    summary_payload = {
        "document_profile": document_profile,
        "document_summary_input": summarize_document_for_prepass(payload, max_paragraphs=24),
        "ending_preview": [
            {
                "paragraph_id": p["paragraph_id"],
                "section": p.get("section", ""),
                "section_bucket": p.get("section_bucket", "other"),
                "text": p.get("text", "")[:900],
            }
            for p in payload.get("paragraphs", [])[-12:]
            if normalize_ws(p.get("text", ""))
        ],
    }
    return call_json_model(client, model, system_prompt, json.dumps(summary_payload, ensure_ascii=False, indent=2), retries=retries)
