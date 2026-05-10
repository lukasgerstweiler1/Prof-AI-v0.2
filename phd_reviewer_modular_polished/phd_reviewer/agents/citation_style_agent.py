from __future__ import annotations

import json
import textwrap
from typing import Any, Dict, List
from openai import OpenAI
from phd_reviewer.pipeline import call_json_model


def run(client: OpenAI, model: str, paragraph_chunk: List[Dict[str, Any]], references_text: str, mode: str, retries: int) -> Dict[str, Any]:
    edit_rule = "Comments only; do not propose edits." if mode == "comments_only" else "Edits are allowed only for safe local citation/reference formatting fixes."
    system_prompt = textwrap.dedent(
        f"""
        You are the CitationStyleAgent in a multi-agent paper-review system.

        Task:
        - Review in-text citations and reference list formatting for consistency.
        - {edit_rule}
        - Never fabricate bibliographic details.
        - Do not claim that an in-text citation is missing from the reference list. Reference-presence checking is handled deterministically elsewhere.
        - Use edits for punctuation, spacing, capitalization, duplicated DOI prefixes, obvious OCR cleanup, superscript artifacts, and very safe style consistency fixes.
        - Reserve comments for issues that likely need human attention, such as suspicious future years, truncated/corrupted entries, incomplete bibliographic fields, retractions/corrections, or unresolved ambiguity.
        - Do not add comments for minor punctuation, capitalization, spacing, or routine formatting fixes if they can be corrected directly.

        Return strict JSON only with keys 'edits' and 'comments'.
        """
    ).strip()
    user_payload = {"paragraphs": paragraph_chunk, "references_text": references_text[:12000]}
    return call_json_model(client, model, system_prompt, json.dumps(user_payload, ensure_ascii=False, indent=2), retries=retries)
