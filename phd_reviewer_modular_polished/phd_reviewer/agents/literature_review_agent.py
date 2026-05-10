from __future__ import annotations

import json
import textwrap
from typing import Any, Dict, List
from openai import OpenAI
from phd_reviewer.pipeline import call_json_model


def run(client: OpenAI, model: str, paragraph_chunk: List[Dict[str, Any]], retries: int) -> Dict[str, Any]:
    system_prompt = textwrap.dedent(
        """
        You are the LiteratureReviewAgent in a multi-agent paper-review system.

        Task:
        - Review a literature review or background-heavy section.
        - Focus on synthesis vs. summary, comparison of sources, thematic organization,
          chronology problems, missing critical evaluation, and unsupported transitions.
        - Use comments only.
        - Do not demand methods/results sections if this appears to be a literature review.

        Return strict JSON only:
        {"edits": [], "comments": [{"paragraph_id":"p0001","anchor_text":"","comment":"specific literature-review note","kind":"literature_review"}]}
        """
    ).strip()
    return call_json_model(client, model, system_prompt, json.dumps({"paragraphs": paragraph_chunk}, ensure_ascii=False, indent=2), retries=retries)
