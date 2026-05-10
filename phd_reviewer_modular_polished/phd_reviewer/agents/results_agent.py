from __future__ import annotations

import json
import textwrap
from typing import Any, Dict, List
from openai import OpenAI
from phd_reviewer.pipeline import call_json_model


def run(client: OpenAI, model: str, paragraph_chunk: List[Dict[str, Any]], retries: int) -> Dict[str, Any]:
    system_prompt = textwrap.dedent(
        """
        You are the ResultsAgent in a multi-agent paper-review system.

        Task:
        - Review results / discussion / conclusion paragraphs for overclaiming, weak evidence language,
          causal overreach, unsupported generalization, and unclear comparison statements.
        - Use comments only.

        Return strict JSON only:
        {"edits": [], "comments": [{"paragraph_id":"p0001","anchor_text":"","comment":"specific results/discussion review note","kind":"results"}]}
        """
    ).strip()
    return call_json_model(client, model, system_prompt, json.dumps({"paragraphs": paragraph_chunk}, ensure_ascii=False, indent=2), retries=retries)
