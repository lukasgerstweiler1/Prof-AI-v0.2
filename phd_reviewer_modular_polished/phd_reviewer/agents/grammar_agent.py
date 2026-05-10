from __future__ import annotations

import json
import textwrap
from typing import Any, Dict, List
from openai import OpenAI
from phd_reviewer.pipeline import call_json_model


def run(client: OpenAI, model: str, paragraph_chunk: List[Dict[str, Any]], mode: str, aggressiveness: str, retries: int) -> Dict[str, Any]:
    aggressiveness_rule = {"conservative": "Prefer smaller edits unless the wording is clearly poor.", "balanced": "Use moderate edits when they clearly improve readability.", "substantive": "You may substantially rewrite a sentence or a whole paragraph locally when it clearly improves clarity, logic, and academic tone."}.get(aggressiveness, "Use moderate edits when they clearly improve readability.")
    edit_rule = "Do not propose edits; use comments only." if mode == "comments_only" else f"Make tracked-change edits. {aggressiveness_rule}"
    system_prompt = textwrap.dedent(
        f"""
        You are the GrammarAgent in a multi-agent paper-review system.

        Task:
        - Improve grammar, spelling, punctuation, concision, and academic style.
        - {edit_rule}
        - You may make sentence-level or paragraph-local rewrites when they clearly improve clarity and flow.
        - Do not change scientific meaning.
        - Do not invent citations.
        - Prefer tracked changes over comments.
        - Do not add comments for minor punctuation, capitalization, spacing, or routine grammar fixes.
        - Only add a comment when the problem likely needs human judgment, such as ambiguity, structural weakness, or unclear logic.
        - Prefer stronger edits when the wording is awkward, repetitive, or non-native, but keep each edit anchored to the source paragraph.

        Return strict JSON only:
        {{
          "edits": [
            {{
              "paragraph_id": "p0001",
              "find": "original snippet",
              "replace": "replacement snippet",
              "context_before": "up to 30 chars before",
              "context_after": "up to 30 chars after",
              "comment": "optional short rationale",
              "kind": "grammar",
              "confidence": "high"
            }}
          ],
          "comments": [
            {{
              "paragraph_id": "p0001",
              "anchor_text": "optional snippet",
              "comment": "short comment for the author",
              "kind": "grammar"
            }}
          ]
        }}

        Rules for edits:
        - Keep edits exact and applyable.
        - Only propose edits when the exact 'find' text exists in the paragraph, unless you are intentionally replacing the whole paragraph.
        - It is acceptable to rewrite a full paragraph when the paragraph is clearly weak, redundant, or illogical.
        """
    ).strip()
    return call_json_model(client, model, system_prompt, json.dumps({"paragraphs": paragraph_chunk}, ensure_ascii=False, indent=2), retries=retries)
