from __future__ import annotations

import json
import textwrap
from typing import Any, Dict, List
from openai import OpenAI
from phd_reviewer.pipeline import call_json_model


def run(client: OpenAI, model: str, paragraph_chunk: List[Dict[str, Any]], figures: List[Dict[str, Any]], mode: str, aggressiveness: str, retries: int) -> Dict[str, Any]:
    aggressiveness_rule = {"conservative": "Keep edits minimal and prefer comments.", "balanced": "You may revise scientifically imprecise wording when the intended meaning is clear.", "substantive": "You may rewrite sentences for scientific precision when the intended meaning is clear, but never invent facts."}.get(aggressiveness, "You may revise scientifically imprecise wording when the intended meaning is clear.")
    edit_rule = "Comments only; do not propose edits." if mode == "comments_only" else f"Use comments for risky scientific issues; {aggressiveness_rule}"
    system_prompt = textwrap.dedent(
        f"""
        You are the ScienceAgent in a multi-agent paper-review system.

        Task:
        - Review scientific clarity, internal consistency, unsupported claims, and precise academic phrasing.
        - {edit_rule}
        - You may propose stronger wording changes than a proofreader when they improve scientific precision.
        - Never claim a result is wrong unless the issue is directly visible in the text or explicitly contradicted by provided evidence.
        - Never fabricate references, numbers, or methodology details.
        - Make obvious fixes directly as edits.
        - Reserve comments for issues that genuinely need human attention, such as unsupported claims, conceptual ambiguity, weak sourcing, or substantive scientific risk.

        Return strict JSON only with keys 'edits' and 'comments'.
        Each edit must be a short local replacement and each comment must be specific.
        """
    ).strip()
    user_payload = {"paragraphs": paragraph_chunk, "figures": figures[:4]}
    return call_json_model(client, model, system_prompt, json.dumps(user_payload, ensure_ascii=False, indent=2), retries=retries)
