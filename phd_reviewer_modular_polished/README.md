# Modular PhD Paper Reviewer

This refactor keeps the working review pipeline but moves each agent into its own file so you can modify prompts and behavior without editing the whole pipeline.

## Project layout

```text
phd_reviewer_modular/
  run_review.py
  ui_streamlit.py
  requirements.txt
  README.md
  phd_reviewer/
    pipeline.py
    agents/
      grammar_agent.py
      science_agent.py
      methods_agent.py
      results_agent.py
      citation_style_agent.py
      overall_review_agent.py
      literature_review_agent.py
      fact_check_agent.py
```

## Install

```bash
pip install -r requirements.txt
```

## CLI

```powershell
python -u .\run_review.py `
  --input .\draft_paper.docx `
  --output .\reviewed_paper_modular.docx `
  --model grok-4-1-fast-reasoning `
  --rounds 2 `
  --agents overall_review_agent,grammar_agent,science_agent,methods_agent,results_agent,citation_style_agent,literature_review_agent,fact_check_agent,reference_verification
```

## UI

```bash
streamlit run ui_streamlit.py
```

The UI gives you:
- a file uploader
- tick boxes for each agent
- a rounds selector
- mode and aggressiveness controls
- a reviewed DOCX download button

## Notes

- `reference_verification` is treated like a selectable pipeline stage even though it is deterministic rather than LLM-based.
- `literature_review_agent` is mainly used when intake classifies the document as a literature review.
- The source of truth for agent behavior is `phd_reviewer/agents/*.py`.


## UI customization

- Edit `ui_help_texts.json` to change field labels, hover help, and messages.
- Edit `ui_agent_catalog.json` to change agent display names, groups, defaults, descriptions, and speed labels.
