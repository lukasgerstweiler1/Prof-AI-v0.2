from __future__ import annotations

import json
import tempfile
from pathlib import Path

import streamlit as st

from phd_reviewer import ReviewConfig, default_enabled_agents, run_review

HELP_FILE = Path(__file__).with_name("ui_help_texts.json")
CATALOG_FILE = Path(__file__).with_name("ui_agent_catalog.json")

DEFAULT_HELP = {
    "page": {
        "title": "PhD Paper Reviewer",
        "caption": "Choose agents, set review rounds, and run the reviewer from your browser.",
    },
    "fields": {},
    "sections": {},
    "messages": {},
}

DEFAULT_CATALOG = {
    "groups": {},
    "agents": {},
}


def load_json_config(path: Path, default: dict) -> dict:
    if not path.exists():
        return default
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        merged = default.copy()
        for key, value in data.items():
            if isinstance(value, dict) and isinstance(merged.get(key), dict):
                merged[key] = {**merged[key], **value}
            else:
                merged[key] = value
        return merged
    except Exception:
        return default


HELP = load_json_config(HELP_FILE, DEFAULT_HELP)
CATALOG = load_json_config(CATALOG_FILE, DEFAULT_CATALOG)
FIELDS = HELP.get("fields", {})
SECTIONS = HELP.get("sections", {})
MESSAGES = HELP.get("messages", {})
PAGE = HELP.get("page", {})
GROUPS = CATALOG.get("groups", {})
AGENTS = CATALOG.get("agents", {})


ALL_AGENT_NAMES = list(AGENTS.keys()) or default_enabled_agents()
DEFAULTS = {
    name
    for name, meta in AGENTS.items()
    if meta.get("default", name in default_enabled_agents())
}
if not DEFAULTS:
    DEFAULTS = set(default_enabled_agents())


SPEED_ORDER = {"fast": 0, "medium": 1, "slow": 2}
SPEED_EMOJI = {"fast": "⚡", "medium": "⏱️", "slow": "🐢"}


st.set_page_config(page_title=PAGE.get("title", "PhD Paper Reviewer"), layout="wide")


def field_label(key: str, fallback: str) -> str:
    return FIELDS.get(key, {}).get("label", fallback)


def field_help(key: str) -> str | None:
    return FIELDS.get(key, {}).get("help")


def agent_meta(agent_name: str) -> dict:
    return AGENTS.get(agent_name, {})


def agent_label(agent_name: str) -> str:
    meta = agent_meta(agent_name)
    base = meta.get("label", agent_name)
    speed = str(meta.get("speed", "medium")).lower().strip()
    speed_label = speed.capitalize() if speed else "Medium"
    return f"{base} ({speed_label})"


def agent_help(agent_name: str) -> str:
    meta = agent_meta(agent_name)
    speed = str(meta.get("speed", "medium")).lower().strip() or "medium"
    emoji = SPEED_EMOJI.get(speed, "⏱️")
    desc = meta.get("help", "")
    return f"Approximate speed: {emoji} {speed}.\n\n{desc}".strip()


def grouped_agent_names() -> dict[str, list[str]]:
    grouped: dict[str, list[str]] = {}
    for name in ALL_AGENT_NAMES:
        group = agent_meta(name).get("group", "other")
        grouped.setdefault(group, []).append(name)
    for names in grouped.values():
        names.sort(key=lambda n: (SPEED_ORDER.get(str(agent_meta(n).get("speed", "medium")).lower(), 1), agent_meta(n).get("label", n)))
    return grouped


st.title(PAGE.get("title", "PhD Paper Reviewer"))
st.caption(PAGE.get("caption", ""))

with st.sidebar:
    st.header(SECTIONS.get("settings_header", "Review settings"))
    st.subheader(SECTIONS.get("basic_settings_header", "Basic settings"))
    rounds = st.number_input(
        field_label("rounds", "Review rounds"),
        min_value=1,
        max_value=5,
        value=2,
        step=1,
        help=field_help("rounds"),
    )
    output_name = st.text_input(
        field_label("output_name", "Output filename"),
        value="reviewed_paper_modular_polished.docx",
        help=field_help("output_name"),
    )

    with st.expander(SECTIONS.get("advanced_settings_header", "Advanced settings"), expanded=False):
        model = st.text_input(field_label("model", "Model"), value="grok-4-1-fast-reasoning", help=field_help("model"))
        mode = st.selectbox(field_label("mode", "Mode"), ["edit_and_comment", "comments_only"], index=0, help=field_help("mode"))
        vision = st.checkbox(field_label("vision", "Enable figure vision"), value=False, help=field_help("vision"))
        st.caption(MESSAGES.get("vision_note", ""))
        editing_aggressiveness = st.selectbox(
            field_label("editing_aggressiveness", "Editing aggressiveness"),
            ["conservative", "balanced", "substantive"],
            index=2,
            help=field_help("editing_aggressiveness"),
        )
        max_workers = st.slider(
            field_label("max_workers", "Parallel workers"),
            min_value=1,
            max_value=8,
            value=4,
            step=1,
            help=field_help("max_workers"),
        )
        crossref_mailto = st.text_input(field_label("crossref_mailto", "Crossref polite email"), value="", help=field_help("crossref_mailto"))
        unpaywall_email = st.text_input(field_label("unpaywall_email", "Unpaywall email"), value="", help=field_help("unpaywall_email"))
        author = st.text_input(field_label("author", "Tracked changes author"), value="AI Reviewer", help=field_help("author"))

st.subheader(SECTIONS.get("document_header", "Document"))
uploaded = st.file_uploader(field_label("upload", "Upload a .docx file"), type=["docx"], help=field_help("upload"))

st.subheader(SECTIONS.get("agents_header", "Choose agents"))
st.caption(MESSAGES.get("speed_legend", "Approximate speed: fast, medium, slow."))

selected_agents: list[str] = []
groups = grouped_agent_names()
for group_key, names in groups.items():
    group_meta = GROUPS.get(group_key, {})
    with st.expander(group_meta.get("label", group_key.replace("_", " ").title()), expanded=True):
        if group_meta.get("help"):
            st.caption(group_meta["help"])
        cols = st.columns(2)
        for idx, agent_name in enumerate(names):
            with cols[idx % 2]:
                checked = st.checkbox(
                    agent_label(agent_name),
                    value=agent_name in DEFAULTS,
                    help=agent_help(agent_name),
                    key=f"agent_{agent_name}",
                )
                if checked:
                    selected_agents.append(agent_name)

if st.button(field_label("run_review", "Run review"), type="primary", disabled=uploaded is None, help=field_help("run_review")):
    if uploaded is None:
        st.error(MESSAGES.get("upload_first", "Please upload a .docx file first."))
    elif not selected_agents:
        st.error(MESSAGES.get("select_agent", "Select at least one agent."))
    else:
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            input_path = tmp / uploaded.name
            input_path.write_bytes(uploaded.read())
            output_path = tmp / output_name

            cfg = ReviewConfig(
                input=str(input_path),
                output=str(output_path),
                payload_json=str(tmp / "review_payload.json"),
                decisions_json=str(tmp / "review_decisions.json"),
                apply_report_json=str(tmp / "review_apply_report.json"),
                crossref_report_json=str(tmp / "review_crossref_report.json"),
                work_dir=str(tmp / "review_workdir"),
                model=model,
                rounds=int(rounds),
                mode=mode,
                vision=vision,
                editing_aggressiveness=editing_aggressiveness,
                crossref_mailto=crossref_mailto,
                unpaywall_email=unpaywall_email,
                author=author,
                enabled_agents=selected_agents,
                log_file=str(tmp / "review.log"),
                max_workers=int(max_workers),
            )

            with st.spinner(MESSAGES.get("running", "Running review... this can take several minutes for longer papers.")):
                result = run_review(cfg)

            st.success(MESSAGES.get("completed", "Review completed."))
            st.subheader(SECTIONS.get("results_header", "Review results"))
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Document type", result["document_profile"].get("document_type", "unknown"))
            col2.metric("Proposed edits", len(result["decisions"].get("edits", [])))
            col3.metric("Proposed comments", len(result["decisions"].get("comments", [])))
            col4.metric("Applied comments", result["apply_report"].get("applied_comments", 0))
            st.write("Applied tracked edits:", result["apply_report"].get("applied_edits", 0))

            st.download_button(
                field_label("download_docx", "Download reviewed DOCX"),
                data=output_path.read_bytes(),
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                help=field_help("download_docx"),
            )
            st.download_button(
                field_label("download_decisions", "Download decisions JSON"),
                data=Path(cfg.decisions_json).read_bytes(),
                file_name="review_decisions.json",
                mime="application/json",
                help=field_help("download_decisions"),
            )
            st.download_button(
                field_label("download_log", "Download log file"),
                data=Path(cfg.log_file).read_bytes(),
                file_name="review.log",
                mime="text/plain",
                help=field_help("download_log"),
            )
