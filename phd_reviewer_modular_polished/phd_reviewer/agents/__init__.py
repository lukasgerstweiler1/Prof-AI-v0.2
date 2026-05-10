from . import grammar_agent, science_agent, methods_agent, results_agent, citation_style_agent, overall_review_agent, literature_review_agent, fact_check_agent

AGENT_REGISTRY = {
    "grammar_agent": grammar_agent.run,
    "science_agent": science_agent.run,
    "methods_agent": methods_agent.run,
    "results_agent": results_agent.run,
    "citation_style_agent": citation_style_agent.run,
    "overall_review_agent": overall_review_agent.run,
    "literature_review_agent": literature_review_agent.run,
    "fact_check_agent": fact_check_agent.run,
}
