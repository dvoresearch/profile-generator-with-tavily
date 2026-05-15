"""
profile_generator.py  —  TAVILY VERSION
NUS Development Office – Prospect Profile Generator

How it works:
  1. Run 10 targeted Tavily web searches on the prospect
  2. Feed all search results into Claude as context
  3. Claude produces a structured JSON profile from the real web data
  4. If Tavily key is missing, falls back to Claude training knowledge
"""

import json
import os
import re
from typing import Optional

import anthropic
import requests


# ─────────────────────────────────────────────────────────────────────────────
# Tavily search
# ─────────────────────────────────────────────────────────────────────────────

def _tavily_search(query: str, api_key: str, max_results: int = 6) -> str:
    """Run one Tavily search and return formatted results as a string."""
    try:
        resp = requests.post(
            "https://api.tavily.com/search",
            json={
                "api_key":      api_key,
                "query":        query,
                "max_results":  max_results,
                "search_depth": "advanced",
                "include_answer": True,
            },
            timeout=15,
        )
        resp.raise_for_status()
        data = resp.json()

        lines = []
        if data.get("answer"):
            lines.append(f"ANSWER: {data['answer']}")
            lines.append("---")
        for r in data.get("results", []):
            lines.append(f"TITLE:   {r.get('title', '')}")
            lines.append(f"URL:     {r.get('url', '')}")
            lines.append(f"CONTENT: {r.get('content', '')[:800]}")
            lines.append("---")
        return "\n".join(lines) if lines else "No results."
    except Exception as e:
        return f"Search error: {e}"


def _gather_research(name: str, api_key: str, log) -> str:
    """Run 10 targeted searches and return all results as one string."""
    queries = [
        f"{name} biography profile",
        f"{name} current role chairman CEO director",
        f"{name} net worth wealth Forbes Bloomberg",
        f"{name} philanthropy donation foundation charity",
        f"{name} NUS National University Singapore",
        f"{name} education university degree",
        f"{name} awards honours recognition",
        f"{name} family spouse children",
        f"{name} lawsuit scandal controversy",
        f"{name} company organisation Singapore",
    ]

    all_results = []
    for i, q in enumerate(queries, 1):
        log(f"Search {i}/10: {q}")
        result = _tavily_search(q, api_key)
        all_results.append(f"=== SEARCH {i}: {q} ===\n{result}")

    return "\n\n".join(all_results)


# ─────────────────────────────────────────────────────────────────────────────
# System prompt
# ─────────────────────────────────────────────────────────────────────────────

SYSTEM_PROMPT = """You are a senior prospect research analyst for the National University of Singapore (NUS) Development Office.

Web search results for the prospect are provided in the conversation. Use them to build an accurate, well-sourced profile. Cite URLs from the search results in the sources field.

CRITICAL: Your ENTIRE response must be one valid JSON object. Start with { and end with }. No prose, no markdown, no explanation outside the JSON.

Detect if the prospect is an individual or a company and use the correct schema.

════════════════════════════════
CONTENT RULES
════════════════════════════════

BULLETS: Every array field contains complete sentence strings.
PRONOUNS: He / She for individuals. It / The [org name] for companies.

SINGLE-HOME RULE – each fact in exactly ONE field:
  Awards        → awards only
  Education     → education only
  Current roles → biography_current_positions only
  Family        → biography_family only

GIVING FORMAT (no exceptions):
  "In [Year], gave to [Organisation], [Amount]."
  Unknown year:   "In [year unknown], gave to [Organisation], [Amount]."
  Unknown amount: "In [Year], gave to [Organisation], amount unknown."
  Nothing found:  ["Not publicly available."]

DEMONSTRATED INTERESTS: hobbies, religion, personal/philanthropic interests only. Not professional.

ADVERSE NEWS: null if nothing found. If found, include as array; last item must be:
  "Note: This is a preliminary search only and may not reflect the complete list of adverse news."

CONNECTORS: only if BOTH are true:
  (1) Close direct personal relationship (co-directors, family, close friends, long-standing business partners)
  (2) Documented NUS connection (NUS Board of Trustees, NUS donor, active NUS alumnus)
  Exclude Prof Tan Eng Chye (NUS President – internal). Max 5. Empty array [] if none qualify.

GIFT IDEAS: always name a specific NUS faculty, school, or programme. Never generic.

ACCURACY: use "Not publicly available." only when the search results contain no information on that field.

════════════════════════════════
JSON SCHEMAS
════════════════════════════════

INDIVIDUAL:
{
  "type": "individual",
  "name": "Full Name",
  "gender": "male or female",
  "key_position": "Title, Organisation",
  "age": "XX (born YYYY) or Not publicly available.",
  "nationality": "Nationality",
  "net_worth": "USD X billion (as of YYYY) or Not publicly available.",
  "education": ["Degree, Institution (Year)"],
  "biography_intro": ["He/She is ...", "He/She founded ..."],
  "biography_current_positions": ["Position, Organisation"],
  "biography_past_positions": ["Position, Organisation (years)"],
  "biography_family": ["He/She is married to Name.", "He/She has X children."],
  "giving": ["In 2020, gave to Organisation, SGD X million."],
  "interests": ["He/She is known to be an avid golfer."],
  "awards": ["Award Name, Year."],
  "other_facts": ["He/She holds dual citizenship in X and Y."],
  "adverse_news": null,
  "connectors": [
    {
      "name_title": "Name, Title",
      "relationship_to_prospect": "Documented personal relationship",
      "nus_connection": "NUS Board of Trustees member since 20XX",
      "recommended_approach": "How to leverage this connection"
    }
  ],
  "gift_ideas": ["Named professorship at NUS Business School."],
  "sources": ["Source title – URL"]
}

COMPANY:
{
  "type": "company",
  "organisation_name": "Full Legal Name",
  "year_established": "YYYY or Not publicly available.",
  "country_of_registration": "Country",
  "annual_revenue": "USD X billion (as of YYYY) or Not publicly available.",
  "biography_intro": ["It is a ...", "It was founded by ..."],
  "biography_subsections": [
    {"label": "Key business lines:", "bullets": ["Line 1", "Line 2"]},
    {"label": "Key CSR programmes:", "bullets": ["Programme description"]}
  ],
  "giving": ["In 2021, gave to Organisation, SGD 2 million."],
  "interests": ["Not publicly available."],
  "awards": ["Award Name, Year."],
  "other_facts": ["It is listed on the SGX under ticker XXX."],
  "adverse_news": null,
  "connectors": [],
  "gift_ideas": ["Corporate-named scholarship at NUS School of Computing."],
  "sources": ["Source title – URL"]
}"""


# ─────────────────────────────────────────────────────────────────────────────
# Main entry point
# ─────────────────────────────────────────────────────────────────────────────

def research_prospect(
    prospect_name: str,
    client: anthropic.Anthropic,
    progress_callback=None,
    tavily_key: str = "",
) -> Optional[dict]:

    def log(msg: str):
        if progress_callback:
            progress_callback(msg)

    # Also check env/secrets as fallback if not passed in
    if not tavily_key:
        try:
            import streamlit as st
            tavily_key = st.secrets.get("TAVILY_API_KEY", "")
        except Exception:
            pass
    if not tavily_key:
        tavily_key = os.environ.get("TAVILY_API_KEY", "")

    if tavily_key:
        log("🌐 Tavily key found — running live web searches…")
        result = _research_with_tavily(prospect_name, client, tavily_key, log)
        if result:
            log("✅ Profile built from live web research.")
            return result
        log("Tavily research failed — falling back to training knowledge…")
    else:
        log("⚠️ No Tavily key — using Claude training knowledge only.")
        log("   Add TAVILY_API_KEY to Streamlit secrets for live web search.")

    # Fallback: training knowledge
    result = _research_knowledge_only(prospect_name, client, log)
    if result:
        log("✅ Profile built from training knowledge.")
    return result


# ─────────────────────────────────────────────────────────────────────────────
# Phase 1 – Tavily + Claude
# ─────────────────────────────────────────────────────────────────────────────

def _research_with_tavily(
    prospect_name: str,
    client,
    tavily_key: str,
    log,
) -> Optional[dict]:
    """Gather web search results then ask Claude to synthesise into JSON."""

    # Gather research
    research_text = _gather_research(prospect_name, tavily_key, log)
    log("Searches complete. Generating profile…")

    user_message = (
        f"Prospect: {prospect_name}\n\n"
        f"Web search results:\n\n"
        f"{research_text}\n\n"
        f"Using the search results above, output a complete JSON prospect profile. "
        f"Your entire response must be one JSON object starting with {{ and ending with }}."
    )

    for attempt in range(1, 4):
        if attempt > 1:
            log(f"Retry {attempt}/3…")
        try:
            raw = _call_claude(user_message, client)
            log(f"Response preview: {raw[:200].replace(chr(10), ' ')}")
            result = _parse_json(raw)
            if result:
                return result
            log("JSON parse failed — trying repair…")
            raw = _repair(raw, client)
            result = _parse_json(raw)
            if result:
                return result
        except Exception as e:
            log(f"Error on attempt {attempt}: {e}")

    return None


# ─────────────────────────────────────────────────────────────────────────────
# Phase 2 – Training knowledge only
# ─────────────────────────────────────────────────────────────────────────────

def _research_knowledge_only(
    prospect_name: str,
    client,
    log,
) -> Optional[dict]:
    """Use Claude training knowledge when Tavily is unavailable."""

    user_message = (
        f"Prospect: {prospect_name}\n\n"
        f"No web search results are available. Use your training knowledge to produce "
        f"a complete JSON profile. Write confidently about what you know. "
        f"Only use 'Not publicly available.' when you genuinely have no information. "
        f"Your entire response must be one JSON object starting with {{ and ending with }}."
    )

    for attempt in range(1, 4):
        if attempt > 1:
            log(f"Retry {attempt}/3…")
        try:
            raw = _call_claude(user_message, client)
            log(f"Response preview: {raw[:200].replace(chr(10), ' ')}")
            result = _parse_json(raw)
            if result:
                return result
            log("JSON parse failed — trying repair…")
            raw = _repair(raw, client)
            result = _parse_json(raw)
            if result:
                return result
        except Exception as e:
            log(f"Error on attempt {attempt}: {e}")

    return None


# ─────────────────────────────────────────────────────────────────────────────
# Claude API calls
# ─────────────────────────────────────────────────────────────────────────────

def _call_claude(user_message: str, client) -> str:
    response = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=8096,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_message}],
    )
    text = ""
    for block in response.content:
        if getattr(block, "type", None) == "text":
            text += getattr(block, "text", "")
    return text.strip()


def _repair(broken: str, client) -> str:
    response = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=8096,
        messages=[
            {
                "role": "user",
                "content": (
                    "The text below should be a JSON object but may be malformed. "
                    "Fix it and return ONLY valid JSON. No explanation. No markdown.\n\n"
                    f"{broken[:7000]}"
                ),
            }
        ],
    )
    text = ""
    for block in response.content:
        if getattr(block, "type", None) == "text":
            text += getattr(block, "text", "")
    return text.strip()


# ─────────────────────────────────────────────────────────────────────────────
# JSON parser
# ─────────────────────────────────────────────────────────────────────────────

def _parse_json(text: str) -> Optional[dict]:
    if not text:
        return None

    text = re.sub(r"```(?:json)?\s*", "", text)
    text = re.sub(r"```\s*$", "", text, flags=re.MULTILINE).strip()

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    start = text.find("{")
    if start == -1:
        return None

    depth = end = 0
    for i, ch in enumerate(text[start:], start):
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                end = i
                break
    if not end:
        return None

    json_str = text[start:end + 1]
    for candidate in (json_str, re.sub(r",\s*([}\]])", r"\1", json_str)):
        try:
            return json.loads(candidate)
        except json.JSONDecodeError:
            continue
    return None


# ─────────────────────────────────────────────────────────────────────────────
# File naming
# ─────────────────────────────────────────────────────────────────────────────

def get_filename(data: dict) -> str:
    if data.get("type") == "company":
        org  = data.get("organisation_name", "Organisation")
        safe = re.sub(r"[^\w\s-]", "", org).strip().replace(" ", "_")
        return f"{safe}_Prospect_Profile.docx"
    name  = data.get("name", "Prospect")
    parts = name.strip().split()
    if len(parts) >= 2:
        fname = re.sub(r"[^\w]", "", parts[0])
        lname = re.sub(r"[^\w]", "", parts[-1])
        return f"{fname}_{lname}_Prospect_Profile.docx"
    return f"{re.sub(r'[^\w]', '', name)}_Prospect_Profile.docx"
