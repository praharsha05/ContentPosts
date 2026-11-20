"""Daily Instagram content bot for @analyst_shubhi - single-file Streamlit app (Gemini)."""
import io, json, re, zipfile, datetime as dt
from zoneinfo import ZoneInfo

import streamlit as st

PIP_NAMES = {"docx": "python-docx", "google": "google-genai", "PIL": "Pillow"}
try:
    from docx import Document
    from docx.shared import Pt
    from google import genai
    from google.genai import types
    from PIL import Image, ImageDraw, ImageFont
except ModuleNotFoundError as e:  # Streamlit Cloud hides the real error, so show it here
    missing = (e.name or "").split(".")[0]
    st.error(f"Missing Python package: **{e.name}**  (pip name: `{PIP_NAMES.get(missing, missing)}`)")
    st.markdown(
        "Fix: make sure **`requirements.txt`** is in your GitHub repo, in the same folder as `app.py`, "
        "and contains:\n```\ngoogle-genai>=1.0.0\npython-docx>=1.1.0\nstreamlit>=1.50.0\nPillow>=10.1.0\n```\n"
        "Then open **Manage app -> Reboot app**."
    )
    st.stop()

st.set_page_config(page_title="Daily Content Bot", page_icon="📊", layout="centered")

IST = ZoneInfo("Asia/Kolkata")


def now_ist() -> dt.datetime:
    return dt.datetime.now(IST)


# ======================= EDIT THIS: your creator profile =======================
DEFAULT_PROFILE = """
Name: Shubhanshi | Handle: @analyst_shubhi | ~70k followers | Data Analyst based in Hyderabad
Mission: help people bridge the gap between learning and getting hired in data.
Audience: aspiring and junior data analysts / data engineers, career switchers, freshers, interview-preppers.
AUDIENCE LEVEL: BEGINNER. Most followers are students, freshers and career switchers in their first months of
learning. Topics must be basic, practical and everyday. Never advanced theory.
Topics she covers: SQL, Power BI/DAX, Excel, Python (pip/uv), statistics, data engineering (Spark, Airflow,
Kafka, dbt, Delta Lake, lakehouse), GenAI / LLMs / AI agents / Claude workflows for analysts, interview prep, roadmaps.

VOICE (learned from her real captions):
- Clean, simple English. Confident, encouraging, practical. No Hinglish, no slang.
- Line 1 is a hook: a bold claim, a contrast ("pip is the past. uv is the future."), a challenge
  ("This query runs with no error... but gives NO output!") or a "Stop doing X" statement.
- Then 1-3 short lines saying exactly what the swipe contains ("Swipe for 7 stages with what to study,
  the tools, a mini project for each stage...").
- Emojis as signposts: 📌 save, 📊 analytics, 🔥 energy, 👇 comments, ✅ lists, 💻 SQL. Not overloaded.
- Honest about AI: "always double-check the results".
- Philosophy: "Don't memorize, understand why and when"; "Build concepts, not just projects."

CAPTION PATTERNS SHE USES:
1. Hook line + what's inside + "📌 Save this and share it with a friend who is starting in data." + a
   question: "Which stage are you on right now? Tell me below 👇".
2. Lead magnet: 'Comment "PDF" / "GENAI" to get the guide/resource.'
3. Challenge post: broken SQL query -> "Can you find the bug?" -> first correct answer gets a pin + shoutout.
4. Tool/cheat-sheet carousels: "Swipe for the cheat sheet 👉".
5. Short hashtag set (3-6) like #SQL #LearnSQL #dataengineering, or none.

HER REAL CAPTIONS (study the structure, density and tone):
---
Stop copy-pasting data into chat. 🚀 Connect Claude to the tools you already use: Excel, Chrome, Drive,
BigQuery, Snowflake, Databricks, Amplitude, Slack. Which one are you connecting first? 👇
Save 🔖 and follow @analyst_shubhi for more data + AI tips.
---
This query runs with no error... but gives NO output! 😳 Can you find the bug? 🔍 Comment your answer below 👇
First correct answer gets a pin and a shoutout in my next video! 🔥 Follow for daily SQL tips for beginners 💻
#SQL #LearnSQL #SQLForBeginners #Freshers #CodingChallenge
---
The Data Engineering Stack: 5 layers, 19 tools, explained simply! Swipe to see what each tool does, a beginner
stack you can start with, and 8 terms to say in an interview. 📌 Save this and share it with a friend learning
data engineering. Which layer do you want me to explain deeper? Tell me below 👇
---
Want to become a Data Engineer? Here is your roadmap! Swipe for 7 stages with what to study, the tools, a mini
project for each stage, and how long it takes (about 10 to 13 months at 12 hours a week).
📌 Save this and share it with a friend who is starting in data. Which stage are you on right now? Tell me below 👇
---
Working with Delta Lake in production? These 10 real-world scenarios are a must-know! 🔥 From duplicate records,
slow MERGE operations, schema evolution, and concurrent writes to CDC, Time Travel, OPTIMIZE vs VACUUM, and more.
This PDF covers practical troubleshooting scenarios that can help you strengthen your Data Engineering skills
and prepare for technical interviews. 📌 Comment "PDF" to get the resource.
---
Top 10 DAX functions behind every Power BI report 📊 Swipe for simple examples you can copy. Save this for your
next project 📌 Which DAX function confused you the most? Tell me below 👇
---
How does AI like Claude actually work? 🤖 It's not magic! 🧠 LLMs predict the next word, use tools, and follow
loops to complete tasks. Brain + Tools + Loop = AI Agent ✨ As a data analyst, I use AI for data cleaning, SQL,
and insights - but always double-check the results! 💡
---
Knowing how to use GenAI isn't the same as understanding how it works! 🚀 I've put together a 15-part GenAI
Interview & Revision Guide covering LLMs, RAG, embeddings, chunking, token optimization, and GenAI architecture.
📌 Build concepts, not just projects. Prepare smarter for your next GenAI interview! Comment "GENAI" to get the guide. 🔥
---
pip is the past. uv is the future. ⚡ pip: waiting. ⏳ uv: done. ⚡ One tool replaces pip, venv, pyenv and pipx.
Swipe for the cheat sheet 👉
---
Statistics is one of the most important foundations for anyone learning Data Analytics, Data Science or Machine
Learning. 📈 I'm sharing these notes to help you build your statistics fundamentals step by step - from descriptive
statistics, mean, median, variance and standard deviation to probability, distributions, confidence intervals,
hypothesis testing, ANOVA, correlation and regression. Don't just memorize formulas. Focus on understanding why
and when to use each concept.
---
AI Agent interviews aren't about memorizing definitions - they're about understanding how the system works. 🤖
From fundamentals -> tools -> multi-agent systems -> memory -> reliability -> production.
---
SQL interviews aren't about memorizing syntax. They're about practicing real problems. 💻
Master these: Joins, GROUP BY, Subqueries, CASE WHEN, Window Functions. Practice more. Fear SQL less. 🔥
---
From Python & SQL to Spark, Airflow, Kafka, Cloud, dbt, Lakehouse, System Design & AI pipelines - this roadmap
covers the complete Data Engineering journey. Perfect for beginners, career switchers, and professionals
preparing for Data Engineering roles. 📌 Save this roadmap & check the link in bio 🖇️

GAP TO FIX (important): many of her posts are broad overviews (lists of tools, roadmaps, "top 10 X"). The
new content must go DEEPER: one concrete problem, worked through on sample data, and what it changes in
an analyst's actual working day.

AVOID: invented statistics, salary numbers, "X% of interviews" claims, tool bashing, get-rich-quick tone.
"""

# One theme per weekday so the feed stays varied. Monday = 0.
WEEKDAY_THEMES = {
    "Mon - SQL deep-dive": "One everyday SQL task solved step by step on a small table (SELECT, WHERE, GROUP BY, JOIN, COUNT, CASE WHEN, "
                           "NULLs, removing duplicates).",
    "Tue - BI & reporting": "Excel / Power BI / Tableau basics: build something stakeholders actually use "
                            "(pivot table, VLOOKUP/XLOOKUP, a KPI card, a simple chart, basic DAX like SUM and CALCULATE).",
    "Wed - A day at work": "A realistic everyday request from a manager ('send me last week's sales by city') and how an analyst "
                           "handles it step by step: clarify, get data, check it, share it.",
    "Thu - Data engineering & Python": "Beginner-friendly data engineering and Python: what a pipeline is, pandas basics, pip vs uv, Git basics, "
                                       "one simple problem and its fix.",
    "Fri - AI for analysts": "Simple ways to use AI (Claude / ChatGPT) for SQL, Excel formulas and cleaning data, with a check step and an "
                             "honest limitation.",
    "Sat - Interview prep": "A common fresher interview question with a simple strong answer and the usual mistakes; "
                            "or a beginner roadmap / cheat sheet with a lead-magnet keyword.",
    "Sun - Foundations & career": "Basic concepts in plain words (mean vs median, what is a KPI, what does a data analyst do), resume, "
                              "portfolio project ideas, learning plans.",
}
# ==============================================================================


# ---------------------------- Greeting banner ----------------------------
def greeting_banner():
    """Always-scrolling greeting strip that changes with the time of day (IST)."""
    hour = now_ist().hour
    if 5 <= hour < 12:
        msg, icons = "Good Morning, Sunshine ", "☀️ 💖 🌞 💕 🧡 💗"
        bg, fg = "linear-gradient(90deg,#fde68a,#fbcfe8,#fde68a)", "#7c2d12"
    elif 12 <= hour < 17:
        msg, icons = "Good Afternoon Indu ", "🌸 🌼 🌻 🌷 🌺 🦋"
        bg, fg = "linear-gradient(90deg,#bbf7d0,#fef9c3,#fbcfe8)", "#14532d"
    elif 17 <= hour < 20:
        msg, icons = "Good Evening Indu ", "🌅 🧡 🍃 ✨ ☕ 🌇"
        bg, fg = "linear-gradient(90deg,#fdba74,#f9a8d4,#c4b5fd)", "#4a044e"
    else:
        msg, icons = "Good Night , Sweet Dreams!", "🌙 ⭐ 🧸 🕯️ ☕ 🛏️ 😴"
        bg, fg = "linear-gradient(90deg,#1e1b4b,#312e81,#1e1b4b)", "#e0e7ff"

    item = f"{icons} &nbsp; <b>{msg}</b> &nbsp; {icons} &nbsp;&nbsp;&nbsp;&nbsp;"
    track = item * 3
    # keep the HTML flush-left, otherwise Markdown treats indented lines as a code block
    html = f"""<style>
.greet-bar {{ overflow:hidden; white-space:nowrap; border-radius:14px; padding:10px 0;
  margin-bottom:12px; background:{bg}; color:{fg}; font-size:1.1rem; }}
.greet-track {{ display:flex; width:max-content; animation:greet-scroll 28s linear infinite; }}
.greet-track span {{ white-space:pre; }}
.greet-bar:hover .greet-track {{ animation-play-state:paused; }}
@keyframes greet-scroll {{ from {{ transform:translateX(0); }} to {{ transform:translateX(-50%); }} }}
</style>
<div class="greet-bar"><div class="greet-track"><span>{track}</span><span>{track}</span></div></div>"""
    st.markdown(html, unsafe_allow_html=True)


# ---------------------------- LLM helpers ----------------------------
def model_chain():
    """Models come only from Streamlit Secrets. Order = priority (first is tried first).
    Accepts a comma-separated string or a TOML list, under GEMINI_MODELS (or GEMINI_MODEL)."""
    try:
        raw = st.secrets.get("GEMINI_MODELS") or st.secrets.get("GEMINI_MODEL") or ""
    except Exception:  # no secrets file at all
        raw = ""
    if isinstance(raw, str):
        raw = raw.split(",")
    chain = []
    for m in raw:
        m = str(m).strip()
        if m and m not in chain:
            chain.append(m)
    if not chain:
        raise RuntimeError(
            'No Gemini model configured. Add this to Secrets: GEMINI_MODELS = "gemini-2.5-flash"'
        )
    return chain


def _retryable(e: Exception) -> bool:
    t = str(e)
    return any(k in t for k in ("429", "RESOURCE_EXHAUSTED", "404", "NOT_FOUND", "503", "UNAVAILABLE",
                                "500 ", "INTERNAL"))


def _is_timeout(e: Exception) -> bool:
    t = (type(e).__name__ + " " + str(e)).lower()
    return any(k in t for k in ("timeout", "timed out", "deadline_exceeded", "504"))


def llm_json(system: str, user: str):
    try:  # 100 s limit per request so the page can never spin forever
        client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"], http_options=types.HttpOptions(timeout=100_000))
    except (AttributeError, TypeError):
        client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
    cfg = types.GenerateContentConfig(
        system_instruction=system, response_mime_type="application/json",
        temperature=0.8, max_output_tokens=24000,  # headroom: thinking tokens count toward this limit
    )
    last = None
    for model in model_chain():
        for attempt in range(2):  # one retry on malformed JSON, same model
            try:
                resp = client.models.generate_content(model=model, contents=user, config=cfg)
            except Exception as e:  # quota / unknown model / overloaded -> try the next model
                last = e
                if _is_timeout(e):  # slow response: stop here instead of waiting through every model
                    raise RuntimeError("TIMEOUT: Gemini took too long to answer.")
                if _retryable(e):
                    break
                raise
            text = re.sub(r"^```(?:json)?|```$", "", (resp.text or "").strip(), flags=re.M).strip()
            try:
                st.session_state["model_used"] = model
                return json.loads(text)
            except json.JSONDecodeError as e:
                last = e
        # fall through to the next model
    raise RuntimeError(f"All Gemini models failed. Last error: {last}")


def pick_topic(profile, used_topics, theme_name, theme_desc, level="Beginner", build_only=False):
    system = ("You are a senior content strategist for a data-analytics Instagram creator. You design posts that "
              "teach ONE concrete thing in depth. Return ONLY valid JSON.")
    user = """
CREATOR PROFILE:
__PROFILE__

Today: __TODAY__.
Today's content theme: __THEME__ -> __THEME_DESC__
Topics already used (do NOT repeat or closely resemble): __USED__

Pick ONE topic for today's post that fits today's theme.

DIFFICULTY: __LEVEL__
- Beginner = absolute beginners and freshers in their first months. The topic must be something a fresher meets in
  the first 3 months of learning, e.g. SELECT/WHERE/GROUP BY, INNER vs LEFT JOIN, COUNT vs COUNT(DISTINCT), handling
  NULLs, removing duplicates, VLOOKUP vs XLOOKUP, pivot tables, picking the right chart, average vs median, what a KPI
  is, cleaning a messy Excel sheet, basic DAX (SUM, CALCULATE), a fresher resume, a first portfolio project,
  common fresher interview questions.
- Beginner+ = a bit further: CASE WHEN, subqueries, simple window functions (ROW_NUMBER, RANK), pandas groupby, DAX
  basics, simple dashboards.
- Intermediate = CTEs, window functions in depth, DAX filter context, query optimisation basics.
- NEVER choose advanced or niche theory for Beginner or Beginner+: no Simpson's paradox, no skewed-distribution
  or standard-deviation outlier debates, no p-values or A/B-test statistics, no query internals, no architecture or
  system design, no machine learning maths.
- Use everyday, relatable situations (food delivery orders, student marks, shop sales, cricket scores, salaries in a
  small company), not finance or research jargon.
- ONE idea per post.

DEPTH RULES (very important):
- The topic must be a SPECIFIC problem an analyst/engineer hits at work, phrased like a real task, e.g.
  "Find customers who bought in January but not February" or "Why your DAX total doesn't match the table",
  NOT a broad overview like "SQL patterns" or "Top 10 tools".
- It must be teachable in depth in one post: a worked example on small sample data, plus what changes in the
  analyst's daily work once they know it.
- Prefer formats that allow depth: Carousel (default), or Reel for demos/bug-hunts.

JSON schema:
{"topic": str, "pillar": str, "format": "Reel"|"Carousel"|"Single image",
 "work_scenario": str,   // one sentence: the stakeholder request / work situation that triggers this
 "angle": str, "why_today": str, "target_pain_point": str}
"""
    user = (user.replace("__PROFILE__", profile).replace("__TODAY__", f"{now_ist():%A, %d %B %Y}")
            .replace("__LEVEL__", level).replace("__THEME__", theme_name).replace("__THEME_DESC__", theme_desc)
            .replace("__USED__", json.dumps(used_topics[-40:])))
    if build_only:
        return system, user
    return llm_json(system, user)


def generate_content(profile, topic, level="Beginner", build_only=False):
    system = ("You write detailed, accurate Instagram educational content for a working data analyst, matching her voice "
              "from the sample captions. You are precise: all code must run as written on the sample data you provide. "
              "You never invent statistics. Return ONLY valid JSON.")
    user = """
CREATOR PROFILE:
__PROFILE__

TOPIC PLAN:
__TOPIC__

Create a DETAILED content pack. This must NOT be a generic overview. Teach this one problem in depth.

LANGUAGE AND LEVEL: __LEVEL__ audience.
- Write for a complete beginner: short, simple words and short sentences. Explain every technical term in plain
  words the first time it appears (e.g. "NULL (a missing value)").
- Use relatable sample data (food orders, student marks, shop sales). Keep tables to 3-5 rows and 3 columns.
- Do not mention advanced concepts. If something is advanced, leave it out instead of explaining it.

REQUIREMENTS
1. Work scenario: start from the real work situation (who asks what, why it matters, what data is involved).
2. Sample data: give a tiny table (3-6 rows) as plain text so the example is concrete, and the expected output
   that the solution produces on exactly that data. For non-code topics (career/AI), use a concrete worked
   example instead (a real prompt + output, a mock resume line, etc.).
3. Solution: step by step, each step says WHAT it does and WHY. State the SQL dialect when it matters.
4. Impact on daily work: be specific about the task, e.g. "the weekly retention report no longer needs a manual
   Excel VLOOKUP merge" or "you catch double-counted orders before the CFO does". Say which step disappears,
   which risk is avoided, who benefits. Generic claims like "saves time" or "boosts productivity" are NOT allowed.
5. Gotchas: 3 real mistakes/edge cases (NULLs, ties, duplicates, wrong join type, timezones, ...) with how to avoid them.
6. Interview angle: how this appears as an interview question and the follow-up to expect.
7. NO invented statistics, percentages of interviews, salaries or company claims. Numbers are allowed only if they
   come from the sample data shown, or are clearly labelled "for example".
8. Voice: match her captions. Hook first, "Swipe for ..." line, 📌 save/share line, ends with a question
   ("Which ... ? Tell me below 👇"). Add a lead magnet keyword only when natural (guide, PDF, cheat sheet).

JSON schema:
{
 "hooks": [3 alternative first lines],
 "work_scenario": str,                 // 3-4 sentences, concrete
 "sample_data": str,                   // plain-text table or worked input, "" if not applicable
 "expected_output": str,               // what the solution returns on that data, "" if not applicable
 "before_after": {"before": str, "after": str, "impact": str},   // daily-work impact, each 1-3 specific sentences
 "gotchas": [3 strings, each: mistake -> how to avoid],
 "interview_angle": str,
 "reel_script": [{"time": "0-3s", "visual": str, "voiceover": str}],   // full spoken lines, 30-45s total; [] if not a Reel
 "carousel_slides": [{"slide": int,
     "layout": "cover"|"points"|"code"|"compare",
     "kicker": str,        // 2-4 word label above the title, e.g. "THE SITUATION", "STEP 1"
     "title": str,         // max 8 words, punchy
     "subtitle": str,      // cover only: the stakeholder ask in one line
     "points": [str],      // 3-4 items, each "Bold lead (2-4 words): one crisp explanation (max 14 words)"
     "code": str,          // code layout only: the code or data table, real line breaks
     "notes": [str],       // code layout only: 2-3 annotations, each max 9 words, one line: "KEYWORD: what it does / why"
     "before": str, "after": str,   // compare layout only, each max 28 words, specific to a work task
     "callout": str}],     // the one-line key takeaway, max 12 words
 "code_or_example": str,               // the full final solution, runnable
 "post_title": str,                    // catchy Instagram post title, max 9 words (see TITLE RULES)
 "alt_titles": [2 alternative titles],
 "caption": str,                       // CRISP, in her voice (see CAPTION RULES), hashtags NOT included
 "cta": str,
 "lead_magnet_keyword": str,           // e.g. "SQL" or "" if none
 "hashtags": [4-6 strings starting with #],
 "story_ideas": [3 short ideas],
 "best_time_to_post": str,
 "cover_text": str                     // 4-6 words
}

TITLE RULES
- Max 9 words. Names the exact task or mistake and promises a clear benefit or a curiosity gap.
- Simple words a beginner understands. No jargon, no ALL CAPS, no fake claims, no numbers you cannot back up.
- Good: "Why Your NOT IN Query Returns Nothing", "VLOOKUP vs XLOOKUP: Which One to Use", "Stop Using Average. Use Median."

CAPTION RULES (crisp, max 70 words, hashtags excluded)
- Line 1: the hook, max 12 words, with one emoji.
- Then 2-3 very short lines saying exactly what the swipe/reel gives ("Swipe for: ...").
- Then one line "📌 Save this and share it with a friend who is learning data." (vary the wording).
- Last line: one question to the audience ending with "Tell me below 👇" (or a comment keyword if there is a lead magnet).
- Short lines (max 60 characters each), blank line between blocks, 3-5 emojis total. No paragraphs.

CAROUSEL RULES
- The slides must carry the SAME specifics as the explanation you give elsewhere. Nothing important may live only
  in the caption: the slide is the crisp version of the explanation.
- CRISP: no paragraphs. Points are "Bold lead: short explanation". No emojis. Every slide has a callout.
- Code and data go in "code"; every code slide has 2-3 "notes" (max 9 words each) saying what a part does and why.
- If format is Carousel, use 9-10 slides in this order:
 1 layout cover: kicker = pillar, title = the work problem as a hook, subtitle = the stakeholder ask in one line,
   points = 3 things the viewer will get, each max 8 words ("Spot the trap: why NULL empties NOT IN"), callout = what to swipe for.
 2 layout points, kicker THE SITUATION: who asks what, what data is involved, where it gets tricky (3 points).
 3 layout code, kicker SAMPLE DATA: the table; notes explain the columns and the tricky rows.
 4 layout points, kicker COMMON MISTAKE: the wrong approach and exactly why it fails (3 points).
 5 layout code, kicker STEP 1: first part of the solution with notes.
 6 layout code, kicker STEP 2: second part of the solution with notes (one clause per line, max ~34 chars per line).
 7 layout code, kicker RESULT: the output on the sample data, notes explain how to read it.
 8 layout compare, kicker IMPACT: before vs after for a specific daily task; callout = the concrete impact.
 9 layout points, kicker GOTCHAS: 3 mistakes, each "Mistake: how to avoid it".
 10 layout points, kicker INTERVIEW ANGLE: how it is asked, the follow-up, 2-3 recap points; callout = the CTA question.
- If the format is not Carousel, carousel_slides is []. Fill reel_script only if format is Reel; otherwise [].
"""
    user = user.replace("__PROFILE__", profile).replace("__TOPIC__", json.dumps(topic, indent=2)).replace("__LEVEL__", level)
    if build_only:
        return system, user
    return llm_json(system, user)


def generate_pack(profile, used_topics, theme_name, theme_desc, level):
    """ONE Gemini request that returns both the topic and the full content pack."""
    s1, u1 = pick_topic(profile, used_topics, theme_name, theme_desc, level, build_only=True)
    s2, u2 = generate_content(profile, "the topic you chose in PART 1", level, build_only=True)
    system = ("You are a senior content strategist and writer for a data-analytics Instagram creator. "
              "Return ONLY valid JSON.")
    user = ("Do TWO PARTS in a single response.\n\n=== PART 1: PICK THE TOPIC ===\n" + u1 +
            "\n\n=== PART 2: WRITE THE CONTENT PACK for the topic you chose in PART 1 ===\n" + u2 +
            "\n\n=== OUTPUT FORMAT (overrides the individual output instructions above) ===\n"
            'Return ONE JSON object and nothing else: {"topic": <the PART 1 JSON object>, '
            '"content": <the PART 2 JSON object>}')
    data = llm_json(system, user)
    if not isinstance(data, dict) or "topic" not in data or "content" not in data:
        raise ValueError("Gemini response was missing 'topic' or 'content'. Click Try again.")
    return data["topic"], data["content"]


def normalize_slide(sl, i=1):
    """Accept both the new schema and the old (type/body) one."""
    sl = dict(sl or {})
    layout = sl.get("layout")
    if not layout:
        layout = "code" if sl.get("type") == "code" else ("cover" if i == 1 else "points")
    if layout == "cards":
        layout = "points"
    points = [str(p) for p in (sl.get("points") or []) if str(p).strip()]
    code = sl.get("code") or (sl.get("body") if layout == "code" else "") or ""
    if not points and layout in ("points", "cover") and sl.get("body"):
        points = [str(sl["body"])]
    return dict(layout=layout, kicker=sl.get("kicker", ""), title=str(sl.get("title", "")),
                subtitle=sl.get("subtitle") or (sl.get("body", "") if layout == "cover" and sl.get("points") else ""),
                points=points, code=str(code), notes=[str(n) for n in (sl.get("notes") or [])],
                before=sl.get("before", ""), after=sl.get("after", ""), callout=sl.get("callout", ""),
                slide=sl.get("slide", i))


# ---------------------------- Word doc ----------------------------
def build_docx(topic, c) -> bytes:
    d = Document()
    d.styles["Normal"].font.size = Pt(11)
    d.add_heading(f"Content Plan - {now_ist():%d %b %Y}", 0)

    d.add_heading("Post title", 1)
    d.add_paragraph(c.get("post_title") or topic.get("topic", "")).runs[0].bold = True
    d.add_heading("Topic", 1)
    d.add_paragraph(topic.get("topic", "")).runs[0].bold = True
    for k in ("pillar", "format", "angle", "why_today", "target_pain_point"):
        p = d.add_paragraph()
        p.add_run(k.replace("_", " ").title() + ": ").bold = True
        p.add_run(str(topic.get(k, "")))

    d.add_heading("Work scenario", 1)
    d.add_paragraph(c.get("work_scenario") or topic.get("work_scenario", ""))

    d.add_heading("Hook options", 1)
    for h in c.get("hooks", []):
        d.add_paragraph(str(h), style="List Number")

    if c.get("sample_data") or c.get("expected_output"):
        d.add_heading("Sample data & expected output", 1)
        for label, key in (("Sample data", "sample_data"), ("Expected output", "expected_output")):
            if c.get(key):
                d.add_paragraph().add_run(label).bold = True
                for r in d.add_paragraph(c[key]).runs:
                    r.font.name = "Consolas"

    ba = c.get("before_after") or {}
    if ba:
        d.add_heading("Impact on daily work", 1)
        for label, key in (("Before", "before"), ("After", "after"), ("Impact", "impact")):
            p = d.add_paragraph(); p.add_run(label + ": ").bold = True; p.add_run(str(ba.get(key, "")))

    if c.get("reel_script"):
        d.add_heading("Reel script", 1)
        t = d.add_table(rows=1, cols=3)
        t.style = "Light Grid Accent 1"
        for i, h in enumerate(("Time", "Visual", "Voiceover")):
            t.rows[0].cells[i].text = h
        for r in c["reel_script"]:
            cells = t.add_row().cells
            cells[0].text = str(r.get("time", ""))
            cells[1].text = str(r.get("visual", ""))
            cells[2].text = str(r.get("voiceover", ""))
    if c.get("carousel_slides"):
        d.add_heading("Carousel slides", 1)
        for i, raw in enumerate(c["carousel_slides"], 1):
            sl = normalize_slide(raw, i)
            head = f"Slide {i}: {sl['title']}" + (f"  [{sl['kicker']}]" if sl["kicker"] else "")
            d.add_paragraph().add_run(head).bold = True
            if sl["subtitle"]:
                d.add_paragraph(sl["subtitle"])
            for p_ in sl["points"]:
                d.add_paragraph(p_, style="List Bullet")
            if sl["code"]:
                for r in d.add_paragraph(sl["code"]).runs:
                    r.font.name = "Consolas"
            for n_ in sl["notes"]:
                d.add_paragraph(n_, style="List Bullet")
            if sl["before"] or sl["after"]:
                d.add_paragraph(f"Before: {sl['before']}")
                d.add_paragraph(f"After: {sl['after']}")
            if sl["callout"]:
                p_ = d.add_paragraph(); p_.add_run("Takeaway: ").bold = True; p_.add_run(sl["callout"])
    if c.get("code_or_example"):
        d.add_heading("Code / worked example", 1)
        for r in d.add_paragraph(c["code_or_example"]).runs:
            r.font.name = "Consolas"

    d.add_heading("Caption", 1)
    d.add_paragraph(c.get("caption", ""))
    p = d.add_paragraph(); p.add_run("CTA: ").bold = True; p.add_run(str(c.get("cta", "")))
    d.add_paragraph(" ".join(c.get("hashtags", [])))
    if c.get("lead_magnet_keyword"):
        p = d.add_paragraph(); p.add_run("Lead magnet keyword: ").bold = True; p.add_run(c["lead_magnet_keyword"])
    if c.get("gotchas"):
        d.add_heading("Gotchas", 1)
        for g in c["gotchas"]:
            d.add_paragraph(str(g), style="List Bullet")
    if c.get("interview_angle"):
        d.add_heading("Interview angle", 1)
        d.add_paragraph(c["interview_angle"])
    d.add_heading("Verify before posting", 1)
    for chk in ("Run the code on the sample data and confirm the output matches.",
                "Remove any number you cannot source.",
                "Check the SQL dialect (PostgreSQL / MySQL / SQL Server) matches what you say."):
        d.add_paragraph(chk, style="List Bullet")
    d.add_heading("Story ideas", 1)
    for s in c.get("story_ideas", []):
        d.add_paragraph(str(s), style="List Bullet")
    p = d.add_paragraph(); p.add_run("Cover text: ").bold = True; p.add_run(str(c.get("cover_text", "")))
    p = d.add_paragraph(); p.add_run("Best time to post: ").bold = True; p.add_run(str(c.get("best_time_to_post", "")))

    buf = io.BytesIO()
    d.save(buf)
    return buf.getvalue()


# ---------------------------- PNG rendering ----------------------------
W, H, M = 1080, 1350, 90
HANDLE = "@analyst_shubhi"
THEMES = {
    "Midnight": dict(bg=((15, 23, 42), (30, 41, 59)), text=(241, 245, 249), muted=(148, 163, 184),
                     accent=(56, 189, 248), code_bg=(2, 6, 23), code_text=(226, 232, 240),
                     kw=(244, 114, 182), st=(134, 239, 172), num=(251, 191, 36), cm=(100, 116, 139)),
    "Clean Light": dict(bg=((250, 250, 247), (240, 240, 235)), text=(17, 24, 39), muted=(107, 114, 128),
                        accent=(79, 70, 229), code_bg=(17, 24, 39), code_text=(229, 231, 235),
                        kw=(167, 139, 250), st=(110, 231, 183), num=(252, 211, 77), cm=(107, 114, 128)),
    "Sunset": dict(bg=((124, 58, 237), (236, 72, 153)), text=(255, 255, 255), muted=(237, 233, 254),
                   accent=(253, 224, 71), code_bg=(49, 18, 77), code_text=(245, 243, 255),
                   kw=(249, 168, 212), st=(167, 243, 208), num=(253, 224, 71), cm=(196, 181, 253)),
    "Forest": dict(bg=((6, 78, 59), (4, 47, 46)), text=(236, 253, 245), muted=(167, 243, 208),
                   accent=(110, 231, 183), code_bg=(2, 24, 22), code_text=(209, 250, 229),
                   kw=(253, 164, 175), st=(125, 211, 252), num=(253, 224, 71), cm=(94, 150, 130)),
    "Ocean": dict(bg=((3, 105, 161), (12, 74, 110)), text=(240, 249, 255), muted=(186, 230, 253),
                  accent=(251, 191, 36), code_bg=(8, 47, 73), code_text=(224, 242, 254),
                  kw=(253, 164, 175), st=(167, 243, 208), num=(253, 224, 71), cm=(125, 160, 190)),
    "Charcoal": dict(bg=((24, 24, 27), (39, 39, 42)), text=(250, 250, 250), muted=(161, 161, 170),
                     accent=(250, 204, 21), code_bg=(9, 9, 11), code_text=(228, 228, 231),
                     kw=(251, 146, 60), st=(74, 222, 128), num=(96, 165, 250), cm=(113, 113, 122)),
    "Neon Cyber": dict(bg=((9, 9, 30), (46, 16, 101)), text=(245, 243, 255), muted=(167, 139, 250),
                       accent=(0, 245, 212), code_bg=(3, 3, 20), code_text=(224, 231, 255),
                       kw=(255, 61, 190), st=(0, 245, 212), num=(254, 240, 138), cm=(110, 110, 180)),
    "Coral Peach": dict(bg=((251, 113, 133), (251, 146, 60)), text=(255, 255, 255), muted=(255, 241, 242),
                        accent=(254, 249, 195), code_bg=(88, 28, 35), code_text=(255, 241, 242),
                        kw=(253, 224, 71), st=(167, 243, 208), num=(253, 186, 116), cm=(253, 164, 175)),
    "Cream Paper": dict(bg=((254, 252, 232), (254, 243, 199)), text=(41, 37, 36), muted=(120, 113, 108),
                        accent=(217, 119, 6), code_bg=(41, 37, 36), code_text=(245, 245, 244),
                        kw=(251, 146, 60), st=(134, 239, 172), num=(253, 224, 71), cm=(140, 133, 128)),
    "Rose Blush": dict(bg=((255, 241, 242), (254, 226, 226)), text=(76, 5, 25), muted=(159, 110, 125),
                       accent=(225, 29, 72), code_bg=(76, 5, 25), code_text=(255, 228, 230),
                       kw=(253, 164, 175), st=(167, 243, 208), num=(253, 224, 71), cm=(190, 140, 155)),
    "Lavender": dict(bg=((237, 233, 254), (221, 214, 254)), text=(46, 16, 101), muted=(109, 92, 160),
                     accent=(124, 58, 237), code_bg=(46, 16, 101), code_text=(245, 243, 255),
                     kw=(249, 168, 212), st=(167, 243, 208), num=(253, 224, 71), cm=(167, 139, 250)),
    "Mono B&W": dict(bg=((255, 255, 255), (245, 245, 245)), text=(0, 0, 0), muted=(115, 115, 115),
                     accent=(0, 0, 0), code_bg=(0, 0, 0), code_text=(255, 255, 255),
                     kw=(250, 204, 21), st=(134, 239, 172), num=(147, 197, 253), cm=(115, 115, 115)),
}
SQL_KW = set("""SELECT FROM WHERE GROUP BY ORDER HAVING JOIN LEFT RIGHT INNER OUTER FULL ON AS WITH OVER PARTITION
AND OR NOT IN IS NULL LIMIT DISTINCT CASE WHEN THEN ELSE END UNION ALL INSERT INTO UPDATE DELETE CREATE TABLE
ASC DESC BETWEEN LIKE EXISTS COUNT SUM AVG MIN MAX ROW_NUMBER RANK DENSE_RANK LAG LEAD DEF IMPORT RETURN FOR
IF ELIF PRINT LAMBDA""".split())
FONT_DIRS = ["/usr/share/fonts/truetype/dejavu/", "/usr/share/fonts/dejavu/", ""]


def font(kind: str, size: int):
    names = {"bold": "DejaVuSans-Bold.ttf", "reg": "DejaVuSans.ttf", "mono": "DejaVuSansMono.ttf",
             "monob": "DejaVuSansMono-Bold.ttf"}
    for d in FONT_DIRS:
        try:
            return ImageFont.truetype(d + names[kind], size)
        except OSError:
            continue
    return ImageFont.load_default(size)  # Pillow >= 10.1 scalable fallback


def clean(text: str) -> str:
    """DejaVu has no emoji glyphs, so convert/remove them (keeps the check/cross marks DejaVu does have)."""
    text = str(text).replace("\u27a1\ufe0f", "\u2192").replace("\u27a1", "\u2192")
    text = re.sub(r"(\d)\ufe0f?\u20e3", r"\1.", text)
    text = re.sub(r"[\U0001F000-\U0001FFFF\u2600-\u2712\u2719-\u27BF\u2B00-\u2BFF\uFE0F\u200d]", "", text)
    return re.sub(r"[ \t]+", " ", text).strip()


def gradient(top, bottom):
    img = Image.new("RGB", (W, H), top)
    px = ImageDraw.Draw(img)
    for y in range(H):
        t = y / (H - 1)
        px.line([(0, y), (W, y)], fill=tuple(int(top[i] + (bottom[i] - top[i]) * t) for i in range(3)))
    return img


def wrap(draw, text, fnt, max_w):
    lines = []
    for pi, para in enumerate(text.split("\n")):
        if pi and para.strip():
            lines.append("")  # paragraph gap
        words, cur = para.split(), ""
        for w in words:
            trial = (cur + " " + w).strip()
            if draw.textlength(trial, font=fnt) <= max_w:
                cur = trial
            else:
                if cur:
                    lines.append(cur)
                cur = w
        lines.append(cur)
    return lines


def fit_text(draw, text, kind, max_w, max_h, start, minimum, spacing=1.3):
    for size in range(start, minimum - 1, -2):
        f = font(kind, size)
        lines = wrap(draw, text, f, max_w)
        if len(lines) * size * spacing <= max_h:
            return f, lines, size
    f = font(kind, minimum)
    return f, wrap(draw, text, f, max_w), minimum


def draw_code(draw, code, th, x, y, w, h):
    code = code.replace("\t", "    ").strip("\n")
    pad = 30
    ix, iw, ih = x + pad, w - 2 * pad, h - 66 - pad
    lines = code.split("\n") or [""]
    size = 40
    while size > 18:
        f = font("mono", size)
        cw = draw.textlength("M", font=f)
        if max(len(l) for l in lines) * cw <= iw and len(lines) * size * 1.32 <= ih:
            break
        size -= 2
    f = font("mono", size)
    cw = draw.textlength("M", font=f)
    maxch = max(10, int(iw // cw))
    flat = []
    for l in lines:  # hard-wrap anything still too long
        indent = len(l) - len(l.lstrip())
        while len(l) > maxch:
            flat.append(l[:maxch]); l = " " * (indent + 2) + l[maxch:]
        flat.append(l)
    panel_h = min(h, int(66 + len(flat) * size * 1.32 + pad))
    draw.rounded_rectangle([x, y, x + w, y + panel_h], radius=28, fill=th["code_bg"])
    for i, col in enumerate([(248, 113, 113), (251, 191, 36), (74, 222, 128)]):
        draw.ellipse([x + 30 + i * 34, y + 26, x + 46 + i * 34, y + 42], fill=col)
    tok = re.compile(r"(--.*|#.*)|('(?:[^']|'')*'|\"[^\"]*\")|(\b\d+(?:\.\d+)?\b)|(\b[A-Za-z_][A-Za-z_0-9]*\b)|(\s+)|(.)")
    ly = y + 66
    for l in flat:
        cx = ix
        for m in tok.finditer(l):
            t = m.group(0)
            if m.group(1): col = th["cm"]
            elif m.group(2): col = th["st"]
            elif m.group(3): col = th["num"]
            elif m.group(4) and t.upper() in SQL_KW: col = th["kw"]
            else: col = th["code_text"]
            draw.text((cx, ly), t, font=f, fill=col)
            cx += draw.textlength(t, font=f)
        ly += size * 1.32

    return panel_h


def blend(c1, c2, t):
    return tuple(int(c1[i] + (c2[i] - c1[i]) * t) for i in range(3))


def lum(c):
    return 0.299 * c[0] + 0.587 * c[1] + 0.114 * c[2]


def mid_bg(th):
    return blend(th["bg"][0], th["bg"][1], 0.5)


def split_lead(p):
    p = clean(p)
    if ":" in p[:70]:
        lead, rest = p.split(":", 1)
        return lead.strip(), rest.strip()
    return "", p


def draw_points(d, pts, th, x, y, w, h, max_size=48, badge_kind="num", compact=False, split=True):
    """Cards with a numbered badge, bold lead and a short explanation. Fills the area."""
    pts = [p for p in pts if p and str(p).strip()][:5]
    if not pts or h < 60:
        return y
    parts = [split_lead(p) if split else ("", clean(p)) for p in pts]
    pad = 16 if compact else 26
    gap = 12 if compact else 20
    card = blend(mid_bg(th), th["text"], 0.09)
    fg = (20, 20, 30) if lum(th["accent"]) > 140 else (255, 255, 255)
    for size in range(max_size, 21, -2):
        badge = int(size * (1.05 if compact else 1.3))
        tx = x + pad + badge + 22
        tw = w - (tx - x) - pad
        fl, fr = font("bold", size), font("reg", size - 4)
        blocks, total = [], 0
        for lead, rest in parts:
            ll = wrap(d, lead, fl, tw) if lead else []
            rl = wrap(d, rest, fr, tw) if rest else []
            if not lead:  # no lead: show the sentence at the larger size
                rl, fr_use = wrap(d, rest, fl, tw), fl
            else:
                fr_use = fr
            body_h = len(ll) * size * 1.25 + len(rl) * (fr_use.size) * 1.32 + (6 if ll and rl else 0)
            ch = max(body_h, badge) + 2 * pad
            blocks.append((ll, rl, ch, fr_use))
            total += ch
        total += gap * (len(pts) - 1)
        if total <= h:
            break
    while total > h and len(blocks) > 1:  # last resort: never overlap the takeaway box
        total -= blocks.pop()[2] + gap
    leftover = max(0, h - total)
    extra = min(26, leftover / max(1, len(pts) - 1)) if len(pts) > 1 else 0
    y0 = y + min(50, max(0, leftover - extra * (len(pts) - 1)) / 2)
    cy = y0
    for i, (ll, rl, ch, fr_use) in enumerate(blocks):
        d.rounded_rectangle([x, cy, x + w, cy + ch], radius=22, fill=card)
        by = cy + (ch - badge) / 2
        d.ellipse([x + pad, by, x + pad + badge, by + badge], fill=th["accent"])
        glyph = "\u2713" if badge_kind == "check" else str(i + 1)
        gf = font("bold", int(badge * 0.5))
        d.text((x + pad + badge / 2, by + badge / 2), glyph, font=gf, fill=fg, anchor="mm")
        ty = cy + (ch - (len(ll) * size * 1.25 + len(rl) * fr_use.size * 1.32 + (6 if ll and rl else 0))) / 2
        tx = x + pad + badge + 22
        for l in ll:
            d.text((tx, ty), l, font=font("bold", size), fill=th["text"]); ty += size * 1.25
        if ll and rl:
            ty += 6
        for l in rl:
            col = th["muted"] if ll else th["text"]
            d.text((tx, ty), l, font=fr_use, fill=col); ty += fr_use.size * 1.32
        cy += ch + gap + extra
    return cy


def callout_box(d, text, th, w):
    text = clean(text)
    pad = 26
    f, lines, size = fit_text(d, text, "bold", w - 2 * pad - 14, 150, 40, 28, 1.25)
    return f, lines, size, int(26 + len(lines) * size * 1.25 + 2 * pad)


def draw_callout(d, th, x, bottom, w, f, lines, size, h):
    top = bottom - h
    tint = blend(mid_bg(th), th["accent"], 0.22)
    d.rounded_rectangle([x, top, x + w, bottom], radius=22, fill=tint)
    d.rounded_rectangle([x, top, x + 12, bottom], radius=6, fill=th["accent"])
    d.text((x + 38, top + 22), "KEY TAKEAWAY", font=font("bold", 22), fill=th["accent"])
    ty = top + 22 + 36
    for l in lines:
        d.text((x + 38, ty), l, font=f, fill=th["text"]); ty += size * 1.25
    return top


def draw_compare(d, before, after, th, x, y, w, h):
    gap = 22
    ch = (h - gap) / 2
    card = blend(mid_bg(th), th["text"], 0.09)
    for i, (label, txt, col) in enumerate((("BEFORE", before, (248, 113, 113)), ("AFTER", after, (74, 222, 128)))):
        cy = y + i * (ch + gap)
        d.rounded_rectangle([x, cy, x + w, cy + ch], radius=24, fill=card)
        d.rounded_rectangle([x, cy, x + 12, cy + ch], radius=6, fill=col)
        d.text((x + 40, cy + 28), label, font=font("bold", 30), fill=col)
        f, lines, size = fit_text(d, clean(txt), "reg", w - 80, ch - 100, 46, 26, 1.35)
        ty = cy + 78
        for l in lines:
            if l == "":
                ty += size * 0.3
                continue
            d.text((x + 40, ty), l, font=f, fill=th["text"]); ty += size * 1.35
    return y + h


def render_slide(sl, idx, total, th):
    img = gradient(*th["bg"])
    d = ImageDraw.Draw(img)
    cw = W - 2 * M
    # header / footer
    d.text((M, 64), HANDLE, font=font("bold", 30), fill=th["accent"])
    pg = f"{idx}/{total}"
    d.text((W - M - d.textlength(pg, font=font("reg", 30)), 64), pg, font=font("reg", 30), fill=th["muted"])
    foot = "Swipe  \u2192" if idx < total else "Save \u2022 Share \u2022 Follow"
    ff = font("bold", 30)
    d.text((W - M - d.textlength(foot, font=ff), H - 88), foot, font=ff, fill=th["muted"])
    d.text((M, H - 86), "DATA ANALYTICS", font=font("reg", 26), fill=th["muted"])

    bottom = H - 125
    if sl["callout"]:
        cf, cl, cs, chh = callout_box(d, sl["callout"], th, cw)
        ctop = draw_callout(d, th, M, bottom, cw, cf, cl, cs, chh)
        content_bottom = ctop - 28
    else:
        content_bottom = bottom

    title = clean(sl["title"])
    if sl["layout"] == "cover":
        d.rounded_rectangle([M, 150, M + 120, 162], radius=6, fill=th["accent"])
        kick = clean(sl["kicker"]).upper()
        if kick:
            d.text((M, 185), kick, font=font("bold", 28), fill=th["accent"])
        f, lines, size = fit_text(d, title, "bold", cw, 300, 84, 54, 1.15)
        y = 240
        for l in lines:
            d.text((M, y), l, font=f, fill=th["text"]); y += size * 1.15
        if sl["subtitle"]:
            fb, bl, bs = fit_text(d, clean(sl["subtitle"]), "reg", cw, 170, 40, 26, 1.35)
            y += 18
            for l in bl:
                d.text((M, y), l, font=fb, fill=th["muted"]); y += bs * 1.35
        y += 30
        draw_points(d, sl["points"], th, M, int(y), cw, int(content_bottom - y), 40, "check", compact=True)
        return img

    kick = clean(sl["kicker"]).upper()
    d.rounded_rectangle([M, 140, M + 100, 150], radius=5, fill=th["accent"])
    if kick:
        d.text((M, 168), kick, font=font("bold", 28), fill=th["accent"])
    f, lines, size = fit_text(d, title, "bold", cw, 160, 64, 42, 1.15)
    y = 212
    for l in lines:
        d.text((M, y), l, font=f, fill=th["text"]); y += size * 1.15
    y = int(y + 26)
    area = content_bottom - y

    if sl["layout"] == "code":
        has_notes = bool(sl["notes"])
        panel_max = int(area * (0.56 if has_notes else 1.0))
        ph = draw_code(d, sl["code"], th, M, y, cw, panel_max)
        y = int(y + ph + 24)
        if has_notes:
            draw_points(d, sl["notes"], th, M, y, cw, int(content_bottom - y), 36, "num", compact=True, split=False)
    elif sl["layout"] == "compare":
        draw_compare(d, sl["before"], sl["after"], th, M, y, cw, area)
    else:
        draw_points(d, sl["points"], th, M, y, cw, area, 50, "num")
    return img


def make_images(topic, c, theme_name):
    th = THEMES[theme_name]
    out = []
    slides = c.get("carousel_slides") or []
    if slides:
        n = len(slides)
        for i, raw in enumerate(slides, 1):
            out.append((f"slide_{i:02d}.png", render_slide(normalize_slide(raw, i), i, n, th)))
    else:  # Reel / single image -> cover thumbnail
        hooks = c.get("hooks") or [""]
        cover = normalize_slide(dict(layout="cover", kicker=topic.get("pillar", ""),
                                     title=c.get("cover_text") or topic.get("topic", ""),
                                     subtitle=hooks[0], callout=c.get("cta", "")), 1)
        out.append(("cover.png", render_slide(cover, 1, 1, th)))
    files = []
    for name, im in out:
        b = io.BytesIO(); im.save(b, "PNG"); files.append((name, b.getvalue()))
    return files


def zip_files(files):
    b = io.BytesIO()
    with zipfile.ZipFile(b, "w", zipfile.ZIP_DEFLATED) as z:
        for n, data in files:
            z.writestr(n, data)
    return b.getvalue()


def pdf_from_pngs(files, title="Instagram carousel"):
    """All slide PNGs -> one PDF, one slide per page, high quality so code text stays sharp."""
    pages = [Image.open(io.BytesIO(data)).convert("RGB") for _, data in files]
    buf = io.BytesIO()
    kw = dict(save_all=True, append_images=pages[1:], resolution=150.0)  # 1080x1350 px -> 7.2 x 9 in pages
    try:
        pages[0].save(buf, "PDF", quality=95, subsampling=0, title=title, author="@analyst_shubhi", **kw)
    except (TypeError, ValueError):  # older Pillow without these options
        buf = io.BytesIO()
        pages[0].save(buf, "PDF", **kw)
    return buf.getvalue()


# ---------------------------- State & caching ----------------------------
@st.cache_resource
def topic_history() -> list:
    """Shared in-memory list of topics used (resets if the app restarts)."""
    return []


@st.cache_data(show_spinner=False)
def plan_for(day: str, version: int, profile: str, theme_name: str, level: str):
    """Cached per day: reopening the app today reuses the plan; 'version' forces a new one."""
    hist = topic_history()
    topic, content = generate_pack(profile, hist, theme_name, WEEKDAY_THEMES[theme_name], level)
    hist.append(topic.get("topic", ""))
    return topic, content, build_docx(topic, content)


@st.cache_data(show_spinner=False)
def render_exports(topic: dict, c: dict, theme_name: str, title: str):
    """Cached so PNGs / ZIP / PDF are not rebuilt on every rerun (e.g. after a download click)."""
    files = make_images(topic, c, theme_name)
    return files, zip_files(files), pdf_from_pngs(files, title)


# ---------------------------- UI ----------------------------
greeting_banner()
st.title("📊 Daily Content Bot - @analyst_shubhi")

if "version" not in st.session_state:
    st.session_state.version = 0
if "applied" not in st.session_state:
    st.session_state.applied = None


def friendly_error(e: Exception) -> str:
    t = str(e)
    if "RESOURCE_EXHAUSTED" in t or "429" in t:
        m = re.search(r"retry in ([0-9hms\.]+)", t)
        wait = f" Google says retry in about {m.group(1)}." if m else ""
        return ("**Gemini free-tier limit reached** (about 20 requests per day per model)." + wait +
                "\n\nOptions: wait and click **Try again**; add more models in Secrets, e.g. "
                '`GEMINI_MODELS = "gemini-3.5-flash,gemini-3.7-flash,gemini-2.5-flash"` (each has its own free quota); '
                "or enable billing in Google AI Studio (a post costs a fraction of a cent).")
    if "TIMEOUT" in t:
        return ("**Gemini took too long to answer** (over 100 seconds). This is usually temporary. "
                "Click **Try again**. If it keeps happening, switch to a faster model in Secrets, e.g. "
                '`GEMINI_MODELS = "gemini-2.5-flash,gemini-3.5-flash"`.')
    return f"Generation failed: {t}"


with st.expander("Creator profile (applies to this session; edit DEFAULT_PROFILE in code to make permanent)"):
    profile = st.text_area("Profile", DEFAULT_PROFILE, height=350)

today_date = now_ist().date()
today = today_date.isoformat()
theme_names = list(WEEKDAY_THEMES)
theme_pick = st.selectbox("Content theme", theme_names, index=today_date.weekday(),
                          help="Defaults to today's weekday theme; change it to explore other content types.")
st.caption(WEEKDAY_THEMES[theme_pick])
level = st.radio("Difficulty", ["Beginner", "Beginner+", "Intermediate"], horizontal=True,
                 help="Beginner = basics a fresher meets in the first months. Keep it on Beginner for your audience.")

current = (profile, theme_pick, level)
if st.session_state.applied is None:  # nothing generated yet: wait for the button (no auto-generation)
    st.info("Nothing is generated automatically. Click the button to create today's content (uses 1 Gemini request).")
    if st.button("✨ Generate", type="primary", width="stretch"):
        st.session_state.applied = current
        st.session_state.version += 1
        st.rerun()
    st.stop()
if current != st.session_state.applied:
    st.warning("Settings changed. Click below to apply (uses 1 Gemini request).")
    if st.button("Generate with new settings", type="primary"):
        st.session_state.applied = current
        st.session_state.version += 1
        st.rerun()

p_, t_, l_ = st.session_state.applied
key = (today, st.session_state.version, p_, t_, l_)

if st.session_state.get("err_key") == key:  # failed earlier: do NOT auto-retry (saves quota)
    st.error(st.session_state.get("err_msg", "Generation failed."))
    if st.button("Try again"):
        st.session_state.err_key = None
        st.rerun()
    st.stop()

try:
    with st.spinner("Picking today's topic and writing your content (1 Gemini request, usually 30-90 seconds)..."):
        topic, c, docx_bytes = plan_for(*key)
except Exception as e:
    st.session_state.err_key = key
    st.session_state.err_msg = friendly_error(e)
    st.error(st.session_state.err_msg)
    if st.button("Try again"):
        st.session_state.err_key = None
        st.rerun()
    st.stop()


def copy_box(text):
    try:
        st.code(text, language=None, wrap_lines=True)  # has a copy button
    except TypeError:  # older Streamlit without wrap_lines
        st.code(text, language=None)


post_title = c.get("post_title") or topic.get("topic", "")
caption_full = (c.get("caption", "").strip() + "\n\n" + " ".join(c.get("hashtags", []))).strip()

st.caption(f"{topic.get('format', '')} · {topic.get('pillar', '')} · best time: {c.get('best_time_to_post', '')}")
st.markdown("### 📝 Post title")
copy_box(post_title)
if c.get("alt_titles"):
    st.caption("Other options: " + "  |  ".join(str(a) for a in c["alt_titles"]))
st.markdown("### ✍️ Caption")
copy_box(caption_full)

b1, b2 = st.columns(2)
b1.download_button(
    "⬇️ Download .docx", docx_bytes, file_name=f"content_plan_{today}.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    width="stretch")
if b2.button("🔄 Different topic", width="stretch"):
    st.session_state.applied = current
    st.session_state.version += 1
    st.rerun()

# ---- PNG export
st.markdown("### 🎨 Slides (PNG)")
theme_name = st.selectbox("Theme", list(THEMES), key="theme")
files, zip_bytes, pdf_bytes = render_exports(topic, c, theme_name, post_title)
dl_zip, dl_pdf = st.columns(2)
dl_zip.download_button(f"⬇️ All {len(files)} PNG(s) (.zip)", zip_bytes,
                       file_name=f"{today}_{theme_name.lower().replace(' ', '_')}_slides.zip",
                       mime="application/zip", width="stretch")
dl_pdf.download_button(f"⬇️ Download PDF ({len(files)} pages)", pdf_bytes,
                       file_name=f"{today}_{theme_name.lower().replace(' ', '_')}_slides.pdf",
                       mime="application/pdf", type="primary", width="stretch")
cols = st.columns(3)
for i, (name, data) in enumerate(files):
    with cols[i % 3]:
        st.image(data, width="stretch")
        st.download_button("PNG", data, file_name=name, mime="image/png", key=f"png_{i}_{theme_name}",
                           width="stretch")

# ---- everything else, collapsed
with st.expander("Full content pack (scenario, slide text, code, gotchas...)"):
    st.markdown(f"**Topic:** {topic.get('topic', '')}")
    st.write(f"**Angle:** {topic.get('angle', '')}")
    st.write(f"**Why today:** {topic.get('why_today', '')}")
    st.write(f"**Work scenario:** {c.get('work_scenario') or topic.get('work_scenario', '')}")
    st.markdown("**Hooks**")
    for h in c.get("hooks", []):
        st.write("- " + str(h))
    if c.get("sample_data") or c.get("expected_output"):
        st.markdown("**Sample data & expected output**")
        if c.get("sample_data"):
            st.code(c["sample_data"])
        if c.get("expected_output"):
            st.code(c["expected_output"])
    ba = c.get("before_after") or {}
    if ba:
        st.markdown("**Impact on daily work**")
        st.markdown(f"**Before:** {ba.get('before', '')}\n\n**After:** {ba.get('after', '')}\n\n**Impact:** {ba.get('impact', '')}")
    if c.get("reel_script"):
        st.markdown("**Reel script**")
        st.table(c["reel_script"])
    if c.get("carousel_slides"):
        st.markdown("**Carousel slides (text)**")
        for i, raw in enumerate(c["carousel_slides"], 1):
            sl = normalize_slide(raw, i)
            st.markdown(f"**Slide {i}: {sl['title']}**" + (f"  ·  _{sl['kicker']}_" if sl["kicker"] else ""))
            for p_ in sl["points"]:
                st.write("- " + p_)
            if sl["code"]:
                st.code(sl["code"])
            for n_ in sl["notes"]:
                st.write("- " + n_)
            if sl["before"] or sl["after"]:
                st.write(f"**Before:** {sl['before']}  \n**After:** {sl['after']}")
            if sl["callout"]:
                st.info(sl["callout"])
    if c.get("code_or_example"):
        st.markdown("**Code / example**")
        st.code(c["code_or_example"])
    if c.get("gotchas"):
        st.markdown("**Gotchas**")
        for g in c["gotchas"]:
            st.write("- " + str(g))
    if c.get("interview_angle"):
        st.write(f"**Interview angle:** {c['interview_angle']}")
    if c.get("lead_magnet_keyword"):
        st.write(f"**Lead magnet keyword:** {c['lead_magnet_keyword']}")
    st.markdown("**Story ideas**")
    for s_ in c.get("story_ideas", []):
        st.write("- " + str(s_))
    st.write(f"**Cover text:** {c.get('cover_text', '')}")
