import os, io, re, csv, html, json, time, tempfile, subprocess
from pathlib import Path

import requests
import streamlit as st
from dotenv import load_dotenv
from google import genai
from google.genai import errors

import importlib
import sys


def _fresh(name):
    """Import one of our own modules and always re-read it from disk, so a long-running server never uses a stale
    cached copy. Order matters: modules that import from another one must be loaded after it is refreshed."""
    mod = sys.modules.get(name)
    return importlib.reload(mod) if mod else importlib.import_module(name)


docx_builder = _fresh("docx_builder")
hb = _fresh("handbook")
preview = _fresh("preview")   # imports from docx_builder, so it comes after it
build_docx = docx_builder.build_docx
preview_html = preview.preview_html

load_dotenv()
client = genai.Client(api_key=os.getenv("GEMINI_API_KEY"))
FALLBACK_MODEL = "gemini-3.8-flash-lite"   # assumed name; change if it 404s
MODEL = "gemini-3.8-flash"

PROMPT_GENERAL = """You are a meeting secretary. This is a meeting recording in Tagalog/Taglish.
Write the Minutes of Meeting in English. Output ONLY valid JSON in this shape:
{"title": "", "date": "", "time": "", "location": "", "purpose": "",
 "attendees": [{"group": "", "members": ["Name - Role"]}],
 "call_to_order": "", "summary": "", "discussion": [], "decisions": [],
 "action_items": [{"task": "", "owner": "", "deadline": ""}], "notes": ""}
- "title": a short subject for the meeting. "purpose": one short line.
- "attendees": group people when the recording makes groups clear (for example "Management",
  "Employee involved"); otherwise use a single group with an empty "group" name.
- "call_to_order": one sentence, e.g. "The meeting was called to order at 10:00 AM by Ms. Santos."
- "summary": 3-5 sentences stating the purpose and overall outcome of the meeting.
- "discussion": DETAILED, one entry per topic, each a full paragraph covering what was raised, who said
  what (if names are stated), the reasons, concerns, numbers/dates mentioned, and how it was resolved.
- "decisions": one entry per decision. "notes": anything else worth recording, or "".
Rules: Do not invent anything. Only use what is said in the recording.
If something is not stated, use "TBD". Keep names exactly as spoken."""


PROMPT_HEARING = """You are an HR secretary. This is a recording of an administrative hearing for a Notice to Explain (NTE)
in Tagalog/Taglish. Write the Minutes of the Meeting in English. Output ONLY valid JSON in this shape:
{"employee": "", "position": "", "date": "", "time": "", "location": "", "purpose": "",
 "attendees": [{"group": "", "members": ["Name - Position"]}],
 "call_to_order": "", "summary": "", "discussion": [], "review_intro": "", "decisions": [],
 "offense": "", "handbook": ""}
- "employee": full name of the employee being heard. "position": their job title.
- "purpose": "Admin Hearing for Notice to Explain regarding <employee>".
- "attendees": two groups when possible: {"group": "Employee Involve: <employee name>", "members": []} and
  {"group": "Management", "members": ["Name - Position", ...]}.
- "call_to_order": one sentence like "The meeting was called to order at 10:00 AM by Ms. Serrano."
- "summary": "The purpose of the meeting was to discuss the Notice to Explain (NTE) issued to <employee> regarding
  <the incident, with ticket or reference numbers if they are stated>."
- "discussion": DETAILED and in order: one paragraph per question or point, covering who asked what, the employee's
  explanation, recommendations from management, and what the employee committed to do.
  Do NOT put management's final decision or sanction in "discussion"; it belongs only in "decisions".
- "review_intro": "Following the discussion, Management reviewed <employee>'s explanation and assessed the facts
  surrounding the issue. After thorough consideration, Management reached the following decision:"
- "decisions": the sanction or decision management reached, only if it is stated, otherwise ["TBD"].
- "offense" and "handbook": leave BOTH as "" unless the speakers explicitly name a Company Employee Handbook
  offense or section number. Never describe the incident here and never make an offense up.
  "handbook" has one section per line.
Rules: Do not invent anything. Only use what is said in the recording.
If something is not stated, use "TBD". Keep names exactly as spoken."""
PROMPTS = {"hearing": PROMPT_HEARING, "general": PROMPT_GENERAL}


def normalize_mom(m, kind="general"):
    """Make sure every field the UI and the Word file need exists, whatever shape the AI returned."""
    m["kind"] = kind
    groups, loose = [], []
    for a in m.get("attendees") or []:
        if isinstance(a, dict) and ("members" in a or "group" in a):
            groups.append({"group": str(a.get("group") or "").strip(),
                           "members": [str(x) for x in (a.get("members") or [])]})
        elif isinstance(a, dict):
            nm, role = a.get("name") or "", a.get("role") or ""
            loose.append(f"{nm} - {role}" if role else str(nm))
        else:
            loose.append(str(a))
    if loose:
        groups.insert(0, {"group": "", "members": loose})
    m["attendees"] = groups
    for k in ("title", "date", "time", "location", "purpose", "call_to_order", "summary", "notes",
              "employee", "position", "review_intro", "offense", "handbook"):
        m[k] = str(m.get(k) or "")
    for k in ("discussion", "decisions"):
        m[k] = [str(x) for x in (m.get(k) or [])]
    m["action_items"] = [a for a in (m.get("action_items") or []) if isinstance(a, dict)]
    if kind == "hearing":
        emp = m["employee"].strip()
        for g in m["attendees"]:   # the employee goes on the "Employee Involve" line only, not again as a member
            if g["group"].lower().startswith("employee"):
                if emp and emp.lower() not in g["group"].lower():
                    g["group"] = f"{g['group'].rstrip(': ')}: {emp}"
                g["members"] = []
        if not m["handbook"].strip():   # an offense title without handbook sections is a guess, not a citation
            m["offense"] = ""
        if emp and not m["purpose"].strip():
            m["purpose"] = f"Admin Hearing for Notice to Explain regarding {emp}"
        if not m["review_intro"].strip():
            m["review_intro"] = (f"Following the discussion, Management reviewed {emp + chr(39) + 's' if emp else 'the'} explanation and "
                                 "assessed the facts surrounding the issue. After thorough consideration, Management reached "
                                 "the following decision:")
    return m


def compress(path):
    """Speech-only mono 16kHz mp3: much smaller upload, faster processing."""
    out = path + ".small.mp3"
    try:
        subprocess.run(
            ["ffmpeg", "-y", "-i", path, "-vn", "-ac", "1", "-ar", "16000", "-b:a", "32k", out],
            check=True, capture_output=True,
        )
        return out
    except Exception:
        return path   # ffmpeg missing/failed: use original


def split_audio(path, limit_mb=24, seg_seconds=1200):
    """Return [path] if it is small enough for one transcription call, else a list of ~20 minute pieces."""
    if os.path.getsize(path) <= limit_mb * 1e6:
        return [path]
    pattern = path + ".part%03d.mp3"
    try:
        subprocess.run(["ffmpeg", "-y", "-i", path, "-f", "segment", "-segment_time", str(seg_seconds),
                        "-c", "copy", pattern], check=True, capture_output=True)
    except Exception:
        return [path]
    parts = sorted(str(p) for p in Path(path).parent.glob(Path(path).name + ".part*.mp3") if p.stat().st_size > 20_000)
    return parts or [path]   # tiny leftover slivers are dropped: the transcription service rejects them


def make_mom_gemini_audio(path, kind="general"):
    path = compress(path)
    f = client.files.upload(file=path)
    while f.state.name == "PROCESSING":
        time.sleep(2)
        f = client.files.get(name=f.name)
    resp = None
    for model in [MODEL, MODEL, MODEL, FALLBACK_MODEL, FALLBACK_MODEL]:
        try:
            resp = client.models.generate_content(
                model=model,
                contents=[f, PROMPTS[kind]],
                config={"response_mime_type": "application/json"},
            )
            break
        except errors.ServerError:   # 503 overloaded: wait and retry
            time.sleep(10)
    client.files.delete(name=f.name)   # delete the audio from the server afterwards
    if resp is None:
        raise RuntimeError("Gemini is still busy. Please try again later.")
    return json.loads(resp.text)


GROQ_KEY = os.getenv("GROQ_API_KEY")
GROQ_URL = "https://api.groq.com/openai/v1"
GROQ_WHISPER = "whisper-large-v3-turbo"
GROQ_LLM = "openai/gpt-oss-120b"   # change if Groq retires it


def groq_transcribe(path, hint="", tries=4):
    for _ in range(tries):
        with open(path, "rb") as fh:
            r = requests.post(
                f"{GROQ_URL}/audio/transcriptions",
                headers={"Authorization": f"Bearer {GROQ_KEY}"},
                files={"file": (os.path.basename(path), fh)},
                data={"model": GROQ_WHISPER, "response_format": "text",
                      **({"prompt": f"Names and terms: {hint}"} if hint else {})},
                timeout=600,
            )
        if r.status_code == 429:   # free-tier audio limit per minute/hour: wait a bit and try again
            m = re.search(r"try again in (?:(\d+)m)?([\d.]+)s", r.text)
            time.sleep(min((int(m.group(1) or 0) * 60 + float(m.group(2))) + 1 if m else 30, 120))
            continue
        break
    r.raise_for_status()
    return r.text


def as_dict(x):
    """The model sometimes wraps the JSON in a list; accept that, reject anything else."""
    if isinstance(x, list) and x and isinstance(x[0], dict):
        x = x[0]
    if not isinstance(x, dict):
        raise ValueError("AI returned an unexpected format")
    return x


def groq_chat(payload, timeout=180, tries=4):
    """Groq chat call. Free tier allows ~8000 tokens/minute, so wait and retry on 429."""
    for _ in range(tries):
        r = requests.post(f"{GROQ_URL}/chat/completions", headers={"Authorization": f"Bearer {GROQ_KEY}"},
                          json=payload, timeout=timeout)
        if r.status_code == 429:
            msg = r.text
            if "too large" in msg.lower():   # one request bigger than the per-minute budget: waiting won't help
                break
            m = re.search(r"try again in (?:(\d+)m)?([\d.]+)s", msg)
            wait = (int(m.group(1) or 0) * 60 + float(m.group(2))) if m else 20
            time.sleep(min(wait + 1, 65))
            continue
        r.raise_for_status()
        return r.json()["choices"][0]["message"]["content"] or ""
    r.raise_for_status()


def mom_from_text(transcript, lang="English", hint="", kind="general"):
    """Text -> MOM JSON. Groq first (fast), Gemini only as fallback."""
    prompt = PROMPTS[kind].replace("This is a meeting recording", "Below is a transcript of a meeting recording")
    prompt = prompt.replace("This is a recording", "Below is a transcript")
    prompt = prompt.replace("Minutes of Meeting in English", f"Minutes of Meeting in {lang}")
    prompt = prompt.replace("Minutes of the Meeting in English", f"Minutes of the Meeting in {lang}")
    if hint:
        prompt += f"\nKnown names/terms (use these exact spellings): {hint}"
    prompt += "\n\nTRANSCRIPT:\n" + transcript
    try:
        content = groq_chat({"model": GROQ_LLM, "messages": [{"role": "user", "content": prompt}],
                             "response_format": {"type": "json_object"}, "max_completion_tokens": 4000}, timeout=120)
        return as_dict(json.loads(content))
    except Exception:
        resp = client.models.generate_content(
            model=MODEL, contents=prompt,
            config={"response_mime_type": "application/json"})
        return as_dict(json.loads(resp.text))


CONVO_PROMPT = """Reformat this raw meeting transcript as a conversation, one speaker turn per line,
like "Speaker 1: ...". Guess where the speaker changes from context (questions and answers,
greetings, replies). Keep the same label for the same person throughout. Use a real name as the
label only if that person is clearly addressed or introduces themselves.
STRICT RULES: do not translate, summarize, add, remove or change any words; keep the original
language (Tagalog/Taglish/English). Output plain text only, no commentary, no markdown."""


def chunk_text(text, size=4000):
    chunks, cur = [], ""
    for sent in re.split(r"(?<=[.!?])\s+", text.strip()):
        if len(cur) + len(sent) > size and cur:
            chunks.append(cur)
            cur = ""
        cur += sent + " "
    if cur.strip():
        chunks.append(cur)
    return chunks


def to_conversation(text):
    """Turn a raw transcript into 'Speaker N: ...' turns. Falls back to raw text on any failure."""
    out, prev = [], ""
    try:
        for ch in chunk_text(text):
            ctx = f"\n\n(Previous turns, for consistent speaker labels:\n{prev})" if prev else ""
            part = groq_chat({"model": GROQ_LLM, "temperature": 0, "reasoning_effort": "low", "max_completion_tokens": 3500,
                              "messages": [{"role": "user", "content": CONVO_PROMPT + ctx + "\n\nTRANSCRIPT:\n" + ch}]}).strip()
            if len(part.split()) < 0.85 * len(ch.split()):   # model dropped/truncated words: keep raw chunk
                part = ch.strip()
            out.append(part)
            prev = "\n".join(part.splitlines()[-3:])
        return "\n".join(out)
    except Exception:
        return text


def make_mom(path, lang="English", hint="", kind="general"):
    if not GROQ_KEY:
        return normalize_mom(make_mom_gemini_audio(path, kind), kind)   # no Groq key: Gemini-only path
    t = time.time()
    small = compress(path)
    st.write(f"1. Audio extracted: {time.time()-t:.0f}s ({os.path.getsize(path)/1e6:.1f}MB -> {os.path.getsize(small)/1e6:.1f}MB)")
    t = time.time()
    parts = split_audio(small)   # Groq accepts at most ~25MB per file, so long recordings go in pieces
    text = "\n".join(groq_transcribe(p, hint) for p in parts)
    st.write(f"2. Transcript: {time.time()-t:.0f}s" + (f" ({len(parts)} parts)" if len(parts) > 1 else ""))
    t = time.time()
    mom = normalize_mom(mom_from_text(text, lang, hint, kind), kind)
    st.write(f"3. Minutes: {time.time()-t:.0f}s")
    t = time.time()
    mom["transcript"] = to_conversation(text.strip())
    st.write(f"4. Conversation format: {time.time()-t:.0f}s")
    return mom


# ---------------------------------------------------------------- attendee text <-> groups
def attendees_to_text(groups):
    """'Group' line, then '- member' lines. Members with no group are plain '- member' lines."""
    out = []
    for g in groups or []:
        if g.get("group"):
            out.append(g["group"])
        out += [f"- {x}" for x in g.get("members", [])]
    return "\n".join(out)


def text_to_attendees(text):
    groups, cur = [], None
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        if line.startswith("-"):
            if cur is None:
                cur = {"group": "", "members": []}
                groups.append(cur)
            cur["members"].append(line.lstrip("- ").strip())
        else:
            cur = {"group": line, "members": []}
            groups.append(cur)
    return groups


def count_attendees(groups):
    return sum(len(g.get("members", [])) + (1 if g.get("group") and not g.get("members") else 0) for g in groups or [])


# ---------------------------------------------------------------- other exports
def build_md(m):
    out = [f"# {m.get('title') or 'Minutes of the Meeting'}",
           f"**Date:** {m.get('date') or 'TBD'}  ", f"**Time:** {m.get('time') or 'TBD'}  ",
           f"**Location:** {m.get('location') or 'TBD'}  ", f"**Purpose:** {m.get('purpose') or 'TBD'}", "", "## Attendees"]
    for g in m.get("attendees", []):
        if g.get("group"):
            out.append(f"- {g['group']}")
        out += [f"  - {x}" if g.get("group") else f"- {x}" for x in g.get("members", [])]
    out += ["", "## 1. Call to Order", m.get("call_to_order") or "", "", "## 2. Purpose of the Meeting", m.get("summary") or "",
            "", "## 3. Discussion Summary"] + [p + "\n" for p in m.get("discussion", [])]
    out += ["## 4. Decisions"] + [f"- {x}" for x in m.get("decisions", [])] + ["", "## 5. Action Items",
            "| Task | Owner | Deadline |", "|---|---|---|"]
    out += [f"| {a.get('task','')} | {a.get('owner','TBD')} | {a.get('deadline','TBD')} |" for a in m.get("action_items", [])]
    if m.get("notes"):
        out += ["", "## Additional Notes", m["notes"]]
    pb = m.get("prepared_by") or {}
    if pb.get("name"):
        out += ["", f"Prepared by: {pb['name']}" + (f", {pb['role']}" if pb.get("role") else "")]
    if m.get("transcript"):
        out += ["", "## Transcript", m["transcript"]]
    return "\n".join(out)


def build_csv(m):
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["Task", "Owner", "Deadline"])
    for a in m.get("action_items", []):
        w.writerow([a.get("task", ""), a.get("owner", "TBD"), a.get("deadline", "TBD")])
    return buf.getvalue().encode("utf-8-sig")   # BOM so Excel opens it correctly


def build_email(m):
    lines_ = [f"Subject: Minutes of Meeting - {m.get('title') or 'Meeting'} ({m.get('date') or 'TBD'})", "",
              "Hi all,", "", "Here is a summary of our meeting.", ""]
    if m.get("summary"):
        lines_ += [m["summary"], ""]
    if m.get("decisions"):
        lines_ += ["Decisions:"] + [f"- {x}" for x in m["decisions"]] + [""]
    if m.get("action_items"):
        lines_ += ["Action items:"] + [
            f"- {a.get('task','')} (Owner: {a.get('owner','TBD')}, Due: {a.get('deadline','TBD')})"
            for a in m["action_items"]] + [""]
    lines_ += ["Please reply if I missed or got anything wrong.", "", "Thanks!"]
    return "\n".join(lines_)


def suggest_sections(case):
    """Ask the AI which Company Employee Handbook sections (from Sanction.pdf) match this hearing.
    It can only answer with indexes from the real list, so it cannot invent a section."""
    text = "\n".join([case.get("summary", ""), *case.get("discussion", []),
                      "Decision: " + "; ".join(case.get("decisions", []))])[:4000]
    prompt = ("You help HR cite the Company Employee Handbook. Below is a numbered list of handbook sections "
              "(index|section|text|degree) and the minutes of an administrative hearing. Choose the sections that clearly "
              "describe the offense the employee is being heard for. Pick the 2 or 3 most specific sections (never more than 3) "
              "and never two sections that say the same thing. If none clearly applies, return an "
              'empty list. Answer ONLY with JSON: {"indices": [<index numbers from the list>]}\n\n'
              "HANDBOOK:\n" + "\n".join(f"{i}|{no}|{t[:80]}|{d}" for i, (no, t, d) in enumerate(hb.SECTIONS)) +
              "\n\nMINUTES:\n" + text)
    content = groq_chat({"model": GROQ_LLM, "temperature": 0, "reasoning_effort": "low", "max_completion_tokens": 600,
                         "response_format": {"type": "json_object"},
                         "messages": [{"role": "user", "content": prompt}]}, timeout=120)
    idx = json.loads(content).get("indices", [])
    return sorted({i for i in idx if isinstance(i, int) and 0 <= i < len(hb.SECTIONS)})


def lines(v):
    return "\n".join(str(x) for x in (v or []))


def to_list(s, sep="\n"):
    return [x.strip() for x in s.split(sep) if x.strip()]


# ---------------------------------------------------------------- UI
st.set_page_config(page_title="MOM Generator", page_icon="📝", layout="centered")

ROOT = Path(__file__).parent
STYLE = (ROOT / "style.css").read_text(encoding="utf-8")
DARK = (ROOT / "dark.css").read_text(encoding="utf-8")
st.markdown(f"<style>{STYLE}</style>", unsafe_allow_html=True)


def stepper(active):
    labels = ["Upload", "Generate", "Review & edit", "Download"]
    parts = []
    for n, lab in enumerate(labels, 1):
        cls = "done" if n < active else ("active" if n == active else "")
        parts.append(f'<div class="st {cls}"><i>{"✓" if n < active else n}</i><span>{lab}</span></div>')
    return '<div class="stepper">' + "".join(parts) + "</div>"


def card_head(num, title, sub=""):
    st.markdown(f'<div class="card-h"><div class="num">{num}</div><div><b>{title}</b>'
                f'{f"<span class=sub>{sub}</span>" if sub else ""}</div></div>', unsafe_allow_html=True)


def sec_head(n, label):
    st.markdown(f'<div class="docsec"><span>{n}</span>{label}</div>', unsafe_allow_html=True)


def stat_cards(m):
    words = len(str(m.get("transcript") or "").split())
    third = ((len(m.get("discussion") or []), "Discussion points") if m.get("kind") == "hearing"
             else (len(m.get("action_items") or []), "Action items"))
    cards = [(count_attendees(m.get("attendees")), "Attendees"), (len(m.get("decisions") or []), "Decisions"),
             third, (f"{words:,}" if words else "–", "Words")]
    return '<div class="stats">' + "".join(f'<div class="stat"><b>{v}</b><span>{k}</span></div>' for v, k in cards) + "</div>"


# Theme: Auto follows the device/browser setting; Light/Dark force it.
if "theme_mode" not in st.session_state:
    st.session_state["theme_mode"] = {"dark": "Dark", "light": "Light"}.get(st.query_params.get("theme"), "Auto")
top_l, top_r = st.columns([3, 2], vertical_alignment="center")
top_l.markdown('<div class="brand"><span class="logo">M</span>MOM Generator</div>', unsafe_allow_html=True)
with top_r:
    mode = st.segmented_control("Theme", ["Auto", "Light", "Dark"], key="theme_mode", label_visibility="collapsed") or "Auto"
if mode == "Dark":
    st.markdown(f"<style>{DARK}</style>", unsafe_allow_html=True)
elif mode == "Auto":
    st.markdown(f"<style>@media (prefers-color-scheme: dark) {{{DARK}}}</style>", unsafe_allow_html=True)

st.markdown("""
<div class="intro">
  <h1>Turn meeting recordings into <span class="grad">clean minutes.</span></h1>
  <p>Upload a recording, check the minutes, and download a Word file in the company template.</p>
</div>""", unsafe_allow_html=True)
progress = st.empty()

mom = st.session_state.get("mom")

# ------------------------------------------------------------ 1-2: upload + details
with st.expander("Recording & details", expanded=not mom) if mom else st.container():
    with st.container(key="card-input"):
        card_head("1", "Upload recording", "MP3, WAV, M4A or MP4")
        f = st.file_uploader("Upload recording", type=["mp3", "wav", "m4a", "mp4"], label_visibility="collapsed")

        card_head("2", "Meeting details", "Optional. Helps the AI get names right and fills the Word file.")
        kind_label = st.radio("Meeting type", ["Admin hearing (HR)", "General meeting"], horizontal=True, key="kind_label")
        kind = "hearing" if kind_label.startswith("Admin") else "general"
        c1, c2 = st.columns(2)
        lang = c1.selectbox("Write the minutes in", ["English", "Tagalog", "Taglish"])
        hint = c2.text_input("Names and terms", placeholder="e.g. Maria Santos, Mabuhay App")
        c3, c4 = st.columns(2)
        prep_name = c3.text_input("Prepared by", placeholder="Your name", key="prep_name")
        prep_role = c4.text_input("Your position", placeholder="e.g. HR Associate", key="prep_role")
        with st.container(key="sigbox"):
            sig_file = st.file_uploader("E-signature (optional)", type=["png", "jpg", "jpeg"], key="sig_file",
                                        help="A clear photo or scan of your signature on white paper, or a PNG with a transparent "
                                             "background. It goes above your name under 'Prepared By' in the Word file. "
                                             "It is only kept while this page is open.")
        sig_bytes = sig_file.getvalue() if sig_file else None
        if sig_bytes:
            st.image(sig_bytes, width=150, caption="Signature preview")

        go = st.button("Generate minutes", type="primary", width="stretch", disabled=not f)
        status_slot = st.container()

        if go and f:
            suffix = os.path.splitext(f.name)[1]
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp.write(f.getbuffer())   # no second in-memory copy: big videos can be hundreds of MB
                path = tmp.name
            with status_slot:
                with st.status("Generating your minutes...", expanded=True) as status:
                    try:
                        st.session_state["mom"] = make_mom(path, lang, hint, kind)
                        st.session_state["gen"] = st.session_state.get("gen", 0) + 1
                        status.update(label="Done! Review and edit below.", state="complete", expanded=False)
                    except Exception as e:
                        status.update(label="Something went wrong", state="error")
                        st.error(f"{type(e).__name__}: {e}")

mom = st.session_state.get("mom")
progress.markdown(stepper(3 if mom else (2 if f else 1)), unsafe_allow_html=True)

# ------------------------------------------------------------ 3: review & edit (same structure as the Word file)
if mom:
    g = st.session_state.get("gen", 0)
    hearing = mom.get("kind") == "hearing"
    with st.container(key="card-review"):
        card_head("3", "Review & edit", "This follows the same sections as the Word file.")
        st.markdown(stat_cards(mom), unsafe_allow_html=True)

        if hearing:
            e1, e2 = st.columns(2)
            employee = e1.text_input("Employee name", mom.get("employee") or "", key=f"emp{g}")
            position = e2.text_input("Position", mom.get("position") or "", key=f"pos{g}")
            title = ""
        else:
            title = st.text_input("Meeting title", mom.get("title") or "", key=f"title{g}")
            employee = position = ""
        d1, d2 = st.columns(2)
        date = d1.text_input("Date", mom.get("date") or "TBD", key=f"date{g}")
        time_ = d2.text_input("Time", mom.get("time") or "TBD", key=f"time{g}")
        loc = st.text_input("Location", mom.get("location") or "TBD", key=f"loc{g}")
        purpose = st.text_input("Purpose", mom.get("purpose") or "TBD", key=f"purpose{g}")

        st.markdown('<div class="docsec"><span>•</span>Attendees</div>'
                    '<div class="hint">A line without a dash is a group (like Management). Start each person with a dash.</div>',
                    unsafe_allow_html=True)
        attendees = st.text_area("Attendees", attendees_to_text(mom.get("attendees")), height=170, key=f"att{g}", label_visibility="collapsed")

        sec_head("1", "Call to Order")
        call = st.text_area("Call to Order", mom.get("call_to_order") or "", height=80, key=f"call{g}", label_visibility="collapsed")

        sec_head("2", "Purpose of the Meeting")
        summary = st.text_area("Purpose of the Meeting", mom.get("summary") or "", height=130, key=f"summary{g}", label_visibility="collapsed")

        sec_head("3", "Discussion Summary")
        st.markdown('<div class="hint">Separate paragraphs with a blank line.</div>', unsafe_allow_html=True)
        discussion = st.text_area("Discussion Summary", "\n\n".join(mom.get("discussion") or []),
                                  height=340, key=f"disc{g}", label_visibility="collapsed")

        review_intro = offense = handbook_text = notes = ""
        action_rows = []
        if hearing:
            sec_head("4", "Management Review and Decision")
            review_intro = st.text_area("Management review", mom.get("review_intro") or "", height=110,
                                        key=f"rev{g}", label_visibility="collapsed")
            st.markdown('<div class="hint">Decision (one per line)</div>', unsafe_allow_html=True)
            decisions = st.text_area("Decision", lines(mom.get("decisions")), height=90, key=f"dec{g}", label_visibility="collapsed")

            sec_head("5", "Handbook reference")
            st.markdown('<div class="hint">Pick the offense sections from the Company Employee Handbook (the Sanction table). '
                        'Leave it empty to leave this section out of the Word file.</div>', unsafe_allow_html=True)
            sel_key = f"hbsel{g}"
            if st.button("Suggest sections with AI", key=f"hbsug{g}", disabled=not GROQ_KEY):
                with st.spinner("Checking the handbook..."):
                    try:
                        found = suggest_sections({"summary": summary, "discussion": to_list(discussion, "\n\n"),
                                                  "decisions": to_list(decisions)})
                        st.session_state[sel_key] = found
                        if not found:
                            st.info("No handbook section clearly matches this case. Please pick one yourself below.")
                    except Exception as e:
                        st.warning(f"Could not get suggestions ({type(e).__name__}). Please pick the sections yourself below.")
            picked = st.multiselect("Handbook sections", options=list(range(len(hb.SECTIONS))), format_func=hb.label,
                                    key=sel_key, placeholder="Search by number or words, e.g. 4.1 or negligence")
            extra = st.text_area("Other lines to cite (optional)", mom.get("handbook") or "", height=80, key=f"hbx{g}",
                                 placeholder="One per line")
            offense = hb.heading(picked)
            handbook_text = "\n".join([hb.citation(i) for i in picked] + [x.strip() for x in extra.splitlines() if x.strip()])
            if offense:
                st.caption(f"Heading in the Word file: 5. {offense}")
        else:
            sec_head("4", "Decisions")
            st.markdown('<div class="hint">One per line.</div>', unsafe_allow_html=True)
            decisions = st.text_area("Decisions", lines(mom.get("decisions")), height=120, key=f"dec{g}", label_visibility="collapsed")

            sec_head("5", "Action Items")
            st.markdown('<div class="hint">Click a cell to edit. Use the + row to add, or select a row and press Delete.</div>', unsafe_allow_html=True)
            actions = st.data_editor(
                [{"task": str(a.get("task", "")), "owner": str(a.get("owner", "TBD")), "deadline": str(a.get("deadline", "TBD"))}
                 for a in (mom.get("action_items") or [])] or [{"task": "", "owner": "TBD", "deadline": "TBD"}],
                num_rows="dynamic", width="stretch", key=f"act{g}",
                column_config={"task": "Task", "owner": "Owner", "deadline": "Deadline"})
            action_rows = [r for r in (actions.to_dict("records") if hasattr(actions, "to_dict") else actions)
                           if str(r.get("task") or "").strip()]

            sec_head("6", "Additional Notes")
            notes = st.text_area("Additional Notes", mom.get("notes") or "", height=100, key=f"notes{g}", label_visibility="collapsed")

        with st.expander("Conversation transcript"):
            st.markdown('<div class="hint">Speakers are guessed by AI from context. You can rename or fix any label. One turn per line.</div>', unsafe_allow_html=True)
            transcript = st.text_area("Transcript", mom.get("transcript") or "(No transcript available)",
                                      height=360, key=f"tr{g}", label_visibility="collapsed")

    edited = {
        "kind": mom.get("kind", "general"),
        "title": f"{employee.strip()} ({position.strip()})" if hearing and employee.strip() and position.strip()
                 else (employee.strip() if hearing and employee.strip() else (title or "Meeting")),
        "employee": employee, "position": position,
        "date": date, "time": time_, "location": loc, "purpose": purpose,
        "attendees": text_to_attendees(attendees), "call_to_order": call, "summary": summary,
        "discussion": to_list(discussion, "\n\n"), "decisions": to_list(decisions),
        "review_intro": review_intro, "offense": offense.strip(), "handbook": handbook_text,
        "action_items": action_rows, "notes": notes,
        "prepared_by": {"name": prep_name.strip(), "role": prep_role.strip()},
        "transcript": transcript if mom.get("transcript") else "",
    }

    # ------------------------------------------------------------ 4: download
    with st.container(key="card-download"):
        card_head("4", "Preview & download", "Check the document first, then download the Word file.")
        with_tr = st.checkbox("Attach the conversation transcript at the end of the Word file", value=bool(edited["transcript"]),
                              disabled=not edited["transcript"])
        use_sig = st.checkbox("Place my e-signature above my name under 'Prepared By'", value=True, disabled=not sig_bytes,
                              help=None if sig_bytes else "Upload your e-signature in step 2 to enable this.")
        edited["signature"] = sig_bytes if (sig_bytes and use_sig) else None
        st.markdown('<div class="hint">Preview. This is how your Word file will look. Spacing and fonts can differ a little in Word.</div>',
                    unsafe_allow_html=True)
        st.markdown(preview_html(edited, include_transcript=with_tr), unsafe_allow_html=True)
        st.download_button("Download Word file (.docx)", build_docx(edited, include_transcript=with_tr),
                           file_name="Minutes of the Meeting.docx", type="primary", width="stretch",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        with st.expander("Follow-up email draft"):
            st.markdown('<div class="hint">A ready-to-send recap built from your edits. Use the copy icon at the top right of the box.</div>', unsafe_allow_html=True)
            st.code(build_email(edited), language=None, wrap_lines=True)

st.markdown('<div class="note">AI-generated draft. Please review before sharing. '
            'Recordings are processed by Groq and Google Gemini.</div>', unsafe_allow_html=True)
