import os, io, re, json, time, tempfile, subprocess
import requests
import streamlit as st
from dotenv import load_dotenv
from google import genai
from google.genai import errors
from docx import Document

load_dotenv()
client = genai.Client(api_key=os.getenv("GEMINI_API_KEY"))
FALLBACK_MODEL = "gemini-3.8-flash-lite"   # assumed name; change if it 404s
MODEL = "gemini-3.8-flash"   # palitan ng kasalukuyang Flash model sa AI Studio kung kailangan

PROMPT = """You are a meeting secretary. This is a meeting recording in Tagalog/Taglish.
Write the Minutes of Meeting in English. Output ONLY valid JSON in this shape:
{"title": "", "date": "", "summary": "", "attendees": [], "agenda": [], "discussion": [],
 "decisions": [], "action_items": [{"task": "", "owner": "", "deadline": ""}]}
"summary" is a 3-5 sentence overview. "discussion" must be DETAILED: one entry per topic,
each a full paragraph covering what was raised, who said what (if names are stated),
the reasons, concerns, numbers/dates mentioned, and how it was resolved.
Rules: Do not invent anything. Only use what is said in the recording.
If something is not stated, use "TBD". Keep names exactly as spoken."""

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

def make_mom_gemini_audio(path):
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
                contents=[f, PROMPT],
                config={"response_mime_type": "application/json"},
            )
            break
        except errors.ServerError:   # 503 overloaded: wait and retry
            time.sleep(10)
    client.files.delete(name=f.name)   # burahin ang audio sa server pagkatapos
    if resp is None:
        raise RuntimeError("Busy pa rin ang Gemini. Subukan ulit mamaya.")
    return json.loads(resp.text)

GROQ_KEY = os.getenv("GROQ_API_KEY")
GROQ_URL = "https://api.groq.com/openai/v1"
GROQ_WHISPER = "whisper-large-v3-turbo"
GROQ_LLM = "openai/gpt-oss-120b"   # change if Groq retires it

def groq_transcribe(path):
    with open(path, "rb") as fh:
        r = requests.post(
            f"{GROQ_URL}/audio/transcriptions",
            headers={"Authorization": f"Bearer {GROQ_KEY}"},
            files={"file": (os.path.basename(path), fh)},
            data={"model": GROQ_WHISPER, "response_format": "text"},
            timeout=600,
        )
    r.raise_for_status()
    return r.text

def mom_from_text(transcript):
    """Text -> MOM JSON. Groq Llama first (fast), Gemini only as fallback."""
    prompt = PROMPT.replace("This is a meeting recording", "Below is a transcript of a meeting recording")
    prompt += "\n\nTRANSCRIPT:\n" + transcript
    try:
        r = requests.post(
            f"{GROQ_URL}/chat/completions",
            headers={"Authorization": f"Bearer {GROQ_KEY}"},
            json={"model": GROQ_LLM, "messages": [{"role": "user", "content": prompt}],
                  "response_format": {"type": "json_object"}},
            timeout=120,
        )
        r.raise_for_status()
        return json.loads(r.json()["choices"][0]["message"]["content"])
    except Exception:
        resp = client.models.generate_content(
            model=MODEL, contents=prompt,
            config={"response_mime_type": "application/json"})
        return json.loads(resp.text)

CONVO_PROMPT = """Reformat this raw meeting transcript as a conversation, one speaker turn per line,
like "Speaker 1: ...". Guess where the speaker changes from context (questions and answers,
greetings, replies). Keep the same label for the same person throughout. Use a real name as the
label only if that person is clearly addressed or introduces themselves.
STRICT RULES: do not translate, summarize, add, remove or change any words; keep the original
language (Tagalog/Taglish/English). Output plain text only, no commentary, no markdown."""

def chunk_text(text, size=7000):
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
            r = requests.post(
                f"{GROQ_URL}/chat/completions",
                headers={"Authorization": f"Bearer {GROQ_KEY}"},
                json={"model": GROQ_LLM, "temperature": 0,
                      "messages": [{"role": "user", "content": CONVO_PROMPT + ctx + "\n\nTRANSCRIPT:\n" + ch}]},
                timeout=180,
            )
            r.raise_for_status()
            part = r.json()["choices"][0]["message"]["content"].strip()
            out.append(part)
            prev = "\n".join(part.splitlines()[-3:])
        return "\n".join(out)
    except Exception:
        return text

def make_mom(path):
    if not GROQ_KEY:
        return make_mom_gemini_audio(path)   # no Groq key: old Gemini-only path
    t = time.time()
    small = compress(path)
    st.write(f"1. Compress: {time.time()-t:.0f}s ({os.path.getsize(path)/1e6:.1f}MB -> {os.path.getsize(small)/1e6:.1f}MB)")
    t = time.time()
    text = groq_transcribe(small)
    st.write(f"2. Groq transcript: {time.time()-t:.0f}s")
    t = time.time()
    mom = mom_from_text(text)
    st.write(f"3. MOM from text: {time.time()-t:.0f}s")
    t = time.time()
    mom["transcript"] = to_conversation(text.strip())
    st.write(f"4. Conversation format: {time.time()-t:.0f}s")
    return mom

def build_docx(mom):
    d = Document()
    d.add_heading(mom.get("title") or "Minutes of Meeting", 0)
    d.add_paragraph(f"Date: {mom.get('date', 'TBD')}")
    if mom.get("summary"):
        d.add_heading("Summary", 1)
        d.add_paragraph(str(mom["summary"]))
    for section in ["attendees", "agenda", "discussion", "decisions"]:
        d.add_heading(section.capitalize(), 1)
        for item in mom.get(section, []):
            d.add_paragraph(str(item), style="List Bullet")
    d.add_heading("Action Items", 1)
    t = d.add_table(rows=1, cols=3)
    t.style = "Table Grid"
    for i, h in enumerate(["Task", "Owner", "Deadline"]):
        t.rows[0].cells[i].text = h
    for ai in mom.get("action_items", []):
        r = t.add_row().cells
        r[0].text = str(ai.get("task", ""))
        r[1].text = str(ai.get("owner", "TBD"))
        r[2].text = str(ai.get("deadline", "TBD"))
    if mom.get("transcript"):
        d.add_page_break()
        d.add_heading("Full Transcript (original language)", 1)
        for para in str(mom["transcript"]).split("\n"):
            if not para.strip():
                continue
            m = re.match(r"^([^:\n]{1,30}):\s*(.*)$", para.strip())
            p = d.add_paragraph()
            if m:
                p.add_run(m.group(1) + ": ").bold = True
                p.add_run(m.group(2))
            else:
                p.add_run(para.strip())
    buf = io.BytesIO()
    d.save(buf)
    return buf.getvalue()

# ---------------------------------------------------------------- UI
st.set_page_config(page_title="MOM Generator", page_icon="📝", layout="centered")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
:root { --ink:#111827; --muted:#6B7280; --line:#E5E7EB; --soft:#F3F4F6; --bg:#FAFAFA; --surface:#fff; --label:#374151; --btnbg:#111827; --btnfg:#fff; --accent:#4F46E5; }
html, body, [class*="css"], .stApp { font-family:'Inter', -apple-system, sans-serif; color:var(--ink); -webkit-font-smoothing:antialiased; }
.stApp { background:var(--bg); }
#MainMenu, footer, header[data-testid="stHeader"] { visibility:hidden; height:0; }
.block-container { max-width:780px; padding-top:3.2rem; padding-bottom:5rem; }

/* hero */
.eyebrow { font-size:.74rem; font-weight:600; letter-spacing:.14em; text-transform:uppercase; color:var(--accent); }
.hero h1 { font-size:clamp(2.1rem,5.5vw,2.9rem); line-height:1.1; font-weight:650; letter-spacing:-.03em; margin:.7rem 0 .7rem; color:var(--ink); }
.hero p { font-size:1.05rem; line-height:1.55; color:var(--muted); max-width:32rem; margin:0; }
.steps { display:flex; gap:1.6rem; margin:1.8rem 0 1.6rem; padding-bottom:1.4rem; border-bottom:1px solid var(--line); flex-wrap:wrap; }
.steps span { font-size:.84rem; color:var(--muted); }
.steps span b { color:var(--ink); font-weight:600; margin-right:.45rem; }

/* uploader */
[data-testid="stFileUploader"] section { background:var(--surface); border:1px dashed #D1D5DB; border-radius:14px; padding:1.8rem; transition:border-color .15s, box-shadow .15s; }
[data-testid="stFileUploader"] section:hover { border-color:var(--ink); box-shadow:0 1px 2px rgba(17,24,39,.05); }
[data-testid="stFileUploader"] button { border-radius:8px; border:1px solid var(--line); background:var(--surface); color:var(--ink); font-weight:500; }

/* buttons */
div.stButton > button, div.stDownloadButton > button { font-size:.92rem; font-weight:550; border-radius:10px; padding:.7rem 1.3rem;
  border:1px solid var(--line); background:var(--surface); color:var(--ink); box-shadow:0 1px 2px rgba(17,24,39,.04); transition:all .15s; }
div.stButton > button:hover, div.stDownloadButton > button:hover { border-color:#9CA3AF; color:var(--ink); }
div.stButton > button[kind="primary"], div.stDownloadButton > button[kind="primary"] { background:var(--btnbg); color:var(--btnfg); border:1px solid var(--btnbg); }
div.stButton > button[kind="primary"]:hover, div.stDownloadButton > button[kind="primary"]:hover { background:var(--btnbg); color:var(--btnfg); opacity:.88; box-shadow:0 4px 14px rgba(0,0,0,.18); }

/* tabs: underline */
.stTabs [data-baseweb="tab-list"] { gap:1.6rem; border-bottom:1px solid var(--line); background:transparent; padding:0; }
.stTabs [data-baseweb="tab"] { padding:.7rem 0; font-size:.92rem; font-weight:500; color:var(--muted); background:transparent; height:auto; }
.stTabs [aria-selected="true"] { color:var(--ink); }
.stTabs [data-baseweb="tab-highlight"] { background:var(--ink); height:2px; }
.stTabs [data-baseweb="tab-border"] { display:none; }
.stTabs [data-baseweb="tab-panel"] { padding-top:1.6rem; }

/* inputs */
.stTextInput input, .stTextArea textarea { background:var(--surface) !important; border:1px solid var(--line) !important; border-radius:10px !important; color:var(--ink) !important; }
.stTextInput input:focus, .stTextArea textarea:focus { border-color:var(--accent) !important; box-shadow:0 0 0 3px rgba(79,70,229,.12) !important; }
label p, .stTextInput label, .stTextArea label { font-size:.84rem !important; font-weight:550 !important; color:var(--label) !important; }
[data-testid="stAlert"] { background:var(--surface); border:1px solid var(--line); border-radius:12px; }
[data-testid="stExpander"], [data-testid="stStatusWidget"], [data-testid="stStatus"] { background:var(--surface); border:1px solid var(--line); border-radius:12px; }
[data-testid="stDataFrame"], [data-testid="stDataEditor"] { border:1px solid var(--line); border-radius:12px; overflow:hidden; }

.card-title { font-size:.9rem; font-weight:600; color:var(--ink); margin:1.4rem 0 .1rem; }
.hint { color:var(--muted); font-size:.82rem; margin-bottom:.5rem; }
</style>
<div class="hero">
  <div class="eyebrow">MOM Generator</div>
  <h1>Meeting minutes, without the manual work.</h1>
  <p>Upload a recording and get structured minutes with the full conversation transcript. Edit anything, then download.</p>
</div>
<div class="steps"><span><b>1</b>Upload</span><span><b>2</b>Generate</span><span><b>3</b>Review &amp; edit</span><span><b>4</b>Download</span></div>
""", unsafe_allow_html=True)

DARK_CSS = """
<style>
:root { --ink:#F4F4F5; --muted:#A1A1AA; --line:#27272F; --soft:#1C1C22; --bg:#0B0B0F; --surface:#14141A;
        --label:#D4D4D8; --accent:#818CF8; --btnbg:#F4F4F5; --btnfg:#0B0B0F; }
.stApp, .stApp p, .stApp li, .stApp label, .stApp h1, .stApp h2, .stApp h3 { color:var(--ink); }
.stApp { background:var(--bg); }
[data-testid="stFileUploader"] section { border-color:#3F3F46; }
[data-testid="stFileUploader"] small { color:var(--muted); }
[data-testid="stDataEditor"] { filter:invert(1) hue-rotate(180deg); }
.stTextInput input, .stTextArea textarea { color:var(--ink) !important; }
</style>
"""

if "dark" not in st.session_state:
    st.session_state["dark"] = st.query_params.get("theme") == "dark"
_, tcol = st.columns([6, 2])
with tcol:
    dark = st.toggle("Dark mode", key="dark")
if dark:
    st.markdown(DARK_CSS, unsafe_allow_html=True)

def lines(v):
    return "\n".join(str(x) for x in (v or []))

def to_list(s, sep="\n"):
    return [x.strip() for x in s.split(sep) if x.strip()]

f = st.file_uploader("Upload recording", type=["mp3", "wav", "m4a", "mp4"], label_visibility="collapsed")

if f and st.button("Generate minutes", type="primary", width="stretch"):
    suffix = os.path.splitext(f.name)[1]
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(f.read())
        path = tmp.name
    with st.status("Generating your minutes...", expanded=True) as status:
        try:
            st.session_state["mom"] = make_mom(path)
            st.session_state["gen"] = st.session_state.get("gen", 0) + 1
            status.update(label="Done! Review and edit below.", state="complete", expanded=False)
        except Exception as e:
            status.update(label="Something went wrong", state="error")
            st.error(f"{type(e).__name__}: {e}")

mom = st.session_state.get("mom")
if mom:
    g = st.session_state.get("gen", 0)
    st.info("Review and edit anything below. Your changes go into the downloaded file.", icon="✏️")

    t1, t2, t3, t4 = st.tabs(["Overview", "Discussion", "Decisions & Actions", "Transcript"])

    with t1:
        title = st.text_input("Meeting title", mom.get("title") or "", key=f"title{g}")
        date = st.text_input("Date", mom.get("date") or "TBD", key=f"date{g}")
        summary = st.text_area("Summary", mom.get("summary") or "", height=140, key=f"summary{g}")
        st.markdown('<div class="card-title">Attendees</div><div class="hint">One per line</div>', unsafe_allow_html=True)
        attendees = st.text_area("Attendees", lines(mom.get("attendees")), height=120, key=f"att{g}", label_visibility="collapsed")
        st.markdown('<div class="card-title">Agenda</div><div class="hint">One per line</div>', unsafe_allow_html=True)
        agenda = st.text_area("Agenda", lines(mom.get("agenda")), height=120, key=f"agenda{g}", label_visibility="collapsed")

    with t2:
        st.markdown('<div class="hint">Separate topics with a blank line.</div>', unsafe_allow_html=True)
        discussion = st.text_area("Discussion", "\n\n".join(str(x) for x in (mom.get("discussion") or [])),
                                  height=420, key=f"disc{g}", label_visibility="collapsed")

    with t3:
        st.markdown('<div class="card-title">Decisions</div><div class="hint">One per line</div>', unsafe_allow_html=True)
        decisions = st.text_area("Decisions", lines(mom.get("decisions")), height=140, key=f"dec{g}", label_visibility="collapsed")
        st.markdown('<div class="card-title">Action items</div><div class="hint">Click a cell to edit. Use the + row to add, or select a row and press Delete.</div>', unsafe_allow_html=True)
        actions = st.data_editor(
            [{"task": str(a.get("task", "")), "owner": str(a.get("owner", "TBD")), "deadline": str(a.get("deadline", "TBD"))}
             for a in (mom.get("action_items") or [])] or [{"task": "", "owner": "TBD", "deadline": "TBD"}],
            num_rows="dynamic", width="stretch", key=f"act{g}",
            column_config={"task": "Task", "owner": "Owner", "deadline": "Deadline"})

    with t4:
        st.markdown('<div class="hint">Speakers are guessed by AI from context. You can rename or fix any label. One turn per line.</div>', unsafe_allow_html=True)
        transcript = st.text_area("Transcript", mom.get("transcript") or "(No transcript available)",
                                  height=420, key=f"tr{g}", label_visibility="collapsed")

    edited = {
        "title": title, "date": date, "summary": summary,
        "attendees": to_list(attendees), "agenda": to_list(agenda),
        "discussion": to_list(discussion, "\n\n"), "decisions": to_list(decisions),
        "action_items": [r for r in (actions.to_dict("records") if hasattr(actions, "to_dict") else actions)
                         if str(r.get("task") or "").strip()],
        "transcript": transcript if mom.get("transcript") else "",
    }
    st.download_button("Download .docx", build_docx(edited), file_name="MOM.docx",
                       type="primary", width="stretch",
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")