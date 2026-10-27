import os, json, time, tempfile, subprocess
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
    mom["transcript"] = text.strip()
    return mom

def build_docx(mom, out="MOM.docx"):
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
            if para.strip():
                d.add_paragraph(para.strip())
    d.save(out)
    return out

st.title("MOM Generator (Libre)")
f = st.file_uploader("Upload recording", type=["mp3", "wav", "m4a", "mp4"])

if f and st.button("Generate MOM"):
    suffix = os.path.splitext(f.name)[1]
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(f.read())
        path = tmp.name
    with st.spinner("Ginagawa ang MOM..."):
        mom = make_mom(path)
        out = build_docx(mom)
    with st.expander("Preview (JSON)"):
        st.json(mom)
    with open(out, "rb") as fh:
        st.download_button("Download MOM (.docx)", fh, file_name="MOM.docx")
    st.success("Tapos na! I-review muna bago i-send.")