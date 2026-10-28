"""On-screen preview of the Word file, built from the same data. Plain HTML, so it needs no Word or LibreOffice."""
import base64
import html
from functools import lru_cache

from docx_builder import COMPANY_LINES, COMPANY_NAME, LOGO, _trim_signature, signature_size

PAGE_TEXT_IN = 6.5   # text width of the Legal page in the Word file (8.5in minus 1in margins)


def _e(x):
    return html.escape(str(x if x is not None else ""))


@lru_cache(maxsize=1)
def _logo_uri():
    return "data:image/png;base64," + base64.b64encode(LOGO.read_bytes()).decode() if LOGO.exists() else ""


def _paras(items, cls="pp-p"):
    return "".join(f'<div class="{cls}">{_e(x)}</div>' for x in items if str(x).strip())


def _rule(thick=False):
    return f'<div class="pp-rule{" thick" if thick else ""}"></div>'


def _head(text):
    return f'<div class="pp-h">{_e(text)}</div>'


def preview_html(m, include_transcript=False):
    out = ['<div class="paper">']

    # letterhead
    logo = f'<img class="pp-logo" src="{_logo_uri()}">' if _logo_uri() else ""
    lines = "".join(f"<div>{_e(x)}</div>" for x in COMPANY_LINES)
    out.append(f'<div class="pp-top"><div class="pp-co"><b>{_e(COMPANY_NAME)}</b>{lines}</div>{logo}</div>')

    # title + details
    out.append('<div class="pp-title">Minutes of the Meeting</div>')
    out.append(f'<div class="pp-title pp-sub">{_e(m.get("title") or "Meeting")}</div>')
    for label, key in (("Date", "date"), ("Time", "time"), ("Location", "location"), ("Purpose", "purpose")):
        out.append(f'<div class="pp-line"><b>{label}</b>: {_e(m.get(key) or "TBD")}</div>')
    out.append('<div class="pp-line"><b>Attendees:</b></div>')
    for g in m.get("attendees") or []:
        name = (g.get("group") or "").strip()
        members = g.get("members") or []
        if name:
            out.append(f'<div class="pp-b1">&#9642;&nbsp; {_e(name)}</div>')
            out += [f'<div class="pp-b2">&bull;&nbsp; {_e(x)}</div>' for x in members]
        else:
            out += [f'<div class="pp-b1">&#9642;&nbsp; {_e(x)}</div>' for x in members]
    out.append(_rule())

    # numbered sections
    out.append(_head("1. Call to Order") + _paras([m.get("call_to_order") or "The meeting was called to order."]))
    out.append(_rule())
    out.append(_head("2. Purpose of the Meeting") + _paras([m.get("summary") or "TBD"]))
    out.append(_rule())
    out.append(_head("3. Discussion Summary") + _paras(m.get("discussion") or ["TBD"]))
    out.append(_rule(thick=True))

    if m.get("kind") == "hearing":
        out.append(_head("4. Management Review and Decision"))
        if (m.get("review_intro") or "").strip():
            out.append(_paras([m["review_intro"]]))
        out.append('<div class="pp-p b">Decision:</div>')
        out.append(_paras(m.get("decisions") or ["TBD"], "pp-p b"))
        offense = (m.get("offense") or "").strip()
        handbook = [x.strip() for x in (m.get("handbook") or "").splitlines() if x.strip()]
        if offense or handbook:
            out.append(_rule(thick=True))
            out.append(_head(f"5. {offense}" if offense else "5. Company Employee Handbook"))
            out.append(_paras(["Based on the Company Employee Handbook:"] + handbook))
    else:
        out.append(_head("4. Decisions") + '<div class="pp-p b">Decision:</div>')
        out.append(_paras(m.get("decisions") or ["No decision recorded."], "pp-p b"))
        n = 4
        if m.get("action_items"):
            n += 1
            rows = "".join(f'<tr><td>{_e(a.get("task"))}</td><td>{_e(a.get("owner") or "TBD")}</td>'
                           f'<td>{_e(a.get("deadline") or "TBD")}</td></tr>' for a in m["action_items"])
            out.append(_rule(thick=True) + _head(f"{n}. Action Items") +
                       f'<table class="pp-t"><tr><th>Task</th><th>Owner</th><th>Deadline</th></tr>{rows}</table>')
        if (m.get("notes") or "").strip():
            n += 1
            out.append(_rule(thick=True) + _head(f"{n}. Additional Notes") + _paras(m["notes"].split("\n")))

    # prepared by (+ e-signature above the name, as in the Word file)
    pb = m.get("prepared_by") or {}
    out.append('<div class="pp-p pp-prep">Prepared By:</div>')
    sig_img = ""
    if m.get("signature"):
        try:
            blob = _trim_signature(m["signature"])
            w_in, _ = signature_size(blob)
            uri = "data:image/png;base64," + base64.b64encode(blob).decode()
            sig_img = f'<img class="pp-sig" src="{uri}" style="width:{w_in / PAGE_TEXT_IN * 100:.1f}%">'
        except Exception:
            sig_img = ""
    out.append(sig_img or '<div class="pp-gap"></div>')
    out.append(f'<div class="pp-line pp-name">{_e(pb.get("name") or "______________________")}</div>')
    if pb.get("role"):
        out.append(f'<div class="pp-line pp-role">{_e(pb["role"])}</div>')

    # optional transcript appendix
    if include_transcript and m.get("transcript"):
        rows = [x for x in str(m["transcript"]).split("\n") if x.strip()]
        shown = rows[:14]
        out.append(_rule(thick=True))
        out.append('<div class="pp-title">Appendix: Conversation Transcript</div>')
        out.append(_paras(shown, "pp-p tr"))
        if len(rows) > len(shown):
            out.append(f'<div class="pp-more">... and {len(rows) - len(shown)} more lines in the Word file</div>')

    out.append("</div>")
    return "".join(out)
