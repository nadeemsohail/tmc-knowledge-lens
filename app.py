"""
TMC Knowledge Lens - search-first, permission-aware document workspace.

Flow:  Search the library  ->  matching documents appear in Sources (left)
       ->  tick the ones you want  ->  summarise / ask cited questions (centre)
       ->  generate a report as Word, PowerPoint, Excel or PDF (right).
"""
import streamlit as st
from openai import OpenAI
import pypdf, docx, openpyxl
from pptx import Presentation
from pptx.util import Pt, Inches
from openpyxl.styles import Font, PatternFill, Alignment
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from xml.sax.saxutils import escape as xml_escape
from collections import Counter
from datetime import datetime
from pathlib import Path
import json, io, re, math, hashlib, time, difflib, os, uuid, html

st.set_page_config(page_title="TMC Knowledge Lens", page_icon="🔍", layout="wide")

# ----------------------------------------------------------------------------
# Constants and storage
# ----------------------------------------------------------------------------
MODEL = st.secrets.get("MODEL", "openai/gpt-oss-120b")
DATA = Path(os.environ.get("KL_DATA_DIR", "kl_data"))   # NOTE: use a persistent volume / DB in production
(DATA / "files").mkdir(parents=True, exist_ok=True)
GROUPS = ["Pre-Sales", "Delivery", "Management"]
PANEL_H = 720

DEFAULT_TAXONOMY = {
    "clients": {"Tata Motors": ["tml", "tata motors ltd", "tata motors limited"]},
    "industries": ["Automotive", "Manufacturing", "Retail", "Oil & Gas", "Pharma", "Banking", "Public Sector", "Other"],
    "sap_modules": ["FI/CO", "MM", "SD", "PP", "PM", "QM", "WM/EWM", "HCM", "BTP", "Basis", "Ariba", "SuccessFactors", "Other"],
    "phases": ["Pre-Sales", "Prepare", "Explore", "Realize", "Deploy", "Run", "Other"],
    "doc_types": ["Proposal", "Blueprint", "Methodology", "Cutover Plan", "Test Plan", "Fit-to-Standard",
                  "RICEFW List", "Status Report", "Training Material", "One-Page Summary", "Other"],
    "glossary": {
        "ricefw": "reports interfaces conversions enhancements forms workflows",
        "fit-to-standard": "fit gap standard process",
        "blueprint": "business blueprint design",
        "cutover": "cutover go-live migration plan",
    },
}

REPORT_TYPES = {
    "Executive summary": "A concise executive summary with key points, decisions and next steps.",
    "Compare documents": "Compare the selected documents: similarities, differences, and a comparison table.",
    "Risk & issue register": "Extract risks and issues. The table must have columns: Risk, Impact, Mitigation, Source.",
    "Decisions & deliverables": "List key decisions, deliverables, owners and dates. Include a table.",
    "Lessons learned": "Summarise lessons learned and reusable assets for future proposals.",
    "Custom (use focus box)": "",
}

# ----------------------------------------------------------------------------
# Styling - search-first workspace, deliberately not a NotebookLM copy
# ----------------------------------------------------------------------------
st.markdown("""
<style>
.stApp { background:#F5F6FA; font-family:'Inter','Segoe UI',sans-serif; }
#MainMenu, footer { visibility:hidden; }
div[data-testid="stVerticalBlockBorderWrapper"] { border-radius:18px; }
.stButton>button, .stDownloadButton>button { border-radius:999px; font-weight:600; }
.stButton>button[kind="primary"] { background:#5B2C6F; border:none; color:#fff; }
.stButton>button[kind="primary"]:hover { background:#7D3C98; }
.topbar { display:flex; align-items:center; gap:18px; background:#fff; padding:12px 26px;
          border-radius:18px; border:1px solid #E3E6EE; margin-bottom:14px; }
.topbar h1 { color:#5B2C6F; margin:0; font-size:1.55rem; }
.topbar p { color:#6B7280; margin:0; font-size:.9rem; }
.chip { display:inline-block; padding:1px 9px; border-radius:999px; background:#EEF0F6;
        color:#3B4252; font-size:.72rem; margin:2px 4px 2px 0; }
.chip.p { background:#F1E6F6; color:#5B2C6F; }
.ok  { background:#E6F4EA; color:#137333; }
.old { background:#FEF3E0; color:#B06000; }
.doc-title { font-weight:600; font-size:.92rem; line-height:1.25; }
.relbar { height:4px; background:#E8EAF0; border-radius:4px; margin:6px 0 2px; }
.relbar>div { height:4px; background:#5B2C6F; border-radius:4px; }
.snip { font-size:.78rem; color:#5F6368; border-left:3px solid #D7DCE8; padding-left:8px; margin-top:4px; }
.glance { background:#F4EEF8; border-left:5px solid #5B2C6F; padding:14px 18px; border-radius:12px; margin-bottom:10px; }
.glance h4 { margin:0 0 6px; color:#5B2C6F; }
.small { font-size:.8rem; color:#6B7280; }
</style>
""", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# LLM client
# ----------------------------------------------------------------------------
API_KEY = st.secrets.get("GROQ_API_KEY") or st.secrets.get("GITHUB_TOKEN")
if not API_KEY:
    st.error("Add GROQ_API_KEY to the app secrets (Settings → Secrets).")
    st.stop()
client = OpenAI(base_url="https://api.groq.com/openai/v1", api_key=API_KEY)


def llm(messages, json_mode=False, temperature=0.1):
    kw = {"response_format": {"type": "json_object"}} if json_mode else {}
    r = client.chat.completions.create(model=MODEL, messages=messages, temperature=temperature, **kw)
    return r.choices[0].message.content or ""


def parse_json(raw):
    try:
        return json.loads(raw)
    except Exception:
        m = re.search(r"\{.*\}", raw, re.S)
        return json.loads(m.group(0)) if m else {}


# ----------------------------------------------------------------------------
# Taxonomy, library persistence, audit
# ----------------------------------------------------------------------------
def load_taxonomy():
    p = DATA / "taxonomy.json"
    return json.loads(p.read_text("utf-8")) if p.exists() else DEFAULT_TAXONOMY


def save_taxonomy(t):
    (DATA / "taxonomy.json").write_text(json.dumps(t, indent=2, ensure_ascii=False), "utf-8")


def load_library():
    p = DATA / "library.json"
    return json.loads(p.read_text("utf-8")) if p.exists() else []


def save_library():
    (DATA / "library.json").write_text(json.dumps(st.session_state.library, ensure_ascii=False), "utf-8")


def audit(action, **detail):
    u = st.session_state.get("user", {"name": "anonymous"})
    row = {"ts": datetime.now().isoformat(timespec="seconds"), "user": u["name"], "action": action, **detail}
    with open(DATA / "audit.jsonl", "a", encoding="utf-8") as f:
        f.write(json.dumps(row, ensure_ascii=False) + "\n")


def read_audit():
    p = DATA / "audit.jsonl"
    if not p.exists():
        return []
    out = []
    for line in p.read_text("utf-8").splitlines():
        try:
            out.append(json.loads(line))
        except Exception:
            pass
    return out


# ----------------------------------------------------------------------------
# Text extraction (returns pages, author, unit-name)
# ----------------------------------------------------------------------------
def chunk_text(text, size=2500):
    pages, cur = [], ""
    for line in text.splitlines():
        if len(cur) + len(line) > size and cur:
            pages.append(cur)
            cur = ""
        cur += line + "\n"
    if cur.strip():
        pages.append(cur)
    return pages


def extract_pdf(b):
    r = pypdf.PdfReader(io.BytesIO(b))
    author = (r.metadata.author if r.metadata and r.metadata.author else "Unknown")
    return [(p.extract_text() or "") for p in r.pages], author, "page"


def extract_docx(b):
    d = docx.Document(io.BytesIO(b))
    lines = [p.text for p in d.paragraphs]
    for t in d.tables:
        for row in t.rows:
            lines.append(" | ".join(c.text for c in row.cells))
    return chunk_text("\n".join(lines)), d.core_properties.author or "Unknown", "section"


def extract_pptx(b):
    prs = Presentation(io.BytesIO(b))
    pages = []
    for s in prs.slides:
        pages.append("\n".join(sh.text for sh in s.shapes if getattr(sh, "has_text_frame", False) and sh.text))
    return pages, "Unknown", "slide"


def extract_xlsx(b):
    wb = openpyxl.load_workbook(io.BytesIO(b), data_only=True, read_only=True)
    lines = []
    for ws in wb.worksheets:
        lines.append(f"# Sheet: {ws.title}")
        for row in ws.iter_rows(values_only=True):
            if any(c is not None for c in row):
                lines.append(" | ".join("" if c is None else str(c) for c in row))
    return chunk_text("\n".join(lines)), "Unknown", "section"


def extract_txt(b):
    return chunk_text(b.decode("utf-8", errors="ignore")), "Unknown", "section"


EXTRACTORS = {"pdf": extract_pdf, "docx": extract_docx, "pptx": extract_pptx, "xlsx": extract_xlsx, "txt": extract_txt, "md": extract_txt}

# <<RETRIEVAL>>
# ----------------------------------------------------------------------------
# Retrieval (BM25 + controlled-vocabulary query expansion). Swap for embeddings
# + a vector DB (pgvector / Azure AI Search) when the library grows.
# ----------------------------------------------------------------------------
STOP = set("the a an and or of to in for on with is are was were be by as at from that this it its".split())


def tok(s):
    return [t for t in re.findall(r"[a-z0-9]+", s.lower()) if t not in STOP and len(t) > 1]


def expand_query(q, tax):
    qt, low = tok(q), q.lower()
    for canon, aliases in tax["clients"].items():
        names = [canon.lower()] + [a.lower() for a in aliases]
        if any(re.search(r"\b" + re.escape(n) + r"\b", low) for n in names):
            qt += tok(canon)
    for k, v in tax["glossary"].items():
        if k in low:
            qt += tok(v)
    return qt


def bm25(query_tokens, texts, k1=1.5, b=0.75):
    toks = [tok(t) for t in texts]
    n = len(toks)
    if not n:
        return []
    avg = sum(map(len, toks)) / n or 1
    df = Counter()
    for t in toks:
        df.update(set(t))
    scores = []
    for t in toks:
        tf, s = Counter(t), 0.0
        for q in set(query_tokens):
            if q in tf:
                idf = math.log(1 + (n - df[q] + 0.5) / (df[q] + 0.5))
                s += idf * tf[q] * (k1 + 1) / (tf[q] + k1 * (1 - b + b * len(t) / avg))
        scores.append(s)
    return scores


def meta_blob(d):
    m = d["meta"]
    return " ".join(str(m.get(k, "")) for k in ("doc_title", "client", "industry", "sap_module", "phase", "doc_type", "summary")) + " " + d["name"]


def make_snippet(text, qtokens, width=240):
    flat = re.sub(r"\s+", " ", text)
    low = flat.lower()
    pos = min([low.find(t) for t in set(qtokens) if low.find(t) >= 0] or [0])
    start = max(0, pos - 60)
    return ("…" if start else "") + flat[start:start + width] + "…"


def search_docs(query, docs, tax):
    qt = expand_query(query, tax)
    chunks = []
    for d in docs:
        chunks.append((d["id"], 0, meta_blob(d)))
        chunks += [(d["id"], i + 1, p) for i, p in enumerate(d["pages"]) if p.strip()]
    scores = bm25(qt, [c[2] for c in chunks])
    agg = {}
    for (did, page, text), s in zip(chunks, scores):
        if s <= 0:
            continue
        a = agg.setdefault(did, {"score": 0.0, "page": None, "ps": 0.0, "text": ""})
        a["score"] = max(a["score"], s)
        if page > 0 and s > a["ps"]:
            a.update(ps=s, page=page, text=text)
    top = max([a["score"] for a in agg.values()], default=1.0)
    items = [{"id": k, "rel": int(100 * a["score"] / top), "page": a["page"],
              "snippet": make_snippet(a["text"], qt) if a["text"] else ""} for k, a in agg.items()]
    return sorted(items, key=lambda x: -x["rel"])


def retrieve(query, docs, tax, k=8):
    chunks = [(d, i + 1, p) for d in docs for i, p in enumerate(d["pages"]) if p.strip()]
    if not chunks:
        return []
    sc = bm25(expand_query(query, tax), [c[2] for c in chunks])
    order = sorted(range(len(chunks)), key=lambda i: -sc[i])
    pick = [i for i in order[:k] if sc[i] > 0]
    have = {chunks[i][0]["id"] for i in pick}
    for d in docs:                      # make sure every selected document is represented
        if d["id"] not in have:
            cand = [i for i in order if chunks[i][0]["id"] == d["id"]]
            if cand:
                pick.append(cand[0])
    return [{"doc_id": chunks[i][0]["id"], "doc_title": chunks[i][0]["meta"]["doc_title"],
             "unit": chunks[i][0]["unit"], "page": chunks[i][1], "text": chunks[i][2]} for i in pick]
# <<END RETRIEVAL>>


def sample_text(d, budget):
    pages, n = d["pages"], len(d["pages"])
    if sum(map(len, pages)) <= budget or n == 1:
        idx = list(range(n))
    else:
        m = min(n, 12)
        idx = sorted({round(i * (n - 1) / (m - 1)) for i in range(m)})
    per = max(800, budget // max(1, len(idx)))
    return "\n".join(f"[{d['unit']} {i + 1}] {pages[i][:per]}" for i in idx)


# ----------------------------------------------------------------------------
# Tagging, authority control, version detection
# ----------------------------------------------------------------------------
def canon_client(name, tax):
    name = (name or "Unknown").strip()
    for canon, aliases in tax["clients"].items():
        if name.lower() == canon.lower() or name.lower() in [a.lower() for a in aliases]:
            return canon
    return name


def pick_from(value, options):
    for o in options:
        if str(value).strip().lower() == o.lower():
            return o
    return "Other"


def auto_tag(filename, text, tax):
    sys = ("You are a library cataloguer for an SAP consultancy. Return ONLY JSON with keys: "
           "client, industry, sap_module, phase, doc_type, doc_title, version, summary (2 sentences). "
           f"industry must be one of {tax['industries']}; sap_module one of {tax['sap_modules']}; "
           f"phase one of {tax['phases']}; doc_type one of {tax['doc_types']}. "
           "client is the organisation the work is for ('TMC Internal' if none). version only if visible, else ''.")
    try:
        m = parse_json(llm([{"role": "system", "content": sys}, {"role": "user", "content": text}], json_mode=True))
    except Exception:
        m = {}
    return {
        "client": canon_client(m.get("client"), tax),
        "industry": pick_from(m.get("industry"), tax["industries"]),
        "sap_module": pick_from(m.get("sap_module"), tax["sap_modules"]),
        "phase": pick_from(m.get("phase"), tax["phases"]),
        "doc_type": pick_from(m.get("doc_type"), tax["doc_types"]),
        "doc_title": (m.get("doc_title") or filename).strip(),
        "version": str(m.get("version") or ""),
        "summary": m.get("summary") or "",
    }


def vkey(d):
    nums = re.findall(r"\d+", d["meta"].get("version") or "") or re.findall(r"\d+", d["name"])
    return tuple(int(n) for n in nums)


def detect_versions(new, library):
    for d in library:
        if d["hash"] == new["hash"]:
            return "duplicate", d
    ns = set(tok(" ".join(new["pages"])[:6000]))
    for d in library:
        if d["meta"]["client"] != new["meta"]["client"] or d["meta"]["doc_type"] != new["meta"]["doc_type"]:
            continue
        ts = difflib.SequenceMatcher(None, d["meta"]["doc_title"].lower(), new["meta"]["doc_title"].lower()).ratio()
        os_ = set(tok(" ".join(d["pages"])[:6000]))
        js = len(ns & os_) / len(ns | os_) if ns | os_ else 0
        if ts > 0.8 or js > 0.6:
            return "version", d
    return None, None


def ingest(file, access, tax):
    raw, ext = file.getvalue(), file.name.rsplit(".", 1)[-1].lower()
    pages, author, unit = EXTRACTORS[ext](raw)
    if not "".join(pages).strip():
        return "warning", f"{file.name}: no readable text (scanned PDF? OCR is needed)."
    meta = auto_tag(file.name, " ".join(pages)[:5000], tax)
    doc = {"id": uuid.uuid4().hex[:10], "name": file.name, "ext": ext, "pages": pages, "unit": unit,
           "author": author, "meta": meta, "hash": hashlib.sha256(raw).hexdigest(),
           "uploaded_by": st.session_state.user["name"], "uploaded_at": datetime.now().isoformat(timespec="seconds"),
           "access": access or ["All"], "sanitized": True, "status": "Current", "superseded_by": None}
    kind, other = detect_versions(doc, st.session_state.library)
    if kind == "duplicate":
        return "warning", f"{file.name} is identical to “{other['meta']['doc_title']}” - not added."
    note = ""
    if kind == "version":
        if vkey(doc) >= vkey(other):
            other["status"], other["superseded_by"] = "Superseded", doc["id"]
            note = f" It supersedes “{other['meta']['doc_title']}”."
        else:
            doc["status"], doc["superseded_by"] = "Superseded", other["id"]
            note = f" A newer version already exists (“{other['meta']['doc_title']}”)."
    path = DATA / "files" / f"{doc['id']}_{re.sub(r'[^A-Za-z0-9._-]', '_', file.name)}"
    path.write_bytes(raw)
    doc["path"] = str(path)
    st.session_state.library.append(doc)
    save_library()
    audit("upload", doc=doc["name"], access=doc["access"])
    return "success", f"Added “{meta['doc_title']}” ({meta['doc_type']}, {meta['client']}).{note}"


# <<EXPORTS>>
# ----------------------------------------------------------------------------
# Report export: Word, PowerPoint, Excel, PDF
# ----------------------------------------------------------------------------
def norm_report(r, fallback):
    secs = [{"heading": str(s.get("heading", "")), "bullets": [str(b) for b in s.get("bullets", [])]}
            for s in r.get("sections", []) if isinstance(s, dict)]
    t = r.get("table") or {}
    cols = [str(c) for c in t.get("columns", [])] if isinstance(t, dict) else []
    rows = [(list(map(str, row)) + [""] * len(cols))[:len(cols)] for row in (t.get("rows", []) if cols else []) if isinstance(row, list)]
    return {"title": str(r.get("title") or fallback), "executive_summary": str(r.get("executive_summary", "")),
            "sections": secs, "table": {"columns": cols, "rows": rows} if cols and rows else None,
            "sources": [str(s) for s in r.get("sources", [])]}


def build_docx(r):
    d = docx.Document()
    d.add_heading(r["title"], 0)
    d.add_paragraph(f"Generated by TMC Knowledge Lens • {datetime.now():%d %b %Y}").italic = True
    d.add_heading("Executive summary", 1)
    d.add_paragraph(r["executive_summary"])
    for s in r["sections"]:
        d.add_heading(s["heading"], 1)
        for b in s["bullets"]:
            d.add_paragraph(b, style="List Bullet")
    if r["table"]:
        t = d.add_table(rows=1, cols=len(r["table"]["columns"]))
        t.style = "Light Grid Accent 1"
        for i, c in enumerate(r["table"]["columns"]):
            t.rows[0].cells[i].text = c
        for row in r["table"]["rows"]:
            cells = t.add_row().cells
            for i, v in enumerate(row):
                cells[i].text = v
    d.add_heading("Sources", 1)
    for s in r["sources"]:
        d.add_paragraph(s, style="List Bullet")
    buf = io.BytesIO()
    d.save(buf)
    return buf.getvalue()


def build_pptx(r):
    prs = Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[0])
    s.shapes.title.text = r["title"]
    s.placeholders[1].text = f"TMC Knowledge Lens • {datetime.now():%d %b %Y}"

    def bullets_slide(title, items, size=18):
        sl = prs.slides.add_slide(prs.slide_layouts[1])
        sl.shapes.title.text = title
        tf = sl.placeholders[1].text_frame
        tf.clear()
        for i, it in enumerate(items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = it
            p.font.size = Pt(size)

    bullets_slide("Executive summary", [r["executive_summary"][:900]], 16 if len(r["executive_summary"]) < 500 else 13)
    for sec in r["sections"]:
        chunks = [sec["bullets"][i:i + 5] for i in range(0, len(sec["bullets"]), 5)] or [[]]
        for n, ch in enumerate(chunks):
            bullets_slide(sec["heading"] + (f" ({n + 1})" if n else ""), [b[:260] for b in ch], 16)
    if r["table"]:
        cols, rows = r["table"]["columns"], r["table"]["rows"][:8]
        sl = prs.slides.add_slide(prs.slide_layouts[5])
        sl.shapes.title.text = "Summary table"
        tbl = sl.shapes.add_table(len(rows) + 1, len(cols), Inches(0.4), Inches(1.6), Inches(9.2), Inches(0.4 * (len(rows) + 1))).table
        for i, c in enumerate(cols):
            tbl.cell(0, i).text = c
        for ri, row in enumerate(rows, 1):
            for ci, v in enumerate(row):
                tbl.cell(ri, ci).text = v[:120]
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.text_frame.paragraphs:
                    p.font.size = Pt(10)
    bullets_slide("Sources", r["sources"] or ["-"], 14)
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def build_xlsx(r):
    wb = openpyxl.Workbook()
    head = Font(bold=True, color="FFFFFF")
    fill = PatternFill("solid", fgColor="5B2C6F")

    def sheet(ws, header, rows, widths):
        ws.append(header)
        for c in ws[1]:
            c.font, c.fill = head, fill
        for row in rows:
            ws.append(row)
        for i, w in enumerate(widths):
            ws.column_dimensions[chr(65 + i)].width = w
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical="top")

    ws = wb.active
    ws.title = "Summary"
    sheet(ws, ["Title", "Executive summary"], [[r["title"], r["executive_summary"]]], [40, 110])
    sheet(wb.create_sheet("Findings"), ["Section", "Point"],
          [[s["heading"], b] for s in r["sections"] for b in s["bullets"]], [32, 120])
    if r["table"]:
        sheet(wb.create_sheet("Table"), r["table"]["columns"], r["table"]["rows"], [30] * len(r["table"]["columns"]))
    sheet(wb.create_sheet("Sources"), ["Source"], [[s] for s in r["sources"]], [100])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_pdf(r):
    ss = getSampleStyleSheet()
    esc = lambda s: xml_escape(s).encode("latin-1", "replace").decode("latin-1")
    story = [Paragraph(esc(r["title"]), ss["Title"]),
             Paragraph(f"Generated by TMC Knowledge Lens • {datetime.now():%d %b %Y}", ss["Italic"]), Spacer(1, 12),
             Paragraph("Executive summary", ss["Heading2"]), Paragraph(esc(r["executive_summary"]), ss["BodyText"])]
    for s in r["sections"]:
        story.append(Paragraph(esc(s["heading"]), ss["Heading2"]))
        story += [Paragraph("• " + esc(b), ss["BodyText"]) for b in s["bullets"]]
    if r["table"]:
        data = [[Paragraph(f"<b>{esc(c)}</b>", ss["BodyText"]) for c in r["table"]["columns"]]]
        data += [[Paragraph(esc(v), ss["BodyText"]) for v in row] for row in r["table"]["rows"]]
        t = Table(data, repeatRows=1)
        t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.4, colors.grey), ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E8DAEF")),
                               ("VALIGN", (0, 0), (-1, -1), "TOP")]))
        story += [Spacer(1, 10), t]
    story.append(Paragraph("Sources", ss["Heading2"]))
    story += [Paragraph("• " + esc(s), ss["BodyText"]) for s in r["sources"]]
    buf = io.BytesIO()
    SimpleDocTemplate(buf, pagesize=A4, leftMargin=48, rightMargin=48, topMargin=48, bottomMargin=48).build(story)
    return buf.getvalue()
# <<END EXPORTS>>


# ----------------------------------------------------------------------------
# AI actions
# ----------------------------------------------------------------------------
def make_summary(docs):
    budget = max(3000, 24000 // len(docs))
    ctx = "\n\n".join(f"### {d['meta']['doc_title']} ({d['meta']['doc_type']}, {d['meta']['client']})\n{sample_text(d, budget)}" for d in docs)
    sys = ("Return ONLY JSON: {\"overview\": one paragraph, \"key_decisions\": [..], \"deliverables\": [..], "
           "\"timeline\": [\"milestone - date\"], \"risks\": [..]}. Use only the sources; use [] when absent. "
           "Mention the document name in brackets where useful.")
    s = parse_json(llm([{"role": "system", "content": sys}, {"role": "user", "content": ctx}], json_mode=True))
    return {"overview": str(s.get("overview", "")), **{k: [str(x) for x in s.get(k, [])] for k in ("key_decisions", "deliverables", "timeline", "risks")},
            "docs": [d["meta"]["doc_title"] for d in docs]}


def answer_question(q, docs, tax):
    hits = retrieve(q, docs, tax, k=8)
    ctx = "\n\n".join(f"[S{i + 1}] {h['doc_title']} - {h['unit']} {h['page']}\n{h['text'][:1800]}" for i, h in enumerate(hits))
    sys = ("You are TMC Knowledge Lens. Answer ONLY from the sources. Cite every claim with [S#]. "
           "If the sources do not contain the answer, say so plainly. Be concise; use bullets for lists.")
    ans = llm([{"role": "system", "content": sys}, {"role": "user", "content": f"Sources:\n{ctx}\n\nQuestion: {q}"}])
    used = sorted({int(n) for n in re.findall(r"\[S(\d+)\]", ans) if 0 < int(n) <= len(hits)}) or list(range(1, min(3, len(hits)) + 1))
    cites = [{"n": n, "title": hits[n - 1]["doc_title"], "loc": f"{hits[n - 1]['unit']} {hits[n - 1]['page']}",
              "snippet": make_snippet(hits[n - 1]["text"], tok(q), 280)} for n in used]
    return ans, cites


def make_report(rtype, focus, docs, tax):
    query = focus or rtype
    hits = retrieve(query, docs, tax, k=12)
    ctx = "\n\n".join(f"[{h['doc_title']} - {h['unit']} {h['page']}]\n{h['text'][:1800]}" for h in hits)
    instr = REPORT_TYPES[rtype]
    sys = ("You write professional consulting reports from source excerpts only. Return ONLY JSON: "
           "{\"title\": str, \"executive_summary\": str (max 120 words), \"sections\": [{\"heading\": str, \"bullets\": [str]}] (3-6 sections, 3-6 bullets each), "
           "\"table\": {\"columns\": [str], \"rows\": [[str]]} or null}. Put (Document name, page) at the end of factual bullets. Never invent facts.")
    raw = llm([{"role": "system", "content": sys},
               {"role": "user", "content": f"Report type: {rtype}. {instr}\nFocus: {focus or 'general'}\n\nSources:\n{ctx}"}], json_mode=True)
    r = norm_report(parse_json(raw), rtype)
    r["sources"] = [f"{d['meta']['doc_title']} ({d['name']}, v{d['meta'].get('version') or '-'}, {d['status']})" for d in docs]
    stem = re.sub(r"[^A-Za-z0-9]+", "_", r["title"])[:50] or "report"
    return {"data": r, "stem": stem, "files": {"docx": build_docx(r), "pptx": build_pptx(r), "xlsx": build_xlsx(r), "pdf": build_pdf(r)}}


# ----------------------------------------------------------------------------
# Session state + sidebar (demo identity; replace with SSO in production)
# ----------------------------------------------------------------------------
ss_ = st.session_state
ss_.setdefault("library", load_library())
ss_.setdefault("selected", set())
ss_.setdefault("results", None)
ss_.setdefault("chat", [])
ss_.setdefault("summary", None)
ss_.setdefault("report", None)
ss_.setdefault("flash", [])
ss_.setdefault("up_key", 0)
ss_.setdefault("rfp_key", 0)
tax = load_taxonomy()

with st.sidebar:
    st.markdown("### 👤 Session")
    uname = st.text_input("Your name", "Nadeem Sohail")
    ugroups = st.multiselect("Access groups", GROUPS, default=["Pre-Sales"],
                             help="Demo only. In production these come from SSO / Azure AD groups.")
    admin_pw = st.secrets.get("ADMIN_PASSWORD")
    is_admin = False
    if st.checkbox("Admin mode"):
        is_admin = admin_pw is None or st.text_input("Admin password", type="password") == admin_pw
        if admin_pw is None:
            st.caption("Set ADMIN_PASSWORD in secrets to protect admin mode.")
    ss_.user = {"name": uname, "groups": ugroups, "admin": is_admin}
    st.markdown("---")
    st.caption("Answers come only from the sources you tick. Every question, report and upload is written to the audit log.")


def can_see(d):
    u = ss_.user
    return u["admin"] or "All" in d["access"] or bool(set(d["access"]) & set(u["groups"]))


def esc(s):
    return html.escape(str(s))


def chip(t, cls=""):
    return f'<span class="chip {cls}">{esc(t)}</span>' if t and t != "Other" else ""


# callbacks ------------------------------------------------------------------
def toggle(did):
    (ss_.selected.add if ss_[f"sel_{did}"] else ss_.selected.discard)(did)


def select_all(ids):
    for i in ids:
        ss_.selected.add(i)
        ss_[f"sel_{i}"] = True


def clear_selection():
    ss_.selected.clear()
    for k in [k for k in ss_ if k.startswith("sel_")]:
        ss_[k] = False


def clear_results():
    ss_.results = None


def rate(idx, val):
    ss_.chat[idx]["rating"] = val
    audit("feedback", rating=val, question=ss_.chat[idx]["q"])


visible = [d for d in ss_.library if can_see(d)]
selected_docs = [d for d in visible if d["id"] in ss_.selected]
by_id = {d["id"]: d for d in ss_.library}

# ----------------------------------------------------------------------------
# Header + tabs
# ----------------------------------------------------------------------------
st.markdown("""
<div class="topbar">
  <img src="https://tmcltd.com/wp-content/uploads/2025/10/TallyMarks-Consulting-TMC-logo.webp" width="72">
  <div><h1>Knowledge Lens</h1><p>Search TMC project knowledge, select what matters, get cited answers and ready-to-send reports.</p></div>
</div>""", unsafe_allow_html=True)

tab_ws, tab_admin = st.tabs(["Workspace", "Admin & pilot metrics"])

# ============================================================================
# WORKSPACE
# ============================================================================
with tab_ws:
    col_src, col_mid, col_studio = st.columns([1.25, 2.3, 1.2], gap="medium")

    # ---------------- LEFT: SOURCES ----------------
    with col_src:
        with st.container(height=PANEL_H, border=True):
            st.markdown("#### Sources")
            for kind, msg in ss_.flash:
                getattr(st, kind)(msg)
            ss_.flash = []

            with st.form("search_form"):
                q = st.text_input("Search", placeholder="Search the library: client, SAP module, topic…", label_visibility="collapsed")
                go = st.form_submit_button("Search library", type="primary", use_container_width=True)

            with st.expander("Filters"):
                f_client = st.multiselect("Client", sorted({d["meta"]["client"] for d in visible}), key="f_client")
                f_type = st.multiselect("Document type", sorted({d["meta"]["doc_type"] for d in visible}), key="f_type")
                f_mod = st.multiselect("SAP module", sorted({d["meta"]["sap_module"] for d in visible}), key="f_mod")
                f_phase = st.multiselect("Phase", sorted({d["meta"]["phase"] for d in visible}), key="f_phase")
                f_old = st.toggle("Include superseded versions", value=False, key="f_old")

            def passes(d):
                m = d["meta"]
                return ((not f_client or m["client"] in f_client) and (not f_type or m["doc_type"] in f_type)
                        and (not f_mod or m["sap_module"] in f_mod) and (not f_phase or m["phase"] in f_phase)
                        and (f_old or d["status"] == "Current"))

            candidates = [d for d in visible if passes(d)]
            if go:
                if q.strip():
                    t0 = time.time()
                    ss_.results = {"label": f"Results for “{q.strip()}”", "items": search_docs(q, visible, tax)}
                    audit("search", query=q.strip(), hits=len(ss_.results["items"]), seconds=round(time.time() - t0, 2))
                else:
                    ss_.results = None

            allowed = {d["id"] for d in candidates}
            if ss_.results:
                shown = [it for it in ss_.results["items"] if it["id"] in allowed]
                label = f"{ss_.results['label']} - {len(shown)} document(s)"
            else:
                shown = [{"id": d["id"], "rel": None, "page": None, "snippet": ""}
                         for d in sorted(candidates, key=lambda x: x["uploaded_at"], reverse=True)]
                label = f"Library - {len(shown)} document(s)"
            shown = shown[:60]

            st.caption(label)
            if ss_.results:
                st.button("Clear search", on_click=clear_results, use_container_width=True)
            if shown:
                a, b = st.columns(2)
                a.button("Select all", on_click=select_all, args=([s["id"] for s in shown],), use_container_width=True)
                b.button("Clear selection", on_click=clear_selection, use_container_width=True)
            else:
                st.info("No documents yet. Add documents below, or loosen the filters." if not ss_.library else "Nothing matches. Try other words or clear the filters.")

            for it in shown:
                d = by_id[it["id"]]
                m = d["meta"]
                key = f"sel_{d['id']}"
                if key not in ss_:
                    ss_[key] = d["id"] in ss_.selected
                c1, c2 = st.columns([1, 9])
                c1.checkbox("select", key=key, on_change=toggle, args=(d["id"],), label_visibility="collapsed")
                badge = '<span class="chip ok">Current</span>' if d["status"] == "Current" else '<span class="chip old">Superseded</span>'
                body = f'<div class="doc-title">{esc(m["doc_title"])}</div>{badge}{chip(m["doc_type"], "p")}{chip(m["client"])}{chip(m["sap_module"])}{chip(m["phase"])}'
                if it["rel"] is not None:
                    body += f'<div class="relbar"><div style="width:{it["rel"]}%"></div></div><div class="small">Relevance {it["rel"]}%' + (f' • {d["unit"]} {it["page"]}' if it["page"] else "") + "</div>"
                    if it["snippet"]:
                        body += f'<div class="snip">{esc(it["snippet"])}</div>'
                c2.markdown(body, unsafe_allow_html=True)

            st.markdown("---")
            with st.expander("Add documents to the library"):
                files = st.file_uploader("Files", type=list(EXTRACTORS), accept_multiple_files=True,
                                         key=f"up_{ss_.up_key}", label_visibility="collapsed")
                access = st.multiselect("Who can see these?", ["All"] + GROUPS, default=["All"])
                ok = st.checkbox("I confirm these files are sanitized and approved for Knowledge Lens")
                if st.button("Add to library", disabled=not (files and ok), use_container_width=True):
                    for f in files:
                        with st.spinner(f"Reading and tagging {f.name}…"):
                            try:
                                ss_.flash.append(ingest(f, access, tax))
                            except Exception as e:
                                ss_.flash.append(("error", f"{f.name}: {e}"))
                    ss_.up_key += 1
                    st.rerun()

    # ---------------- CENTRE: SUMMARY + CHAT ----------------
    with col_mid:
        with st.container(height=PANEL_H, border=True):
            n = len(selected_docs)
            h1, h2 = st.columns([3, 2])
            h1.markdown(f"#### Project at a glance  \n<span class='small'>{n} source(s) selected</span>", unsafe_allow_html=True)
            if h2.button("Summarize selected", type="primary", disabled=n == 0, use_container_width=True):
                with st.spinner("Summarizing…"):
                    try:
                        ss_.summary = make_summary(selected_docs)
                        audit("summary", docs=[d["name"] for d in selected_docs])
                    except Exception as e:
                        st.error(f"Could not summarize: {e}")

            if ss_.summary:
                s = ss_.summary
                st.markdown(f'<div class="glance"><h4>Summary of {len(s["docs"])} document(s)</h4>{esc(s["overview"])}</div>', unsafe_allow_html=True)
                cols = st.columns(2)
                for i, (k, t) in enumerate([("key_decisions", "Key decisions"), ("deliverables", "Deliverables"), ("timeline", "Timeline"), ("risks", "Risks")]):
                    if s[k]:
                        with cols[i % 2]:
                            st.markdown(f"**{t}**")
                            st.markdown("\n".join(f"- {x}" for x in s[k]))
            else:
                st.info("Search, tick the documents you need, then press Summarize selected.")

            st.markdown("##### Ask the selected documents")
            chat_box = st.container(height=330, border=False)
            prompt = st.chat_input("Ask a question about the selected documents…", disabled=n == 0)
            if prompt:
                t0 = time.time()
                try:
                    ans, cites = answer_question(prompt, selected_docs, tax)
                    sec = round(time.time() - t0, 1)
                    ss_.chat.append({"q": prompt, "a": ans, "cites": cites, "seconds": sec, "rating": None})
                    audit("ask", question=prompt, seconds=sec, docs=len(selected_docs))
                except Exception as e:
                    ss_.chat.append({"q": prompt, "a": f"Sorry, the question failed: {e}", "cites": [], "seconds": 0, "rating": 0})
            with chat_box:
                if not ss_.chat:
                    st.caption("Answers include citations like [S1] that point to the document and page.")
                for i, c in enumerate(ss_.chat):
                    with st.chat_message("user"):
                        st.write(c["q"])
                    with st.chat_message("assistant"):
                        st.write(c["a"])
                        if c["cites"]:
                            with st.expander("Sources"):
                                for x in c["cites"]:
                                    st.markdown(f"**[S{x['n']}] {x['title']}** - {x['loc']}")
                                    st.caption(x["snippet"])
                        if c["rating"] is None:
                            r1, r2, _ = st.columns([1, 1, 6])
                            r1.button("👍", key=f"up_{i}", on_click=rate, args=(i, 1))
                            r2.button("👎", key=f"dn_{i}", on_click=rate, args=(i, -1))
                        elif c["rating"] != 0:
                            st.caption("Thanks for the feedback.")

    # ---------------- RIGHT: STUDIO ----------------
    with col_studio:
        with st.container(height=PANEL_H, border=True):
            st.markdown("#### Studio")

            with st.expander("Find similar past work", expanded=False):
                rfp = st.file_uploader("Upload an RFP or brief", type=["pdf", "docx", "pptx", "txt"], key=f"rfp_{ss_.rfp_key}", label_visibility="collapsed")
                if st.button("Find similar documents", disabled=rfp is None, use_container_width=True):
                    pages, _, _ = EXTRACTORS[rfp.name.rsplit(".", 1)[-1].lower()](rfp.getvalue())
                    words = Counter(t for t in tok(" ".join(pages)[:20000]) if len(t) > 3)
                    query = " ".join(w for w, _ in words.most_common(40))
                    ss_.results = {"label": f"Similar to “{rfp.name}”", "items": search_docs(query, visible, tax)}
                    audit("similar_search", file=rfp.name)
                    st.success("Matches are listed under Sources.")

            st.markdown("**Generate report**")
            rtype = st.selectbox("Report type", list(REPORT_TYPES), label_visibility="collapsed")
            focus = st.text_area("Focus", placeholder="Optional: what should the report concentrate on?", height=80)
            if st.button("Generate report", type="primary", disabled=not selected_docs, use_container_width=True):
                with st.spinner("Writing the report in 4 formats…"):
                    try:
                        ss_.report = make_report(rtype, focus.strip(), selected_docs, tax)
                        audit("report", rtype=rtype, docs=[d["name"] for d in selected_docs])
                    except Exception as e:
                        st.error(f"Could not generate the report: {e}")
            if not selected_docs:
                st.caption("Select at least one source to enable reports.")

            if ss_.report:
                rp = ss_.report
                st.success(rp["data"]["title"])
                MIME = {"docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "pdf": "application/pdf"}
                LAB = {"docx": "Word", "pptx": "PowerPoint", "xlsx": "Excel", "pdf": "PDF"}
                g = st.columns(2)
                for i, ext in enumerate(["docx", "pptx", "xlsx", "pdf"]):
                    g[i % 2].download_button(LAB[ext], rp["files"][ext], file_name=f"{rp['stem']}.{ext}", mime=MIME[ext],
                                             key=f"dl_{ext}", use_container_width=True)
                with st.expander("Preview"):
                    st.write(rp["data"]["executive_summary"])
                    for s in rp["data"]["sections"]:
                        st.markdown(f"**{s['heading']}**")
                        st.markdown("\n".join(f"- {b}" for b in s["bullets"]))

            if selected_docs:
                st.markdown("**Selected sources**")
                for d in selected_docs:
                    m = d["meta"]
                    with st.expander(m["doc_title"]):
                        st.markdown(f"**Client:** {esc(m['client'])}  \n**Type:** {esc(m['doc_type'])}  \n**SAP module:** {esc(m['sap_module'])}  \n"
                                    f"**Phase:** {esc(m['phase'])}  \n**Version:** {esc(m['version'] or '-')}  \n**Author:** {esc(d['author'])}  \n"
                                    f"**Status:** {d['status']}  \n**Added by:** {esc(d['uploaded_by'])}", unsafe_allow_html=False)
                        if m["summary"]:
                            st.caption(m["summary"])
                        if d["status"] == "Superseded" and d.get("superseded_by") in by_id:
                            st.warning(f"Newer version: {by_id[d['superseded_by']]['meta']['doc_title']}")
                        p = Path(d.get("path", ""))
                        if p.exists():
                            st.download_button("Download original", p.read_bytes(), file_name=d["name"], key=f"orig_{d['id']}")

# ============================================================================
# ADMIN & PILOT METRICS
# ============================================================================
with tab_admin:
    if not ss_.user["admin"]:
        st.info("Turn on Admin mode in the sidebar to manage the library, vocabulary and pilot metrics.")
    else:
        import pandas as pd
        log = read_audit()
        asks = [r for r in log if r["action"] == "ask"]
        fb = [r for r in log if r["action"] == "feedback"]
        st.markdown("#### Pilot metrics")
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Documents", len(ss_.library))
        m2.metric("Questions asked", len(asks))
        m3.metric("Avg. answer time", f"{(sum(r['seconds'] for r in asks) / len(asks)):.1f}s" if asks else "-")
        m4.metric("Helpful answers", f"{100 * sum(r['rating'] == 1 for r in fb) / len(fb):.0f}%" if fb else "-")
        m5.metric("Reports generated", sum(r["action"] == "report" for r in log))
        st.caption("Use these alongside the manual baseline measured in November (search test, summary test).")

        st.markdown("#### Library")
        if ss_.library:
            df = pd.DataFrame([{"id": d["id"], "title": d["meta"]["doc_title"], "file": d["name"], "client": d["meta"]["client"],
                                "type": d["meta"]["doc_type"], "module": d["meta"]["sap_module"], "phase": d["meta"]["phase"],
                                "status": d["status"], "access": ", ".join(d["access"]), "added by": d["uploaded_by"]} for d in ss_.library])
            st.dataframe(df, use_container_width=True, hide_index=True)
            pick = st.selectbox("Manage a document", [d["id"] for d in ss_.library], format_func=lambda i: by_id[i]["meta"]["doc_title"])
            d = by_id[pick]
            a1, a2, a3 = st.columns(3)
            if a1.button("Mark as current", use_container_width=True):
                d["status"], d["superseded_by"] = "Current", None
                save_library(); audit("status", doc=d["name"], status="Current"); st.rerun()
            if a2.button("Mark as superseded", use_container_width=True):
                d["status"] = "Superseded"
                save_library(); audit("status", doc=d["name"], status="Superseded"); st.rerun()
            if a3.button("Delete document", use_container_width=True):
                ss_.library = [x for x in ss_.library if x["id"] != pick]
                ss_.selected.discard(pick)
                Path(d.get("path", "")).unlink(missing_ok=True)
                save_library(); audit("delete", doc=d["name"]); st.rerun()
            new_access = st.multiselect("Access for this document", ["All"] + GROUPS, default=d["access"], key=f"acc_{pick}")
            if st.button("Save access") and new_access:
                d["access"] = new_access
                save_library(); audit("access", doc=d["name"], access=new_access); st.rerun()

        st.markdown("#### Controlled vocabulary")
        st.caption("Client aliases map variants (TML, Tata Motors Ltd) to one record. The glossary expands TMC terms in searches.")
        txt = st.text_area("taxonomy.json", json.dumps(tax, indent=2, ensure_ascii=False), height=300)
        if st.button("Save vocabulary"):
            try:
                t = json.loads(txt)
                assert all(k in t for k in DEFAULT_TAXONOMY)
                save_taxonomy(t); audit("taxonomy_update"); st.success("Saved.")
            except Exception as e:
                st.error(f"Invalid vocabulary: {e}")

        st.markdown("#### Audit log")
        if log:
            ldf = pd.DataFrame(log).tail(300).iloc[::-1]
            st.dataframe(ldf, use_container_width=True, hide_index=True)
            st.download_button("Download audit log (CSV)", pd.DataFrame(log).to_csv(index=False), file_name="audit_log.csv")
