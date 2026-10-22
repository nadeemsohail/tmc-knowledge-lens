import streamlit as st
from openai import OpenAI
import pypdf
import docx
from pptx import Presentation
from langchain_text_splitters import RecursiveCharacterTextSplitter
import json
import datetime
import io

# 1. Page Configuration
st.set_page_config(page_title="TMC Knowledge Lens", page_icon="🔍", layout="wide")

# Custom CSS to mimic NotebookLM's clean, card-based look
st.markdown("""
    <style>
    .stApp { background-color: #FAFAFA; font-family: 'Inter', sans-serif; }
    header {visibility: hidden;}
    .card {
        background-color: white; padding: 20px; border-radius: 12px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05); margin-bottom: 15px;
        border: 1px solid #E5E7EB;
    }
    .stButton>button {
        background-color: #5B2C6F; color: white; border-radius: 8px;
        font-weight: 600; padding: 8px 16px; width: 100%;
        transition: 0.2s; border: none;
    }
    .stButton>button:hover { background-color: #7D3C98; color: white; }
    .stTextInput>div>div>input { 
        border-radius: 20px; border: 1px solid #D1D5DB; padding: 10px 20px;
    }
    h3 { font-size: 1.2rem; font-weight: 600; color: #1F2937; margin-bottom: 10px;}
    </style>
""", unsafe_allow_html=True)

# 2. Initialize Groq API Client
try:
    client = OpenAI(base_url="https://api.groq.com/openai/v1", api_key=st.secrets["GITHUB_TOKEN"])
except Exception as e:
    st.error(f"API Key error: {e}")
    st.stop()

# Session State Initialization
if "metadata" not in st.session_state: st.session_state.metadata = {}
if "summary" not in st.session_state: st.session_state.summary = ""
if "doc_text" not in st.session_state: st.session_state.doc_text = ""
if "file_name" not in st.session_state: st.session_state.file_name = ""
if "file_bytes" not in st.session_state: st.session_state.file_bytes = None
if "chat_history" not in st.session_state: st.session_state.chat_history = []

# --- TEXT EXTRACTION HELPER FUNCTIONS ---
def extract_text_from_pdf(file_bytes):
    reader = pypdf.PdfReader(io.BytesIO(file_bytes))
    text = ""
    for page in reader.pages:
        text += (page.extract_text() or "") + "\n"
    return text, reader.metadata

def extract_text_from_docx(file_bytes):
    doc = docx.Document(io.BytesIO(file_bytes))
    text = "\n".join([para.text for para in doc.paragraphs])
    return text, None

def extract_text_from_pptx(file_bytes):
    prs = Presentation(io.BytesIO(file_bytes))
    text = ""
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text += shape.text + "\n"
    return text, None

def extract_text_from_txt(file_bytes):
    return file_bytes.decode("utf-8", errors="ignore"), None

# --- HEADER ---
col_logo, col_title = st.columns([1, 11], vertical_alignment="center")
with col_logo:
    st.image("https://tmcltd.com/wp-content/uploads/2025/10/TallyMarks-Consulting-TMC-logo.webp", width=100)
with col_title:
    st.markdown('<h2 style="color: #5B2C6F; margin-bottom: 0;">Knowledge Lens</h2>', unsafe_allow_html=True)
    st.markdown('<p style="color: gray; margin-top: 0; font-size: 0.9rem;">Enterprise AI Workspace</p>', unsafe_allow_html=True)

st.markdown("<hr style='margin-top: 5px; margin-bottom: 20px;'>", unsafe_allow_html=True)

# --- THREE COLUMN LAYOUT ---
col_sources, col_chat, col_studio = st.columns([1.2, 2, 1.2])

# ==========================================
# LEFT COLUMN: SOURCES
# ==========================================
with col_sources:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### 📁 Sources")
    
    # UPDATED: Accept multiple file types
    uploaded_file = st.file_uploader(
        "Add a source file", 
        type=["pdf", "docx", "pptx", "txt"], 
        label_visibility="collapsed"
    )
    
    if uploaded_file is not None:
        if st.session_state.file_name != uploaded_file.name:
            with st.spinner("Processing document..."):
                st.session_state.file_name = uploaded_file.name
                file_bytes = uploaded_file.getvalue()
                st.session_state.file_bytes = file_bytes
                
                # Determine file type and extract text
                file_ext = uploaded_file.name.split(".")[-1].lower()
                native_author = "Unknown"
                
                if file_ext == "pdf":
                    raw_text, meta = extract_text_from_pdf(file_bytes)
                    if meta and meta.author: native_author = meta.author
                elif file_ext == "docx":
                    raw_text, _ = extract_text_from_docx(file_bytes)
                elif file_ext == "pptx":
                    raw_text, _ = extract_text_from_pptx(file_bytes)
                elif file_ext == "txt":
                    raw_text, _ = extract_text_from_txt(file_bytes)
                
                # Split into chunks
                splitter = RecursiveCharacterTextSplitter(chunk_size=2000, chunk_overlap=200)
                chunks = splitter.split_text(raw_text)
                st.session_state.doc_text = "\n".join(chunks[:20])
                
                # AI Auto-Tagging
                sys_prompt = f"""
                Read the text. Return a JSON object with:
                "client": (e.g., "Tata Motors" or "Other"),
                "module": (e.g., "S/4HANA" or "All"),
                "doc_type": (e.g., "Blueprint", "Proposal", "Presentation", etc.),
                "doc_title": (Real title of the doc),
                "doc_author": (Author or '{native_author}'),
                "summary": "2-sentence summary.",
                "key_decisions": "Bullet points.",
                "timeline": "Key dates.",
                "risks": "Risks mentioned."
                """
                try:
                    res = client.chat.completions.create(
                        model="openai/gpt-oss-120b",
                        messages=[{"role": "system", "content": sys_prompt}, {"role": "user", "content": st.session_state.doc_text[:6000]}],
                        response_format={ "type": "json_object" }
                    )
                    meta = json.loads(res.choices[0].message.content)
                    meta["submitter"] = "Nadeem Sohail" # Hardcoded for demo
                    meta["dept"] = "AI Engineering"
                    st.session_state.metadata = meta
                    st.session_state.summary = meta.get("summary", "")
                except Exception as e:
                    st.error(f"AI Parsing Error: {e}")
                st.rerun()

    if st.session_state.file_name:
        st.success(f"✅ {st.session_state.file_name}")
        st.caption(f"Client: {st.session_state.metadata.get('client', 'Unknown')} | Type: {st.session_state.metadata.get('doc_type', 'Unknown')}")
    st.markdown('</div>', unsafe_allow_html=True)


# ==========================================
# MIDDLE COLUMN: CHAT / Q&A
# ==========================================
with col_chat:
    st.markdown("### 💬 Ask the Lens")
    
    # Display Chat History
    for chat in st.session_state.chat_history:
        with st.chat_message("user"):
            st.write(chat["question"])
        with st.chat_message("assistant"):
            st.write(chat["answer"])
            st.caption(f"Source: {st.session_state.file_name}")

    # Chat Input
    with st.form(key="chat_form", clear_on_submit=True):
        query = st.text_input("Ask a question...", placeholder="e.g., What are the main risks?", label_visibility="collapsed")
        submit = st.form_submit_button("Send")
        
        if submit and query:
            if not st.session_state.file_name:
                st.warning("Please upload a document first.")
            else:
                with st.spinner("Searching..."):
                    prompt = f"Answer based ONLY on this document:\n\n{st.session_state.doc_text}\n\nQuestion: {query}"
                    response = client.chat.completions.create(
                        model="openai/gpt-oss-120b",
                        messages=[{"role": "user", "content": prompt}],
                        temperature=0.1
                    )
                    answer = response.choices[0].message.content
                    
                    st.session_state.chat_history.append({"question": query, "answer": answer})
                    st.rerun()


# ==========================================
# RIGHT COLUMN: STUDIO
# ==========================================
with col_studio:
    st.markdown("### 🎛️ Studio")
    
    if st.session_state.file_name:
        # 1. Project at a Glance
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown("#### 📄 At a Glance")
        st.info(st.session_state.summary)
        
        with st.expander("View Key Details"):
            st.markdown(f"**Decisions:**\n{st.session_state.metadata.get('key_decisions', 'N/A')}")
            st.markdown(f"**Timeline:**\n{st.session_state.metadata.get('timeline', 'N/A')}")
            st.markdown(f"**Risks:**\n{st.session_state.metadata.get('risks', 'N/A')}")
        st.markdown('</div>', unsafe_allow_html=True)
        
        # 2. Custom Report Generator
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown("#### 📑 Generate Report")
        report_prompt = st.text_area("What should the report focus on?", placeholder="e.g., Extract all Finance details.")
        
        if st.button("Generate"):
            if report_prompt:
                with st.spinner("Writing..."):
                    prompt = f"Write a report on '{report_prompt}' based on:\n{st.session_state.doc_text}"
                    res = client.chat.completions.create(model="openai/gpt-oss-120b", messages=[{"role": "user", "content": prompt}])
                    report_text = res.choices[0].message.content
                    st.download_button("📥 Download Report", report_text, file_name="report.txt")
        st.markdown('</div>', unsafe_allow_html=True)

        # 3. Downloads
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown("#### 📥 Export")
        st.download_button(
            label="Download Original File",
            data=st.session_state.file_bytes,
            file_name=st.session_state.file_name,
            mime="application/octet-stream",
            use_container_width=True
        )
        st.markdown('</div>', unsafe_allow_html=True)
        
    else:
        st.info("Upload a document to unlock studio features.")
