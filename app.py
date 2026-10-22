import streamlit as st
from openai import OpenAI
import pypdf
import docx
from pptx import Presentation
from langchain_text_splitters import RecursiveCharacterTextSplitter
import json
import io

# 1. Page Configuration
st.set_page_config(page_title="TMC Knowledge Lens", page_icon="🔍", layout="wide")

# Custom CSS for a sleek, modern look
st.markdown("""
    <style>
    /* Main App Styling */
    .stApp { background-color: #F8F9FA; font-family: 'Inter', sans-serif; }
    header {visibility: hidden;}
    
    /* Modern Cards */
    .card {
        background-color: white; padding: 20px; border-radius: 12px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05); margin-bottom: 15px;
        border: 1px solid #E5E7EB;
    }
    
    /* Buttons */
    .stButton>button {
        background-color: #5B2C6F; color: white; border-radius: 8px;
        font-weight: 600; padding: 8px 16px; width: 100%;
        transition: 0.2s; border: none;
    }
    .stButton>button:hover { background-color: #7D3C98; color: white; }
    
    /* Chat Input */
    .stTextInput>div>div>input { 
        border-radius: 20px; border: 1px solid #D1D5DB; padding: 10px 20px;
    }
    
    /* Summary Banner */
    .summary-banner {
        background-color: #EBF5FB;
        border-left: 5px solid #2980B9;
        padding: 15px;
        border-radius: 8px;
        margin-bottom: 20px;
    }
    </style>
""", unsafe_allow_html=True)

# 2. Initialize Groq API Client
try:
    client = OpenAI(base_url="https://api.groq.com/openai/v1", api_key=st.secrets["GITHUB_TOKEN"])
except Exception as e:
    st.error(f"API Key error: {e}")
    st.stop()

# 3. Session State Initialization
if "documents" not in st.session_state: st.session_state.documents = []
if "chat_history" not in st.session_state: st.session_state.chat_history = []
if "active_summary" not in st.session_state: st.session_state.active_summary = ""

# --- TEXT EXTRACTION HELPERS ---
def extract_pdf(file_bytes):
    reader = pypdf.PdfReader(io.BytesIO(file_bytes))
    return "\n".join([(page.extract_text() or "") for page in reader.pages]), getattr(reader.metadata, 'author', 'Unknown')
def extract_docx(file_bytes):
    doc = docx.Document(io.BytesIO(file_bytes))
    return "\n".join([para.text for para in doc.paragraphs]), "Unknown"
def extract_pptx(file_bytes):
    prs = Presentation(io.BytesIO(file_bytes))
    text = ""
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"): text += shape.text + "\n"
    return text, "Unknown"
def extract_txt(file_bytes):
    return file_bytes.decode("utf-8", errors="ignore"), "Unknown"

# --- HEADER ---
st.markdown("""
    <div style="display: flex; align-items: center; background-color: white; padding: 15px 30px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); margin-bottom: 20px;">
        <img src="https://tmcltd.com/wp-content/uploads/2025/10/TallyMarks-Consulting-TMC-logo.webp" width="80" style="margin-right: 20px;">
        <div>
            <h1 style="color: #5B2C6F; margin: 0; font-size: 2rem;">Knowledge Lens</h1>
            <p style="color: #6B7280; margin: 0; font-size: 1rem;">Enterprise AI Document Workspace</p>
        </div>
    </div>
""", unsafe_allow_html=True)

# --- LAYOUT ---
col_sources, col_center, col_studio = st.columns([1.2, 2.5, 1.2])

# ==========================================
# LEFT: SOURCES PANEL
# ==========================================
with col_sources:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### 📁 Sources")
    
    uploaded_files = st.file_uploader("Add sources", type=["pdf", "docx", "pptx", "txt"], accept_multiple_files=True, label_visibility="collapsed")
    
    if uploaded_files:
        for file in uploaded_files:
            # Check if file is already in session state
            if not any(d['name'] == file.name for d in st.session_state.documents):
                with st.spinner(f"Processing {file.name}..."):
                    bytes_data = file.getvalue()
                    ext = file.name.split(".")[-1].lower()
                    
                    if ext == "pdf": text, author = extract_pdf(bytes_data)
                    elif ext == "docx": text, author = extract_docx(bytes_data)
                    elif ext == "pptx": text, author = extract_pptx(bytes_data)
                    else: text, author = extract_txt(bytes_data)
                    
                    # Auto-Tagging
                    res = client.chat.completions.create(
                        model="openai/gpt-oss-120b",
                        messages=[
                            {"role": "system", "content": "Return JSON: {'client': 'Name', 'doc_type': 'Type', 'doc_title': 'Title', 'summary': '2 sentences'}"},
                            {"role": "user", "content": text[:4000]}
                        ],
                        response_format={ "type": "json_object" }
                    )
                    meta = json.loads(res.choices[0].message.content)
                    
                    st.session_state.documents.append({
                        "name": file.name,
                        "bytes": bytes_data,
                        "text": text,
                        "meta": meta,
                        "author": author
                    })
                    
        st.rerun() # Refresh to show the new files cleanly

    # Display Document Selection
    if st.session_state.documents:
        st.markdown("---")
        st.caption("Select sources to search from:")
        
        selected_docs = []
        for i, doc in enumerate(st.session_state.documents):
            col_chk, col_info = st.columns([1, 8])
            with col_chk:
                if st.checkbox("", key=f"chk_{i}", value=True):
                    selected_docs.append(doc)
            with col_info:
                st.markdown(f"**📄 {doc['meta'].get('doc_title', doc['name'])}**")
                st.caption(f"{doc['meta'].get('doc_type', 'Document')} • {doc['meta'].get('client', 'Unknown')}")
        
        if st.button("Load / Refresh Summary"):
            # Generate a combined summary of all selected docs
            combined_text = "\n".join([d['text'][:2000] for d in selected_docs])
            res = client.chat.completions.create(
                model="openai/gpt-oss-120b",
                messages=[{"role": "user", "content": f"Summarize these documents in 3 sentences: {combined_text}"}]
            )
            st.session_state.active_summary = res.choices[0].message.content
            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)


# ==========================================
# CENTER: SUMMARY & CHAT
# ==========================================
with col_center:
    # 1. Top Summary Banner
    if st.session_state.active_summary:
        st.markdown(f'''
            <div class="summary-banner">
                <h3 style="margin-top: 0; color: #2980B9;">📌 Project at a Glance</h3>
                <p style="font-size: 1.1rem; margin-bottom: 0;">{st.session_state.active_summary}</p>
            </div>
        ''', unsafe_allow_html=True)
    else:
        st.info("👈 Upload documents and click 'Load / Refresh Summary' to see the Project at a Glance.")

    # 2. Chat Interface
    st.markdown("### 💬 Ask the Lens")
    
    for chat in st.session_state.chat_history:
        with st.chat_message("user"): st.write(chat["question"])
        with st.chat_message("assistant"):
            st.write(chat["answer"])
            st.caption(f"Source(s): {', '.join(chat['sources'])}")

    with st.form(key="chat_form", clear_on_submit=True):
        query = st.text_input("Ask a question...", placeholder="e.g., Compare the blueprints...", label_visibility="collapsed")
        submit = st.form_submit_button("Send")
        
        if submit and query:
            if not selected_docs:
                st.warning("Please select at least one source document.")
            else:
                with st.spinner("Searching across selected documents..."):
                    # Combine text from ONLY selected documents
                    search_context = "\n\n".join([d['text'] for d in selected_docs])
                    # Truncate if too long for Groq
                    search_context = search_context[:30000] 
                    
                    prompt = f"Answer based ONLY on these documents:\n\n{search_context}\n\nQuestion: {query}"
                    response = client.chat.completions.create(
                        model="openai/gpt-oss-120b",
                        messages=[{"role": "user", "content": prompt}],
                        temperature=0.1
                    )
                    answer = response.choices[0].message.content
                    
                    st.session_state.chat_history.append({
                        "question": query, 
                        "answer": answer, 
                        "sources": [d['name'] for d in selected_docs]
                    })
                    st.rerun()


# ==========================================
# RIGHT: STUDIO
# ==========================================
with col_studio:
    st.markdown("### 🎛️ Studio")
    
    if st.session_state.documents:
        # 1. Reports
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown("#### 📑 Generate Report")
        report_prompt = st.text_area("Focus?", placeholder="e.g., Extract all risks.")
        
        if st.button("Generate"):
            if report_prompt and selected_docs:
                with st.spinner("Writing..."):
                    search_context = "\n\n".join([d['text'] for d in selected_docs])[:30000]
                    prompt = f"Write a report on '{report_prompt}' based on:\n{search_context}"
                    res = client.chat.completions.create(model="openai/gpt-oss-120b", messages=[{"role": "user", "content": prompt}])
                    st.download_button("📥 Download Report", res.choices[0].message.content, file_name="report.txt")
        st.markdown('</div>', unsafe_allow_html=True)

        # 2. Detailed Metadata View
        if selected_docs:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown("#### 🏷️ Source Metadata")
            for doc in selected_docs:
                with st.expander(f"📄 {doc['meta'].get('doc_title', doc['name'])}"):
                    st.write(f"**Client:** {doc['meta'].get('client', 'N/A')}")
                    st.write(f"**Type:** {doc['meta'].get('doc_type', 'N/A')}")
                    st.write(f"**Author:** {doc['author']}")
                    st.download_button("Download File", doc['bytes'], file_name=doc['name'])
            st.markdown('</div>', unsafe_allow_html=True)
