import streamlit as st
from openai import OpenAI
from pypdf import PdfReader
from langchain_text_splitters import RecursiveCharacterTextSplitter
import json
import datetime

# 1. Page Configuration
st.set_page_config(page_title="TMC Knowledge Lens", page_icon="🔍", layout="wide")

# Custom CSS for a modern, professional look
st.markdown("""
    <style>
    .stApp { background-color: #F8F9FA; }
    .stButton>button {
        background-color: #5B2C6F; color: white; border-radius: 8px;
        font-weight: bold; padding: 10px 24px; transition: 0.3s;
    }
    .stButton>button:hover { background-color: #7D3C98; color: white; }
    .stTextInput>div>div>input { border-radius: 8px; border: 2px solid #5B2C6F; }
    .stTabs [data-baseweb="tab-list"] { gap: 24px; }
    .stTabs [data-baseweb="tab"] {
        background-color: #FFFFFF; border-radius: 8px 8px 0px 0px;
        padding: 10px 20px; border: 1px solid #E0E0E0;
    }
    .stTabs [aria-selected="true"] { background-color: #5B2C6F; color: white; }
    </style>
""", unsafe_allow_html=True)

# Header
col1, col2 = st.columns([1, 8])
with col1:
    st.image("https://tmcltd.com/wp-content/uploads/2025/10/TallyMarks-Consulting-TMC-logo.webp", width=120)
with col2:
    st.title("TMC Knowledge Lens")
    st.markdown("<p style='color: gray; font-size: 1.1em;'>Librarian-Grade AI for Project Documentation</p>", unsafe_allow_html=True)

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

# 3. Sidebar: Upload & Submitter Info
with st.sidebar:
    st.image("https://tmcltd.com/wp-content/uploads/2025/10/TallyMarks-Consulting-TMC-logo.webp", width=150)
    st.markdown("---")
    
    st.header("👤 Submitter Info")
    submitter_name = st.text_input("Your Name", placeholder="e.g., Nadeem Sohail")
    submitter_email = st.text_input("Email", placeholder="e.g., nadeem@tmcltd.com")
    submitter_dept = st.text_input("Department", placeholder="e.g., AI Engineering")
    
    st.divider()
    
    st.header("📄 Upload Document")
    uploaded_file = st.file_uploader("Upload a PDF", type="pdf")

# 4. Process the Uploaded PDF
if uploaded_file is not None:
    if st.session_state.file_name != uploaded_file.name:
        with st.spinner("🔍 AI is reading, cataloging, and summarizing..."):
            st.session_state.file_name = uploaded_file.name
            st.session_state.file_bytes = uploaded_file.getvalue()
            
            # Extract PDF Native Metadata
            pdf_reader = PdfReader(uploaded_file)
            native_meta = pdf_reader.metadata
            native_author = native_meta.author if native_meta and native_meta.author else "Unknown"
            native_date = native_meta.creation_date if native_meta and native_meta.creation_date else "Unknown"
            
            # Extract Text
            raw_text = ""
            for page in pdf_reader.pages:
                raw_text += page.extract_text() + "\n"
            
            text_splitter = RecursiveCharacterTextSplitter(chunk_size=2000, chunk_overlap=200)
            chunks = text_splitter.split_text(raw_text)
            st.session_state.doc_text = "\n".join(chunks[:15]) 
            
            # AI Auto-Tagging & Summary
            sys_prompt = f"""
            You are a librarian AI. Read the text and return a JSON object with these exact keys:
            "client": (e.g., "Tata Motors" or "All"),
            "module": (e.g., "S/4HANA" or "All"),
            "doc_type": (e.g., "Blueprint", "Proposal", or "All"),
            "doc_title": (The real title of the document, inferred from the text),
            "doc_author": (Author if mentioned in text, else use '{native_author}'),
            "summary": "A 2-sentence summary of the project.",
            "key_decisions": "Bullet points of key decisions made.",
            "timeline": "Key dates or phases mentioned.",
            "risks": "Any risks mentioned."
            """
            
            tag_response = client.chat.completions.create(
                model="openai/gpt-oss-120b",
                messages=[{"role": "system", "content": sys_prompt}, {"role": "user", "content": st.session_state.doc_text[:6000]}],
                response_format={ "type": "json_object" }
            )
            
            try:
                meta = json.loads(tag_response.choices[0].message.content)
                meta["native_date"] = native_date
                meta["submitter"] = submitter_name
                meta["email"] = submitter_email
                meta["dept"] = submitter_dept
                meta["upload_date"] = datetime.date.today().strftime("%Y-%m-%d")
                st.session_state.metadata = meta
                st.session_state.summary = meta.get("summary", "")
            except:
                st.error("Failed to parse AI response. Please try another document.")
            
            st.rerun() 

# 5. Main Area: Tabs
if st.session_state.file_name:
    tab1, tab2, tab3 = st.tabs(["📋 Project at a Glance", "💬 Search Q&A", "📑 Custom Report"])
    
    # TAB 1: Project at a Glance
    with tab1:
        st.subheader(f"📄 {st.session_state.metadata.get('doc_title', st.session_state.file_name)}")
        
        # Metadata Grid
        col_a, col_b, col_c = st.columns(3)
        col_a.metric("Client", st.session_state.metadata.get('client', 'N/A'))
        col_b.metric("SAP Module", st.session_state.metadata.get('module', 'N/A'))
        col_c.metric("Document Type", st.session_state.metadata.get('doc_type', 'N/A'))
        
        st.markdown("---")
        st.info(f"**Summary:** {st.session_state.summary}")
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("#### 🎯 Key Decisions")
            st.write(st.session_state.metadata.get('key_decisions', 'None found.'))
            st.markdown("#### ⚠️ Risks")
            st.write(st.session_state.metadata.get('risks', 'None found.'))
        with col2:
            st.markdown("#### 🗓️ Timeline")
            st.write(st.session_state.metadata.get('timeline', 'None found.'))
            st.markdown("#### 👤 Document Info")
            st.write(f"**Author:** {st.session_state.metadata.get('doc_author', 'Unknown')}")
            st.write(f"**Date:** {st.session_state.metadata.get('native_date', 'Unknown')}")
            st.write(f"**Submitted By:** {st.session_state.metadata.get('submitter', 'N/A')} ({st.session_state.metadata.get('dept', 'N/A')})")

        # Download original PDF
        st.download_button(
            label="📥 Download Original PDF",
            data=st.session_state.file_bytes,
            file_name=st.session_state.file_name,
            mime="application/pdf"
        )

    # TAB 2: Search Q&A
    with tab2:
        st.subheader("Ask the Knowledge Lens")
        with st.form(key="search_form"):
            query = st.text_input("Ask a question about the uploaded document...", placeholder="e.g., What is the cutover strategy?")
            submit_button = st.form_submit_button("Search")

        if submit_button and query:
            with st.spinner("Analyzing..."):
                prompt = f"Use the document excerpts to answer the question. If not found, say 'I don't know'.\n\nDocument:\n{st.session_state.doc_text}\n\nQuestion: {query}"
                response = client.chat.completions.create(
                    model="openai/gpt-oss-120b",
                    messages=[{"role": "system", "content": "You are a precise assistant."}, {"role": "user", "content": prompt}],
                    temperature=0.1
                )
                answer = response.choices[0].message.content
                
                st.success("✅ **Answer found:**")
                st.markdown(answer)
                st.caption(f"Source Citation: `{st.session_state.file_name}`")
                
                # Download Answer
                st.download_button(
                    label="📥 Download Answer as Text",
                    data=f"Question: {query}\n\nAnswer:\n{answer}",
                    file_name="Knowledge_Lens_Answer.txt",
                    mime="text/plain"
                )

    # TAB 3: Custom Report
    with tab3:
        st.subheader("Generate Custom Report")
        st.write("Select specific aspects of the document you want to extract into a report.")
        
        report_prompt = st.text_area(
            "What should the report focus on?",
            placeholder="e.g., Extract all mentions of 'Finance' and 'Cutover' into a structured summary."
        )
        
        if st.button("Generate Report"):
            if report_prompt:
                with st.spinner("Generating custom report..."):
                    prompt = f"Based on the document, generate a detailed report focusing on: {report_prompt}\n\nDocument:\n{st.session_state.doc_text}"
                    response = client.chat.completions.create(
                        model="openai/gpt-oss-120b",
                        messages=[{"role": "system", "content": "You are a professional report writer."}, {"role": "user", "content": prompt}],
                        temperature=0.2
                    )
                    report_text = response.choices[0].message.content
                    st.markdown(report_text)
                    
                    st.download_button(
                        label="📥 Download Report",
                        data=report_text,
                        file_name="Custom_Report.txt",
                        mime="text/plain"
                    )
            else:
                st.warning("Please enter a prompt for the report.")
else:
    st.info("👈 Please upload a PDF document in the sidebar to begin.")
