import streamlit as st
from openai import OpenAI
from pypdf import PdfReader
from langchain_text_splitters import RecursiveCharacterTextSplitter
import json

# 1. Page Configuration (Wide layout, custom icon, title)
st.set_page_config(
    page_title="TMC Knowledge Lens", 
    page_icon="🔍", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for a modern, branded look
st.markdown("""
    <style>
    /* Main background */
    .stApp {
        background-color: #F8F9FA;
    }
    /* Primary buttons */
    .stButton>button {
        background-color: #5B2C6F; /* TMC Purple */
        color: white;
        border-radius: 8px;
        font-weight: bold;
        padding: 10px 24px;
        transition: 0.3s;
    }
    .stButton>button:hover {
        background-color: #7D3C98;
        color: white;
    }
    /* Text input */
    .stTextInput>div>div>input {
        border-radius: 8px;
        border: 2px solid #5B2C6F;
    }
    /* Success and Info boxes */
    .stSuccess {
        background-color: #D5F5E3;
        border-left: 5px solid #27AE60;
    }
    .stInfo {
        background-color: #EBF5FB;
        border-left: 5px solid #2980B9;
    }
    </style>
""", unsafe_allow_html=True)

# Header with Logo
col1, col2 = st.columns([1, 8])
with col1:
    st.image("https://tmcltd.com/wp-content/uploads/2025/10/TallyMarks-Consulting-TMC-logo.webp", width=120)
with col2:
    st.title("TMC Knowledge Lens")
    st.markdown("<p style='color: gray; font-size: 1.1em;'>Librarian-Grade AI for Project Documentation</p>", unsafe_allow_html=True)

# 2. Initialize Groq API Client
try:
    client = OpenAI(
        base_url="https://api.groq.com/openai/v1",
        api_key=st.secrets["GITHUB_TOKEN"],
    )
except Exception as e:
    st.error(f"API Key error: {e}")
    st.stop()

# Initialize Session State
if "auto_client" not in st.session_state: st.session_state.auto_client = "All"
if "auto_module" not in st.session_state: st.session_state.auto_module = "All"
if "auto_type" not in st.session_state: st.session_state.auto_type = "All"
if "project_summary" not in st.session_state: st.session_state.project_summary = ""

# 3. Sidebar
with st.sidebar:
    st.image("https://tmcltd.com/wp-content/uploads/2025/10/TallyMarks-Consulting-TMC-logo.webp", width=150)
    st.markdown("---")
    
    st.header("📄 1. Upload Document")
    uploaded_file = st.file_uploader("Upload a PDF to the Knowledge Lens", type="pdf")
    
    st.markdown("---")
    
    st.header("🏷️ 2. Auto-Extracted Metadata")
    st.caption("The AI automatically tags your document:")
    
    # Sidebar Dropdowns
    client_options = ["All", "Tata Motors", "TML", "Other"]
    client_index = client_options.index(st.session_state.auto_client) if st.session_state.auto_client in client_options else 0
    client_name = st.selectbox("Client", client_options, index=client_index)
    
    module_options = ["All", "FI", "CO", "MM", "SD", "S/4HANA"]
    module_index = module_options.index(st.session_state.auto_module) if st.session_state.auto_module in module_options else 0
    module = st.selectbox("SAP Module", module_options, index=module_index)
    
    type_options = ["All", "Blueprint", "BRD", "Proposal", "Cutover Plan"]
    type_index = type_options.index(st.session_state.auto_type) if st.session_state.auto_type in type_options else 0
    doc_type = st.selectbox("Document Type", type_options, index=type_index)

# 4. Process the Uploaded PDF (The "Wow" Factor)
document_text = ""
if uploaded_file is not None:
    if "processed_file" not in st.session_state or st.session_state.processed_file != uploaded_file.name:
        with st.spinner("🔍 AI is reading, cataloging, and summarizing the document..."):
            pdf_reader = PdfReader(uploaded_file)
            raw_text = ""
            for page in pdf_reader.pages:
                raw_text += page.extract_text() + "\n"
                
            text_splitter = RecursiveCharacterTextSplitter(chunk_size=2000, chunk_overlap=200)
            chunks = text_splitter.split_text(raw_text)
            document_text = "\n".join(chunks[:15]) # Read more chunks
            
            # AI Auto-Tagging
            sys_prompt = """
            You are a librarian AI. Read the text and return a JSON object with exactly these keys:
            "client": (guess from text, must be one of: "Tata Motors", "TML", "Other"),
            "module": (guess from text, must be one of: "FI", "CO", "MM", "SD", "S/4HANA", or "All"),
            "doc_type": (guess from text, must be one of: "Blueprint", "BRD", "Proposal", "Cutover Plan", or "All"),
            "summary": "A 3-sentence 'Project at a Glance' summary detailing the main objective."
            """
            
            tag_response = client.chat.completions.create(
                model="openai/gpt-oss-120b",
                messages=[{"role": "system", "content": sys_prompt}, {"role": "user", "content": document_text[:6000]}],
                response_format={ "type": "json_object" }
            )
            
            try:
                metadata = json.loads(tag_response.choices[0].message.content)
                st.session_state.auto_client = metadata.get("client", "All")
                st.session_state.auto_module = metadata.get("module", "All")
                st.session_state.auto_type = metadata.get("doc_type", "All")
                st.session_state.project_summary = metadata.get("summary", "No summary generated.")
            except:
                st.session_state.project_summary = "Could not auto-generate summary."
                
            st.session_state.processed_file = uploaded_file.name
            st.rerun() 

    # Display the "At a Glance" Banner
    if st.session_state.project_summary:
        st.info(f"📄 **Project at a Glance:** {st.session_state.project_summary}")
        st.success(f"✅ **Document Cataloged:** Client={st.session_state.auto_client} | Module={st.session_state.auto_module} | Type={st.session_state.auto_type}")

# 5. Q&A Section
st.markdown("<br>", unsafe_allow_html=True)
st.markdown("### 💬 Ask the Knowledge Lens")
with st.form(key="search_form", clear_on_submit=False):
    query = st.text_input("Ask a question about the uploaded document...", placeholder="e.g., What are the key risks mentioned in this blueprint?")
    submit_button = st.form_submit_button("Search Knowledge Base")

# 6. Handle Submission
if submit_button and query:
    if not uploaded_file:
        st.warning("⚠️ Please upload a PDF document first in the sidebar.")
    else:
        with st.spinner("Analyzing..."):
            try:
                prompt = f"""
                You are an AI assistant for TMC Consulting. 
                Use the document excerpts to answer the question. 
                If the answer is not in the document, say "I don't know based on the provided document."
                Format the answer clearly using bullet points if necessary.
                
                Document Excerpts:
                {document_text}
                
                Question: {query}
                """
                
                response = client.chat.completions.create(
                    model="openai/gpt-oss-120b", 
                    messages=[
                        {"role": "system", "content": "You are a precise, professional knowledge management assistant."},
                        {"role": "user", "content": prompt}
                    ],
                    temperature=0.1
                )
                
                answer = response.choices[0].message.content
                st.success("✅ **Answer found:**")
                st.markdown(answer)
                st.caption(f"Source Citation: `{uploaded_file.name}`")
                    
            except Exception as e:
                st.error(f"Error: {e}")
