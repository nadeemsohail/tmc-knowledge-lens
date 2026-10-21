import streamlit as st
from openai import OpenAI
from pypdf import PdfReader
from langchain_text_splitters import RecursiveCharacterTextSplitter

# 1. Page Configuration
st.set_page_config(page_title="TMC Knowledge Lens", layout="wide")

st.title("🔍 TMC Knowledge Lens")
st.subheader("AI Search & Q&A Layer for TMC Project Documentation")

# 2. Initialize Groq API Client
try:
    client = OpenAI(
        base_url="https://api.groq.com/openai/v1",
        api_key=st.secrets["GITHUB_TOKEN"],
    )
except Exception as e:
    st.error(f"API Key not configured properly: {e}")
    st.stop()

# 3. Sidebar: Metadata Filters & Document Upload
with st.sidebar:
    st.header("1. Upload Document")
    uploaded_file = st.file_uploader("Upload a project PDF", type="pdf")
    
    st.divider()
    
    st.header("2. Filter by Metadata")
    client_name = st.selectbox("Client", ["All", "Tata Motors", "TML", "Other"])
    module = st.selectbox("SAP Module", ["All", "FI", "CO", "MM", "SD", "S/4HANA"])
    doc_type = st.selectbox("Document Type", ["All", "Blueprint", "BRD", "Proposal", "Cutover Plan"])

# 4. Process the Uploaded PDF
document_text = ""
if uploaded_file is not None:
    with st.spinner("Reading and processing PDF..."):
        # Read PDF
        pdf_reader = PdfReader(uploaded_file)
        raw_text = ""
        for page in pdf_reader.pages:
            raw_text += page.extract_text() + "\n"
            
        # Split into chunks (useful for large documents, though we pass it all for now)
        text_splitter = RecursiveCharacterTextSplitter(
            chunk_size=2000,
            chunk_overlap=200,
            length_function=len,
        )
        chunks = text_splitter.split_text(raw_text)
        
        # For this pilot, we'll join the first few chunks to fit in the prompt
        document_text = "\n".join(chunks[:10]) 
        
        st.success(f"✅ Document loaded! Extracted {len(pdf_reader.pages)} pages.")

# 5. Form for Enter Key Submission
st.markdown("### Ask a question")
with st.form(key="search_form"):
    query = st.text_input("Ask a question about the uploaded document (e.g., 'What is the cutover strategy?')")
    submit_button = st.form_submit_button("Search Knowledge Base")

# 6. Handle Submission
if submit_button and query:
    if not uploaded_file:
        st.warning("⚠️ Please upload a PDF document first in the sidebar.")
    else:
        with st.spinner("AI is reading the document and searching..."):
            try:
                prompt = f"""
                You are an AI assistant for TMC Consulting. 
                Use the following document excerpts to answer the user's question.
                If the answer is not in the document, say "I don't know based on the provided documents."
                
                Document Excerpts:
                {document_text}
                
                Question: {query}
                """
                
                # Call Groq API
                response = client.chat.completions.create(
                    model="openai/gpt-oss-120b", 
                    messages=[
                        {"role": "system", "content": "You are a helpful knowledge management assistant."},
                        {"role": "user", "content": prompt}
                    ],
                    temperature=0.2
                )
                
                # Extract answer
                answer = response.choices[0].message.content
                
                # Display results
                st.info(f"Searching for: '{query}' (Filtered by: Client={client_name}, Module={module}, Type={doc_type})")
                st.success("✅ **Answer found:**")
                st.write(answer)
                
                st.markdown("**Source Citations:**")
                st.markdown(f"- `{uploaded_file.name}`")
                    
            except Exception as e:
                st.error(f"An error occurred: {e}")

elif submit_button and not query:
    st.warning("Please enter a question.")
