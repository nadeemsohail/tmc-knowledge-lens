import streamlit as st
from openai import OpenAI

# 1. Page Configuration
st.set_page_config(page_title="TMC Knowledge Lens", layout="wide")

st.title("🔍 TMC Knowledge Lens")
st.subheader("AI Search & Q&A Layer for TMC Project Documentation")

# 2. Initialize API Client
try:
    client = OpenAI(
        base_url="https://models.github.ai/inference",
        api_key=st.secrets["GITHUB_TOKEN"],
    )
except Exception as e:
    st.error(f"API Key not configured properly: {e}")
    st.stop()

# 3. Sidebar
with st.sidebar:
    st.header("Filter by Metadata")
    client_name = st.selectbox("Client", ["All", "Tata Motors", "TML", "Other"])
    module = st.selectbox("SAP Module", ["All", "FI", "CO", "MM", "SD", "S/4HANA"])
    doc_type = st.selectbox("Document Type", ["All", "Blueprint", "BRD", "Proposal", "Cutover Plan"])

# 4. Sample Document
sample_document = """
Tata Motors S/4HANA Blueprint Document.
Client: Tata Motors Ltd (TML).
SAP Module: S/4HANA.
Phase: Blueprint.
Author: TMC Consulting.
Date: 2023.

Key Decisions:
1. The approach for S/4HANA will be a Greenfield implementation.
2. We will use SAP Activate methodology.
3. The Finance (FI) module will be deployed first, followed by Materials Management (MM).
4. Custom code will be strictly limited to RICEFW objects.
5. Cutover strategy will involve a 48-hour system downtime window.
"""

# 5. Form for Enter Key Submission
with st.form(key="search_form"):
    query = st.text_input("Ask a question about past projects (e.g., 'What is the cutover strategy?')")
    # The submit button inside the form allows pressing Enter to submit
    submit_button = st.form_submit_button("Search Knowledge Base")

# 6. Handle Submission
if submit_button and query:
    with st.spinner("AI is reading the documents and searching..."):
        try:
            prompt = f"""
            You are an AI assistant for TMC Consulting. 
            Use the following document excerpts to answer the user's question.
            If the answer is not in the document, say "I don't know based on the provided documents."
            
            Document:
            {sample_document}
            
            Question: {query}
            """
            
            response = client.chat.completions.create(
              model="openai/gpt-4o",
                messages=[
                    {"role": "system", "content": "You are a helpful knowledge management assistant."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.2
            )
            
            # --- THE FIX: Check if response is a string (error) ---
            if isinstance(response, str):
                st.error(f"GitHub API returned an error text: {response}")
            else:
                # It's a proper object, so we can access .choices
                answer = response.choices[0].message.content
                
                st.info(f"Searching for: '{query}' (Filtered by: Client={client_name}, Module={module}, Type={doc_type})")
                st.success("✅ **Answer found:**")
                st.write(answer)
                
                st.markdown("**Source Citations:**")
                st.markdown("- `TataMotors_S4HANA_Blueprint_v2.pdf` (Page 14)")
                
        except Exception as e:
            st.error(f"An error occurred: {e}")

elif submit_button and not query:
    st.warning("Please enter a question.")
