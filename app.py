import streamlit as st
from openai import OpenAI

st.set_page_config(page_title="TMC Knowledge Lens", layout="wide")

st.title("🔍 TMC Knowledge Lens")
st.subheader("AI Search & Q&A Layer for TMC Project Documentation")

# Initialize the OpenAI client using GitHub Models
# We pull the token safely from Streamlit Secrets
try:
    client = OpenAI(
        base_url="https://models.inference.ai.azure.com",
        api_key=st.secrets["GITHUB_TOKEN"],
    )
except Exception as e:
    st.error(f"API Key not configured properly: {e}")
    st.stop()

# Sidebar for metadata filtering (Taxonomy)
with st.sidebar:
    st.header("Filter by Metadata")
    client_name = st.selectbox("Client", ["All", "Tata Motors", "TML", "Other"])
    module = st.selectbox("SAP Module", ["All", "FI", "CO", "MM", "SD", "S/4HANA"])
    doc_type = st.selectbox("Document Type", ["All", "Blueprint", "BRD", "Proposal", "Cutover Plan"])

# A sample document to act as our "Knowledge Base" for this test
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

# Main search interface
query = st.text_input("Ask a question about past projects (e.g., 'What is the cutover strategy?')")

if st.button("Search Knowledge Base"):
    if query:
        with st.spinner("AI is reading the documents and searching..."):
            try:
                # Create the prompt for the AI
                prompt = f"""
                You are an AI assistant for TMC Consulting. 
                Use the following document excerpts to answer the user's question.
                If the answer is not in the document, say "I don't know based on the provided documents."
                
                Document:
                {sample_document}
                
                Question: {query}
                """
                
                # Call GitHub Models (using GPT-4o-mini for speed)
                response = client.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=[
                        {"role": "system", "content": "You are a helpful knowledge management assistant."},
                        {"role": "user", "content": prompt}
                    ],
                    temperature=0.2
                )
                
                answer = response.choices[0].message.content
                
                # Display results
                st.info(f"Searching for: '{query}' (Filtered by: Client={client_name}, Module={module}, Type={doc_type})")
                st.success("✅ **Answer found:**")
                st.write(answer)
                
                st.markdown("**Source Citations:**")
                st.markdown("- `TataMotors_S4HANA_Blueprint_v2.pdf` (Page 14)")
                
            except Exception as e:
                st.error(f"An error occurred: {e}")
    else:
        st.warning("Please enter a question.")
