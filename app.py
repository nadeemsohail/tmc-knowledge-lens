import streamlit as st

st.set_page_config(page_title="TMC Knowledge Lens", layout="wide")

st.title("🔍 TMC Knowledge Lens")
st.subheader("AI Search & Q&A Layer for TMC Project Documentation")

st.markdown("""
Welcome to the **TMC Knowledge Lens** pilot interface. 
This tool uses library-grade taxonomy to search and summarize project documents.
""")

# Sidebar for metadata filtering (Taxonomy)
with st.sidebar:
    st.header("Filter by Metadata")
    client = st.selectbox("Client", ["All", "Tata Motors", "TML", "Other"])
    module = st.selectbox("SAP Module", ["All", "FI", "CO", "MM", "SD", "S/4HANA"])
    doc_type = st.selectbox("Document Type", ["All", "Blueprint", "BRD", "Proposal", "Cutover Plan"])

# Main search interface
query = st.text_input("Ask a question about past projects (e.g., 'What was the blueprint approach for S/4HANA?')")

if st.button("Search Knowledge Base"):
    if query:
        st.info(f"Searching for: '{query}' (Filtered by: Client={client}, Module={module}, Type={doc_type})")
        st.success("✅ **Answer found:** (This is a placeholder. The AI backend will be connected here.)")
        st.markdown("**Source Citations:**")
        st.markdown("- [TataMotors_S4HANA_Blueprint_v2.pdf, Page 14](https://example.com)")
    else:
        st.warning("Please enter a question.")
