import streamlit as st
import json
import os
from excel_parser import parse_excel, calc_rags, DIM_ORDER
from rfp_reader import read_docx
from ai_engine import generate_content
from ppt_builder import build_ppt

# ── Page Config ──
st.set_page_config(
    page_title="Deal Deliverability Review Generator",
    page_icon="🔒",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Custom CSS ──
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1A1A2E 0%, #2D1B4E 100%);
        padding: 2rem 2rem 1.5rem 2rem;
        border-radius: 0 0 12px 12px;
        margin: -1rem -1rem 2rem -1rem;
        border-bottom: 3px solid #A100FF;
    }
    .main-header h1 {
        color: white !important;
        font-size: 1.8rem !important;
        margin-bottom: 0.2rem !important;
    }
    .main-header p {
        color: #9CA3AF !important;
        font-size: 0.9rem !important;
    }
    .accent-text { color: #A100FF; font-weight: 700; letter-spacing: 2px; font-size: 0.7rem; }
    .rag-green { color: #059669; font-weight: 700; }
    .rag-amber { color: #D97706; font-weight: 700; }
    .rag-red { color: #DC2626; font-weight: 700; }
    .dim-card {
        border: 1px solid #E2E4E9;
        border-radius: 8px;
        padding: 1rem;
        margin-bottom: 0.8rem;
        background: white;
    }
    .stButton > button {
        width: 100%;
    }
</style>
""", unsafe_allow_html=True)

# ── Header ──
st.markdown("""
<div class="main-header">
    <p class="accent-text">ACCENTURE SECURITY PRACTICE</p>
    <h1>Deal Deliverability Review Generator</h1>
    <p>Upload Excel assessment + RFP document → AI analyzes both → Download completed PPT deck</p>
</div>
""", unsafe_allow_html=True)

# ── Sidebar: AI Configuration ──
with st.sidebar:
    st.markdown("### ⚙️ AI Configuration")
    
    ai_provider = st.selectbox(
        "AI Provider",
        ["Google Gemini (Free)", "Groq (Free)", "Cohere (Free Trial)"],
        help="Select which AI model to use for content generation"
    )
    
    # Map display names to internal keys
    provider_map = {
        "Google Gemini (Free)": "gemini",
        "Groq (Free)": "groq",
        "Cohere (Free Trial)": "cohere"
    }
    provider_key = provider_map[ai_provider]
    
    # API Key input
    api_key_env = {
        "gemini": "GEMINI_API_KEY",
        "groq": "GROQ_API_KEY",
        "cohere": "COHERE_API_KEY"
    }
    
    # Check if key exists in secrets/env first
    env_key = os.environ.get(api_key_env[provider_key], "")
    secrets_key = ""
    try:
        secrets_key = st.secrets.get(api_key_env[provider_key], "")
    except Exception:
        pass
    
    stored_key = secrets_key or env_key
    
    if stored_key:
        st.success(f"✅ {ai_provider} API key configured")
        api_key = stored_key
    else:
        api_key = st.text_input(
            f"{ai_provider} API Key",
            type="password",
            help=f"Enter your API key. Get one free from the provider's website."
        )
    
    st.markdown("---")
    st.markdown("### 📋 How to get free API keys")
    st.markdown("""
    **Google Gemini:**  
    → [aistudio.google.com](https://aistudio.google.com)  
    → Click "Get API Key" → Create
    
    **Groq:**  
    → [console.groq.com](https://console.groq.com)  
    → Sign up → API Keys → Create
    
    **Cohere:**  
    → [dashboard.cohere.com](https://dashboard.cohere.com)  
    → Sign up → API Keys
    """)
    
    st.markdown("---")
    st.markdown("### 📊 About")
    st.markdown("""
    This tool automates the creation of Deal Deliverability Review 
    PowerPoint decks for the Accenture Security Practice.
    
    **Built by:** Raghad Altawil  
    **Version:** 1.0
    """)


# ── Main Content ──
col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📊 Filled Excel Assessment")
    xl_file = st.file_uploader(
        "Upload the completed .xlsx or .xlsm file",
        type=["xlsx", "xlsm"],
        key="excel_upload"
    )

with col2:
    st.markdown("### 📄 RFP Document")
    rfp_file = st.file_uploader(
        "Upload the RFP .docx file",
        type=["docx"],
        key="rfp_upload"
    )

# ── Process ──
if xl_file and rfp_file:
    st.markdown("---")
    
    # Parse button
    if st.button("⚡ Analyze & Generate Deliverability Review", type="primary"):
        
        if not api_key:
            st.error("❌ Please enter your API key in the sidebar first.")
            st.stop()
        
        # Step 1: Parse Excel
        with st.status("🔄 Processing...", expanded=True) as status:
            st.write("📊 Parsing Excel assessment...")
            try:
                qs, ov, risks = parse_excel(xl_file)
                if not qs:
                    st.error("❌ No data found in 02_Assessment sheet. Check your Excel file.")
                    st.stop()
                rags = calc_rags(qs)
                st.write(f"✅ Found {len(qs)} questions across 5 dimensions")
            except Exception as e:
                st.error(f"❌ Excel parsing failed: {e}")
                st.stop()
            
            # Step 2: Read RFP
            st.write("📄 Reading RFP document...")
            try:
                rfp_text = read_docx(rfp_file)
                if len(rfp_text) < 50:
                    st.error("❌ Could not extract RFP text. Ensure it's a valid .docx file.")
                    st.stop()
                st.write(f"✅ Extracted {len(rfp_text):,} characters from RFP")
            except Exception as e:
                st.error(f"❌ RFP reading failed: {e}")
                st.stop()
            
            # Step 3: AI Generation
            st.write(f"🤖 AI analyzing both documents ({ai_provider})...")
            try:
                ai_result = generate_content(
                    provider=provider_key,
                    api_key=api_key,
                    questions=qs,
                    overview=ov,
                    risks=risks,
                    rags=rags,
                    rfp_text=rfp_text
                )
                st.write("✅ AI content generated successfully")
            except Exception as e:
                st.error(f"❌ AI generation failed: {e}")
                st.stop()
            
            # Step 4: Build PPT
            st.write("📑 Building PowerPoint deck...")
            try:
                ppt_bytes = build_ppt(ai_result, rags, risks)
                st.write("✅ PowerPoint generated successfully")
            except Exception as e:
                st.error(f"❌ PPT generation failed: {e}")
                st.stop()
            
            status.update(label="✅ Complete!", state="complete", expanded=False)
        
        # Store results in session state
        st.session_state['ai_result'] = ai_result
        st.session_state['rags'] = rags
        st.session_state['ppt_bytes'] = ppt_bytes
        st.session_state['qs'] = qs

# ── Display Results ──
if 'ai_result' in st.session_state:
    rags = st.session_state['rags']
    ai_result = st.session_state['ai_result']
    ppt_bytes = st.session_state['ppt_bytes']
    
    st.markdown("---")
    
    # Download button — prominent at top
    st.download_button(
        label="📥 Download Deliverability Review PPT",
        data=ppt_bytes,
        file_name="Deal_Deliverability_Review.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        type="primary"
    )
    
    # RAG Dashboard
    st.markdown("### 🎯 RAG Scoring")
    
    rag_colors = {"GREEN": "🟢", "AMBER": "🟡", "RED": "🔴"}
    overall = rags['overall']
    decision = {
        "GREEN": "Ready to Proceed",
        "AMBER": "Proceed with Caution", 
        "RED": "Do Not Proceed — Escalation Required"
    }
    
    rag_col1, rag_col2 = st.columns([1, 3])
    with rag_col1:
        st.metric("Overall RAG", f"{rag_colors[overall]} {overall}")
    with rag_col2:
        st.info(f"**Decision:** {decision[overall]}")
    
    # Dimension RAGs
    dim_cols = st.columns(5)
    for i, dim in enumerate(DIM_ORDER):
        dr = rags['dimRags'][dim]
        with dim_cols[i]:
            short_name = dim.split(" & ")[0] if " & " in dim else dim.split(" ")[0]
            st.metric(f"D{i+1}", f"{rag_colors[dr]} {dr}")
    
    # AI Content Preview
    st.markdown("### 📝 Generated Content Preview")
    
    with st.expander("Key Justification", expanded=True):
        st.write(ai_result.get('key_justification', ''))
        st.caption(f"Opportunity Value: **{ai_result.get('opportunity_value', '')}**")
    
    with st.expander("Deal Overview"):
        for line in ai_result.get('deal_overview', []):
            st.write(line)
    
    with st.expander("Positive Notes"):
        for note in ai_result.get('positive_notes', []):
            st.write(f"• {note}")
    
    with st.expander("Dimension Analysis"):
        for i, dim in enumerate(ai_result.get('dimensions', [])):
            dim_name = DIM_ORDER[i] if i < len(DIM_ORDER) else dim.get('name', '')
            dr = rags['dimRags'].get(dim_name, 'GREEN')
            st.markdown(f"**{rag_colors[dr]} Dimension {i+1}: {dim_name}**")
            for bullet in dim.get('bullets', []):
                st.write(f"  • {bullet}")
            st.caption(f"Comments: {dim.get('comments', '')}")
            st.markdown("---")
    
    with st.expander("Critical Blockers & Amber Actions"):
        col_r, col_a = st.columns(2)
        with col_r:
            st.markdown("**🔴 Critical Blockers**")
            st.write(ai_result.get('red_summary', 'None identified'))
        with col_a:
            st.markdown("**🟡 Amber Actions**")
            st.write(ai_result.get('amber_summary', 'None identified'))
    
    with st.expander("Assumptions & Next Steps"):
        col_as, col_ns = st.columns(2)
        with col_as:
            st.markdown("**Key Assumptions**")
            for i, a in enumerate(ai_result.get('assumptions', [])):
                st.write(f"{i+1}. {a}")
        with col_ns:
            st.markdown("**Next Steps**")
            for ns in ai_result.get('next_steps', []):
                st.write(f"**{ns['title']}** — {ns['desc']}")
                st.caption(f"Owner: {ns['owner']}")
    
    # Copy JSON button
    with st.expander("📋 Raw JSON (for debugging)"):
        st.json({**ai_result, 'calculated_rags': rags})

elif xl_file and rfp_file:
    st.info("👆 Click the button above to start the analysis.")
elif xl_file or rfp_file:
    st.warning("⚠️ Please upload both files to proceed.")
else:
    st.markdown("""
    ### 👋 Getting Started
    
    1. **Upload** your filled Excel assessment (.xlsx/.xlsm) on the left
    2. **Upload** the RFP document (.docx) on the right  
    3. **Configure** your AI provider and API key in the sidebar
    4. **Click** "Analyze & Generate" — that's it!
    
    The tool will parse both documents, run AI analysis across all 5 deliverability 
    dimensions, and generate a complete PowerPoint deck ready for leadership review.
    """)
