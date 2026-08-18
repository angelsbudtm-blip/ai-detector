import streamlit as st
from pptx import Presentation
from pypdf import PdfReader
import psycopg2
import re
import math
from difflib import SequenceMatcher

# ----------------------------------------------------
# Database Connection (Supabase PostgreSQL via Method 2)
# ----------------------------------------------------
def get_db_connection():
    """Establishes connection to Supabase using separate secret keys."""
    pg = st.secrets["postgres"]
    return psycopg2.connect(
        host=pg["host"],
        port=pg["port"],
        dbname=pg["dbname"],
        user=pg["user"],
        password=pg["password"]
    )

def init_db():
    """Creates database tables in Supabase if they don't exist yet."""
    conn = get_db_connection()
    c = conn.cursor()
    # Table for uploaded files
    c.execute('''
        CREATE TABLE IF NOT EXISTS presentations (
            id SERIAL PRIMARY KEY,
            filename TEXT,
            file_type TEXT,
            total_slides INTEGER,
            overall_ai_score REAL,
            upload_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    ''')
    # Table for individual slide/page contents
    c.execute('''
        CREATE TABLE IF NOT EXISTS slide_records (
            id SERIAL PRIMARY KEY,
            presentation_id INTEGER REFERENCES presentations(id) ON DELETE CASCADE,
            slide_number INTEGER,
            extracted_text TEXT,
            ai_score REAL
        );
    ''')
    conn.commit()
    c.close()
    conn.close()

# Initialize DB structure on startup
try:
    init_db()
except Exception as e:
    st.error(f"Database connection error: {e}")

# ----------------------------------------------------
# Security / Login Handler
# ----------------------------------------------------
def check_password():
    """Simple single username/password authentication interface."""
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False

    if not st.session_state["authenticated"]:
        st.title("🔒 AI Slide Detector Login")
        user_input = st.text_input("Username")
        pass_input = st.text_input("Password", type="password")
        
        if st.button("Login", type="primary"):
            # Set your desired app username & password here
            if user_input == "admin" and pass_input == "ABEdu@5603":
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("Invalid Username or Password")
        return False
    return True

if not check_password():
    st.stop()  # Stop app rendering until logged in

# ----------------------------------------------------
# Helper Functions: AI Detection & Matching
# ----------------------------------------------------
def analyze_text_ai_score(text: str) -> float:
    """Enhanced AI Content Detection calibrated for modern LLM text patterns."""
    if not text or len(text.strip()) < 15:
        return 0.0

    words = re.findall(r'\b[a-zA-Z]+\b', text.lower())
    sentences = [s.strip() for s in re.split(r'[.!?]+', text) if len(s.strip()) > 3]

    word_count = len(words)
    sentence_count = len(sentences)

    if word_count < 8 or sentence_count == 0:
        return 0.0

    # 1. Expanded AI Phrase & Pattern Dictionary (Weight: 40%)
    ai_phrases = [
        "in today's", "fast-paced", "delve into", "testament to", "tapestry of",
        "fostering a", "crucial role", "paramount to", "pivotal role", "furthermore",
        "transformative impact", "key takeaways", "important to note", "in summary",
        "moreover", "seamlessly", "ever-evolving", "rich tapestry", "beacon of",
        "vital component", "underscores the", "by leveraging", "it is worth noting",
        "plays a vital", "holistic approach", "game-changer", "in conclusion",
        "embark on", "delving", "demystify", "unravel", "navigate the", "harnessing",
        "in order to", "serves as a", "it is essential", "shed light on"
    ]
    
    text_lower = text.lower()
    phrase_hits = sum(1 for phrase in ai_phrases if phrase in text_lower)
    phrase_score = min(100.0, phrase_hits * 30.0)

    # 2. Pacing & Coefficient of Variation (Weight: 30%)
    sentence_lengths = [len(re.findall(r'\b\w+\b', s)) for s in sentences]
    avg_sentence_len = sum(sentence_lengths) / sentence_count
    
    # AI targets structured sentence length (12 to 24 words per sentence)
    if 12 <= avg_sentence_len <= 24:
        avg_len_score = 85.0
    elif 9 <= avg_sentence_len <= 28:
        avg_len_score = 60.0
    else:
        avg_len_score = 25.0

    # Standard deviation ratio relative to sentence length
    std_dev = math.sqrt(sum((l - avg_sentence_len) ** 2 for l in sentence_lengths) / sentence_count) if sentence_count > 1 else 0.0
    cv = (std_dev / avg_sentence_len) if avg_sentence_len > 0 else 1.0

    # Predictable AI pacing typically sits in the 0.20 to 0.55 range
    if 0.20 <= cv <= 0.55:
        cv_score = 90.0
    elif 0.10 <= cv < 0.20 or 0.55 < cv <= 0.75:
        cv_score = 60.0
    else:
        cv_score = 20.0

    pacing_score = (avg_len_score * 0.5) + (cv_score * 0.5)

    # 3. Vocabulary Formality & Length (Weight: 20%)
    avg_char_length = sum(len(w) for w in words) / word_count
    if 5.1 <= avg_char_length <= 6.8:
        formality_score = 85.0
    elif 4.7 <= avg_char_length < 5.1 or 6.8 < avg_char_length <= 7.5:
        formality_score = 55.0
    else:
        formality_score = 25.0

    # 4. Structural Connectors & Bullet Formatting (Weight: 10%)
    connectors = ["additionally,", "consequently,", "however,", "overall,", "specifically,", "ultimately,", "thus,", "therefore,"]
    connector_hits = sum(1 for c in connectors if c in text_lower)
    colon_hits = len(re.findall(r':\s', text))
    
    structural_score = min(100.0, (connector_hits * 25.0) + (colon_hits * 15.0))

    # Calculate Weighted Final Score
    raw_score = (phrase_score * 0.40) + (pacing_score * 0.30) + (formality_score * 0.20) + (structural_score * 0.10)

    # Floor Score Calibration for direct AI phrase hits
    if phrase_hits >= 2:
        raw_score = max(raw_score, 78.0)
    elif phrase_hits == 1:
        raw_score = max(raw_score, 52.0)

    return round(min(100.0, max(0.0, raw_score)), 1)


def check_duplicate_in_db(slide_text: str, threshold=0.75):
    """Compares current page/slide content against all records stored in Supabase Cloud DB."""
    if not slide_text.strip() or len(slide_text.split()) < 5:
        return []

    conn = get_db_connection()
    c = conn.cursor()
    c.execute('''
        SELECT p.filename, s.slide_number, s.extracted_text 
        FROM slide_records s
        JOIN presentations p ON s.presentation_id = p.id
    ''')
    records = c.fetchall()
    c.close()
    conn.close()

    matches = []
    for filename, slide_num, stored_text in records:
        ratio = SequenceMatcher(None, slide_text.lower(), stored_text.lower()).ratio()
        if ratio >= threshold:
            matches.append({
                "filename": filename,
                "slide_number": slide_num,
                "similarity_pct": round(ratio * 100, 1)
            })
    return matches


def save_to_database(filename, file_type, total_slides, overall_score, slides_data):
    """Saves file metadata and per-page content to Supabase Cloud DB."""
    conn = get_db_connection()
    c = conn.cursor()
    c.execute('''
        INSERT INTO presentations (filename, file_type, total_slides, overall_ai_score)
        VALUES (%s, %s, %s, %s) RETURNING id;
    ''', (filename, file_type, total_slides, overall_score))
    
    pres_id = c.fetchone()[0]

    for s in slides_data:
        c.execute('''
            INSERT INTO slide_records (presentation_id, slide_number, extracted_text, ai_score)
            VALUES (%s, %s, %s, %s);
        ''', (pres_id, s["slide_num"], s["text"], s["ai_score"]))
    
    conn.commit()
    c.close()
    conn.close()


def parse_file(uploaded_file):
    """Parses text page-by-page from either PPTX or PDF files."""
    filename = uploaded_file.name.lower()
    pages_data = []

    if filename.endswith(".pptx"):
        prs = Presentation(uploaded_file)
        for idx, slide in enumerate(prs.slides, start=1):
            slide_text = []
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        text = paragraph.text.strip()
                        if text:
                            slide_text.append(text)
            full_text = " ".join(slide_text)
            pages_data.append({
                "slide_num": idx,
                "text": full_text,
                "word_count": len(re.findall(r'\b\w+\b', full_text))
            })

    elif filename.endswith(".pdf"):
        reader = PdfReader(uploaded_file)
        for idx, page in enumerate(reader.pages, start=1):
            text = page.extract_text() or ""
            clean_text = " ".join(text.split())
            pages_data.append({
                "slide_num": idx,
                "text": clean_text,
                "word_count": len(re.findall(r'\b\w+\b', clean_text))
            })

    return pages_data

# ----------------------------------------------------
# Streamlit UI Interface
# ----------------------------------------------------
st.set_page_config(page_title="Cloud AI Slide & PDF Detector", layout="wide")

# Header with Logout Button
col_title, col_logout = st.columns([0.85, 0.15])
with col_title:
    st.title("📊 Cloud AI Content Detector (PPTX & PDF)")
with col_logout:
    st.write("")
    if st.button("Logout"):
        st.session_state["authenticated"] = False
        st.rerun()

tab1, tab2 = st.tabs(["Analyze Document", "Shared Cloud History"])

with tab1:
    uploaded_file = st.file_uploader("Upload PowerPoint (.pptx) or PDF (.pdf)", type=["pptx", "pdf"])

    if uploaded_file is not None:
        if st.button("Check File", type="primary"):
            file_ext = uploaded_file.name.split(".")[-1].upper()
            
            with st.spinner(f"Analyzing {file_ext} pages & querying Supabase DB..."):
                pages = parse_file(uploaded_file)
                
                if not pages:
                    st.warning("No readable text found in document.")
                else:
                    total_pages = len(pages)
                    total_ai = 0.0
                    scannable_count = 0
                    analyzed = []

                    for p in pages:
                        score = analyze_text_ai_score(p["text"])
                        p["ai_score"] = score
                        p["duplicates"] = check_duplicate_in_db(p["text"])
                        analyzed.append(p)

                        if p["word_count"] >= 5:
                            total_ai += score
                            scannable_count += 1

                    overall_pct = round(total_ai / scannable_count, 1) if scannable_count > 0 else 0.0

                    # Save to Cloud DB
                    save_to_database(uploaded_file.name, file_ext, total_pages, overall_pct, analyzed)

                    # Display Summary Metrics
                    c1, c2, c3 = st.columns(3)
                    c1.metric("Overall AI Score", f"{overall_pct}%")
                    c2.metric("Total Pages / Slides", total_pages)
                    c3.metric("Scannable Pages", scannable_count)

                    st.markdown("---")
                    st.subheader("Page-by-Page Breakdown & Duplicate Report")

                    for page in analyzed:
                        p_num = page["slide_num"]
                        score = page["ai_score"]
                        dups = page["duplicates"]

                        # Status Badge
                        if score >= 70:
                            badge = "🔴 High AI"
                        elif score >= 40:
                            badge = "🟡 Moderate AI"
                        else:
                            badge = "🟢 Likely Human"

                        match_status = f" | ⚠️ Matched in DB ({len(dups)})" if dups else ""

                        with st.expander(f"Page/Slide {p_num} — AI Rating: {score}% ({badge}){match_status}"):
                            st.write(f"**Word Count:** {page['word_count']}")
                            st.info(page["text"] if page["text"] else "*(Empty Page)*")

                            if dups:
                                st.warning("⚠️ Content Similarity Detected with Previously Saved Records:")
                                for d in dups:
                                    st.write(f"- **{d['similarity_pct']}% match** with file `{d['filename']}` (Page {d['slide_number']})")

with tab2:
    st.subheader("Cloud History (Supabase Database)")
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("SELECT id, filename, file_type, total_slides, overall_ai_score, upload_time FROM presentations ORDER BY upload_time DESC")
        records = c.fetchall()
        c.close()
        conn.close()

        if records:
            for r in records:
                file_type = r[2] if r[2] else "PPTX"
                formatted_time = str(r[5]).split('.')[0] if r[5] else ""
                st.write(f"**File:** `{r[1]}` ({file_type}) | **Total Pages:** {r[3]} | **AI Score:** {r[4]}% | **Uploaded:** {formatted_time}")
        else:
            st.info("No documents saved in cloud database yet.")
    except Exception as e:
        st.error(f"Failed to retrieve history: {e}")
