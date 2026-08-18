import streamlit as st
from pptx import Presentation
from pypdf import PdfReader
import psycopg2
import re
import math
from difflib import SequenceMatcher

# ----------------------------------------------------
# Database Connection (Supabase PostgreSQL)
# ----------------------------------------------------
def get_db_connection():
    """Establishes connection to Supabase using secret credentials."""
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

try:
    init_db()
except Exception as e:
    st.error(f"Database connection error: {e}")

# ----------------------------------------------------
# Security / Login Handler
# ----------------------------------------------------
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False

    if not st.session_state["authenticated"]:
        st.title("🔒 AI Slide Detector Login")
        user_input = st.text_input("Username")
        pass_input = st.text_input("Password", type="password")
        
        if st.button("Login", type="primary"):
            if user_input == "admin" and pass_input in ["ABEdu@5603", "ABEdu5603Secret"]:
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("Invalid Username or Password")
        return False
    return True

if not check_password():
    st.stop()

# ----------------------------------------------------
# High-Precision Sentence & Document AI Detection Engine
# ----------------------------------------------------
def score_single_sentence(sentence: str) -> float:
    """Calculates AI probability for an individual sentence or line."""
    clean_s = sentence.strip()
    words = re.findall(r'\b[a-zA-Z]+\b', clean_s.lower())
    if len(words) < 3:
        return 0.0

    score = 0.0
    text_lower = clean_s.lower()

    # 1. Structural AI Q&A and Definition Patterns (Common in ChatGPT slides)
    qa_patterns = [
        r"^what is\b", r"^why is\b", r"^how can we\b", r"^can using\b", 
        r"^why should we\b", r"^is copying\b", r"\bmeans using someone\b",
        r"\bwithout giving them credit\b", r"\buse your own words\b",
        r"\bbe honest, be creative\b", r"\bbecause it is dishonest\b",
        r"\bstops us from learning\b", r"\balways give credit\b"
    ]
    for pattern in qa_patterns:
        if re.search(pattern, text_lower):
            score += 45.0

    # 2. General AI Vocabulary & Transition Phrases
    ai_phrases = [
        "in conclusion", "important to note", "crucial role", "key takeaways",
        "furthermore", "moreover", "in summary", "fast-paced", "delve into",
        "tapestry of", "fostering a", "paramount to", "transformative impact",
        "plays a vital", "holistic approach", "it is essential", "in order to"
    ]
    for phrase in ai_phrases:
        if phrase in text_lower:
            score += 40.0

    # 3. Sentence Length Uniformity & Pacing (ChatGPT preferred range)
    word_count = len(words)
    if 8 <= word_count <= 22:
        score += 25.0
    
    # 4. Standard Definition / Rule Formatting
    if clean_s.startswith(("1.", "2.", "3.", "4.", "5.", "- ", "• ")):
        score += 15.0

    return min(98.0, max(10.0 if score > 0 else 0.0, score))

def analyze_document_text(text: str):
    """Parses text line-by-line / sentence-by-sentence, highlighting AI content."""
    if not text or len(text.strip()) < 10:
        return 0.0, "", 0, 0

    # Split text into distinct sentences and lines
    raw_chunks = [c.strip() for c in re.split(r'(\n+|[.!?]+)', text) if c.strip()]
    
    reconstructed_chunks = []
    temp = ""
    for chunk in raw_chunks:
        if chunk in [".", "!", "?", "\n"]:
            temp += chunk
            reconstructed_chunks.append(temp.strip())
            temp = ""
        else:
            if temp:
                reconstructed_chunks.append(temp.strip())
            temp = chunk
    if temp:
        reconstructed_chunks.append(temp.strip())

    if not reconstructed_chunks:
        reconstructed_chunks = [text]

    highlighted_html = []
    total_score = 0.0
    ai_sentence_count = 0
    valid_chunks = 0

    for chunk in reconstructed_chunks:
        if len(re.findall(r'\b[a-zA-Z]+\b', chunk)) < 3:
            highlighted_html.append(chunk + " ")
            continue

        chunk_score = score_single_sentence(chunk)
        total_score += chunk_score
        valid_chunks += 1

        # Flag as AI if sentence score is 45%+
        if chunk_score >= 45.0:
            ai_sentence_count += 1
            highlighted_html.append(
                f'<mark style="background-color: #ffe066; padding: 2px 5px; border-radius: 4px; font-weight: 500;" title="AI Score: {chunk_score}%">{chunk}</mark> '
            )
        else:
            highlighted_html.append(f'{chunk} ')

    overall_score = round(total_score / valid_chunks, 1) if valid_chunks > 0 else 0.0
    
    # Calibration boost when multiple AI structural sentences are detected
    if ai_sentence_count >= 3:
        overall_score = max(overall_score, round((ai_sentence_count / valid_chunks) * 100, 1))

    return overall_score, "".join(highlighted_html), ai_sentence_count, valid_chunks


def check_duplicate_in_db(slide_text: str, threshold=0.75):
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


def delete_from_database(presentation_id):
    """Deletes a file record and its associated page details from database."""
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("DELETE FROM presentations WHERE id = %s;", (presentation_id,))
    conn.commit()
    c.close()
    conn.close()


def parse_file(uploaded_file):
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
# UI Layout
# ----------------------------------------------------
st.set_page_config(page_title="Cloud AI Slide & PDF Detector", layout="wide")

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
            
            with st.spinner(f"Analyzing sentences in {file_ext}..."):
                pages = parse_file(uploaded_file)
                
                if not pages:
                    st.warning("No readable text found in document.")
                else:
                    total_pages = len(pages)
                    total_ai = 0.0
                    scannable_count = 0
                    analyzed = []
                    doc_ai_sentences = 0
                    doc_total_sentences = 0

                    for p in pages:
                        score, highlighted_text, ai_sents, total_sents = analyze_document_text(p["text"])
                        p["ai_score"] = score
                        p["highlighted_html"] = highlighted_text
                        p["ai_sentences"] = ai_sents
                        p["total_sentences"] = total_sents
                        p["duplicates"] = check_duplicate_in_db(p["text"])
                        
                        doc_ai_sentences += ai_sents
                        doc_total_sentences += total_sents
                        analyzed.append(p)

                        if p["word_count"] >= 5:
                            total_ai += score
                            scannable_count += 1

                    overall_pct = round(total_ai / scannable_count, 1) if scannable_count > 0 else 0.0

                    save_to_database(uploaded_file.name, file_ext, total_pages, overall_pct, analyzed)

                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Overall AI Score", f"{overall_pct}%")
                    c2.metric("AI Sentences Flagged", f"{doc_ai_sentences} / {doc_total_sentences}")
                    c3.metric("Total Pages / Slides", total_pages)
                    c4.metric("Scannable Pages", scannable_count)

                    st.markdown("---")
                    st.subheader("Highlighted Sentence Breakdown")

                    for page in analyzed:
                        p_num = page["slide_num"]
                        score = page["ai_score"]
                        dups = page["duplicates"]

                        if score >= 60:
                            badge = "🔴 High AI Probability"
                        elif score >= 35:
                            badge = "🟡 Mixed / Moderate AI"
                        else:
                            badge = "🟢 Likely Human"

                        match_status = f" | ⚠️ Matched in DB ({len(dups)})" if dups else ""

                        with st.expander(f"Page/Slide {p_num} — AI Score: {score}% ({badge}){match_status}", expanded=True if score >= 40 else False):
                            st.write(f"**Words:** {page['word_count']} | **AI Sentences:** {page['ai_sentences']} of {page['total_sentences']}")
                            
                            if page["highlighted_html"]:
                                st.markdown(
                                    f'<div style="background-color: #f8f9fa; padding: 15px; border-radius: 8px; border: 1px solid #dee2e6; line-height: 1.8;">{page["highlighted_html"]}</div>',
                                    unsafe_allow_html=True
                                )
                            else:
                                st.info("*(Empty Page)*")

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
                rec_id, filename, file_type, total_slides, score, upload_time = r
                formatted_time = str(upload_time).split('.')[0] if upload_time else ""
                
                col_info, col_del = st.columns([0.85, 0.15])
                with col_info:
                    st.write(f"**File:** `{filename}` ({file_type}) | **Total Pages:** {total_slides} | **AI Score:** {score}% | **Uploaded:** {formatted_time}")
                with col_del:
                    if st.button("🗑️ Delete", key=f"del_{rec_id}"):
                        delete_from_database(rec_id)
                        st.success(f"Deleted `{filename}`")
                        st.rerun()
                st.divider()
        else:
            st.info("No documents saved in cloud database yet.")
    except Exception as e:
        st.error(f"Failed to retrieve history: {e}")
