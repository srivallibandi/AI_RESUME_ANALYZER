import os
import streamlit as st
import pdfplumber
import docx
import sqlite3
import matplotlib.pyplot as plt
from fpdf import FPDF
from docx import Document

# -------------------------------
# Database Setup (SQLite)
# -------------------------------
conn = sqlite3.connect("resumes.db")
c = conn.cursor()
c.execute("""CREATE TABLE IF NOT EXISTS resumes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT,
    job_role TEXT,
    ats_score REAL,
    match_score REAL,
    recommendations TEXT,
    interview_questions TEXT,
    salary_prediction TEXT
)""")
conn.commit()

# -------------------------------
# Font handling
# -------------------------------
# Build an absolute path to the font file based on this script's location,
# so it works no matter what directory Streamlit is launched from.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FONT_PATH = os.path.join(BASE_DIR, "DejaVuSans.ttf")
UNICODE_FONT_AVAILABLE = os.path.isfile(FONT_PATH)

# -------------------------------
# Helper Functions
# -------------------------------
def extract_text_from_pdf(uploaded_file):
    text = ""
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            if page.extract_text():
                text += page.extract_text() + "\n"
    return text

def extract_text_from_docx(uploaded_file):
    doc = docx.Document(uploaded_file)
    return "\n".join([para.text for para in doc.paragraphs])

def calculate_ats_score(resume_text, job_description):
    resume_words = set(resume_text.lower().split())
    jd_words = set(job_description.lower().split())
    if len(jd_words) == 0:
        return 0
    match = resume_words.intersection(jd_words)
    return round(len(match) / len(jd_words) * 100, 2)

def generate_recommendations(job_role):
    recs = {
        "Data Analyst": ["Learn SQL", "Practice Tableau dashboards", "Add Python projects"],
        "Software Engineer": ["Strengthen DSA", "Contribute to GitHub projects", "Add cloud experience"],
        "AI Engineer": ["Build ML models", "Add NLP projects", "Learn TensorFlow/PyTorch"]
    }
    return recs.get(job_role, ["Keep improving your resume"])

def generate_interview_questions(job_role):
    questions = {
        "Data Analyst": ["Explain normalization in SQL", "How do you handle missing data?"],
        "Software Engineer": ["What is polymorphism?", "Explain REST APIs"],
        "AI Engineer": ["Difference between supervised and unsupervised learning?", "Explain transformers in NLP"]
    }
    return questions.get(job_role, ["Tell me about yourself"])

def predict_salary(job_role):
    salaries = {
        "Data Analyst": "Rs. 5-8 LPA",
        "Software Engineer": "Rs. 6-12 LPA",
        "AI Engineer": "Rs. 10-20 LPA"
    }
    return salaries.get(job_role, "Rs. 5-10 LPA")

def export_pdf_report(name, job_role, ats_score, match_score, recommendations, interview_questions, salary_prediction):
    pdf = FPDF()
    pdf.add_page()

    if UNICODE_FONT_AVAILABLE:
        # Unicode font found - supports any special characters (e.g. rupee symbol, accents)
        pdf.add_font("DejaVu", "", FONT_PATH, uni=True)
        pdf.set_font("DejaVu", size=12)
    else:
        # Fall back to a built-in core font so the app doesn't crash.
        # Core fonts (Arial/Helvetica) only support latin-1, so make sure
        # predict_salary() etc. don't contain characters like ₹.
        pdf.set_font("Arial", size=12)

    pdf.cell(200, 10, f"Resume Analysis Report for {name}", ln=True, align="C")
    pdf.multi_cell(0, 10, f"Job Role: {job_role}\nATS Score: {ats_score}%\nMatch Score: {match_score}%\n\nRecommendations:\n" + "\n".join(recommendations))
    pdf.multi_cell(0, 10, "\nInterview Questions:\n" + "\n".join(interview_questions))
    pdf.multi_cell(0, 10, f"\nSalary Prediction: {salary_prediction}")
    pdf.output("resume_report.pdf")

def export_word_report(name, job_role, ats_score, match_score, recommendations, interview_questions, salary_prediction):
    doc = Document()
    doc.add_heading(f"Resume Analysis Report for {name}", 0)
    doc.add_paragraph(f"Job Role: {job_role}")
    doc.add_paragraph(f"ATS Score: {ats_score}%")
    doc.add_paragraph(f"Match Score: {match_score}%")
    doc.add_heading("Recommendations", level=1)
    for rec in recommendations:
        doc.add_paragraph(rec, style="List Bullet")
    doc.add_heading("Interview Questions", level=1)
    for q in interview_questions:
        doc.add_paragraph(q, style="List Bullet")
    doc.add_paragraph(f"Salary Prediction: {salary_prediction}")
    doc.save("resume_report.docx")

# -------------------------------
# Streamlit UI
# -------------------------------
st.set_page_config(page_title="AI Resume Analyzer", layout="wide")
st.title("📄 AI Resume Analyzer Dashboard")

name = st.text_input("Candidate Name")
job_role = st.selectbox("Select Job Role", ["Data Analyst", "Software Engineer", "AI Engineer"])
uploaded_resume = st.file_uploader("Upload Resume (PDF/DOCX)", type=["pdf", "docx"])
job_description = st.text_area("Paste Job Description")

# Default JD fallback
default_jds = {
    "Data Analyst": "Looking for a Data Analyst with skills in SQL, Python, Tableau, and statistics.",
    "Software Engineer": "Seeking a Software Engineer with experience in Java, C++, system design, and algorithms.",
    "AI Engineer": "Hiring AI Engineer with expertise in machine learning, NLP, and TensorFlow/PyTorch."
}
if not job_description.strip():
    job_description = default_jds.get(job_role, "")

if st.button("Analyze Resume"):
    if uploaded_resume and name:
        # Extract resume text
        if uploaded_resume.type == "application/pdf":
            resume_text = extract_text_from_pdf(uploaded_resume)
        else:
            resume_text = extract_text_from_docx(uploaded_resume)

        # ATS Score
        ats_score = calculate_ats_score(resume_text, job_description)
        match_score = ats_score  # simplified

        # Recommendations, Interview Qs, Salary
        recommendations = generate_recommendations(job_role)
        interview_questions = generate_interview_questions(job_role)
        salary_prediction = predict_salary(job_role)

        # Save to DB
        c.execute("INSERT INTO resumes (name, job_role, ats_score, match_score, recommendations, interview_questions, salary_prediction) VALUES (?, ?, ?, ?, ?, ?, ?)",
                  (name, job_role, ats_score, match_score, ", ".join(recommendations), ", ".join(interview_questions), salary_prediction))
        conn.commit()

        # Dashboard
        st.subheader("📊 Analysis Results")
        st.write(f"**ATS Score:** {ats_score}%")
        st.write(f"**Match Score:** {match_score}%")
        st.write(f"**Salary Prediction:** {salary_prediction}")

        # Charts
        fig, ax = plt.subplots()
        ax.bar(["ATS Score", "Match Score"], [ats_score, match_score], color=["blue", "green"])
        st.pyplot(fig)

        # Recommendations
        st.subheader("✅ Recommendations")
        for rec in recommendations:
            st.write("- " + rec)

        # Interview Questions
        st.subheader("🎤 Interview Questions")
        for q in interview_questions:
            st.write("- " + q)

        # Export Reports with Download Buttons
        export_pdf_report(name, job_role, ats_score, match_score, recommendations, interview_questions, salary_prediction)
        with open("resume_report.pdf", "rb") as f:
            st.download_button("📥 Download PDF Report", f, file_name="resume_report.pdf")

        export_word_report(name, job_role, ats_score, match_score, recommendations, interview_questions, salary_prediction)
        with open("resume_report.docx", "rb") as f:
            st.download_button("📥 Download Word Report", f, file_name="resume_report.docx")