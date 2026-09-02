import io
import json
import os
import re

import google.generativeai as genai
import streamlit as st
from docx import Document
from pypdf import PdfReader

MAX_FILE_BYTES = 3 * 1024 * 1024
MAX_TEXT_CHARS = 12_000
MAX_ANALYSES_PER_SESSION = 15
MODELS_TO_TRY = (
    "gemini-2.5-flash",
    "gemini-2.5-flash-lite",
    "gemini-3.5-flash-lite",
    "gemini-flash-latest",
)


def secret(name: str, default: str = "") -> str:
    try:
        value = st.secrets.get(name)
        if value:
            return str(value)
    except Exception:
        pass
    return os.getenv(name, default)


def extract_pdf(data: bytes) -> str:
    reader = PdfReader(io.BytesIO(data))
    pages = []
    for page in reader.pages:
        pages.append(page.extract_text() or "")
    return "\n".join(pages)


def extract_docx(data: bytes) -> str:
    doc = Document(io.BytesIO(data))
    return "\n".join(p.text for p in doc.paragraphs)


def extract_resume(uploaded) -> str:
    name = (uploaded.name or "").lower()
    data = uploaded.getvalue()
    if len(data) > MAX_FILE_BYTES:
        raise ValueError("File is larger than 3 MB. Upload a smaller resume.")
    if name.endswith(".pdf"):
        text = extract_pdf(data)
    elif name.endswith(".docx"):
        text = extract_docx(data)
    else:
        raise ValueError("Use a PDF or DOCX resume.")
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) < 80:
        raise ValueError("Could not read enough text from the file. Try another export of the resume.")
    return text[:MAX_TEXT_CHARS]


def parse_json_payload(raw: str) -> dict:
    raw = (raw or "").strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)
    return json.loads(raw)


def analyze(resume_text: str, jd_text: str) -> dict:
    api_key = secret("GOOGLE_API_KEY")
    if not api_key:
        raise RuntimeError("GOOGLE_API_KEY is missing. Add it in Streamlit Secrets.")
    genai.configure(api_key=api_key)

    prompt = f"""You are JDFit, an India-first resume vs job-description analyst.
Hiring context: Indian IT hiring (freshers to mid-level), Naukri-style JDs, degrees like B.E./B.Tech/M.Tech, internships, notice period, CTC.
This is an AI estimate, not Naukri, Workday, or any real ATS.

Return ONLY valid JSON with this shape:
{{
  "ats_score": <integer 0-100>,
  "score_summary": "<2 sentences>",
  "matched_skills": ["<skill>", "..."],
  "missing_skills": ["<skill>", "..."],
  "resume_suggestions": ["<actionable edit>", "..."],
  "interview_questions": ["<question>", "..."],
  "recommended_job_roles": ["<role title>", "..."]
}}

Rules:
- Score how well THIS resume matches THIS JD (keywords, skills, seniority, tools).
- Prefer skills recruiters search (languages, frameworks, SQL, cloud, DSA) over fluffy adjectives.
- 5-10 matched_skills, 5-10 missing_skills, 5-8 resume_suggestions, 8-12 interview_questions (mix behavioral and technical), 3 recommended_job_roles.
- Suggestions must tell the candidate what to add or rewrite. No generic "be confident".
- If the JD is outside India, still analyze it; keep advice practical.

RESUME:
{resume_text}

JOB DESCRIPTION:
{jd_text[:MAX_TEXT_CHARS]}
"""

    last_error = None
    for model_name in MODELS_TO_TRY:
        try:
            model = genai.GenerativeModel(
                model_name,
                generation_config={
                    "temperature": 0.3,
                    "response_mime_type": "application/json",
                },
            )
            response = model.generate_content(prompt)
            payload = parse_json_payload(response.text)
            payload["model_used"] = model_name
            return payload
        except Exception as exc:
            last_error = exc
            continue
    raise RuntimeError(f"Gemini did not return usable JSON. Last error: {last_error}")


def gate_password() -> bool:
    expected = secret("APP_PASSWORD")
    if not expected:
        return True
    if st.session_state.get("jdfit_ok"):
        return True
    st.title("JDFit")
    st.caption("Private test — enter the password you were given.")
    entered = st.text_input("Password", type="password")
    if st.button("Enter"):
        if entered == expected:
            st.session_state["jdfit_ok"] = True
            st.rerun()
        st.error("Wrong password.")
    return False


def main():
    st.set_page_config(page_title="JDFit", page_icon=":clipboard:", layout="centered")
    if not gate_password():
        return

    if "analyses" not in st.session_state:
        st.session_state.analyses = 0

    st.title("JDFit")
    st.caption(
        "See how well your resume fits a job description. "
        "AI estimate only — not a real ATS. Resumes are processed in memory and not saved."
    )

    resume_file = st.file_uploader("Resume (PDF or DOCX)", type=["pdf", "docx"])
    jd_text = st.text_area("Paste the job description", height=220, placeholder="Paste the full JD from Naukri, LinkedIn, or the company site.")

    if st.button("Analyze fit", type="primary"):
        if st.session_state.analyses >= MAX_ANALYSES_PER_SESSION:
            st.warning("Session limit reached. Refresh later so the free Gemini quota lasts for other testers.")
            return
        if not resume_file:
            st.error("Upload a resume.")
            return
        if not jd_text or len(jd_text.strip()) < 40:
            st.error("Paste a fuller job description.")
            return
        try:
            with st.spinner("Reading resume and scoring against the JD..."):
                resume_text = extract_resume(resume_file)
                result = analyze(resume_text, jd_text.strip())
            st.session_state.analyses += 1
            st.session_state.result = result
        except Exception as exc:
            st.error(str(exc))
            return

    result = st.session_state.get("result")
    if not result:
        return

    score = int(result.get("ats_score") or 0)
    st.subheader(f"Fit score: {score} / 100")
    st.progress(min(max(score, 0), 100) / 100)
    if result.get("score_summary"):
        st.write(result["score_summary"])

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Matched skills**")
        for item in result.get("matched_skills") or []:
            st.markdown(f"- {item}")
    with c2:
        st.markdown("**Missing skills**")
        for item in result.get("missing_skills") or []:
            st.markdown(f"- {item}")

    st.markdown("**What to change on the resume**")
    for item in result.get("resume_suggestions") or []:
        st.markdown(f"- {item}")

    st.markdown("**Interview questions to practice**")
    for item in result.get("interview_questions") or []:
        st.markdown(f"- {item}")

    st.markdown("**Nearby roles to consider**")
    for item in result.get("recommended_job_roles") or []:
        st.markdown(f"- {item}")

    st.caption("JDFit MVP · free Gemini quota · do not upload confidential employer documents you cannot share.")


if __name__ == "__main__":
    main()
