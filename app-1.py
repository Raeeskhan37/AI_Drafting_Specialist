import os
import re
import io
import json
import base64
import hashlib
import textwrap
from datetime import datetime

import streamlit as st

# Optional dependencies
try:
    from groq import Groq
except Exception:
    Groq = None

try:
    from docx import Document
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
except Exception:
    Document = None

try:
    from fpdf import FPDF
except Exception:
    FPDF = None

try:
    from PIL import Image, ImageDraw, ImageFont
except Exception:
    Image = None


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="DraftForge — AI Document Composer",
    page_icon="✦",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# CUSTOM CSS
# ============================================================

st.markdown(
    """
<style>

    /* ---------- GLOBAL ---------- */

    .stApp {
        background:
            radial-gradient(circle at 10% 0%, rgba(99,102,241,0.08), transparent 28%),
            radial-gradient(circle at 90% 10%, rgba(14,165,233,0.07), transparent 25%),
            #f7f8fc;
    }

    .main .block-container {
        max-width: 1400px;
        padding-top: 1.2rem;
        padding-bottom: 3rem;
    }

    h1, h2, h3 {
        letter-spacing: -0.02em;
    }

    /* ---------- HEADER ---------- */

    .df-header {
        background: linear-gradient(
            135deg,
            #111827 0%,
            #1e293b 55%,
            #312e81 100%
        );
        border-radius: 24px;
        padding: 30px 34px;
        color: white;
        margin-bottom: 24px;
        box-shadow: 0 14px 35px rgba(15,23,42,0.16);
    }

    .df-brand {
        font-size: 2.1rem;
        font-weight: 800;
        margin-bottom: 4px;
    }

    .df-tagline {
        font-size: 1rem;
        opacity: 0.82;
    }

    .df-status {
        display: inline-block;
        padding: 7px 12px;
        border-radius: 999px;
        background: rgba(255,255,255,0.12);
        border: 1px solid rgba(255,255,255,0.16);
        font-size: 0.82rem;
        margin-top: 16px;
    }

    /* ---------- STEP BAR ---------- */

    .step-card {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 16px;
        padding: 15px 18px;
        min-height: 82px;
        box-shadow: 0 5px 18px rgba(15,23,42,0.05);
    }

    .step-number {
        font-size: 0.75rem;
        font-weight: 700;
        color: #6366f1;
        text-transform: uppercase;
        letter-spacing: 0.08em;
    }

    .step-title {
        font-size: 1rem;
        font-weight: 750;
        color: #111827;
        margin-top: 3px;
    }

    .step-text {
        font-size: 0.78rem;
        color: #6b7280;
        margin-top: 3px;
    }

    /* ---------- DOCUMENT CARDS ---------- */

    .doc-card {
        background: white;
        border: 2px solid #e5e7eb;
        border-radius: 18px;
        padding: 22px;
        min-height: 150px;
        box-shadow: 0 7px 22px rgba(15,23,42,0.05);
        transition: all 0.2s ease;
    }

    .doc-card.selected {
        border-color: #6366f1;
        background: #f5f5ff;
        box-shadow: 0 10px 28px rgba(99,102,241,0.12);
    }

    .doc-icon {
        font-size: 2rem;
    }

    .doc-title {
        font-size: 1.12rem;
        font-weight: 750;
        margin-top: 8px;
        color: #111827;
    }

    .doc-desc {
        font-size: 0.82rem;
        color: #6b7280;
        line-height: 1.45;
        margin-top: 5px;
    }

    /* ---------- SECTION ---------- */

    .section-header {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 18px;
        padding: 19px 22px;
        margin: 20px 0 12px;
        box-shadow: 0 5px 18px rgba(15,23,42,0.04);
    }

    .section-title {
        font-size: 1.18rem;
        font-weight: 800;
        color: #111827;
    }

    .section-subtitle {
        font-size: 0.83rem;
        color: #6b7280;
        margin-top: 3px;
    }

    /* ---------- INDEX CARDS ---------- */

    .index-card {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 15px;
        padding: 14px 17px;
        margin-bottom: 8px;
    }

    .index-number {
        color: #6366f1;
        font-weight: 800;
    }

    .index-name {
        font-weight: 650;
        color: #1f2937;
    }

    /* ---------- SELECTED PANEL ---------- */

    .selected-panel {
        background: #111827;
        color: white;
        border-radius: 18px;
        padding: 18px 20px;
        margin-top: 15px;
        box-shadow: 0 10px 28px rgba(15,23,42,0.14);
    }

    .selected-title {
        font-weight: 800;
        font-size: 1rem;
        margin-bottom: 10px;
    }

    .selected-pill {
        display: inline-block;
        background: rgba(255,255,255,0.11);
        border: 1px solid rgba(255,255,255,0.15);
        border-radius: 999px;
        padding: 6px 10px;
        margin: 3px;
        font-size: 0.75rem;
    }

    /* ---------- VOICE AREA ---------- */

    .voice-box {
        background: #f8fafc;
        border: 1px dashed #cbd5e1;
        border-radius: 15px;
        padding: 12px 15px;
        margin: 8px 0 10px;
    }

    .voice-title {
        font-weight: 700;
        color: #334155;
        font-size: 0.88rem;
    }

    .voice-help {
        font-size: 0.76rem;
        color: #64748b;
    }

    /* ---------- GENERATED DOCUMENT ---------- */

    .document-shell {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 20px;
        padding: 25px;
        box-shadow: 0 10px 30px rgba(15,23,42,0.07);
    }

    /* ---------- SIDEBAR ---------- */

    [data-testid="stSidebar"] {
        background: #111827;
    }

    [data-testid="stSidebar"] * {
        color: #e5e7eb;
    }

    [data-testid="stSidebar"] .stButton button {
        border-color: #374151;
        background: #1f2937;
    }

    /* ---------- BUTTONS ---------- */

    .stButton > button {
        border-radius: 11px;
        font-weight: 650;
        min-height: 42px;
    }

    /* ---------- DIVIDER ---------- */

    .soft-divider {
        height: 1px;
        background: #e5e7eb;
        margin: 22px 0;
    }

</style>
""",
    unsafe_allow_html=True,
)


# ============================================================
# CONSTANTS
# ============================================================

PROFILE_FILE = "user_profile.json"
HISTORY_FILE = "draftforge_history.json"

GROQ_MODEL = "openai/gpt-oss-120b"
WHISPER_MODEL = "whisper-large-v3"

DOCUMENT_TYPES = ["Email", "Letter", "Inquiry"]

INQUIRY_TYPES = ["E&D Inquiry", "FFI Inquiry"]

ED_INDEXES = [
    "Inquiry Reference No.",
    "Subject",
    "Brief of the Inquiry",
    "Articles of Charge / Allegations",
    "Statement of the Accused",
    "Questions / Answers with the Accused",
    "Statements of Witnesses / Officials",
    "Documentary Evidence / Record Examined",
    "Defence / Written Explanation",
    "Findings",
    "Findings on Each Charge",
    "Discussion / Analysis",
    "Conclusion",
    "Recommendations",
    "Documents Recorded",
    "Inquiry Committee",
]

DOCUMENTS_RECORDED = [
    "CNICF",
    "Birth Certificate (BC)",
    "Marriage Certificate",
    "CNICs",
    "Domicile",
    "Affidavit",
    "Complaint / Application",
    "Written Explanation",
    "Official Record",
    "Statement of Accused",
    "Statement of Witness",
    "Other",
]

COMMITTEE_ROLES = [
    "Convener of Inquiry",
    "Member 1",
    "Member 2",
    "Departmental Representative",
]


# ============================================================
# SESSION STATE
# ============================================================

DEFAULT_STATE = {
    "document_type": "Email",
    "inquiry_type": "E&D Inquiry",
    "index_data": [],
    "generated_draft": "",
    "editable_draft": "",
    "document_editor": "",
    "editor_sync": "",
    "edit_instruction": "",
    "edit_instruction_sync": "",
    "email_instruction": "",
    "letter_instruction": "",
    "email_voice_hash": "",
    "letter_voice_hash": "",
    "profile": {},
    "history": [],
    "show_history": False,
    "show_profile": False,
    "generation_counter": 0,
}

for key, value in DEFAULT_STATE.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# PROFILE / HISTORY
# ============================================================

def load_user_profile():
    default = {
        "name": "",
        "designation": "",
        "contact_no": "",
        "current_station": "",
    }

    try:
        if os.path.exists(PROFILE_FILE):
            with open(PROFILE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)

            if isinstance(data, dict):
                for key in default:
                    default[key] = data.get(key, "")
    except Exception:
        pass

    return default


def save_user_profile(profile):
    try:
        with open(PROFILE_FILE, "w", encoding="utf-8") as f:
            json.dump(profile, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


def load_history():
    try:
        if os.path.exists(HISTORY_FILE):
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, list) else []
    except Exception:
        pass

    return []


def save_history(history):
    try:
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(history[-30:], f, ensure_ascii=False, indent=2)
    except Exception:
        pass


if not st.session_state.profile:
    st.session_state.profile = load_user_profile()

if not st.session_state.history:
    st.session_state.history = load_history()


# ============================================================
# UTILITY FUNCTIONS
# ============================================================

def get_groq_client():
    api_key = None

    try:
        api_key = st.secrets.get("GROQ_API_KEY")
    except Exception:
        pass

    api_key = api_key or os.getenv("GROQ_API_KEY")

    if not api_key or Groq is None:
        return None

    try:
        return Groq(api_key=api_key)
    except Exception:
        return None


def clean_text(text):
    if text is None:
        return ""

    return str(text).strip()


def safe_filename(text):
    text = re.sub(r"[^A-Za-z0-9._-]+", "_", text)
    return text[:100] or "document"


def pdf_safe(text):
    if text is None:
        return ""

    replacements = {
        "\u2018": "'",
        "\u2019": "'",
        "\u201c": '"',
        "\u201d": '"',
        "\u2013": "-",
        "\u2014": "-",
        "\u2026": "...",
        "\u00a0": " ",
        "\u2022": "-",
        "\u2192": "->",
        "\u00ad": "",
    }

    for old, new in replacements.items():
        text = text.replace(old, new)

    return text.encode("latin-1", "replace").decode("latin-1")


def normalize_heading(text):
    return re.sub(r"\s+", " ", text.strip().lower())


def is_qa_heading(line):
    prefix = "Questions / Answers with the Accused"
    return normalize_heading(line).startswith(normalize_heading(prefix))


def parse_markdown_table(lines, start_index):
    rows = []
    i = start_index

    while i < len(lines):
        line = lines[i].strip()

        if not line.startswith("|"):
            break

        cells = [c.strip() for c in line.strip("|").split("|")]

        if all(
            re.fullmatch(r":?-+:?", c.replace(" ", ""))
            for c in cells
        ):
            i += 1
            continue

        if len(cells) >= 2:
            rows.append((cells[0], cells[1]))

        i += 1

    return rows, i


def is_heading_line(line):
    stripped = line.strip()

    if not stripped:
        return False

    if stripped.startswith("#"):
        return True

    if re.match(r"^\d+[\.\)]\s+", stripped):
        return True

    known = [
        "Subject",
        "Inquiry Reference No.",
        "Brief of the Inquiry",
        "Articles of Charge / Allegations",
        "Statement of the Accused",
        "Statements of Witnesses / Officials",
        "Documentary Evidence / Record Examined",
        "Defence / Written Explanation",
        "Findings",
        "Findings on Each Charge",
        "Discussion / Analysis",
        "Conclusion",
        "Recommendations",
        "Documents Recorded",
        "Inquiry Committee",
    ]

    return any(
        normalize_heading(stripped) == normalize_heading(item)
        for item in known
    )


def get_annexure_label(index):
    # 0 -> A, 25 -> Z, 26 -> AA
    result = ""

    n = index + 1

    while n:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result

    return f"Annex-{result}"


# ============================================================
# VOICE TRANSCRIPTION
# ============================================================

def transcribe_audio(audio_file):
    client = get_groq_client()

    if client is None:
        st.error(
            "Groq API is not configured. Please add GROQ_API_KEY "
            "to Streamlit Secrets."
        )
        return ""

    try:
        audio_bytes = audio_file.getvalue()

        if not audio_bytes:
            return ""

        transcription = client.audio.transcriptions.create(
            file=("voice_input.wav", audio_bytes),
            model=WHISPER_MODEL,
            response_format="text",
        )

        if isinstance(transcription, str):
            return transcription.strip()

        return getattr(transcription, "text", "").strip()

    except Exception as e:
        st.error(f"Voice transcription failed: {e}")
        return ""


def voice_input_block(
    label,
    text_value,
    state_hash_key,
    widget_key,
    help_text="You can speak naturally. Your transcription will be added to the text below.",
):
    st.markdown(
        f"""
        <div class="voice-box">
            <div class="voice-title">🎤 {label}</div>
            <div class="voice-help">{help_text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    audio = st.audio_input(
        "Record voice",
        key=widget_key,
    )

    if audio is not None:
        try:
            audio_bytes = audio.getvalue()
            current_hash = hashlib.sha256(audio_bytes).hexdigest()

            if current_hash != st.session_state.get(state_hash_key, ""):
                with st.spinner("Transcribing your voice..."):
                    transcript = transcribe_audio(audio)

                if transcript:
                    if text_value.strip():
                        text_value = (
                            text_value.rstrip()
                            + "\n"
                            + transcript.strip()
                        )
                    else:
                        text_value = transcript.strip()

                    st.session_state[state_hash_key] = current_hash

        except Exception as e:
            st.warning(f"Could not process recording: {e}")

    return text_value


# ============================================================
# AI GENERATION
# ============================================================

def call_groq(prompt):
    client = get_groq_client()

    if client is None:
        return None

    try:
        response = client.chat.completions.create(
            model=GROQ_MODEL,
            messages=[
                {
                    "role": "system",
                    "content": (
                        "You are a professional official-document drafting "
                        "assistant. Follow the user's supplied information "
                        "strictly. Never invent facts."
                    ),
                },
                {
                    "role": "user",
                    "content": prompt,
                },
            ],
            temperature=0.15,
        )

        return response.choices[0].message.content.strip()

    except Exception as e:
        st.error(f"AI generation failed: {e}")
        return None


def build_profile_signature():
    profile = st.session_state.profile

    name = clean_text(profile.get("name", ""))
    designation = clean_text(profile.get("designation", ""))
    contact = clean_text(profile.get("contact_no", ""))
    station = clean_text(profile.get("current_station", ""))

    if not any([name, designation, contact, station]):
        return ""

    lines = []

    if name:
        lines.append(name)

    if designation:
        lines.append(designation)

    if contact:
        lines.append(f"Contact No.: {contact}")

    if station:
        lines.append(f"Current Station: {station}")

    return "\n".join(lines)


def generate_email(instruction):
    if not instruction.strip():
        st.warning("Please provide instructions or information for the email.")
        return None

    prompt = f"""
Draft a professional official email based ONLY on the information below.

STRICT RULES:
1. Correct spelling, grammar, punctuation and obvious voice-transcription errors.
2. Preserve the user's intended meaning.
3. Do not invent names, dates, reference numbers, events, allegations,
   commitments, facts or other information.
4. Do not add unsupported facts.
5. Use concise and professional official English.
6. Do not create a sender signature. The application will append the
   official profile automatically.
7. Return only the email body.

USER INFORMATION:
{instruction}
"""

    return call_groq(prompt)


def generate_letter(instruction):
    if not instruction.strip():
        st.warning("Please provide instructions or information for the letter.")
        return None

    prompt = f"""
Draft a professional official letter based ONLY on the information below.

STRICT RULES:
1. Correct spelling, grammar, punctuation and obvious voice-transcription errors.
2. Preserve the user's intended meaning.
3. Do not invent names, dates, reference numbers, events, allegations,
   commitments, facts or other information.
4. Do not add unsupported facts.
5. Use formal official English.
6. Do not create a sender signature. The application will append the
   official profile automatically.
7. Return only the letter content.

USER INFORMATION:
{instruction}
"""

    return call_groq(prompt)


# ============================================================
# E&D DATA
# ============================================================

def get_index_display_name(index_name, occurrence):
    if index_name in [
        "Statement of the Accused",
        "Questions / Answers with the Accused",
    ] and occurrence > 1:
        return f"{index_name} No. {occurrence}"

    return index_name


def selected_sections_from_data():
    sections = []

    occurrence_counter = {}

    for item in st.session_state.index_data:
        name = item.get("name", "")
        occurrence_counter[name] = occurrence_counter.get(name, 0) + 1

        display_name = get_index_display_name(
            name,
            occurrence_counter[name],
        )

        sections.append(
            {
                "base_name": name,
                "display_name": display_name,
                "content": item.get("content", ""),
                "documents": item.get("documents", []),
                "committee": item.get("committee", []),
            }
        )

    return sections


def build_documents_recorded_text(documents):
    if not documents:
        return "Documents Recorded\n\nNo documents were selected."

    lines = ["Documents Recorded", ""]

    for i, document_name in enumerate(documents):
        lines.append(
            f"{i + 1}. {document_name} — {get_annexure_label(i)}"
        )

    return "\n".join(lines)


def build_committee_text(committee):
    lines = ["Inquiry Committee", ""]

    for member in committee:
        role = clean_text(member.get("role", ""))
        erp = clean_text(member.get("erp", ""))
        name = clean_text(member.get("name", ""))
        designation = clean_text(member.get("designation", ""))

        if role:
            lines.append(role)

        if erp:
            lines.append(f"ERP#: {erp}")

        if name:
            lines.append(f"Name: {name}")

        if designation:
            lines.append(f"Designation: {designation}")

        lines.append("")

    return "\n".join(lines).strip()


def generate_ed_inquiry():
    sections = selected_sections_from_data()

    if not sections:
        st.warning("Please add at least one inquiry index.")
        return None

    # --------------------------------------------------------
    # Deterministic handling for Documents Recorded-only
    # --------------------------------------------------------

    if all(
        section["base_name"] == "Documents Recorded"
        for section in sections
    ):
        documents = []

        for section in sections:
            for doc in section.get("documents", []):
                if doc not in documents:
                    documents.append(doc)

        return build_documents_recorded_text(documents)

    # --------------------------------------------------------
    # Deterministic handling for Inquiry Committee-only
    # --------------------------------------------------------

    if all(
        section["base_name"] == "Inquiry Committee"
        for section in sections
    ):
        committee = []

        for section in sections:
            committee.extend(section.get("committee", []))

        return build_committee_text(committee)

    # --------------------------------------------------------
    # Build exact selected-index allowlist
    # --------------------------------------------------------

    selected_names = [
        section["display_name"]
        for section in sections
    ]

    selected_allowlist = "\n".join(
        f"- {name}"
        for name in selected_names
    )

    information_parts = []

    for section in sections:
        base_name = section["base_name"]
        display_name = section["display_name"]

        if base_name == "Documents Recorded":
            documents = section.get("documents", [])

            if documents:
                content = build_documents_recorded_text(documents)
            else:
                content = "No documents were selected."

        elif base_name == "Inquiry Committee":
            committee = section.get("committee", [])

            if committee:
                content = build_committee_text(committee)
            else:
                content = "No committee details were provided."

        else:
            content = clean_text(section.get("content", ""))

            if not content:
                content = "No information was provided for this index."

        information_parts.append(
            f"""
INDEX:
{display_name}

USER-SUPPLIED INFORMATION:
{content}
"""
        )

    normal_information = "\n".join(information_parts)

    inquiry_date = datetime.now().strftime("%d %B %Y")

    prompt = f"""
Prepare the final E&D Inquiry Report using ONLY the sections explicitly
selected by the user.

INQUIRY DATE:
{inquiry_date}

============================================================
ABSOLUTE SECTION CONTROL
============================================================

THE FOLLOWING ARE THE ONLY SECTIONS THAT MAY APPEAR:

{selected_allowlist}

You MUST NOT create any other heading, section, chapter or summary.

For example, DO NOT add:
- Introduction
- Background
- Summary of Evidence
- Findings
- Findings on Charges
- Discussion
- Conclusion
- Recommendations
- Inquiry Committee
- Documents Recorded
- Any other section

UNLESS that exact section appears in the selected-section list above.

If only "Documents Recorded" is selected, the final report must contain
ONLY "Documents Recorded" and its selected documents.

If only "Inquiry Committee" is selected, the final report must contain
ONLY "Inquiry Committee" and its supplied committee details.

Do not mention sections that were not selected.

============================================================
CONTENT RULES
============================================================

1. Use professional official English.
2. Correct spelling, grammar, punctuation and obvious voice errors.
3. Preserve the user's intended meaning.
4. Never invent names, dates, allegations, evidence, witnesses, findings,
   recommendations, reference numbers or events.
5. Do not infer unsupported facts.
6. Do not add a generic introduction or conclusion.
7. Preserve the order of the selected indexes.
8. Preserve repeated indexes and their numbering.
9. A selected index with no supplied information may state:
   "No information was provided for this index."
10. Do not create a heading for any unselected index.

============================================================
QUESTIONS / ANSWERS
============================================================

Whenever a selected section is:
Questions / Answers with the Accused
or a numbered occurrence of it,

format the content as a two-column Markdown table:

| Questions | Answers |
|---|---|
| 1. Question | Answer |
| 2. Question | Answer |

Keep every question paired with its corresponding answer.
Do not invent answers.

============================================================
DOCUMENTS RECORDED
============================================================

Only include Documents Recorded when it is selected.

Use ONLY the documents supplied by the user.

Annexures MUST be assigned automatically in alphabetical order:
Annex-A, Annex-B, Annex-C, etc.

Do not invent additional documents.
Do not leave gaps in annexure lettering.

============================================================
INQUIRY COMMITTEE
============================================================

Only include Inquiry Committee when it is selected.

Use only the supplied:
- Convener of Inquiry
- Member 1
- Member 2
- Departmental Representative
- ERP#
- Name
- Designation

Do not invent committee members or details.

============================================================
SELECTED INDEX INFORMATION
============================================================

{normal_information}

============================================================

Return ONLY the final report.

Do not add commentary before or after the report.
"""

    return call_groq(prompt)


# ============================================================
# EDIT AI
# ============================================================

def modify_document_with_ai(document, instruction):
    if not document.strip():
        st.warning("There is no generated document to modify.")
        return None

    if not instruction.strip():
        st.warning("Please enter instructions for the AI.")
        return None

    prompt = f"""
Modify the official document below according to the user's instruction.

STRICT RULES:
1. Preserve all factual information.
2. Do not invent names, dates, allegations, evidence, findings,
   recommendations or events.
3. Do not add new facts.
4. Preserve the meaning of statements.
5. Correct grammar and improve official wording where requested.
6. Do not create new sections unless the user's instruction explicitly
   requests modification of an existing section.
7. Preserve the existing section structure.
8. For inquiry reports, NEVER introduce sections that are not already
   present in the document.
9. If the document contains a Questions / Answers table, preserve the
   question-answer pairing.
10. Return only the modified document.

USER'S EDITING INSTRUCTION:
{instruction}

CURRENT DOCUMENT:
{document}
"""

    return call_groq(prompt)


# ============================================================
# OUTPUT PARSING
# ============================================================

def parse_document_for_display(text):
    """
    Converts the generated text into Streamlit-friendly chunks.
    Markdown tables are detected for Q&A.
    """

    lines = text.splitlines()
    chunks = []
    i = 0

    while i < len(lines):
        line = lines[i]

        if is_qa_heading(line):
            chunks.append(("heading", line.strip()))
            i += 1

            # Skip blank lines
            while i < len(lines) and not lines[i].strip():
                i += 1

            if i < len(lines) and lines[i].strip().startswith("|"):
                rows, new_i = parse_markdown_table(lines, i)

                if rows:
                    chunks.append(("table", rows))
                    i = new_i
                    continue

        if line.strip():
            if is_heading_line(line):
                chunks.append(("heading", line.strip()))
            else:
                chunks.append(("paragraph", line.strip()))

        i += 1

    return chunks


# ============================================================
# DOCX EXPORT
# ============================================================

def export_docx(text, filename="DraftForge_Document.docx"):
    if Document is None:
        raise RuntimeError(
            "python-docx is not installed. Install it with: pip install python-docx"
        )

    doc = Document()

    section = doc.sections[0]
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(11)

    chunks = parse_document_for_display(text)

    for kind, content in chunks:

        if kind == "heading":
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT

            run = p.add_run(content)
            run.bold = True
            run.underline = True
            run.font.name = "Arial"
            run.font.size = Pt(11)

        elif kind == "paragraph":
            p = doc.add_paragraph(content)
            p.paragraph_format.space_after = Pt(7)

        elif kind == "table":
            table = doc.add_table(
                rows=1,
                cols=2,
            )

            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.style = "Table Grid"

            hdr = table.rows[0].cells
            hdr[0].text = "Questions"
            hdr[1].text = "Answers"

            for cell in hdr:
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.bold = True

            for question, answer in content:
                cells = table.add_row().cells
                cells[0].text = question
                cells[1].text = answer

                for cell in cells:
                    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    return buffer.getvalue()


# ============================================================
# PDF EXPORT
# ============================================================

def pdf_wrap_text(pdf, text, font_size=10):
    """
    Robustly wraps text before FPDF multi_cell.

    This specifically avoids the previous:
    'Not enough horizontal space to render a single character'
    problem caused by very long/unbreakable strings.
    """

    text = pdf_safe(text)

    if not text:
        return [""]

    usable_width = pdf.w - pdf.l_margin - pdf.r_margin

    # Approximate character capacity.
    # We intentionally wrap aggressively to prevent overflow.
    estimated_chars = max(
        15,
        int(usable_width / max(font_size * 0.42, 1))
    )

    wrapped = []

    for paragraph in text.split("\n"):
        paragraph = paragraph.strip()

        if not paragraph:
            wrapped.append("")
            continue

        pieces = textwrap.wrap(
            paragraph,
            width=estimated_chars,
            break_long_words=True,
            break_on_hyphens=False,
            replace_whitespace=False,
        )

        if not pieces:
            wrapped.append("")
        else:
            wrapped.extend(pieces)

    return wrapped


def export_pdf(text):
    if FPDF is None:
        raise RuntimeError(
            "fpdf2 is not installed. Install it with: pip install fpdf2"
        )

    pdf = FPDF(
        orientation="P",
        unit="mm",
        format="A4",
    )

    pdf.set_auto_page_break(
        auto=True,
        margin=15,
    )

    pdf.set_margins(
        left=18,
        top=17,
        right=18,
    )

    pdf.add_page()

    pdf.set_font(
        "Arial",
        size=10,
    )

    chunks = parse_document_for_display(text)

    usable_width = (
        pdf.w
        - pdf.l_margin
        - pdf.r_margin
    )

    for kind, content in chunks:

        if kind == "heading":
            pdf.set_x(pdf.l_margin)
            pdf.set_font("Arial", "B", 10)

            lines = pdf_wrap_text(
                pdf,
                content,
                10,
            )

            for wrapped_line in lines:
                if wrapped_line:
                    pdf.set_x(pdf.l_margin)
                    pdf.cell(
                        usable_width,
                        6,
                        wrapped_line,
                        ln=True,
                    )

            pdf.ln(1)

            pdf.set_font("Arial", size=10)

        elif kind == "paragraph":
            pdf.set_x(pdf.l_margin)

            wrapped_lines = pdf_wrap_text(
                pdf,
                content,
                10,
            )

            for wrapped_line in wrapped_lines:
                pdf.set_x(pdf.l_margin)

                if wrapped_line:
                    pdf.multi_cell(
                        usable_width,
                        5.5,
                        wrapped_line,
                    )
                else:
                    pdf.ln(3)

            pdf.ln(1)

        elif kind == "table":
            pdf.set_x(pdf.l_margin)

            col1 = usable_width * 0.42
            col2 = usable_width * 0.58

            # Header
            pdf.set_font("Arial", "B", 9)

            pdf.set_x(pdf.l_margin)
            pdf.cell(
                col1,
                7,
                "Questions",
                border=1,
            )
            pdf.cell(
                col2,
                7,
                "Answers",
                border=1,
                ln=True,
            )

            pdf.set_font("Arial", size=9)

            for question, answer in content:

                q_lines = pdf_wrap_text(
                    pdf,
                    question,
                    9,
                )

                a_lines = pdf_wrap_text(
                    pdf,
                    answer,
                    9,
                )

                row_height = max(
                    len(q_lines),
                    len(a_lines),
                ) * 5

                row_height = max(
                    row_height,
                    6,
                )

                # If row won't fit, add page
                if pdf.get_y() + row_height > (
                    pdf.h - pdf.b_margin
                ):
                    pdf.add_page()

                    pdf.set_font("Arial", "B", 9)

                    pdf.set_x(pdf.l_margin)
                    pdf.cell(
                        col1,
                        7,
                        "Questions",
                        border=1,
                    )
                    pdf.cell(
                        col2,
                        7,
                        "Answers",
                        border=1,
                        ln=True,
                    )

                    pdf.set_font("Arial", size=9)

                start_x = pdf.l_margin
                start_y = pdf.get_y()

                # Question cell
                pdf.set_xy(start_x, start_y)
                pdf.multi_cell(
                    col1,
                    5,
                    "\n".join(q_lines),
                    border=1,
                )

                q_end_y = pdf.get_y()

                # Answer cell
                pdf.set_xy(
                    start_x + col1,
                    start_y,
                )

                pdf.multi_cell(
                    col2,
                    5,
                    "\n".join(a_lines),
                    border=1,
                )

                a_end_y = pdf.get_y()

                pdf.set_y(
                    max(
                        q_end_y,
                        a_end_y,
                    )
                )

            pdf.ln(2)

    return bytes(pdf.output())


# ============================================================
# TXT EXPORT
# ============================================================

def export_txt(text):
    return text.encode(
        "utf-8",
    )


# ============================================================
# PNG EXPORT
# ============================================================

def export_png(text):
    if Image is None:
        raise RuntimeError(
            "Pillow is not installed. Install it with: pip install pillow"
        )

    width = 1600
    margin = 80

    try:
        font = ImageFont.truetype(
            "DejaVuSans.ttf",
            30,
        )

        bold_font = ImageFont.truetype(
            "DejaVuSans-Bold.ttf",
            31,
        )
    except Exception:
        font = ImageFont.load_default()
        bold_font = font

    chunks = parse_document_for_display(text)

    lines = []

    for kind, content in chunks:

        if kind == "heading":
            wrapped = textwrap.wrap(
                content,
                width=75,
                break_long_words=True,
            )

            for line in wrapped:
                lines.append(
                    ("heading", line)
                )

            lines.append(
                ("normal", "")
            )

        elif kind == "paragraph":
            wrapped = textwrap.wrap(
                content,
                width=88,
                break_long_words=True,
            )

            for line in wrapped:
                lines.append(
                    ("normal", line)
                )

            lines.append(
                ("normal", "")
            )

        elif kind == "table":
            lines.append(
                ("heading", "Questions | Answers")
            )

            for q, a in content:
                row = f"{q} | {a}"

                wrapped = textwrap.wrap(
                    row,
                    width=88,
                    break_long_words=True,
                )

                for line in wrapped:
                    lines.append(
                        ("normal", line)
                    )

                lines.append(
                    ("normal", "")
                )

    line_height = 45
    height = max(
        400,
        margin * 2 + line_height * len(lines),
    )

    image = Image.new(
        "RGB",
        (width, height),
        "white",
    )

    draw = ImageDraw.Draw(image)

    y = margin

    for kind, line in lines:

        current_font = (
            bold_font
            if kind == "heading"
            else font
        )

        draw.text(
            (margin, y),
            line,
            fill="black",
            font=current_font,
        )

        y += line_height

    buffer = io.BytesIO()

    image.crop(
        (
            0,
            0,
            width,
            min(
                height,
                y + margin,
            ),
        )
    ).save(
        buffer,
        format="PNG",
    )

    buffer.seek(0)

    return buffer.getvalue()


# ============================================================
# HISTORY
# ============================================================

def save_to_history(document, document_type):
    if not document:
        return

    entry = {
        "id": datetime.now().strftime(
            "%Y%m%d%H%M%S%f"
        ),
        "date": datetime.now().strftime(
            "%d %B %Y %I:%M %p"
        ),
        "type": document_type,
        "preview": document[:250],
        "content": document,
    }

    st.session_state.history.append(entry)

    if len(st.session_state.history) > 30:
        st.session_state.history = (
            st.session_state.history[-30:]
        )

    save_history(
        st.session_state.history
    )


# ============================================================
# HEADER
# ============================================================

st.markdown(
    """
<div class="df-header">
    <div class="df-brand">✦ DraftForge</div>
    <div class="df-tagline">
        AI Document Composer — create professional official documents faster
    </div>
    <div class="df-status">
        ● AI Document Workspace
    </div>
</div>
""",
    unsafe_allow_html=True,
)


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.markdown(
        """
        <div style="
            font-size:1.35rem;
            font-weight:800;
            margin-bottom:4px;
        ">
            ✦ DraftForge
        </div>
        <div style="
            font-size:0.78rem;
            color:#9ca3af;
            margin-bottom:20px;
        ">
            Official document workspace
        </div>
        """,
        unsafe_allow_html=True,
    )

    if st.button(
        "🏠  Workspace",
        use_container_width=True,
    ):
        st.session_state.show_history = False
        st.session_state.show_profile = False

    if st.button(
        "📚  My Documents",
        use_container_width=True,
    ):
        st.session_state.show_history = True
        st.session_state.show_profile = False

    if st.button(
        "👤  My Profile",
        use_container_width=True,
    ):
        st.session_state.show_profile = True
        st.session_state.show_history = False

    st.markdown("---")

    with st.expander("💡 Tips & Templates"):
        st.markdown(
            """
            **Document templates**
            
            • Email  
            • Letter  
            • E&D Inquiry  

            **Useful features**
            
            • Voice + text input  
            • Repeated inquiry indexes  
            • Automatic annexures  
            • AI editing  
            • PDF / DOCX / TXT / PNG export  
            """
        )

    with st.expander("ℹ️ About DraftForge"):
        st.markdown(
            """
            **About the Developer**

            Developed by: **Raees Khan**

            Assistant Director, NADRA
            """
        )


# ============================================================
# PROFILE SCREEN
# ============================================================

if st.session_state.show_profile:

    st.markdown(
        """
        <div class="section-header">
            <div class="section-title">👤 My Profile</div>
            <div class="section-subtitle">
                Your profile is automatically used at the end of Email and Letter documents.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    profile = st.session_state.profile

    c1, c2 = st.columns(2)

    with c1:
        profile_name = st.text_input(
            "Name",
            value=profile.get("name", ""),
            key="profile_name",
        )

        profile_designation = st.text_input(
            "Designation",
            value=profile.get("designation", ""),
            key="profile_designation",
        )

    with c2:
        profile_contact = st.text_input(
            "Contact No.",
            value=profile.get("contact_no", ""),
            key="profile_contact",
        )

        profile_station = st.text_input(
            "Current Station",
            value=profile.get("current_station", ""),
            key="profile_station",
        )

    if st.button(
        "💾 Save Profile",
        type="primary",
    ):
        new_profile = {
            "name": profile_name,
            "designation": profile_designation,
            "contact_no": profile_contact,
            "current_station": profile_station,
        }

        if save_user_profile(new_profile):
            st.session_state.profile = new_profile
            st.success("Profile saved successfully.")
        else:
            st.error("Could not save profile.")

    st.stop()


# ============================================================
# HISTORY SCREEN
# ============================================================

if st.session_state.show_history:

    st.markdown(
        """
        <div class="section-header">
            <div class="section-title">📚 My Documents</div>
            <div class="section-subtitle">
                Your recently generated documents are stored locally.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not st.session_state.history:
        st.info(
            "No documents have been generated yet."
        )
    else:
        for index, entry in enumerate(
            reversed(st.session_state.history)
        ):

            with st.expander(
                f"📄 {entry.get('type', 'Document')} — "
                f"{entry.get('date', '')}"
            ):
                st.caption(
                    entry.get(
                        "preview",
                        "",
                    )
                )

                if st.button(
                    "Open Document",
                    key=f"history_open_{index}",
                ):
                    original = entry.get(
                        "content",
                        "",
                    )

                    st.session_state.editable_draft = original
                    st.session_state.generated_draft = original
                    st.session_state.editor_sync = original

                    st.session_state.show_history = False

                    st.rerun()

    st.stop()


# ============================================================
# STEP 1 — DOCUMENT TYPE
# ============================================================

st.markdown(
    """
<div class="section-header">
    <div class="section-title">① Choose your document</div>
    <div class="section-subtitle">
        Select the type of document you want DraftForge to create.
    </div>
</div>
""",
    unsafe_allow_html=True,
)

type_columns = st.columns(3)

doc_info = {
    "Email": (
        "📧",
        "Email",
        "Professional official email",
    ),
    "Letter": (
        "📄",
        "Letter",
        "Formal official correspondence",
    ),
    "Inquiry": (
        "🔎",
        "Inquiry",
        "E&D / FFI inquiry documents",
    ),
}

for col, doc_type in zip(
    type_columns,
    DOCUMENT_TYPES,
):
    icon, title, description = doc_info[doc_type]

    selected_class = (
        "selected"
        if st.session_state.document_type == doc_type
        else ""
    )

    with col:
        st.markdown(
            f"""
            <div class="doc-card {selected_class}">
                <div class="doc-icon">{icon}</div>
                <div class="doc-title">{title}</div>
                <div class="doc-desc">{description}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if st.button(
            f"Use {title}",
            key=f"doc_type_{doc_type}",
            use_container_width=True,
            type=(
                "primary"
                if st.session_state.document_type == doc_type
                else "secondary"
            ),
        ):
            st.session_state.document_type = doc_type

            if doc_type != "Inquiry":
                st.session_state.index_data = []

            st.rerun()


# ============================================================
# STEP 2 — INPUT
# ============================================================

st.markdown(
    """
<div class="section-header">
    <div class="section-title">② Provide your information</div>
    <div class="section-subtitle">
        Type naturally, speak naturally, or combine both.
    </div>
</div>
""",
    unsafe_allow_html=True,
)


# ============================================================
# EMAIL
# ============================================================

if st.session_state.document_type == "Email":

    st.info(
        "💡 Tell DraftForge what the email should say. "
        "You do not need to write perfect English."
    )

    email_text = st.session_state.email_instruction

    email_text = voice_input_block(
        "Voice input",
        email_text,
        "email_voice_hash",
        "email_voice_input",
    )

    email_text = st.text_area(
        "Email instructions / information",
        value=email_text,
        height=220,
        placeholder=(
            "Example: Inform the regional office that the backup "
            "internet connection is unavailable and request immediate restoration."
        ),
        key="email_instruction",
    )

    st.caption(
        "DraftForge will correct grammar and convert your natural-language "
        "instructions into professional official English."
    )

    # Generate intentionally LAST
    if st.button(
        "✨ Generate Email",
        type="primary",
        use_container_width=True,
    ):
        with st.spinner("Drafting your official email..."):
            result = generate_email(
                st.session_state.email_instruction
            )

        if result:
            signature = build_profile_signature()

            if signature:
                result = (
                    result.rstrip()
                    + "\n\n"
                    + signature
                )

            st.session_state.generation_counter += 1
            st.session_state.generated_draft = result
            st.session_state.editable_draft = result
            st.session_state.editor_sync = result

            save_to_history(
                result,
                "Email",
            )

            st.success("Email generated successfully.")
            st.rerun()


# ============================================================
# LETTER
# ============================================================

elif st.session_state.document_type == "Letter":

    st.info(
        "💡 Explain the purpose of the letter in your own words. "
        "DraftForge will prepare the formal version."
    )

    letter_text = st.session_state.letter_instruction

    letter_text = voice_input_block(
        "Voice input",
        letter_text,
        "letter_voice_hash",
        "letter_voice_input",
    )

    letter_text = st.text_area(
        "Letter instructions / information",
        value=letter_text,
        height=250,
        placeholder=(
            "Example: Write to the concerned office requesting replacement "
            "of the defective printer because operations are being affected."
        ),
        key="letter_instruction",
    )

    st.caption(
        "The official profile will be appended automatically at the end."
    )

    # Generate intentionally LAST
    if st.button(
        "✨ Generate Letter",
        type="primary",
        use_container_width=True,
    ):
        with st.spinner("Drafting your official letter..."):
            result = generate_letter(
                st.session_state.letter_instruction
            )

        if result:
            signature = build_profile_signature()

            if signature:
                result = (
                    result.rstrip()
                    + "\n\n"
                    + signature
                )

            st.session_state.generation_counter += 1
            st.session_state.generated_draft = result
            st.session_state.editable_draft = result
            st.session_state.editor_sync = result

            save_to_history(
                result,
                "Letter",
            )

            st.success("Letter generated successfully.")
            st.rerun()


# ============================================================
# INQUIRY
# ============================================================

elif st.session_state.document_type == "Inquiry":

    inquiry_columns = st.columns(2)

    with inquiry_columns[0]:
        st.markdown(
            """
            <div class="index-card">
                <div class="index-number">Inquiry Type</div>
                <div class="index-name">
                    Select the inquiry format
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with inquiry_columns[1]:
        selected_inquiry = st.selectbox(
            "Inquiry type",
            INQUIRY_TYPES,
            index=INQUIRY_TYPES.index(
                st.session_state.inquiry_type
            ),
            label_visibility="collapsed",
        )

        st.session_state.inquiry_type = selected_inquiry

    # FFI
    if st.session_state.inquiry_type == "FFI Inquiry":

        st.info(
            "🚧 FFI Inquiry is currently Under Construction / Under Process."
        )

    # E&D
    else:

        st.markdown(
            """
            <div class="section-header">
                <div class="section-title">📑 Add Inquiry Sections</div>
                <div class="section-subtitle">
                    Add only the sections you need. The final report will contain
                    only the sections you actually add.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        selected_names = [
            item["name"]
            for item in st.session_state.index_data
        ]

        # ----------------------------------------------------
        # Add index
        # ----------------------------------------------------

        index_col1, index_col2 = st.columns(
            [4, 1]
        )

        with index_col1:

            chosen_index = st.selectbox(
                "Select an inquiry section to add",
                ED_INDEXES,
                key="new_ed_index",
            )

        with index_col2:

            st.markdown(
                "<br>",
                unsafe_allow_html=True,
            )

            if st.button(
                "＋ Add",
                use_container_width=True,
                type="primary",
            ):

                st.session_state.index_data.append(
                    {
                        "name": chosen_index,
                        "content": "",
                        "documents": [],
                        "committee": [],
                    }
                )

                st.rerun()

        # ----------------------------------------------------
        # Selected section overview
        # ----------------------------------------------------

        if st.session_state.index_data:

            st.markdown(
                '<div class="selected-panel">'
                '<div class="selected-title">'
                '✓ Selected Inquiry Sections'
                '</div>',
                unsafe_allow_html=True,
            )

            occurrence_counter = {}

            for position, item in enumerate(
                st.session_state.index_data
            ):
                name = item["name"]

                occurrence_counter[name] = (
                    occurrence_counter.get(name, 0) + 1
                )

                display_name = get_index_display_name(
                    name,
                    occurrence_counter[name],
                )

                st.markdown(
                    f'<span class="selected-pill">'
                    f'{position + 1}. {display_name}'
                    f'</span>',
                    unsafe_allow_html=True,
                )

            st.markdown(
                "</div>",
                unsafe_allow_html=True,
            )

        # ----------------------------------------------------
        # Input sections
        # ----------------------------------------------------

        if not st.session_state.index_data:
            st.info(
                "Start by selecting an inquiry section above."
            )

        occurrence_counter = {}

        for position, item in enumerate(
            st.session_state.index_data
        ):

            base_name = item["name"]

            occurrence_counter[base_name] = (
                occurrence_counter.get(base_name, 0) + 1
            )

            occurrence = occurrence_counter[base_name]

            display_name = get_index_display_name(
                base_name,
                occurrence,
            )

            st.markdown(
                f"""
                <div class="section-header">
                    <div class="section-title">
                        {position + 1}. {display_name}
                    </div>
                    <div class="section-subtitle">
                        Provide information for this selected index.
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            # ------------------------------------------------
            # Documents Recorded
            # ------------------------------------------------

            if base_name == "Documents Recorded":

                current_documents = item.get(
                    "documents",
                    [],
                )

                selected_documents = st.multiselect(
                    "Select documents to record",
                    DOCUMENTS_RECORDED,
                    default=current_documents,
                    key=f"documents_{position}",
                )

                item["documents"] = selected_documents

                if selected_documents:

                    st.markdown(
                        "#### 📎 Annexure Preview"
                    )

                    preview_rows = []

                    for i, document in enumerate(
                        selected_documents
                    ):
                        preview_rows.append(
                            {
                                "Document": document,
                                "Annexure": get_annexure_label(i),
                            }
                        )

                    st.table(
                        preview_rows
                    )

                continue

            # ------------------------------------------------
            # Inquiry Committee
            # ------------------------------------------------

            if base_name == "Inquiry Committee":

                st.caption(
                    "Enter details only for committee members that actually exist."
                )

                committee = item.get(
                    "committee",
                    [],
                )

                while len(committee) < len(
                    COMMITTEE_ROLES
                ):
                    role = COMMITTEE_ROLES[
                        len(committee)
                    ]

                    committee.append(
                        {
                            "role": role,
                            "erp": "",
                            "name": "",
                            "designation": "",
                        }
                    )

                for member_index, role in enumerate(
                    COMMITTEE_ROLES
                ):

                    member = committee[
                        member_index
                    ]

                    st.markdown(
                        f"**{role}**"
                    )

                    c1, c2, c3 = st.columns(3)

                    with c1:
                        member["erp"] = st.text_input(
                            "ERP#",
                            value=member.get(
                                "erp",
                                "",
                            ),
                            key=(
                                f"committee_erp_"
                                f"{position}_{member_index}"
                            ),
                        )

                    with c2:
                        member["name"] = st.text_input(
                            "Name",
                            value=member.get(
                                "name",
                                "",
                            ),
                            key=(
                                f"committee_name_"
                                f"{position}_{member_index}"
                            ),
                        )

                    with c3:
                        member["designation"] = st.text_input(
                            "Designation",
                            value=member.get(
                                "designation",
                                "",
                            ),
                            key=(
                                f"committee_designation_"
                                f"{position}_{member_index}"
                            ),
                        )

                item["committee"] = committee

                continue

            # ------------------------------------------------
            # Normal indexes
            # ------------------------------------------------

            current_content = item.get(
                "content",
                "",
            )

            # Voice hash unique to each inquiry input
            voice_hash_key = (
                f"inquiry_voice_hash_{position}"
            )

            widget_key = (
                f"inquiry_voice_{position}"
            )

            current_content = voice_input_block(
                "Voice input",
                current_content,
                voice_hash_key,
                widget_key,
            )

            current_content = st.text_area(
                "Type or edit information",
                value=current_content,
                height=190,
                placeholder=(
                    "Speak or type the information for this section..."
                ),
                key=f"inquiry_text_{position}",
            )

            item["content"] = current_content

            # ------------------------------------------------
            # Remove index
            # ------------------------------------------------

            if st.button(
                "🗑 Remove this section",
                key=f"remove_index_{position}",
            ):
                st.session_state.index_data.pop(
                    position
                )
                st.rerun()

        # ----------------------------------------------------
        # Generate Inquiry
        # ----------------------------------------------------

        if st.session_state.index_data:

            st.markdown(
                "<div class='soft-divider'></div>",
                unsafe_allow_html=True,
            )

            st.caption(
                "The Generate button is intentionally placed after all "
                "selected sections and their inputs."
            )

            if st.button(
                "✨ Generate Inquiry Report",
                type="primary",
                use_container_width=True,
            ):

                with st.spinner(
                    "Preparing your E&D Inquiry Report..."
                ):
                    result = generate_ed_inquiry()

                if result:
                    st.session_state.generation_counter += 1
                    st.session_state.generated_draft = result
                    st.session_state.editable_draft = result
                    st.session_state.editor_sync = result

                    save_to_history(
                        result,
                        "E&D Inquiry",
                    )

                    st.success(
                        "Inquiry Report generated successfully."
                    )

                    st.rerun()


# ============================================================
# STEP 3 — GENERATED DOCUMENT
# ============================================================

if st.session_state.editable_draft:

    st.markdown(
        """
        <div class="section-header">
            <div class="section-title">③ Review & improve your document</div>
            <div class="section-subtitle">
                Edit the document directly or ask AI to improve it.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    left, right = st.columns(
        [1.65, 1],
        gap="large",
    )

    # --------------------------------------------------------
    # Document editor
    # --------------------------------------------------------

    with left:

        st.markdown(
            """
            <div class="document-shell">
                <strong>📄 Generated Document</strong>
            </div>
            """,
            unsafe_allow_html=True,
        )

        # IMPORTANT:
        # Synchronize widget state BEFORE creating widget.
        if st.session_state.get(
            "editor_sync",
            "",
        ):
            st.session_state.document_editor = (
                st.session_state.editor_sync
            )

            st.session_state.editor_sync = ""

        elif not st.session_state.get(
            "document_editor",
            "",
        ):
            st.session_state.document_editor = (
                st.session_state.editable_draft
                or st.session_state.generated_draft
            )

        edited_document = st.text_area(
            "Document",
            height=650,
            key="document_editor",
            label_visibility="collapsed",
        )

        if st.button(
            "💾 Save Changes",
            use_container_width=True,
            type="primary",
        ):
            st.session_state.editable_draft = (
                edited_document
            )

            st.session_state.generated_draft = (
                edited_document
            )

            save_to_history(
                edited_document,
                st.session_state.document_type,
            )

            st.success(
                "Changes saved successfully."
            )

    # --------------------------------------------------------
    # AI editing
    # --------------------------------------------------------

    with right:

        st.markdown(
            """
            <div class="section-header">
                <div class="section-title">✨ AI Editing</div>
                <div class="section-subtitle">
                    Tell AI what you want changed.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        # ----------------------------------------------------
        # Critical widget-state fix:
        # apply sync BEFORE widget creation.
        # ----------------------------------------------------

        if "edit_instruction_sync" in st.session_state:
            st.session_state.edit_instruction = (
                st.session_state.pop(
                    "edit_instruction_sync"
                )
            )

        edit_instruction = st.text_area(
            "AI editing instruction",
            height=190,
            placeholder=(
                "Example:\n"
                "Make the language more formal.\n\n"
                "Or:\n"
                "Correct the grammar without changing the meaning."
            ),
            key="edit_instruction",
        )

        if st.button(
            "✨ Improve with AI",
            use_container_width=True,
            type="primary",
        ):

            with st.spinner(
                "AI is improving your document..."
            ):
                modified = modify_document_with_ai(
                    st.session_state.editable_draft,
                    edit_instruction,
                )

            if modified:

                st.session_state.editable_draft = (
                    modified
                )

                st.session_state.generated_draft = (
                    modified
                )

                st.session_state.editor_sync = (
                    modified
                )

                # DO NOT modify edit_instruction directly
                # after the widget has been created.
                st.session_state.edit_instruction_sync = ""

                save_to_history(
                    modified,
                    st.session_state.document_type,
                )

                st.success(
                    "Document improved successfully."
                )

                st.rerun()

        st.markdown(
            "<div class='soft-divider'></div>",
            unsafe_allow_html=True,
        )

        # ----------------------------------------------------
        # Restore original
        # ----------------------------------------------------

        if st.button(
            "↩ Restore Original",
            use_container_width=True,
        ):

            # Find latest matching history entry.
            original = None

            for entry in reversed(
                st.session_state.history
            ):
                if (
                    entry.get("type")
                    == st.session_state.document_type
                ):
                    original = entry.get(
                        "content",
                        "",
                    )
                    break

            if original:
                st.session_state.editable_draft = original
                st.session_state.generated_draft = original
                st.session_state.editor_sync = original
                st.rerun()

        # ----------------------------------------------------
        # Export
        # ----------------------------------------------------

        st.markdown(
            """
            <div class="section-header">
                <div class="section-title">📤 Export Document</div>
                <div class="section-subtitle">
                    Download the current version of your document.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        export_text = st.session_state.editable_draft

        filename_base = safe_filename(
            f"DraftForge_{st.session_state.document_type}"
        )

        e1, e2 = st.columns(2)

        with e1:

            try:
                pdf_bytes = export_pdf(
                    export_text
                )

                st.download_button(
                    "📕 PDF",
                    data=pdf_bytes,
                    file_name=f"{filename_base}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )

            except Exception as e:
                st.error(
                    f"PDF export failed: {e}"
                )

        with e2:

            try:
                docx_bytes = export_docx(
                    export_text
                )

                st.download_button(
                    "📘 DOCX",
                    data=docx_bytes,
                    file_name=f"{filename_base}.docx",
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.wordprocessingml.document"
                    ),
                    use_container_width=True,
                )

            except Exception as e:
                st.error(
                    f"DOCX export failed: {e}"
                )

        e3, e4 = st.columns(2)

        with e3:

            txt_bytes = export_txt(
                export_text
            )

            st.download_button(
                "📄 TXT",
                data=txt_bytes,
                file_name=f"{filename_base}.txt",
                mime="text/plain",
                use_container_width=True,
            )

        with e4:

            try:
                png_bytes = export_png(
                    export_text
                )

                st.download_button(
                    "🖼 PNG",
                    data=png_bytes,
                    file_name=f"{filename_base}.png",
                    mime="image/png",
                    use_container_width=True,
                )

            except Exception as e:
                st.error(
                    f"PNG export failed: {e}"
                )


# ============================================================
# FOOTER
# ============================================================

st.markdown(
    """
    <div style="
        text-align:center;
        color:#9ca3af;
        font-size:0.75rem;
        margin-top:45px;
        padding-top:20px;
        border-top:1px solid #e5e7eb;
    ">
        ✦ DraftForge — AI Document Composer
        <br>
        Developed by Raees Khan — Assistant Director, NADRA
    </div>
    """,
    unsafe_allow_html=True,
)
