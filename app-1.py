import os
import io
import re
import json
import hashlib
import textwrap
from datetime import datetime

import streamlit as st


# ============================================================
# OPTIONAL PACKAGES
# ============================================================

try:
    from groq import Groq
except Exception:
    Groq = None

try:
    from docx import Document
    from docx.shared import Pt, Inches
    from docx.enum.table import WD_TABLE_ALIGNMENT
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
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="DraftForge — AI Document Composer",
    page_icon="✦",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# FILES / MODELS
# ============================================================

PROFILE_FILE = "user_profile.json"
HISTORY_FILE = "draftforge_history.json"

GROQ_MODEL = "openai/gpt-oss-120b"
WHISPER_MODEL = "whisper-large-v3"

DOCUMENT_TYPES = [
    "Email",
    "Letter",
    "Inquiry",
]

INQUIRY_TYPES = [
    "E&D Inquiry",
    "FFI Inquiry",
]

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

DEFAULTS = {
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
}

for key, value in DEFAULTS.items():

    if key not in st.session_state:

        if isinstance(value, list):
            st.session_state[key] = []

        elif isinstance(value, dict):
            st.session_state[key] = {}

        else:
            st.session_state[key] = value


# ============================================================
# FILE FUNCTIONS
# ============================================================

def load_json(filename, default):

    try:

        if os.path.exists(filename):

            with open(
                filename,
                "r",
                encoding="utf-8",
            ) as f:

                return json.load(f)

    except Exception:
        pass

    return default


def save_json(filename, data):

    try:

        with open(
            filename,
            "w",
            encoding="utf-8",
        ) as f:

            json.dump(
                data,
                f,
                indent=2,
                ensure_ascii=False,
            )

        return True

    except Exception:

        return False


# ============================================================
# LOAD USER DATA
# ============================================================

if not st.session_state.profile:

    st.session_state.profile = load_json(
        PROFILE_FILE,
        {
            "Name": "",
            "Designation": "",
            "Contact No.": "",
            "Current Station": "",
        },
    )


if not st.session_state.history:

    st.session_state.history = load_json(
        HISTORY_FILE,
        [],
    )


# ============================================================
# STYLING
# ============================================================

st.markdown(
    """
    <style>

    .stApp {
        background-color: #f8fafc;
    }

    .block-container {
        max-width: 1400px;
        padding-top: 1.5rem;
        padding-bottom: 3rem;
    }

    /* ---------------- INPUTS ---------------- */

    div[data-testid="stTextInput"] input,
    div[data-testid="stTextArea"] textarea {
        background-color: #ffffff !important;
        color: #111827 !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 10px !important;
    }

    div[data-testid="stTextInput"] label,
    div[data-testid="stTextArea"] label {
        color: #334155 !important;
        font-weight: 650 !important;
    }

    div[data-testid="stTextInput"] input::placeholder,
    div[data-testid="stTextArea"] textarea::placeholder {
        color: #94a3b8 !important;
        opacity: 1 !important;
    }

    div[data-testid="stTextInput"] input:focus,
    div[data-testid="stTextArea"] textarea:focus {
        border-color: #6366f1 !important;
        box-shadow: 0 0 0 2px rgba(99,102,241,0.12) !important;
    }

    /* ---------------- BUTTONS ---------------- */

    .stButton > button {
        border-radius: 10px !important;
        font-weight: 700 !important;
        min-height: 42px !important;
    }

    /* ---------------- COMMITTEE ROLE CARDS ---------------- */

    .committee-role-card {
        background-color: #e0f2fe;
        border: 2px solid #7dd3fc;
        border-radius: 12px;
        padding: 14px 16px;
        margin-top: 18px;
        margin-bottom: 8px;
    }

    .committee-role-title {
        color: #0c4a6e;
        font-size: 19px;
        font-weight: 800;
        margin-bottom: 4px;
    }

    .committee-role-description {
        color: #334155;
        font-size: 13px;
        font-weight: 500;
    }

    .committee-tip {
        background-color: #eff6ff;
        border-left: 4px solid #2563eb;
        border-radius: 8px;
        padding: 12px 14px;
        color: #1e3a8a;
        margin-top: 15px;
    }

    hr {
        margin-top: 25px;
        margin-bottom: 25px;
    }

    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# AI FUNCTIONS
# ============================================================

def get_groq_client():

    api_key = None

    try:
        api_key = st.secrets.get(
            "GROQ_API_KEY"
        )
    except Exception:
        pass

    api_key = (
        api_key
        or os.getenv("GROQ_API_KEY")
    )

    if not api_key:

        st.error(
            "GROQ_API_KEY is not configured. "
            "Please add it to Streamlit Secrets."
        )

        return None

    if Groq is None:

        st.error(
            "Groq package is not installed."
        )

        return None

    return Groq(
        api_key=api_key
    )


def clean_ai_text(text):

    if not text:
        return ""

    text = str(text).strip()

    text = re.sub(
        r"^```(?:text|markdown)?\s*",
        "",
        text,
        flags=re.IGNORECASE,
    )

    text = re.sub(
        r"\s*```$",
        "",
        text,
    )

    return text.strip()


def profile_signature(profile):

    lines = []

    name = profile.get(
        "Name",
        "",
    ).strip()

    designation = profile.get(
        "Designation",
        "",
    ).strip()

    contact = profile.get(
        "Contact No.",
        "",
    ).strip()

    station = profile.get(
        "Current Station",
        "",
    ).strip()

    if name:
        lines.append(name)

    if designation:
        lines.append(designation)

    if contact:
        lines.append(
            f"Contact No.: {contact}"
        )

    if station:
        lines.append(
            f"Current Station: {station}"
        )

    if not lines:
        return ""

    return "\n\n" + "\n".join(lines)


def generate_ai_document(
    document_type,
    instruction,
    profile,
):

    client = get_groq_client()

    if client is None:
        return ""

    system_prompt = f"""
You are DraftForge, an AI assistant for professional
official document drafting.

DOCUMENT TYPE:
{document_type}

Convert the user's natural-language instructions into
professional official English.

STRICT RULES:

1. Correct spelling, grammar, punctuation and obvious
   voice-transcription errors.

2. Preserve the user's intended meaning.

3. Do not invent names, dates, facts, allegations,
   evidence, events or reference numbers.

4. Do not add unsupported information.

5. Do not add a sender signature.

6. Return only the finished document.

7. Use professional official English.

8. Do not explain the drafting process.
"""

    user_prompt = f"""
Prepare the {document_type} based ONLY on this information:

{instruction}
"""

    try:

        response = client.chat.completions.create(
            model=GROQ_MODEL,
            temperature=0.15,
            messages=[
                {
                    "role": "system",
                    "content": system_prompt,
                },
                {
                    "role": "user",
                    "content": user_prompt,
                },
            ],
        )

        result = clean_ai_text(
            response.choices[0].message.content
        )

        if document_type in [
            "Email",
            "Letter",
        ]:

            signature = profile_signature(
                profile
            )

            if signature:
                result += signature

        return result

    except Exception as e:

        st.error(
            f"AI generation failed: {e}"
        )

        return ""


# ============================================================
# INQUIRY HELPERS
# ============================================================

def documents_text(documents):

    if not documents:

        return (
            "No information was provided "
            "for this index."
        )

    result = []

    for number, document in enumerate(
        documents
    ):

        annex_letter = chr(
            ord("A") + number
        )

        result.append(
            f"Annex-{annex_letter}: {document}"
        )

    return "\n".join(result)


def committee_text(committee):

    result = []

    for member in committee:

        role = member.get(
            "role",
            "",
        ).strip()

        erp = member.get(
            "erp",
            "",
        ).strip()

        name = member.get(
            "name",
            "",
        ).strip()

        designation = member.get(
            "designation",
            "",
        ).strip()

        if not (
            erp
            or name
            or designation
        ):
            continue

        parts = []

        if erp:
            parts.append(
                f"ERP#: {erp}"
            )

        if name:
            parts.append(
                f"Name: {name}"
            )

        if designation:
            parts.append(
                f"Designation: {designation}"
            )

        result.append(
            f"{role} — "
            + ", ".join(parts)
        )

    if not result:

        return (
            "No information was provided "
            "for this index."
        )

    return "\n".join(result)


def qa_markdown(rows):

    result = [
        "| Questions | Answers |",
        "|---|---|",
    ]

    number = 1

    for row in rows:

        question = (
            row.get(
                "question",
                "",
            )
            or ""
        ).strip()

        answer = (
            row.get(
                "answer",
                "",
            )
            or ""
        ).strip()

        if not question and not answer:
            continue

        question = question.replace(
            "|",
            "\\|",
        )

        answer = answer.replace(
            "|",
            "\\|",
        )

        result.append(
            f"| {number}. {question} | {answer} |"
        )

        number += 1

    if len(result) == 2:

        result.append(
            "| No information was provided for this index. | |"
        )

    return "\n".join(result)


def generate_inquiry_report(
    index_data,
    inquiry_type,
):

    client = get_groq_client()

    if client is None:
        return ""

    sections = []

    for item in index_data:

        name = item.get(
            "name",
            "",
        )

        if name == "Documents Recorded":

            content = documents_text(
                item.get(
                    "documents",
                    [],
                )
            )

        elif name == "Inquiry Committee":

            content = committee_text(
                item.get(
                    "committee",
                    [],
                )
            )

        elif name == (
            "Questions / Answers with the Accused"
        ):

            content = qa_markdown(
                item.get(
                    "qa_rows",
                    [],
                )
            )

        else:

            content = (
                item.get(
                    "content",
                    "",
                )
                or ""
            ).strip()

            if not content:

                content = (
                    "No information was provided "
                    "for this index."
                )

        sections.append(
            {
                "name": name,
                "content": content,
            }
        )

    source = []

    for section in sections:

        source.append(
            f"""
INDEX:
{section["name"]}

USER INFORMATION:
{section["content"]}
"""
        )

    system_prompt = f"""
You are DraftForge preparing an official
{inquiry_type} inquiry report.

VERY IMPORTANT:

Generate ONLY the indexes supplied by the user.

Never create additional indexes.

Do NOT add Introduction, Background, Summary of Evidence,
Findings, Conclusion, Recommendations, Inquiry Committee,
Documents Recorded or any other section unless that exact
index was supplied.

Preserve the exact order of the selected indexes.

Do not invent facts, names, dates, allegations, evidence,
findings or recommendations.

Correct spelling, grammar, punctuation and obvious
voice-transcription errors.

For Questions / Answers with the Accused, preserve this
format:

| Questions | Answers |
|---|---|
| 1. Question | Answer |

For Documents Recorded, preserve Annex-A, Annex-B,
Annex-C etc.

For Inquiry Committee, preserve all supplied roles
and details.

Return ONLY the finished inquiry report.
"""

    user_prompt = f"""
The user selected ONLY these indexes:

{"".join(source)}

Generate the final report.

Remember:

ONLY selected indexes.
SAME ORDER.
NO INVENTED SECTIONS.
"""

    try:

        response = client.chat.completions.create(
            model=GROQ_MODEL,
            temperature=0.10,
            messages=[
                {
                    "role": "system",
                    "content": system_prompt,
                },
                {
                    "role": "user",
                    "content": user_prompt,
                },
            ],
        )

        return clean_ai_text(
            response.choices[0].message.content
        )

    except Exception as e:

        st.error(
            f"Inquiry generation failed: {e}"
        )

        return ""


def ai_edit_document(
    document,
    instruction,
):

    client = get_groq_client()

    if client is None:
        return ""

    system_prompt = """
You are DraftForge's document editing assistant.

Edit the document according to the user's instruction.

Preserve meaning unless the user explicitly requests a
change.

Do not invent facts, names, dates, evidence or events.

Correct grammar and spelling when appropriate.

Return ONLY the complete revised document.
"""

    user_prompt = f"""
CURRENT DOCUMENT:

{document}

EDITING INSTRUCTION:

{instruction}

Return the complete revised document.
"""

    try:

        response = client.chat.completions.create(
            model=GROQ_MODEL,
            temperature=0.10,
            messages=[
                {
                    "role": "system",
                    "content": system_prompt,
                },
                {
                    "role": "user",
                    "content": user_prompt,
                },
            ],
        )

        return clean_ai_text(
            response.choices[0].message.content
        )

    except Exception as e:

        st.error(
            f"AI editing failed: {e}"
        )

        return ""


# ============================================================
# HISTORY
# ============================================================

def save_history(
    title,
    document,
):

    if not document:
        return

    record = {
        "title": title,
        "document": document,
        "created": datetime.now().strftime(
            "%Y-%m-%d %H:%M"
        ),
    }

    history = st.session_state.history

    history.append(record)

    st.session_state.history = history[-50:]

    save_json(
        HISTORY_FILE,
        st.session_state.history,
    )


# ============================================================
# EXPORT HELPERS
# ============================================================

def safe_pdf_text(text):

    return (
        str(text)
        .replace("–", "-")
        .replace("—", "-")
        .replace("’", "'")
        .replace("“", '"')
        .replace("”", '"')
        .replace("•", "-")
        .encode(
            "latin-1",
            "replace",
        )
        .decode(
            "latin-1"
        )
    )


def create_pdf(text):

    if FPDF is None:

        st.error(
            "FPDF is not installed."
        )

        return None

    try:

        pdf = FPDF()

        pdf.set_auto_page_break(
            auto=True,
            margin=18,
        )

        pdf.add_page()

        left = 15
        right = 15

        pdf.set_left_margin(left)
        pdf.set_right_margin(right)

        usable = (
            pdf.w
            - left
            - right
        )

        pdf.set_font(
            "Arial",
            "B",
            15,
        )

        pdf.cell(
            usable,
            10,
            "DraftForge Document",
            ln=True,
            align="C",
        )

        pdf.ln(5)

        pdf.set_font(
            "Arial",
            "",
            10,
        )

        for line in text.splitlines():

            if not line.strip():

                pdf.ln(4)

                continue

            if (
                "|" in line
                and line.count("|") >= 2
                and not re.match(
                    r"^\s*\|?\s*:?-+:?\s*\|",
                    line,
                )
            ):

                parts = [
                    x.strip()
                    for x in line.strip(
                        "|"
                    ).split("|")
                ]

                if len(parts) >= 2:

                    col1 = usable * 0.42
                    col2 = usable * 0.58

                    x = left
                    y = pdf.get_y()

                    pdf.set_xy(
                        x,
                        y,
                    )

                    pdf.multi_cell(
                        col1,
                        7,
                        safe_pdf_text(
                            parts[0]
                        ),
                        border=1,
                    )

                    row_height = (
                        pdf.get_y()
                        - y
                    )

                    pdf.set_xy(
                        x + col1,
                        y,
                    )

                    pdf.multi_cell(
                        col2,
                        7,
                        safe_pdf_text(
                            parts[1]
                        ),
                        border=1,
                    )

                    continue

            wrapped = textwrap.wrap(
                line,
                width=100,
                break_long_words=True,
                break_on_hyphens=True,
            )

            for wrapped_line in wrapped:

                pdf.set_x(left)

                pdf.multi_cell(
                    usable,
                    6,
                    safe_pdf_text(
                        wrapped_line
                    ),
                )

        return bytes(
            pdf.output(
                dest="S"
            )
        )

    except Exception as e:

        st.error(
            f"PDF export failed: {e}"
        )

        return None


def create_docx(text):

    if Document is None:

        st.error(
            "python-docx is not installed."
        )

        return None

    try:

        doc = Document()

        section = doc.sections[0]

        section.top_margin = Inches(0.7)
        section.bottom_margin = Inches(0.7)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

        lines = text.splitlines()

        i = 0

        while i < len(lines):

            line = lines[i]

            if (
                "|" in line
                and i + 1 < len(lines)
                and "|" in lines[i + 1]
                and re.match(
                    r"^\|?\s*:?-+:?\s*\|",
                    lines[i + 1],
                )
            ):

                rows = []

                while (
                    i < len(lines)
                    and "|" in lines[i]
                ):

                    current = lines[i].strip()

                    if re.match(
                        r"^\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)+\|?$",
                        current,
                    ):

                        i += 1
                        continue

                    parts = [
                        p.strip()
                        for p in current.strip(
                            "|"
                        ).split("|")
                    ]

                    if len(parts) >= 2:

                        rows.append(
                            parts[:2]
                        )

                    i += 1

                if rows:

                    table = doc.add_table(
                        rows=len(rows),
                        cols=2,
                    )

                    table.style = "Table Grid"

                    table.alignment = (
                        WD_TABLE_ALIGNMENT.CENTER
                    )

                    for r, row in enumerate(
                        rows
                    ):

                        for c in range(2):

                            cell = table.cell(
                                r,
                                c,
                            )

                            cell.text = (
                                row[c]
                                if c < len(row)
                                else ""
                            )

                            for paragraph in (
                                cell.paragraphs
                            ):

                                for run in (
                                    paragraph.runs
                                ):

                                    run.font.name = (
                                        "Arial"
                                    )

                                    run.font.size = (
                                        Pt(10)
                                    )

                    doc.add_paragraph()

                    continue

            paragraph = doc.add_paragraph()

            run = paragraph.add_run(line)

            run.font.name = "Arial"
            run.font.size = Pt(11)

            i += 1

        output = io.BytesIO()

        doc.save(output)

        output.seek(0)

        return output.getvalue()

    except Exception as e:

        st.error(
            f"DOCX export failed: {e}"
        )

        return None


def create_png(text):

    if Image is None:

        st.error(
            "Pillow is not installed."
        )

        return None

    try:

        font = None

        candidates = [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
        ]

        for path in candidates:

            if os.path.exists(path):

                font = ImageFont.truetype(
                    path,
                    26,
                )

                break

        if font is None:

            font = ImageFont.load_default()

        lines = []

        for line in text.splitlines():

            wrapped = textwrap.wrap(
                line,
                width=80,
                break_long_words=True,
            )

            if wrapped:

                lines.extend(wrapped)

            else:

                lines.append("")

        width = 1400

        line_height = 40

        height = (
            len(lines) * line_height
            + 150
        )

        image = Image.new(
            "RGB",
            (width, height),
            "white",
        )

        draw = ImageDraw.Draw(image)

        draw.text(
            (50, 35),
            "DraftForge",
            font=font,
            fill="black",
        )

        y = 90

        for line in lines:

            draw.text(
                (50, y),
                line,
                font=font,
                fill="black",
            )

            y += line_height

        output = io.BytesIO()

        image.save(
            output,
            format="PNG",
        )

        output.seek(0)

        return output.getvalue()

    except Exception as e:

        st.error(
            f"PNG export failed: {e}"
        )

        return None


# ============================================================
# HEADER
# ============================================================

st.title("✦ DraftForge")

st.caption(
    "AI Document Composer — create professional official "
    "documents faster"
)

st.success(
    "● AI Document Workspace"
)


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.title("✦ DraftForge")

    st.caption(
        "Official document workspace"
    )

    st.divider()

    st.subheader(
        "💡 Tips & Templates"
    )

    st.write(
        "Write naturally. DraftForge will convert your "
        "instructions into professional official English."
    )

    st.write(
        "🎤 You can speak instead of typing."
    )

    st.write(
        "📄 Inquiry sections can be added in any order."
    )

    st.write(
        "🔁 The same inquiry index can be added multiple times."
    )

    st.divider()

    if st.button(
        "📚 My Documents",
        use_container_width=True,
    ):

        st.session_state.show_history = not (
            st.session_state.show_history
        )

        st.session_state.show_profile = False

        st.rerun()

    if st.button(
        "👤 My Profile",
        use_container_width=True,
    ):

        st.session_state.show_profile = not (
            st.session_state.show_profile
        )

        st.session_state.show_history = False

        st.rerun()

    st.divider()

    st.subheader(
        "✦ About DraftForge"
    )

    st.caption(
        "AI-assisted drafting workspace for professional "
        "official correspondence and inquiry documentation."
    )

    st.caption(
        "Supported: Email • Letter • E&D Inquiry • FFI Inquiry"
    )


# ============================================================
# PROFILE
# ============================================================

if st.session_state.show_profile:

    st.header(
        "👤 My Profile"
    )

    st.caption(
        "This information is automatically appended to "
        "Email and Letter documents."
    )

    profile = st.session_state.profile

    name = st.text_input(
        "Name",
        value=profile.get(
            "Name",
            "",
        ),
        key="profile_name",
    )

    designation = st.text_input(
        "Designation",
        value=profile.get(
            "Designation",
            "",
        ),
        key="profile_designation",
    )

    contact = st.text_input(
        "Contact No.",
        value=profile.get(
            "Contact No.",
            "",
        ),
        key="profile_contact",
    )

    station = st.text_input(
        "Current Station",
        value=profile.get(
            "Current Station",
            "",
        ),
        key="profile_station",
    )

    if st.button(
        "💾 Save Profile",
        type="primary",
    ):

        st.session_state.profile = {
            "Name": name,
            "Designation": designation,
            "Contact No.": contact,
            "Current Station": station,
        }

        save_json(
            PROFILE_FILE,
            st.session_state.profile,
        )

        st.success(
            "Profile saved successfully."
        )


# ============================================================
# HISTORY
# ============================================================

if st.session_state.show_history:

    st.header(
        "📚 My Documents"
    )

    if not st.session_state.history:

        st.info(
            "No saved documents yet."
        )

    else:

        for n, record in enumerate(
            reversed(
                st.session_state.history
            )
        ):

            title = record.get(
                "title",
                "Untitled",
            )

            created = record.get(
                "created",
                "",
            )

            with st.expander(
                f"📄 {title} — {created}"
            ):

                st.text(
                    record.get(
                        "document",
                        "",
                    )
                )

                if st.button(
                    "Open Document",
                    key=f"open_history_{n}",
                ):

                    original = record.get(
                        "document",
                        "",
                    )

                    st.session_state.generated_draft = original
                    st.session_state.editable_draft = original
                    st.session_state.editor_sync = original

                    st.session_state.show_history = False

                    st.rerun()


# ============================================================
# CHOOSE DOCUMENT
# ============================================================

st.header(
    "① Choose your document"
)

st.caption(
    "Select the type of document you want DraftForge to create."
)

doc_columns = st.columns(3)

cards = [
    (
        "📧",
        "Email",
        "Professional official email",
    ),
    (
        "📄",
        "Letter",
        "Formal official correspondence",
    ),
    (
        "🔎",
        "Inquiry",
        "E&D / FFI inquiry documents",
    ),
]

for column, card in zip(
    doc_columns,
    cards,
):

    icon, title, description = card

    with column:

        st.subheader(
            f"{icon} {title}"
        )

        st.caption(
            description
        )

        if (
            st.session_state.document_type
            == title
        ):

            st.success(
                "✓ Selected"
            )

        else:

            if st.button(
                f"Select {title}",
                key=f"choose_{title}",
                use_container_width=True,
            ):

                st.session_state.document_type = title

                if title != "Inquiry":

                    st.session_state.index_data = []

                st.rerun()


# ============================================================
# EMAIL
# ============================================================

if st.session_state.document_type == "Email":

    st.header(
        "② Provide your information"
    )

    st.caption(
        "Type naturally, speak naturally, or combine both."
    )

    st.info(
        "💡 Tell DraftForge what the email should say. "
        "You do not need to write perfect English."
    )

    st.subheader(
        "🎤 Voice input"
    )

    st.caption(
        "You can speak naturally. Your transcription will be "
        "added to the text below."
    )

    audio = st.audio_input(
        "Record voice",
        key="email_audio",
    )

    if audio is not None:

        try:

            audio_bytes = audio.getvalue()

            audio_hash = hashlib.sha256(
                audio_bytes
            ).hexdigest()

            if (
                audio_hash
                != st.session_state.email_voice_hash
            ):

                client = get_groq_client()

                if client:

                    transcription = (
                        client.audio.transcriptions.create(
                            file=(
                                "voice.wav",
                                audio_bytes,
                            ),
                            model=WHISPER_MODEL,
                        )
                    )

                    transcript = (
                        getattr(
                            transcription,
                            "text",
                            "",
                        )
                        or ""
                    ).strip()

                    if transcript:

                        existing = (
                            st.session_state.email_instruction
                            or ""
                        ).strip()

                        if existing:

                            st.session_state.email_instruction = (
                                existing
                                + "\n"
                                + transcript
                            )

                        else:

                            st.session_state.email_instruction = (
                                transcript
                            )

                    st.session_state.email_voice_hash = (
                        audio_hash
                    )

        except Exception as e:

            st.error(
                f"Voice transcription failed: {e}"
            )

    email_instruction = st.text_area(
        "Email instructions / information",
        key="email_instruction",
        height=180,
        placeholder=(
            "Example: Inform the regional office that "
            "the backup internet connection is unavailable "
            "and request immediate restoration."
        ),
    )

    st.caption(
        "DraftForge will correct grammar and convert your "
        "natural-language instructions into professional "
        "official English."
    )

    st.divider()

    if st.button(
        "✦ Generate Email",
        type="primary",
        use_container_width=True,
    ):

        if not email_instruction.strip():

            st.warning(
                "Please provide some information first."
            )

        else:

            with st.spinner(
                "DraftForge is preparing your email..."
            ):

                result = generate_ai_document(
                    "Email",
                    email_instruction,
                    st.session_state.profile,
                )

            if result:

                st.session_state.generated_draft = result
                st.session_state.editable_draft = result
                st.session_state.editor_sync = result

                save_history(
                    "Email",
                    result,
                )

                st.rerun()


# ============================================================
# LETTER
# ============================================================

elif st.session_state.document_type == "Letter":

    st.header(
        "② Provide your information"
    )

    st.caption(
        "Type naturally, speak naturally, or combine both."
    )

    st.info(
        "💡 Tell DraftForge what the letter should say. "
        "You do not need to write perfect English."
    )

    st.subheader(
        "🎤 Voice input"
    )

    st.caption(
        "You can speak naturally. Your transcription will be "
        "added to the text below."
    )

    audio = st.audio_input(
        "Record voice",
        key="letter_audio",
    )

    if audio is not None:

        try:

            audio_bytes = audio.getvalue()

            audio_hash = hashlib.sha256(
                audio_bytes
            ).hexdigest()

            if (
                audio_hash
                != st.session_state.letter_voice_hash
            ):

                client = get_groq_client()

                if client:

                    transcription = (
                        client.audio.transcriptions.create(
                            file=(
                                "voice.wav",
                                audio_bytes,
                            ),
                            model=WHISPER_MODEL,
                        )
                    )

                    transcript = (
                        getattr(
                            transcription,
                            "text",
                            "",
                        )
                        or ""
                    ).strip()

                    if transcript:

                        existing = (
                            st.session_state.letter_instruction
                            or ""
                        ).strip()

                        if existing:

                            st.session_state.letter_instruction = (
                                existing
                                + "\n"
                                + transcript
                            )

                        else:

                            st.session_state.letter_instruction = (
                                transcript
                            )

                    st.session_state.letter_voice_hash = (
                        audio_hash
                    )

        except Exception as e:

            st.error(
                f"Voice transcription failed: {e}"
            )

    letter_instruction = st.text_area(
        "Letter instructions / information",
        key="letter_instruction",
        height=180,
        placeholder=(
            "Example: Write a formal letter requesting "
            "immediate repair of the office backup connection."
        ),
    )

    st.caption(
        "DraftForge will convert your natural-language "
        "instructions into professional official English."
    )

    st.divider()

    if st.button(
        "✦ Generate Letter",
        type="primary",
        use_container_width=True,
    ):

        if not letter_instruction.strip():

            st.warning(
                "Please provide some information first."
            )

        else:

            with st.spinner(
                "DraftForge is preparing your letter..."
            ):

                result = generate_ai_document(
                    "Letter",
                    letter_instruction,
                    st.session_state.profile,
                )

            if result:

                st.session_state.generated_draft = result
                st.session_state.editable_draft = result
                st.session_state.editor_sync = result

                save_history(
                    "Letter",
                    result,
                )

                st.rerun()


# ============================================================
# INQUIRY
# ============================================================

elif st.session_state.document_type == "Inquiry":

    st.header(
        "② Build your inquiry"
    )

    st.caption(
        "Select the inquiry type and add only the sections "
        "you actually need."
    )

    inquiry_type = st.selectbox(
        "Inquiry type",
        INQUIRY_TYPES,
        index=INQUIRY_TYPES.index(
            st.session_state.inquiry_type
        ),
    )

    if inquiry_type != st.session_state.inquiry_type:

        st.session_state.inquiry_type = inquiry_type
        st.session_state.index_data = []
        st.session_state.generated_draft = ""
        st.session_state.editable_draft = ""

        st.rerun()

    if inquiry_type == "FFI Inquiry":

        st.warning(
            "FFI Inquiry is currently under construction. "
            "The E&D Inquiry workflow is available."
        )

    st.subheader(
        "➕ Add inquiry section"
    )

    c1, c2 = st.columns(
        [4, 1]
    )

    with c1:

        selected_index = st.selectbox(
            "Select section",
            ED_INDEXES,
            key="new_inquiry_index",
        )

    with c2:

        st.write("")

        if st.button(
            "＋ Add",
            use_container_width=True,
        ):

            new_item = {
                "name": selected_index,
                "content": "",
            }

            if selected_index == (
                "Documents Recorded"
            ):

                new_item["documents"] = []

            elif selected_index == (
                "Inquiry Committee"
            ):

                new_item["committee"] = [
                    {
                        "role": role,
                        "erp": "",
                        "name": "",
                        "designation": "",
                    }
                    for role in COMMITTEE_ROLES
                ]

            elif selected_index == (
                "Questions / Answers with the Accused"
            ):

                new_item["qa_rows"] = [
                    {
                        "question": "",
                        "answer": "",
                    }
                ]

            st.session_state.index_data.append(
                new_item
            )

            st.rerun()

    # ========================================================
    # SELECTED INDEXES
    # ========================================================

    if st.session_state.index_data:

        st.subheader(
            "Selected inquiry sections"
        )

        st.caption(
            "Complete only the sections you selected. "
            "They will appear in the same order."
        )

        occurrence_counter = {}

        for position, item in enumerate(
            st.session_state.index_data
        ):

            base_name = item["name"]

            occurrence_counter[
                base_name
            ] = occurrence_counter.get(
                base_name,
                0,
            ) + 1

            occurrence = occurrence_counter[
                base_name
            ]

            display_name = base_name

            if (
                base_name in [
                    "Statement of the Accused",
                    "Questions / Answers with the Accused",
                ]
                and occurrence > 1
            ):

                display_name = (
                    f"{base_name} "
                    f"No. {occurrence}"
                )

            st.markdown("---")

            st.subheader(
                display_name
            )

            # =================================================
            # DOCUMENTS RECORDED
            # =================================================

            if base_name == (
                "Documents Recorded"
            ):

                st.caption(
                    "Select the documents that were recorded "
                    "or examined."
                )

                documents = st.multiselect(
                    "Documents",
                    DOCUMENTS_RECORDED,
                    default=item.get(
                        "documents",
                        [],
                    ),
                    key=f"documents_{position}",
                )

                item["documents"] = documents

                if st.button(
                    "🗑 Remove this section",
                    key=f"remove_documents_{position}",
                ):

                    del st.session_state.index_data[
                        position
                    ]

                    st.rerun()

                continue

            # =================================================
            # INQUIRY COMMITTEE
            # =================================================

            if base_name == "Inquiry Committee":

                st.markdown(
                    """
                    <div class="committee-tip">
                        <b>How to complete this section:</b>
                        Each shaded box below represents a different
                        committee role. Enter the ERP#, Name and
                        Designation inside the box belonging to that role.
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                committee = item.get(
                    "committee",
                    [],
                )

                # Safety conversion for older saved session data.
                if not isinstance(
                    committee,
                    list,
                ):

                    committee = []

                # Make sure all four roles exist.
                existing_roles = {
                    member.get(
                        "role",
                        "",
                    )
                    for member in committee
                    if isinstance(
                        member,
                        dict,
                    )
                }

                for role in COMMITTEE_ROLES:

                    if role not in existing_roles:

                        committee.append(
                            {
                                "role": role,
                                "erp": "",
                                "name": "",
                                "designation": "",
                            }
                        )

                for member_index, member in enumerate(
                    committee
                ):

                    role = member.get(
                        "role",
                        "",
                    )

                    if role not in COMMITTEE_ROLES:
                        continue

                    # -------------------------------------------------
                    # ROLE CARD
                    # -------------------------------------------------

                    st.markdown(
                        f"""
                        <div class="committee-role-card">
                            <div class="committee-role-title">
                                👤 {role}
                            </div>
                            <div class="committee-role-description">
                                Enter the details of the person serving in this role.
                            </div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

                    role_key = (
                        role.lower()
                        .replace(
                            " ",
                            "_",
                        )
                        .replace(
                            "#",
                            "no",
                        )
                    )

                    col1, col2, col3 = st.columns(3)

                    with col1:

                        erp = st.text_input(
                            "ERP#",
                            value=member.get(
                                "erp",
                                "",
                            ),
                            key=(
                                f"committee_"
                                f"{position}_"
                                f"{role_key}_erp"
                            ),
                            placeholder="Enter ERP#",
                        )

                    with col2:

                        name = st.text_input(
                            "Name",
                            value=member.get(
                                "name",
                                "",
                            ),
                            key=(
                                f"committee_"
                                f"{position}_"
                                f"{role_key}_name"
                            ),
                            placeholder="Enter full name",
                        )

                    with col3:

                        designation = st.text_input(
                            "Designation",
                            value=member.get(
                                "designation",
                                "",
                            ),
                            key=(
                                f"committee_"
                                f"{position}_"
                                f"{role_key}_designation"
                            ),
                            placeholder="Enter designation",
                        )

                    member["erp"] = erp
                    member["name"] = name
                    member["designation"] = designation

                item["committee"] = committee

                st.info(
                    "Each role is shown in its own shaded box. "
                    "For example, the ERP#, Name and Designation "
                    "directly below “Convener of Inquiry” belong "
                    "to the Convener."
                )

                if st.button(
                    "🗑 Remove this section",
                    key=f"remove_committee_{position}",
                ):

                    del st.session_state.index_data[
                        position
                    ]

                    st.rerun()

                continue

            # =================================================
            # QUESTIONS / ANSWERS
            # =================================================

            if base_name == (
                "Questions / Answers with the Accused"
            ):

                rows = item.get(
                    "qa_rows",
                    [],
                )

                if not rows:

                    rows.append(
                        {
                            "question": "",
                            "answer": "",
                        }
                    )

                for row_index, row in enumerate(
                    rows
                ):

                    q_col, a_col = st.columns(2)

                    with q_col:

                        row["question"] = st.text_area(
                            f"Question {row_index + 1}",
                            value=row.get(
                                "question",
                                "",
                            ),
                            key=(
                                f"question_"
                                f"{position}_"
                                f"{row_index}"
                            ),
                            height=100,
                            placeholder="Enter question...",
                        )

                    with a_col:

                        row["answer"] = st.text_area(
                            f"Answer {row_index + 1}",
                            value=row.get(
                                "answer",
                                "",
                            ),
                            key=(
                                f"answer_"
                                f"{position}_"
                                f"{row_index}"
                            ),
                            height=100,
                            placeholder="Enter answer...",
                        )

                item["qa_rows"] = rows

                if st.button(
                    "＋ Add Question / Answer",
                    key=f"add_qa_{position}",
                ):

                    rows.append(
                        {
                            "question": "",
                            "answer": "",
                        }
                    )

                    item["qa_rows"] = rows

                    st.rerun()

            else:

                # =================================================
                # VOICE + TEXT
                # =================================================

                text_key = (
                    f"inquiry_text_{position}"
                )

                audio_key = (
                    f"inquiry_audio_{position}"
                )

                hash_key = (
                    f"inquiry_voice_hash_{position}"
                )

                if text_key not in st.session_state:

                    st.session_state[
                        text_key
                    ] = item.get(
                        "content",
                        "",
                    )

                st.caption(
                    "🎤 Voice input"
                )

                audio = st.audio_input(
                    "Record voice",
                    key=audio_key,
                )

                # Process voice BEFORE text widget.

                if audio is not None:

                    try:

                        audio_bytes = (
                            audio.getvalue()
                        )

                        audio_hash = hashlib.sha256(
                            audio_bytes
                        ).hexdigest()

                        if (
                            audio_hash
                            != st.session_state.get(
                                hash_key,
                                "",
                            )
                        ):

                            client = get_groq_client()

                            if client:

                                transcription = (
                                    client.audio.transcriptions.create(
                                        file=(
                                            "voice.wav",
                                            audio_bytes,
                                        ),
                                        model=WHISPER_MODEL,
                                    )
                                )

                                transcript = (
                                    getattr(
                                        transcription,
                                        "text",
                                        "",
                                    )
                                    or ""
                                ).strip()

                                if transcript:

                                    existing = (
                                        st.session_state.get(
                                            text_key,
                                            "",
                                        )
                                        or ""
                                    ).strip()

                                    if existing:

                                        st.session_state[
                                            text_key
                                        ] = (
                                            existing
                                            + "\n"
                                            + transcript
                                        )

                                    else:

                                        st.session_state[
                                            text_key
                                        ] = transcript

                                st.session_state[
                                    hash_key
                                ] = audio_hash

                    except Exception as e:

                        st.error(
                            f"Voice transcription failed: {e}"
                        )

                current_content = st.text_area(
                    display_name,
                    key=text_key,
                    height=180,
                    placeholder=(
                        "Enter information for this section "
                        "in your own words..."
                    ),
                )

                item["content"] = current_content

            # =================================================
            # REMOVE NORMAL SECTION
            # =================================================

            if st.button(
                "🗑 Remove this section",
                key=f"remove_section_{position}",
            ):

                del st.session_state.index_data[
                    position
                ]

                st.rerun()

    else:

        st.info(
            "No inquiry sections added yet. "
            "Use the selector above to add the sections "
            "you need."
        )

    # ========================================================
    # FINAL GENERATE BUTTON
    # ========================================================

    st.divider()

    if st.button(
        "✦ Generate Inquiry Report",
        type="primary",
        use_container_width=True,
    ):

        if not st.session_state.index_data:

            st.warning(
                "Please add at least one inquiry section."
            )

        else:

            with st.spinner(
                "DraftForge is preparing your inquiry report..."
            ):

                result = generate_inquiry_report(
                    st.session_state.index_data,
                    st.session_state.inquiry_type,
                )

            if result:

                st.session_state.generated_draft = result
                st.session_state.editable_draft = result
                st.session_state.editor_sync = result

                save_history(
                    "Inquiry Report",
                    result,
                )

                st.rerun()


# ============================================================
# GENERATED DOCUMENT WORKSPACE
# ============================================================

if st.session_state.generated_draft:

    st.divider()

    st.header(
        "③ Generate & Export"
    )

    st.caption(
        "Review, edit and export your generated document."
    )

    # ========================================================
    # EDITOR SYNC BEFORE WIDGET
    # ========================================================

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
        "Document Editor",
        key="document_editor",
        height=550,
    )

    # ========================================================
    # SAVE / RESTORE
    # ========================================================

    col1, col2 = st.columns(2)

    with col1:

        if st.button(
            "💾 Save Changes",
            type="primary",
            use_container_width=True,
        ):

            st.session_state.editable_draft = (
                edited_document
            )

            st.session_state.generated_draft = (
                edited_document
            )

            st.success(
                "Changes saved."
            )

    with col2:

        if st.button(
            "↩ Restore Original",
            use_container_width=True,
        ):

            original = (
                st.session_state.generated_draft
            )

            st.session_state.editable_draft = original
            st.session_state.editor_sync = original

            st.rerun()

    # ========================================================
    # AI EDIT
    # ========================================================

    st.subheader(
        "✦ AI Editing Assistant"
    )

    if st.session_state.get(
        "edit_instruction_sync",
        "",
    ):

        st.session_state.edit_instruction = (
            st.session_state.pop(
                "edit_instruction_sync"
            )
        )

    edit_instruction = st.text_area(
        "Tell DraftForge what you want to change",
        key="edit_instruction",
        height=120,
        placeholder=(
            "Example: Make this more concise and formal."
        ),
    )

    if st.button(
        "✦ Apply AI Edit",
        use_container_width=True,
    ):

        if not edit_instruction.strip():

            st.warning(
                "Please describe the change you want."
            )

        else:

            with st.spinner(
                "Applying requested changes..."
            ):

                modified = ai_edit_document(
                    edited_document,
                    edit_instruction,
                )

            if modified:

                st.session_state.editable_draft = modified
                st.session_state.generated_draft = modified
                st.session_state.editor_sync = modified
                st.session_state.edit_instruction_sync = ""

                st.rerun()

    # ========================================================
    # EXPORT
    # ========================================================

    st.subheader(
        "📤 Export"
    )

    export1, export2, export3, export4 = st.columns(4)

    current_document = (
        st.session_state.editable_draft
        or edited_document
    )

    with export1:

        if st.button(
            "📄 PDF",
            use_container_width=True,
        ):

            pdf = create_pdf(
                current_document
            )

            if pdf:

                st.download_button(
                    "⬇ Download PDF",
                    data=pdf,
                    file_name="DraftForge_Document.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )

    with export2:

        if st.button(
            "📝 DOCX",
            use_container_width=True,
        ):

            docx = create_docx(
                current_document
            )

            if docx:

                st.download_button(
                    "⬇ Download DOCX",
                    data=docx,
                    file_name="DraftForge_Document.docx",
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.wordprocessingml.document"
                    ),
                    use_container_width=True,
                )

    with export3:

        st.download_button(
            "⬇ TXT",
            data=current_document,
            file_name="DraftForge_Document.txt",
            mime="text/plain",
            use_container_width=True,
        )

    with export4:

        if Image is not None:

            if st.button(
                "🖼 PNG",
                use_container_width=True,
            ):

                png = create_png(
                    current_document
                )

                if png:

                    st.download_button(
                        "⬇ Download PNG",
                        data=png,
                        file_name="DraftForge_Document.png",
                        mime="image/png",
                        use_container_width=True,
                    )


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "✦ DraftForge — AI Document Composer"
)

st.caption(
    "Developed by: Raees Khan — Assistant Director, NADRA"
                )
