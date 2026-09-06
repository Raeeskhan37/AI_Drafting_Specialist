import os
import io
import re
import json
import hashlib
import textwrap
from datetime import datetime

import streamlit as st

try:
    from groq import Groq
except Exception:
    Groq = None

try:
    from docx import Document
    from docx.shared import Pt, Inches
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
# CONSTANTS
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
# DEFAULT SESSION STATE
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
    "show_about": False,
}

for key, value in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# CSS
# ============================================================

st.markdown(
    """
    <style>
    .main .block-container {
        max-width: 1250px;
        padding-top: 1.5rem;
        padding-bottom: 4rem;
    }

    [data-testid="stSidebar"] {
        border-right: 1px solid rgba(128,128,128,.18);
    }

    .hero-box {
        padding: 1.5rem 1.7rem;
        border: 1px solid rgba(128,128,128,.20);
        border-radius: 18px;
        margin-bottom: 1.2rem;
        background: linear-gradient(
            135deg,
            rgba(120,120,120,.08),
            rgba(120,120,120,.025)
        );
    }

    .hero-title {
        font-size: 2rem;
        font-weight: 750;
        margin-bottom: .25rem;
    }

    .hero-subtitle {
        font-size: 1rem;
        opacity: .75;
    }

    .footer-note {
        text-align: center;
        opacity: .6;
        font-size: .82rem;
        margin-top: 2rem;
    }

    div[data-testid="stButton"] button {
        border-radius: 10px;
        min-height: 42px;
    }

    div[data-testid="stDownloadButton"] button {
        border-radius: 10px;
        min-height: 42px;
    }

    textarea {
        font-size: 1rem !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# FILE HELPERS
# ============================================================

def load_json(filename, default):
    try:
        if os.path.exists(filename):
            with open(filename, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass

    return default


def save_json(filename, data):
    try:
        with open(filename, "w", encoding="utf-8") as f:
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
# PROFILE
# ============================================================

def profile_signature(profile):
    if not isinstance(profile, dict):
        return ""

    name = str(
        profile.get("Name", "") or ""
    ).strip()

    designation = str(
        profile.get("Designation", "") or ""
    ).strip()

    contact = str(
        profile.get("Contact No.", "") or ""
    ).strip()

    station = str(
        profile.get("Current Station", "") or ""
    ).strip()

    parts = []

    if name:
        parts.append(name)

    if designation:
        parts.append(designation)

    if contact:
        parts.append(
            f"Contact No.: {contact}"
        )

    if station:
        parts.append(
            f"Current Station: {station}"
        )

    return "\n".join(parts)


# ============================================================
# GROQ
# ============================================================

def get_groq_client():
    if Groq is None:
        return None

    api_key = None

    try:
        api_key = st.secrets.get(
            "GROQ_API_KEY"
        )
    except Exception:
        pass

    if not api_key:
        api_key = os.getenv(
            "GROQ_API_KEY"
        )

    if not api_key:
        return None

    try:
        return Groq(
            api_key=api_key
        )
    except Exception:
        return None


# ============================================================
# AI TEXT CLEANING
# ============================================================

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


# ============================================================
# VOICE TRANSCRIPTION
# ============================================================

def transcribe_audio(
    audio_bytes,
    filename="voice.wav",
):
    client = get_groq_client()

    if client is None:
        st.error(
            "Groq API key is not configured."
        )
        return ""

    try:
        audio_file = io.BytesIO(
            audio_bytes
        )

        audio_file.name = filename

        result = client.audio.transcriptions.create(
            file=(
                filename,
                audio_file,
            ),
            model=WHISPER_MODEL,
            response_format="text",
        )

        if isinstance(result, str):
            return result.strip()

        if hasattr(result, "text"):
            return str(
                result.text
            ).strip()

        return str(result).strip()

    except Exception as e:
        st.error(
            f"Voice transcription failed: {e}"
        )
        return ""


def process_voice_input(
    audio,
    state_key,
):
    if audio is None:
        return ""

    try:
        audio_bytes = audio.getvalue()
    except Exception:
        return ""

    if not audio_bytes:
        return ""

    audio_hash = hashlib.sha256(
        audio_bytes
    ).hexdigest()

    if (
        st.session_state.get(
            state_key
        )
        == audio_hash
    ):
        return ""

    st.session_state[state_key] = (
        audio_hash
    )

    return transcribe_audio(
        audio_bytes
    )


# ============================================================
# AI DOCUMENT GENERATION
# ============================================================

def generate_ai_document(
    document_type,
    instruction,
    profile,
):
    client = get_groq_client()

    if client is None:
        raise RuntimeError(
            "Groq API key is missing. "
            "Add GROQ_API_KEY to Streamlit Secrets."
        )

    signature = profile_signature(
        profile
    )

    prompt = f"""
You are DraftForge, an AI assistant for drafting professional official correspondence.

Document type:
{document_type}

User's raw instructions:
{instruction}

Sender profile:
{signature}

Prepare a professional official English document.

Rules:
1. Correct spelling, grammar, punctuation and obvious speech-to-text errors.
2. Preserve the user's intended meaning.
3. Do not invent names, dates, reference numbers, allegations, evidence,
   witnesses, events or other facts.
4. Do not add unsupported information.
5. Use clear, formal and professional official English.
6. Do not add a sender signature because the application adds it separately.
7. Return only the actual document.
"""

    response = client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[
            {
                "role": "system",
                "content": (
                    "You are a professional official-document "
                    "drafting assistant. Never fabricate facts."
                ),
            },
            {
                "role": "user",
                "content": prompt,
            },
        ],
        temperature=0.2,
    )

    return clean_ai_text(
        response.choices[0].message.content
    )


# ============================================================
# INQUIRY HELPERS
# ============================================================

def documents_text(documents):
    if not documents:
        return (
            "No information was provided "
            "for this index."
        )

    lines = []

    annex_letter = ord("A")

    for document in documents:
        if not document:
            continue

        label = (
            f"Annex-{chr(annex_letter)}"
        )

        lines.append(
            f"{label}: {document}"
        )

        annex_letter += 1

    if not lines:
        return (
            "No information was provided "
            "for this index."
        )

    return "\n".join(lines)


def committee_text(committee):
    result = []

    if isinstance(committee, dict):

        for role in COMMITTEE_ROLES:

            member = committee.get(
                role,
                {},
            )

            if not isinstance(
                member,
                dict,
            ):
                continue

            erp = str(
                member.get(
                    "ERP#",
                    "",
                )
                or ""
            ).strip()

            name = str(
                member.get(
                    "Name",
                    "",
                )
                or ""
            ).strip()

            designation = str(
                member.get(
                    "Designation",
                    "",
                )
                or ""
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

    elif isinstance(
        committee,
        list,
    ):

        for member in committee:

            if not isinstance(
                member,
                dict,
            ):
                continue

            role = str(
                member.get(
                    "role",
                    "",
                )
                or ""
            ).strip()

            erp = str(
                member.get(
                    "erp",
                    "",
                )
                or ""
            ).strip()

            name = str(
                member.get(
                    "name",
                    "",
                )
                or ""
            ).strip()

            designation = str(
                member.get(
                    "designation",
                    "",
                )
                or ""
            ).strip()

            if not (
                role
                or erp
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
                f"{role or 'Committee Member'} — "
                + ", ".join(parts)
            )

    if not result:
        return (
            "No information was provided "
            "for this index."
        )

    return "\n".join(result)


def qa_markdown(rows):
    if not rows:
        return (
            "No information was provided "
            "for this index."
        )

    lines = [
        "| Questions | Answers |",
        "|---|---|",
    ]

    for number, row in enumerate(
        rows,
        start=1,
    ):

        if not isinstance(
            row,
            dict,
        ):
            continue

        question = str(
            row.get(
                "question",
                "",
            )
            or ""
        ).strip()

        answer = str(
            row.get(
                "answer",
                "",
            )
            or ""
        ).strip()

        if not question and not answer:
            continue

        question = (
            question
            .replace(
                "|",
                "\\|",
            )
            .replace(
                "\n",
                " ",
            )
        )

        answer = (
            answer
            .replace(
                "|",
                "\\|",
            )
            .replace(
                "\n",
                " ",
            )
        )

        lines.append(
            f"| {number}. {question} | {answer} |"
        )

    if len(lines) == 2:
        return (
            "No information was provided "
            "for this index."
        )

    return "\n".join(lines)


def normalize_index_data():

    data = st.session_state.get(
        "index_data",
        [],
    )

    if not isinstance(
        data,
        list,
    ):
        data = []

    for item in data:

        if not isinstance(
            item,
            dict,
        ):
            continue

        if (
            item.get("type")
            == "Inquiry Committee"
        ):

            committee = item.get(
                "committee"
            )

            if isinstance(
                committee,
                list,
            ):

                converted = {
                    role: {
                        "ERP#": "",
                        "Name": "",
                        "Designation": "",
                    }
                    for role in COMMITTEE_ROLES
                }

                for member in committee:

                    if not isinstance(
                        member,
                        dict,
                    ):
                        continue

                    role = member.get(
                        "role",
                        "",
                    )

                    if role in converted:

                        converted[role] = {
                            "ERP#": member.get(
                                "erp",
                                "",
                            ),
                            "Name": member.get(
                                "name",
                                "",
                            ),
                            "Designation": member.get(
                                "designation",
                                "",
                            ),
                        }

                item["committee"] = (
                    converted
                )

            elif not isinstance(
                committee,
                dict,
            ):

                item["committee"] = {
                    role: {
                        "ERP#": "",
                        "Name": "",
                        "Designation": "",
                    }
                    for role in COMMITTEE_ROLES
                }

            else:

                for role in COMMITTEE_ROLES:

                    if role not in committee:
                        committee[role] = {
                            "ERP#": "",
                            "Name": "",
                            "Designation": "",
                        }

    st.session_state.index_data = data


# ============================================================
# INQUIRY REPORT
# ============================================================

def generate_inquiry_report(
    index_data,
    inquiry_type,
):
    client = get_groq_client()

    if client is None:
        raise RuntimeError(
            "Groq API key is missing. "
            "Add GROQ_API_KEY to Streamlit Secrets."
        )

    selected_sections = []

    for item in index_data:

        if not isinstance(
            item,
            dict,
        ):
            continue

        index_name = str(
            item.get(
                "type",
                "",
            )
        ).strip()

        if not index_name:
            continue

        content = item.get(
            "content",
            "",
        )

        if (
            index_name
            == "Documents Recorded"
        ):

            content = documents_text(
                item.get(
                    "documents",
                    [],
                )
            )

        elif (
            index_name
            == "Inquiry Committee"
        ):

            content = committee_text(
                item.get(
                    "committee",
                    {},
                )
            )

        elif (
            index_name
            == "Questions / Answers with the Accused"
        ):

            content = qa_markdown(
                item.get(
                    "qa_rows",
                    [],
                )
            )

        content = str(
            content or ""
        ).strip()

        if not content:
            content = (
                "No information was provided "
                "for this index."
            )

        selected_sections.append(
            {
                "name": index_name,
                "content": content,
            }
        )

    if not selected_sections:
        raise ValueError(
            "Please add at least one inquiry "
            "index before generating."
        )

    payload = "\n\n".join(
        [
            f"INDEX {i}: {section['name']}\n"
            f"CONTENT:\n{section['content']}"
            for i, section in enumerate(
                selected_sections,
                start=1,
            )
        ]
    )

    prompt = f"""
Prepare an official {inquiry_type} report from the information supplied below.

CRITICAL RULES:

1. Output ONLY the indexes supplied below.
2. Keep exactly the same order as supplied.
3. Do NOT create any missing index.
4. Do NOT add Introduction.
5. Do NOT add Summary of Evidence.
6. Do NOT add Findings unless Findings was supplied.
7. Do NOT add Conclusion unless Conclusion was supplied.
8. Do NOT add Recommendations unless Recommendations was supplied.
9. Do NOT add Inquiry Committee unless it was supplied.
10. Do NOT add Documents Recorded unless it was supplied.
11. Never invent facts.
12. Preserve names, dates, allegations, evidence and other facts exactly
    as provided, while improving grammar and official wording.
13. If an index has no information, write:
    "No information was provided for this index."
14. Statement of the Accused and Questions / Answers may occur multiple times.
15. Keep repeated indexes separately.
16. Do not merge repeated indexes.
17. Do not create additional headings.

The final report must contain ONLY the selected indexes.

SUPPLIED INDEXES:

{payload}
"""

    response = client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[
            {
                "role": "system",
                "content": (
                    "You are an official inquiry-report "
                    "drafting assistant. Follow the supplied "
                    "section structure exactly."
                ),
            },
            {
                "role": "user",
                "content": prompt,
            },
        ],
        temperature=0.1,
    )

    return clean_ai_text(
        response.choices[0].message.content
    )


# ============================================================
# AI EDITING
# ============================================================

def ai_edit_document(
    document,
    instruction,
):
    client = get_groq_client()

    if client is None:
        raise RuntimeError(
            "Groq API key is missing."
        )

    prompt = f"""
Edit the following official document according to the user's instruction.

USER INSTRUCTION:
{instruction}

DOCUMENT:
{document}

Rules:
- Preserve the original meaning.
- Do not invent facts.
- Do not remove important factual information unless explicitly requested.
- Correct grammar, spelling and punctuation.
- Maintain professional official English.
- Return only the revised document.
"""

    response = client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[
            {
                "role": "system",
                "content": (
                    "You are a careful professional document "
                    "editor. Never fabricate facts."
                ),
            },
            {
                "role": "user",
                "content": prompt,
            },
        ],
        temperature=0.1,
    )

    return clean_ai_text(
        response.choices[0].message.content
    )


# ============================================================
# HISTORY
# ============================================================

def save_history(
    title,
    document,
):
    history = load_json(
        HISTORY_FILE,
        [],
    )

    if not isinstance(
        history,
        list,
    ):
        history = []

    history.insert(
        0,
        {
            "title": title,
            "document": document,
            "date": datetime.now().strftime(
                "%Y-%m-%d %H:%M"
            ),
        },
    )

    history = history[:30]

    save_json(
        HISTORY_FILE,
        history,
    )

    st.session_state.history = history


# ============================================================
# TXT EXPORT
# ============================================================

def create_txt(text):
    return str(text).encode(
        "utf-8"
    )


# ============================================================
# DOCX EXPORT
# ============================================================

def create_docx(text):

    if Document is None:
        raise RuntimeError(
            "python-docx is not installed."
        )

    doc = Document()

    section = doc.sections[0]

    section.top_margin = Inches(
        0.7
    )

    section.bottom_margin = Inches(
        0.7
    )

    section.left_margin = Inches(
        0.8
    )

    section.right_margin = Inches(
        0.8
    )

    for line in str(text).splitlines():

        if not line.strip():
            doc.add_paragraph()
            continue

        paragraph = doc.add_paragraph()

        paragraph.paragraph_format.space_after = Pt(
            6
        )

        run = paragraph.add_run(
            line
        )

        run.font.name = "Arial"
        run.font.size = Pt(11)

    output = io.BytesIO()

    doc.save(output)

    output.seek(0)

    return output.getvalue()


# ============================================================
# PDF — FIND UNICODE FONT
# ============================================================

def find_unicode_font():

    possible_paths = [

        # Linux / Streamlit Cloud
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSansCondensed.ttf",

        # Ubuntu
        "/usr/share/fonts/dejavu/DejaVuSans.ttf",

        # Common local locations
        os.path.join(
            os.getcwd(),
            "DejaVuSans.ttf",
        ),

        os.path.join(
            os.path.dirname(
                os.path.abspath(__file__)
            ),
            "DejaVuSans.ttf",
        ),

        # Windows
        r"C:\Windows\Fonts\DejaVuSans.ttf",
        r"C:\Windows\Fonts\arial.ttf",

        # macOS
        "/Library/Fonts/Arial.ttf",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
    ]

    for path in possible_paths:

        if os.path.isfile(path):
            return path

    # Last attempt: search common font folders.
    search_roots = [
        "/usr/share/fonts",
        "/usr/local/share/fonts",
    ]

    for root in search_roots:

        if not os.path.isdir(root):
            continue

        try:

            for dirpath, _, filenames in os.walk(
                root
            ):

                for filename in filenames:

                    lower = filename.lower()

                    if (
                        lower == "dejavusans.ttf"
                        or lower == "dejavusanscondensed.ttf"
                    ):

                        return os.path.join(
                            dirpath,
                            filename,
                        )

        except Exception:
            pass

    return None


# ============================================================
# PDF — SAFE TEXT
# ============================================================

def safe_pdf_text(
    text,
    width=88,
):
    """
    Wrap PDF text safely.

    Important:
    We deliberately DO NOT strip Unicode here.
    The Unicode font handles characters such as:
    — - ’ “ ” • etc.
    """

    result = []

    for raw_line in str(
        text
    ).splitlines():

        line = raw_line.rstrip()

        if not line:

            result.append("")

            continue

        wrapped = textwrap.wrap(
            line,
            width=width,
            break_long_words=True,
            break_on_hyphens=True,
            replace_whitespace=False,
            drop_whitespace=False,
        )

        if not wrapped:
            result.append("")
        else:
            result.extend(
                wrapped
            )

    return result


# ============================================================
# PDF — FALLBACK CHARACTER CLEANING
# ============================================================

def pdf_ascii_fallback(text):

    replacements = {
        "\u2010": "-",
        "\u2011": "-",
        "\u2012": "-",
        "\u2013": "-",
        "\u2014": "-",
        "\u2015": "-",

        "\u2018": "'",
        "\u2019": "'",
        "\u201a": "'",
        "\u201b": "'",

        "\u201c": '"',
        "\u201d": '"',
        "\u201e": '"',
        "\u201f": '"',

        "\u2022": "-",
        "\u2023": "-",

        "\u00a0": " ",

        "\u2026": "...",

        "\u00d7": "x",

        "\u2212": "-",

        "\u00a9": "(c)",
        "\u00ae": "(R)",
        "\u2122": "(TM)",
    }

    text = str(text)

    for old, new in replacements.items():
        text = text.replace(
            old,
            new,
        )

    return text


# ============================================================
# PDF EXPORT
# ============================================================

def create_pdf(text):

    if FPDF is None:
        raise RuntimeError(
            "fpdf is not installed."
        )

    font_path = find_unicode_font()

    # --------------------------------------------------------
    # Preferred route: Unicode PDF
    # --------------------------------------------------------

    if font_path:

        try:

            pdf = FPDF()

            pdf.set_auto_page_break(
                auto=True,
                margin=15,
            )

            pdf.add_page()

            # fpdf2 / compatible versions
            try:

                pdf.add_font(
                    "DraftForgeUnicode",
                    "",
                    font_path,
                )

            except TypeError:

                # Older PyFPDF compatibility
                pdf.add_font(
                    "DraftForgeUnicode",
                    "",
                    font_path,
                    uni=True,
                )

            pdf.set_font(
                "DraftForgeUnicode",
                size=11,
            )

            usable_width = (
                pdf.w
                - pdf.l_margin
                - pdf.r_margin
            )

            lines = safe_pdf_text(
                text,
                width=88,
            )

            for line in lines:

                pdf.set_x(
                    pdf.l_margin
                )

                if line == "":

                    pdf.ln(5)

                else:

                    pdf.multi_cell(
                        usable_width,
                        6,
                        line,
                        border=0,
                        align="L",
                    )

            output = pdf.output()

            if isinstance(
                output,
                bytearray,
            ):
                output = bytes(
                    output
                )

            if isinstance(
                output,
                str,
            ):
                output = output.encode(
                    "latin-1"
                )

            return output

        except Exception:
            # If Unicode font route fails,
            # continue to safe fallback below.
            pass

    # --------------------------------------------------------
    # Fallback route
    # --------------------------------------------------------

    fallback_text = pdf_ascii_fallback(
        text
    )

    pdf = FPDF()

    pdf.set_auto_page_break(
        auto=True,
        margin=15,
    )

    pdf.add_page()

    pdf.set_font(
        "Helvetica",
        size=11,
    )

    usable_width = (
        pdf.w
        - pdf.l_margin
        - pdf.r_margin
    )

    lines = safe_pdf_text(
        fallback_text,
        width=88,
    )

    for line in lines:

        pdf.set_x(
            pdf.l_margin
        )

        if line == "":

            pdf.ln(5)

        else:

            # Encode safely for Helvetica.
            safe_line = (
                line.encode(
                    "latin-1",
                    "replace",
                )
                .decode(
                    "latin-1"
                )
            )

            pdf.multi_cell(
                usable_width,
                6,
                safe_line,
                border=0,
                align="L",
            )

    output = pdf.output()

    if isinstance(
        output,
        bytearray,
    ):
        output = bytes(
            output
        )

    if isinstance(
        output,
        str,
    ):
        output = output.encode(
            "latin-1"
        )

    return output


# ============================================================
# PNG EXPORT
# ============================================================

def create_png(text):

    if Image is None:
        raise RuntimeError(
            "Pillow is not installed."
        )

    try:

        font = ImageFont.truetype(
            "DejaVuSans.ttf",
            24,
        )

    except Exception:

        font = ImageFont.load_default()

    lines = safe_pdf_text(
        text,
        width=65,
    )

    line_height = 38

    margin = 60

    height = max(
        400,
        margin * 2
        + len(lines)
        * line_height,
    )

    width = 1400

    image = Image.new(
        "RGB",
        (
            width,
            height,
        ),
        "white",
    )

    draw = ImageDraw.Draw(
        image
    )

    y = margin

    for line in lines:

        draw.text(
            (
                margin,
                y,
            ),
            line,
            fill="black",
            font=font,
        )

        y += line_height

    output = io.BytesIO()

    image.save(
        output,
        format="PNG",
    )

    output.seek(0)

    return output.getvalue()


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.markdown(
        "## ✦ DraftForge"
    )

    st.caption(
        "AI Document Composer"
    )

    st.divider()

    st.markdown(
        "### Workspace"
    )

    if st.button(
        "✦ New Document",
        use_container_width=True,
    ):

        st.session_state.document_type = (
            "Email"
        )

        st.session_state.inquiry_type = (
            "E&D Inquiry"
        )

        st.session_state.index_data = []

        st.session_state.generated_draft = ""

        st.session_state.editable_draft = ""

        st.session_state.document_editor = ""

        st.session_state.editor_sync = ""

        st.session_state.email_instruction = ""

        st.session_state.letter_instruction = ""

        st.rerun()

    if st.button(
        "🗂 My Documents",
        use_container_width=True,
    ):

        st.session_state.show_history = True
        st.session_state.show_profile = False
        st.session_state.show_about = False

    if st.button(
        "👤 My Profile",
        use_container_width=True,
    ):

        st.session_state.show_profile = True
        st.session_state.show_history = False
        st.session_state.show_about = False

    if st.button(
        "ℹ About DraftForge",
        use_container_width=True,
    ):

        st.session_state.show_about = True
        st.session_state.show_history = False
        st.session_state.show_profile = False

    st.divider()

    st.markdown(
        "### 💡 Tips & Templates"
    )

    st.info(
        "Write naturally; DraftForge converts your "
        "instructions into professional official English."
    )

    st.info(
        "🎤 Speak instead of typing."
    )

    st.info(
        "📄 Inquiry sections can be added in any order."
    )

    st.info(
        "🔁 The same inquiry index can be added multiple times."
    )

    st.divider()

    st.caption(
        "Developed by: Raees Khan\n\n"
        "Assistant Director, NADRA"
    )


# ============================================================
# HERO
# ============================================================

st.markdown(
    """
    <div class="hero-box">
        <div class="hero-title">✦ DraftForge</div>
        <div class="hero-subtitle">
            AI-powered workspace for professional official correspondence
            and inquiry documentation.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# PROFILE
# ============================================================

if st.session_state.show_profile:

    st.markdown(
        "## 👤 My Profile"
    )

    profile = load_json(
        PROFILE_FILE,
        {
            "Name": "",
            "Designation": "",
            "Contact No.": "",
            "Current Station": "",
        },
    )

    with st.form(
        "profile_form"
    ):

        name = st.text_input(
            "Name",
            value=profile.get(
                "Name",
                "",
            ),
        )

        designation = st.text_input(
            "Designation",
            value=profile.get(
                "Designation",
                "",
            ),
        )

        contact = st.text_input(
            "Contact No.",
            value=profile.get(
                "Contact No.",
                "",
            ),
        )

        station = st.text_input(
            "Current Station",
            value=profile.get(
                "Current Station",
                "",
            ),
        )

        submitted = st.form_submit_button(
            "Save Profile",
            use_container_width=True,
        )

        if submitted:

            new_profile = {
                "Name": name.strip(),
                "Designation": designation.strip(),
                "Contact No.": contact.strip(),
                "Current Station": station.strip(),
            }

            if save_json(
                PROFILE_FILE,
                new_profile,
            ):

                st.session_state.profile = (
                    new_profile
                )

                st.success(
                    "Profile saved successfully."
                )

            else:

                st.error(
                    "Unable to save profile."
                )

    st.divider()


# ============================================================
# HISTORY
# ============================================================

if st.session_state.show_history:

    st.markdown(
        "## 🗂 My Documents"
    )

    history = load_json(
        HISTORY_FILE,
        [],
    )

    if not history:

        st.info(
            "No saved documents yet. "
            "Generated documents will appear here."
        )

    else:

        for i, item in enumerate(
            history
        ):

            if not isinstance(
                item,
                dict,
            ):
                continue

            title = item.get(
                "title",
                "Untitled Document",
            )

            date = item.get(
                "date",
                "",
            )

            with st.expander(
                f"{title} — {date}"
            ):

                saved_document = item.get(
                    "document",
                    "",
                )

                st.text_area(
                    "Document",
                    saved_document,
                    height=250,
                    key=f"history_{i}",
                )

                if st.button(
                    "Restore",
                    key=f"restore_{i}",
                ):

                    st.session_state.generated_draft = (
                        saved_document
                    )

                    st.session_state.editable_draft = (
                        saved_document
                    )

                    st.session_state.editor_sync = (
                        saved_document
                    )

                    st.session_state.show_history = False

                    st.rerun()


# ============================================================
# ABOUT
# ============================================================

if st.session_state.show_about:

    st.markdown(
        "## ℹ About DraftForge"
    )

    st.write(
        "DraftForge is an AI-assisted drafting workspace "
        "for professional official correspondence and "
        "inquiry documentation."
    )

    st.markdown(
        """
        **Supported document types**

        - Email
        - Letter
        - E&D Inquiry
        - FFI Inquiry — Under Construction / Under Process
        """
    )

    st.markdown(
        """
        **Developer**

        Developed by: Raees Khan  
        Assistant Director, NADRA
        """
    )


# ============================================================
# DOCUMENT TYPE
# ============================================================

st.markdown(
    "## ① Choose Document"
)

st.caption(
    "Start by selecting what you want DraftForge to prepare."
)

doc_cols = st.columns(3)

for i, document_type in enumerate(
    DOCUMENT_TYPES
):

    with doc_cols[i]:

        selected = (
            st.session_state.document_type
            == document_type
        )

        label = (
            f"✓ {document_type}"
            if selected
            else document_type
        )

        if st.button(
            label,
            key=f"document_type_{document_type}",
            use_container_width=True,
        ):

            st.session_state.document_type = (
                document_type
            )

            if document_type != "Inquiry":

                st.session_state.index_data = []

            st.session_state.generated_draft = ""
            st.session_state.editable_draft = ""
            st.session_state.document_editor = ""
            st.session_state.editor_sync = ""

            st.rerun()


document_type = (
    st.session_state.document_type
)


# ============================================================
# EMAIL
# ============================================================

if document_type == "Email":

    st.markdown(
        "## ② Provide Information"
    )

    st.caption(
        "Describe what you want to communicate. "
        "You can type, speak, or use both."
    )

    profile = load_json(
        PROFILE_FILE,
        {},
    )

    # IMPORTANT:
    # Audio is processed BEFORE the text widget.
    audio = st.audio_input(
        "🎤 Speak your email instructions",
        key="email_audio",
    )

    transcript = process_voice_input(
        audio,
        "email_voice_hash",
    )

    if transcript:

        current = st.session_state.get(
            "email_instruction",
            "",
        ).strip()

        if current:

            st.session_state.email_instruction = (
                current
                + "\n"
                + transcript
            )

        else:

            st.session_state.email_instruction = (
                transcript
            )

        st.success(
            "Voice input added to the instruction box."
        )

    email_instruction = st.text_area(
        "Email Instructions",
        key="email_instruction",
        height=220,
        placeholder=(
            "Example: Inform the regional office that "
            "the network issue has been resolved and "
            "request confirmation."
        ),
    )

    st.caption(
        "The microphone works even when the "
        "instruction box is empty."
    )

    # Generate remains LAST.
    if st.button(
        "✦ Generate Email",
        type="primary",
        use_container_width=True,
    ):

        if not email_instruction.strip():

            st.warning(
                "Please provide information by typing or speaking."
            )

        else:

            with st.spinner(
                "DraftForge is preparing your email..."
            ):

                try:

                    final_document = (
                        generate_ai_document(
                            "Email",
                            email_instruction,
                            profile,
                        )
                    )

                    signature = (
                        profile_signature(
                            profile
                        )
                    )

                    if signature:

                        final_document += (
                            "\n\n"
                            + signature
                        )

                    st.session_state.generated_draft = (
                        final_document
                    )

                    st.session_state.editable_draft = (
                        final_document
                    )

                    st.session_state.editor_sync = (
                        final_document
                    )

                    save_history(
                        "Email",
                        final_document,
                    )

                    st.rerun()

                except Exception as e:

                    st.error(
                        str(e)
                    )


# ============================================================
# LETTER
# ============================================================

elif document_type == "Letter":

    st.markdown(
        "## ② Provide Information"
    )

    st.caption(
        "Describe the purpose and required contents "
        "of the letter. Use text, voice, or both."
    )

    profile = load_json(
        PROFILE_FILE,
        {},
    )

    # Audio BEFORE text widget.
    audio = st.audio_input(
        "🎤 Speak your letter instructions",
        key="letter_audio",
    )

    transcript = process_voice_input(
        audio,
        "letter_voice_hash",
    )

    if transcript:

        current = st.session_state.get(
            "letter_instruction",
            "",
        ).strip()

        if current:

            st.session_state.letter_instruction = (
                current
                + "\n"
                + transcript
            )

        else:

            st.session_state.letter_instruction = (
                transcript
            )

        st.success(
            "Voice input added to the instruction box."
        )

    letter_instruction = st.text_area(
        "Letter Instructions",
        key="letter_instruction",
        height=220,
        placeholder=(
            "Example: Draft a letter to the concerned "
            "office regarding timely resolution of pending cases."
        ),
    )

    st.caption(
        "The microphone works even when the "
        "instruction box is empty."
    )

    # Generate remains LAST.
    if st.button(
        "✦ Generate Letter",
        type="primary",
        use_container_width=True,
    ):

        if not letter_instruction.strip():

            st.warning(
                "Please provide information by typing or speaking."
            )

        else:

            with st.spinner(
                "DraftForge is preparing your letter..."
            ):

                try:

                    final_document = (
                        generate_ai_document(
                            "Letter",
                            letter_instruction,
                            profile,
                        )
                    )

                    signature = (
                        profile_signature(
                            profile
                        )
                    )

                    if signature:

                        final_document += (
                            "\n\n"
                            + signature
                        )

                    st.session_state.generated_draft = (
                        final_document
                    )

                    st.session_state.editable_draft = (
                        final_document
                    )

                    st.session_state.editor_sync = (
                        final_document
                    )

                    save_history(
                        "Letter",
                        final_document,
                    )

                    st.rerun()

                except Exception as e:

                    st.error(
                        str(e)
                    )


# ============================================================
# INQUIRY
# ============================================================

elif document_type == "Inquiry":

    st.markdown(
        "## ② Provide Information"
    )

    inquiry_type = st.radio(
        "Inquiry Type",
        INQUIRY_TYPES,
        index=(
            INQUIRY_TYPES.index(
                st.session_state.inquiry_type
            )
            if st.session_state.inquiry_type
            in INQUIRY_TYPES
            else 0
        ),
        horizontal=True,
    )

    st.session_state.inquiry_type = (
        inquiry_type
    )

    if inquiry_type == "FFI Inquiry":

        st.warning(
            "FFI Inquiry is currently "
            "Under Construction / Under Process."
        )

    st.markdown(
        "### Add Inquiry Sections"
    )

    st.caption(
        "Select an index and add it to the report. "
        "Indexes can be added in any order."
    )

    normalize_index_data()

    selected_index = st.selectbox(
        "Select Inquiry Index",
        ED_INDEXES,
        key="new_inquiry_index",
    )

    if st.button(
        "＋ Add Selected Index",
        use_container_width=True,
    ):

        new_item = {
            "type": selected_index,
            "content": "",
        }

        if selected_index == "Documents Recorded":

            new_item["documents"] = []

        elif selected_index == "Inquiry Committee":

            new_item["committee"] = {
                role: {
                    "ERP#": "",
                    "Name": "",
                    "Designation": "",
                }
                for role in COMMITTEE_ROLES
            }

        elif (
            selected_index
            == "Questions / Answers with the Accused"
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

    st.divider()

    if st.session_state.index_data:

        st.markdown(
            "### Selected Inquiry Sections"
        )

        st.caption(
            "Only the sections shown below will appear "
            "in the generated report."
        )

        for position, item in enumerate(
            st.session_state.index_data
        ):

            if not isinstance(
                item,
                dict,
            ):
                continue

            index_name = item.get(
                "type",
                "",
            )

            with st.container(
                border=True
            ):

                header_cols = st.columns(
                    [
                        0.8,
                        5.5,
                        1.2,
                    ]
                )

                with header_cols[0]:

                    st.markdown(
                        f"### {position + 1}"
                    )

                with header_cols[1]:

                    st.markdown(
                        f"### {index_name}"
                    )

                with header_cols[2]:

                    if st.button(
                        "Remove",
                        key=f"remove_index_{position}",
                    ):

                        st.session_state.index_data.pop(
                            position
                        )

                        st.rerun()

                # ------------------------------------------------
                # DOCUMENTS RECORDED
                # ------------------------------------------------

                if (
                    index_name
                    == "Documents Recorded"
                ):

                    st.caption(
                        "Select the documents examined or recorded. "
                        "Annexures will be assigned automatically."
                    )

                    current_documents = item.get(
                        "documents",
                        [],
                    )

                    selected_documents = st.multiselect(
                        "Documents",
                        DOCUMENTS_RECORDED,
                        default=current_documents,
                        key=f"documents_{position}",
                    )

                    item["documents"] = (
                        selected_documents
                    )

                    if selected_documents:

                        st.info(
                            "\n".join(
                                [
                                    f"Annex-{chr(65+i)} — {doc}"
                                    for i, doc in enumerate(
                                        selected_documents
                                    )
                                ]
                            )
                        )

                # ------------------------------------------------
                # INQUIRY COMMITTEE
                # ------------------------------------------------

                elif (
                    index_name
                    == "Inquiry Committee"
                ):

                    st.caption(
                        "Enter the details of each committee member."
                    )

                    committee = item.setdefault(
                        "committee",
                        {
                            role: {
                                "ERP#": "",
                                "Name": "",
                                "Designation": "",
                            }
                            for role in COMMITTEE_ROLES
                        },
                    )

                    for role_index, role in enumerate(
                        COMMITTEE_ROLES
                    ):

                        st.markdown(
                            f"**{role}**"
                        )

                        current = committee.setdefault(
                            role,
                            {
                                "ERP#": "",
                                "Name": "",
                                "Designation": "",
                            },
                        )

                        erp_key = (
                            f"committee_erp_"
                            f"{position}_{role_index}"
                        )

                        name_key = (
                            f"committee_name_"
                            f"{position}_{role_index}"
                        )

                        designation_key = (
                            f"committee_designation_"
                            f"{position}_{role_index}"
                        )

                        if (
                            erp_key
                            not in st.session_state
                        ):

                            st.session_state[
                                erp_key
                            ] = current.get(
                                "ERP#",
                                "",
                            )

                        if (
                            name_key
                            not in st.session_state
                        ):

                            st.session_state[
                                name_key
                            ] = current.get(
                                "Name",
                                "",
                            )

                        if (
                            designation_key
                            not in st.session_state
                        ):

                            st.session_state[
                                designation_key
                            ] = current.get(
                                "Designation",
                                "",
                            )

                        c1, c2, c3 = st.columns(
                            3
                        )

                        with c1:

                            erp = st.text_input(
                                "ERP#",
                                key=erp_key,
                            )

                        with c2:

                            name = st.text_input(
                                "Name",
                                key=name_key,
                            )

                        with c3:

                            designation = st.text_input(
                                "Designation",
                                key=designation_key,
                            )

                        committee[role] = {
                            "ERP#": erp,
                            "Name": name,
                            "Designation": designation,
                        }

                        if (
                            role_index
                            < len(
                                COMMITTEE_ROLES
                            )
                            - 1
                        ):

                            st.divider()

                    item["committee"] = (
                        committee
                    )

                # ------------------------------------------------
                # QUESTIONS / ANSWERS
                # ------------------------------------------------

                elif (
                    index_name
                    == "Questions / Answers with the Accused"
                ):

                    st.caption(
                        "Enter each question and its corresponding answer."
                    )

                    rows = item.setdefault(
                        "qa_rows",
                        [
                            {
                                "question": "",
                                "answer": "",
                            }
                        ],
                    )

                    for row_index, row in enumerate(
                        rows
                    ):

                        q_col, a_col = st.columns(
                            [1, 1]
                        )

                        with q_col:

                            q_key = (
                                f"qa_question_"
                                f"{position}_{row_index}"
                            )

                            if (
                                q_key
                                not in st.session_state
                            ):

                                st.session_state[
                                    q_key
                                ] = row.get(
                                    "question",
                                    "",
                                )

                            question = st.text_area(
                                f"Question {row_index + 1}",
                                key=q_key,
                                height=100,
                            )

                        with a_col:

                            a_key = (
                                f"qa_answer_"
                                f"{position}_{row_index}"
                            )

                            if (
                                a_key
                                not in st.session_state
                            ):

                                st.session_state[
                                    a_key
                                ] = row.get(
                                    "answer",
                                    "",
                                )

                            answer = st.text_area(
                                f"Answer {row_index + 1}",
                                key=a_key,
                                height=100,
                            )

                        row["question"] = (
                            question
                        )

                        row["answer"] = (
                            answer
                        )

                    add_q_col, remove_q_col = (
                        st.columns(2)
                    )

                    with add_q_col:

                        if st.button(
                            "＋ Add Question",
                            key=f"add_question_{position}",
                            use_container_width=True,
                        ):

                            rows.append(
                                {
                                    "question": "",
                                    "answer": "",
                                }
                            )

                            st.rerun()

                    with remove_q_col:

                        if (
                            len(rows) > 1
                            and st.button(
                                "− Remove Last Question",
                                key=f"remove_question_{position}",
                                use_container_width=True,
                            )
                        ):

                            rows.pop()

                            st.rerun()

                    item["qa_rows"] = rows

                # ------------------------------------------------
                # NORMAL INQUIRY INDEX
                # ------------------------------------------------

                else:

                    audio_key = (
                        f"inquiry_audio_{position}"
                    )

                    hash_key = (
                        f"inquiry_voice_hash_{position}"
                    )

                    text_key = (
                        f"inquiry_text_{position}"
                    )

                    # Audio MUST be processed before
                    # the text widget is instantiated.
                    audio = st.audio_input(
                        "🎤 Speak information for this index",
                        key=audio_key,
                    )

                    transcript = process_voice_input(
                        audio,
                        hash_key,
                    )

                    if transcript:

                        current = st.session_state.get(
                            text_key,
                            item.get(
                                "content",
                                "",
                            ),
                        )

                        current = str(
                            current or ""
                        ).strip()

                        if current:

                            st.session_state[
                                text_key
                            ] = (
                                current
                                + "\n"
                                + transcript
                            )

                        else:

                            st.session_state[
                                text_key
                            ] = transcript

                        st.success(
                            "Voice input added."
                        )

                    if (
                        text_key
                        not in st.session_state
                    ):

                        st.session_state[
                            text_key
                        ] = item.get(
                            "content",
                            "",
                        )

                    content = st.text_area(
                        "Information",
                        key=text_key,
                        height=170,
                        placeholder=(
                            "Provide information naturally. "
                            "DraftForge will convert it into "
                            "professional official wording."
                        ),
                    )

                    item["content"] = (
                        content
                    )

        st.divider()

        # --------------------------------------------------------
        # GENERATE
        # --------------------------------------------------------

        st.markdown(
            "## ③ Generate & Export"
        )

        st.caption(
            "The final report will contain only the inquiry "
            "sections you added above, in exactly the same order."
        )

        if st.button(
            "✦ Generate Inquiry Report",
            type="primary",
            use_container_width=True,
        ):

            with st.spinner(
                "DraftForge is preparing the inquiry report..."
            ):

                try:

                    report = (
                        generate_inquiry_report(
                            st.session_state.index_data,
                            inquiry_type,
                        )
                    )

                    st.session_state.generated_draft = (
                        report
                    )

                    st.session_state.editable_draft = (
                        report
                    )

                    st.session_state.editor_sync = (
                        report
                    )

                    save_history(
                        inquiry_type,
                        report,
                    )

                    st.rerun()

                except Exception as e:

                    st.error(
                        str(e)
                    )

    else:

        st.info(
            "No inquiry sections have been added yet. "
            "Select an index above and click "
            "“Add Selected Index”."
        )


# ============================================================
# GENERATED DOCUMENT WORKSPACE
# ============================================================

if st.session_state.get(
    "generated_draft",
    "",
).strip():

    st.divider()

    st.markdown(
        "## ✎ Document Workspace"
    )

    st.caption(
        "Review your generated document, make manual changes, "
        "or ask DraftForge to edit it."
    )

    # --------------------------------------------------------
    # Editor synchronization
    # --------------------------------------------------------

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
        key="document_editor",
        height=520,
    )

    # --------------------------------------------------------
    # Save / Restore / TXT
    # --------------------------------------------------------

    action_cols = st.columns(3)

    with action_cols[0]:

        if st.button(
            "💾 Save Changes",
            use_container_width=True,
        ):

            st.session_state.editable_draft = (
                edited_document
            )

            st.session_state.generated_draft = (
                edited_document
            )

            save_history(
                "Edited Document",
                edited_document,
            )

            st.success(
                "Changes saved."
            )

    with action_cols[1]:

        if st.button(
            "↩ Restore Original",
            use_container_width=True,
        ):

            original = (
                st.session_state.generated_draft
            )

            st.session_state.editable_draft = (
                original
            )

            st.session_state.editor_sync = (
                original
            )

            st.rerun()

    with action_cols[2]:

        st.download_button(
            "⬇ Download TXT",
            data=create_txt(
                edited_document
            ),
            file_name="DraftForge_Document.txt",
            mime="text/plain",
            use_container_width=True,
        )

    # --------------------------------------------------------
    # AI EDIT
    # --------------------------------------------------------

    st.markdown(
        "### ✨ AI Edit"
    )

    if st.session_state.get(
        "edit_instruction_sync",
        "",
    ):

        st.session_state.edit_instruction = (
            st.session_state.edit_instruction_sync
        )

        st.session_state.edit_instruction_sync = ""

    edit_instruction = st.text_area(
        "Tell DraftForge what you want changed",
        key="edit_instruction",
        height=120,
        placeholder=(
            "Example: Make this more concise and formal."
        ),
    )

    if st.button(
        "✨ Apply AI Edit",
        use_container_width=True,
    ):

        if not edit_instruction.strip():

            st.warning(
                "Please enter an editing instruction."
            )

        else:

            with st.spinner(
                "DraftForge is editing your document..."
            ):

                try:

                    modified = ai_edit_document(
                        edited_document,
                        edit_instruction,
                    )

                    st.session_state.editable_draft = (
                        modified
                    )

                    st.session_state.generated_draft = (
                        modified
                    )

                    st.session_state.editor_sync = (
                        modified
                    )

                    st.session_state.edit_instruction_sync = (
                        ""
                    )

                    st.rerun()

                except Exception as e:

                    st.error(
                        str(e)
                    )

    st.divider()

    # ========================================================
    # EXPORT
    # ========================================================

    st.markdown(
        "### 📤 Export"
    )

    export_cols = st.columns(3)

    # --------------------------------------------------------
    # DOCX
    # --------------------------------------------------------

    with export_cols[0]:

        try:

            docx_data = create_docx(
                edited_document
            )

            st.download_button(
                "📄 Download DOCX",
                data=docx_data,
                file_name="DraftForge_Document.docx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.document"
                ),
                use_container_width=True,
            )

        except Exception as e:

            st.error(
                f"DOCX export unavailable: {e}"
            )

    # --------------------------------------------------------
    # PDF
    # --------------------------------------------------------

    with export_cols[1]:

        try:

            pdf_data = create_pdf(
                edited_document
            )

            st.download_button(
                "📕 Download PDF",
                data=pdf_data,
                file_name="DraftForge_Document.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

        except Exception as e:

            st.error(
                f"PDF export unavailable: {e}"
            )

    # --------------------------------------------------------
    # PNG
    # --------------------------------------------------------

    with export_cols[2]:

        try:

            png_data = create_png(
                edited_document
            )

            st.download_button(
                "🖼 Download PNG",
                data=png_data,
                file_name="DraftForge_Document.png",
                mime="image/png",
                use_container_width=True,
            )

        except Exception as e:

            st.error(
                f"PNG export unavailable: {e}"
            )


# ============================================================
# FOOTER
# ============================================================

st.markdown(
    """
    <div class="footer-note">
        DraftForge — AI Document Composer
        <br>
        Developed by: Raees Khan — Assistant Director, NADRA
    </div>
    """,
    unsafe_allow_html=True,
)
