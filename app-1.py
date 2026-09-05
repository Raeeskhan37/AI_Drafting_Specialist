# ============================================================
# DraftForge — AI Document Composer
# Complete replacement app.py
# ============================================================

import io
import os
import re
import sqlite3
from datetime import datetime

import streamlit as st
import requests

from groq import Groq

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

from fpdf import FPDF

from PIL import Image, ImageDraw, ImageFont


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="DraftForge — AI Document Composer",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# CONSTANTS
# ============================================================

GROQ_MODEL = "openai/gpt-oss-120b"
WHISPER_MODEL = "whisper-large-v3"
GEMINI_MODEL = "gemini-2.0-flash"

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
    "Other",
]

COMMITTEE_ROLES = [
    "Convener of Inquiry",
    "Member 1",
    "Member 2",
    "Departmental Representative",
]


# ============================================================
# CSS
# ============================================================

st.markdown(
    """
<style>

.main-title {
    font-size: 38px;
    font-weight: 800;
    margin-bottom: 0px;
}

.subtitle {
    font-size: 17px;
    color: #666;
    margin-top: 0px;
    margin-bottom: 25px;
}

.section-title {
    font-size: 25px;
    font-weight: 750;
    margin-top: 15px;
    margin-bottom: 10px;
}

.index-card {
    border: 1px solid #d9d9d9;
    border-radius: 12px;
    padding: 15px;
    margin-bottom: 15px;
    background-color: rgba(128,128,128,0.04);
}

.help-box {
    padding: 12px 15px;
    border-radius: 10px;
    background-color: rgba(0, 120, 255, 0.07);
    border-left: 4px solid #1683ff;
    margin-bottom: 15px;
}

.warning-box {
    padding: 15px;
    border-radius: 10px;
    background-color: rgba(255, 165, 0, 0.10);
    border-left: 4px solid orange;
}

.generated-box {
    border: 1px solid #d5d5d5;
    border-radius: 12px;
    padding: 20px;
    background-color: rgba(128,128,128,0.03);
}

</style>
""",
    unsafe_allow_html=True,
)


# ============================================================
# SESSION STATE
# ============================================================

DEFAULT_STATE = {
    "document_type": "Email",
    "inquiry_type": "E&D Inquiry",
    "ed_instances": [],
    "generated_draft": "",
    "history": [],
}

for key, value in DEFAULT_STATE.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# SECRETS
# ============================================================

def get_secret(name, default=None):
    try:
        value = st.secrets.get(name)
        if value:
            return value
    except Exception:
        pass

    return os.getenv(name, default)


GROQ_API_KEY = get_secret("GROQ_API_KEY")
GEMINI_API_KEY = get_secret("GEMINI_API_KEY")


# ============================================================
# GROQ CLIENT
# ============================================================

@st.cache_resource
def get_groq_client():
    key = get_secret("GROQ_API_KEY")

    if not key:
        return None

    try:
        return Groq(api_key=key)
    except Exception:
        return None


# ============================================================
# VOICE TRANSCRIPTION
# ============================================================

def transcribe_audio(audio_file):
    """
    Send recorded audio to Groq Whisper.
    """

    client = get_groq_client()

    if client is None:
        raise RuntimeError(
            "GROQ_API_KEY is not configured in Streamlit Secrets."
        )

    try:
        audio_bytes = audio_file.getvalue()

        audio_buffer = io.BytesIO(audio_bytes)
        audio_buffer.name = "recording.wav"

        result = client.audio.transcriptions.create(
            file=audio_buffer,
            model=WHISPER_MODEL,
            response_format="text",
        )

        if hasattr(result, "text"):
            return result.text.strip()

        return str(result).strip()

    except Exception as exc:
        raise RuntimeError(
            f"Voice transcription failed: {exc}"
        )


# ============================================================
# TEXT HELPERS
# ============================================================

def append_text(existing, new_text):
    existing = (existing or "").strip()
    new_text = (new_text or "").strip()

    if not new_text:
        return existing

    if not existing:
        return new_text

    return existing + "\n" + new_text


def clean_markup(text):
    if not text:
        return ""

    text = re.sub(r"\*\*<u>(.*?)</u>\*\*", r"\1", text)
    text = re.sub(r"<u>(.*?)</u>", r"\1", text)
    text = text.replace("**", "")
    text = text.replace("__", "")

    return text


def strip_markdown(text):
    return clean_markup(text)


# ============================================================
# E&D INSTANCE HELPERS
# ============================================================

def new_instance_id():
    return len(st.session_state.ed_instances) + 1


def select_ed_index(index_name):
    instance_id = new_instance_id()

    st.session_state.ed_instances.append(
        {
            "id": instance_id,
            "index": index_name,
            "text": "",
        }
    )


def remove_ed_instance(instance_id):
    st.session_state.ed_instances = [
        item
        for item in st.session_state.ed_instances
        if item["id"] != instance_id
    ]


def reset_inquiry():
    st.session_state.ed_instances = []
    st.session_state.generated_draft = ""


def display_index_name(index_name, occurrence_number):
    if index_name in [
        "Statement of the Accused",
        "Questions / Answers with the Accused",
    ]:
        if occurrence_number > 1:
            return f"{index_name} No. {occurrence_number}"

    return index_name


def occurrence_number_for(index_name, instance_id):
    count = 0

    for item in st.session_state.ed_instances:
        if item["index"] == index_name:
            count += 1

        if item["id"] == instance_id:
            return count

    return 1


# ============================================================
# E&D MANIFEST
# ============================================================

def build_ed_manifest():
    grouped = {}

    for item in st.session_state.ed_instances:
        grouped.setdefault(item["index"], []).append(item)

    lines = []

    for item in st.session_state.ed_instances:
        occurrence = occurrence_number_for(
            item["index"],
            item["id"],
        )

        heading = display_index_name(
            item["index"],
            occurrence,
        )

        lines.append(f"### {heading}")
        lines.append(item.get("text", "").strip())
        lines.append("")

    return "\n".join(lines)


# ============================================================
# DOCUMENTS RECORDED
# ============================================================

def render_documents_recorded():
    st.markdown("### Documents Recorded")

    selected = []

    cols = st.columns(2)

    for i, document in enumerate(DOCUMENTS_RECORDED):
        with cols[i % 2]:
            checked = st.checkbox(
                document,
                key=f"doc_recorded_{i}",
            )

            if checked:
                selected.append(document)

    return selected


# ============================================================
# INQUIRY COMMITTEE
# ============================================================

def render_committee():
    st.markdown("### Inquiry Committee")

    committee = []

    for role in COMMITTEE_ROLES:

        st.markdown(f"**{role}**")

        col1, col2, col3 = st.columns([1, 2, 2])

        with col1:
            erp = st.text_input(
                "ERP#",
                key=f"erp_{role}",
            )

        with col2:
            name = st.text_input(
                "Name",
                key=f"name_{role}",
            )

        with col3:
            designation = st.text_input(
                "Designation",
                key=f"designation_{role}",
            )

        committee.append(
            {
                "role": role,
                "erp": erp,
                "name": name,
                "designation": designation,
            }
        )

    return committee


# ============================================================
# LANGUAGE RULES
# ============================================================

LANGUAGE_CORRECTION_RULES = """
Correct spelling, grammar, punctuation and sentence structure.

Correct obvious typing and speech-transcription errors where the
intended meaning is clear.

Convert informal wording into professional official language.

Do NOT change the substance of allegations, statements, evidence,
findings or recommendations.

Do NOT invent names, dates, allegations, evidence, witnesses,
documents, findings or facts.

If the user's statement is incomplete or unclear, preserve the
available meaning rather than inventing missing information.

For statements of accused persons and witnesses, preserve their
actual meaning faithfully.
"""


# ============================================================
# AI PROMPTS
# ============================================================

ED_SYSTEM_PROMPT = f"""
You are DraftForge, an AI assistant for preparing formal departmental
inquiry reports.

{LANGUAGE_CORRECTION_RULES}

Prepare a professional Departmental Inquiry Report.

The final report must use these major headings exactly as supplied.

Every major heading must be represented as:

**<u>HEADING</u>**

Do not create facts that are not present in the user's information.

The report must contain:

DEPARTMENTAL INQUIRY REPORT

Inquiry Reference No.
Date

Then the supplied E&D sections in EXACTLY the order supplied.

Important:
- Preserve the order.
- Do not omit supplied sections.
- Do not create additional inquiry sections.
- Do not duplicate sections.
- Do not merge separate occurrences of the same index.
- Statement of the Accused and Questions / Answers must remain separate.
- Use professional official language.
- Preserve facts and meaning.
- Do not make unsupported findings.
- If the information says a charge is not established, do not turn it
  into an established charge.
"""


NORMAL_SYSTEM_PROMPT = f"""
You are DraftForge, an AI official-document drafting assistant.

{LANGUAGE_CORRECTION_RULES}

Create a polished professional document from the user's instructions.

Never invent facts.

Do not invent names, dates, reference numbers, events, commitments,
attachments or claims.

Preserve the intended meaning.

Use formal professional government/official correspondence language.
"""


# ============================================================
# GROQ GENERATION
# ============================================================

def generate_with_groq(system_prompt, user_prompt):
    client = get_groq_client()

    if client is None:
        return None

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

    return response.choices[0].message.content.strip()


# ============================================================
# GEMINI FALLBACK
# ============================================================

def generate_with_gemini(system_prompt, user_prompt):
    api_key = get_secret("GEMINI_API_KEY")

    if not api_key:
        return None

    url = (
        "https://generativelanguage.googleapis.com/v1beta/models/"
        f"{GEMINI_MODEL}:generateContent?key={api_key}"
    )

    payload = {
        "contents": [
            {
                "parts": [
                    {
                        "text": (
                            system_prompt
                            + "\n\n"
                            + user_prompt
                        )
                    }
                ]
            }
        ]
    }

    response = requests.post(
        url,
        json=payload,
        timeout=120,
    )

    response.raise_for_status()

    data = response.json()

    try:
        return (
            data["candidates"][0]
            ["content"]
            ["parts"][0]
            ["text"]
            .strip()
        )
    except Exception:
        return None


def generate_ai(system_prompt, user_prompt):

    try:
        result = generate_with_groq(
            system_prompt,
            user_prompt,
        )

        if result:
            return result

    except Exception as exc:
        st.warning(
            f"Groq generation failed. Trying fallback: {exc}"
        )

    try:
        result = generate_with_gemini(
            system_prompt,
            user_prompt,
        )

        if result:
            return result

    except Exception as exc:
        st.error(
            f"AI generation failed: {exc}"
        )

    return None


# ============================================================
# E&D HEADING FORMATTER
# ============================================================

def format_ed_headings(text):

    if not text:
        return text

    known_headings = [
        "DEPARTMENTAL INQUIRY REPORT",
        "Inquiry Reference No.",
        "Date",
        *ED_INDEXES,
    ]

    for heading in known_headings:

        pattern = re.compile(
            rf"(?m)^(\s*){re.escape(heading)}\s*$",
            re.IGNORECASE,
        )

        text = pattern.sub(
            lambda m: (
                f"{m.group(1)}"
                f"**<u>{heading}</u>**"
            ),
            text,
        )

    # Handle numbered accused/Q&A headings
    text = re.sub(
        r"(?m)^(\s*)Statement of the Accused No\. (\d+)\s*$",
        r"\1**<u>Statement of the Accused No. \2</u>**",
        text,
        flags=re.IGNORECASE,
    )

    text = re.sub(
        r"(?m)^(\s*)Questions / Answers with the Accused No\. (\d+)\s*$",
        r"\1**<u>Questions / Answers with the Accused No. \2</u>**",
        text,
        flags=re.IGNORECASE,
    )

    return text


# ============================================================
# E&D GENERATION
# ============================================================

def generate_ed_report(
    reference_no,
    documents_recorded,
    committee,
):

    today = datetime.now().strftime("%d %B %Y")

    manifest = build_ed_manifest()

    documents_text = "\n".join(
        f"- {x}" for x in documents_recorded
    )

    committee_lines = []

    for member in committee:

        committee_lines.append(
            f"{member['role']}: "
            f"ERP# {member['erp']}; "
            f"Name: {member['name']}; "
            f"Designation: {member['designation']}"
        )

    committee_text = "\n".join(committee_lines)

    user_prompt = f"""
Prepare the Departmental Inquiry Report.

Inquiry Reference No.:
{reference_no}

Date:
{today}

The following E&D indexes were selected by the user.
Keep them in EXACTLY this order:

{manifest}

Documents Recorded:
{documents_text if documents_text else "No documents were selected."}

Inquiry Committee:
{committee_text}

Remember:
- Preserve the exact selected order.
- Preserve multiple occurrences separately.
- Correct language without changing facts.
- Do not invent missing information.
- Use bold + underline markup for every major heading.
"""

    result = generate_ai(
        ED_SYSTEM_PROMPT,
        user_prompt,
    )

    if not result:
        return None

    result = format_ed_headings(result)

    return result


# ============================================================
# EMAIL / LETTER GENERATION
# ============================================================

def generate_normal_document(
    document_type,
    recipient,
    subject,
    instructions,
):

    today = datetime.now().strftime("%d %B %Y")

    if document_type == "Email":

        structure = """
Prepare a professional official email.

Include:
To
Subject
Dear Sir/Madam,
Main body
Regards
"""

    else:

        structure = """
Prepare a professional official letter.

Include:
Date
To
Subject
Salutation
Main body
Closing
"""

    user_prompt = f"""
Document Type:
{document_type}

Date:
{today}

Recipient:
{recipient}

Subject:
{subject}

User's instructions:
{instructions}

{structure}

Correct the language while preserving the user's facts and intended
meaning.
"""

    return generate_ai(
        NORMAL_SYSTEM_PROMPT,
        user_prompt,
    )


# ============================================================
# PDF HELPERS
# ============================================================

def clean_pdf_text(text):

    replacements = {
        "–": "-",
        "—": "-",
        "’": "'",
        "‘": "'",
        "“": '"',
        "”": '"',
        "•": "-",
        "…": "...",
        "→": "->",
        "\u00a0": " ",
    }

    for old, new in replacements.items():
        text = text.replace(old, new)

    text = clean_markup(text)

    return (
        text
        .encode("latin-1", "ignore")
        .decode("latin-1")
    )


def split_long_word(pdf, word, max_width):

    if not word:
        return [""]

    if pdf.get_string_width(word) <= max_width:
        return [word]

    pieces = []
    current = ""

    for char in word:

        test = current + char

        if pdf.get_string_width(test) <= max_width:
            current = test

        else:

            if current:
                pieces.append(current)

            current = char

    if current:
        pieces.append(current)

    return pieces


def wrap_pdf_line(pdf, text, max_width):

    words = text.split()

    if not words:
        return [""]

    lines = []
    current = ""

    for word in words:

        if pdf.get_string_width(word) > max_width:

            if current:
                lines.append(current)
                current = ""

            pieces = split_long_word(
                pdf,
                word,
                max_width,
            )

            lines.extend(pieces[:-1])
            current = pieces[-1]

            continue

        test = (
            word
            if not current
            else current + " " + word
        )

        if pdf.get_string_width(test) <= max_width:
            current = test

        else:

            if current:
                lines.append(current)

            current = word

    if current:
        lines.append(current)

    return lines


def is_heading_line(line):

    stripped = line.strip()

    if stripped == "DEPARTMENTAL INQUIRY REPORT":
        return True

    if re.match(
        r"^(Inquiry Reference No\.|Date)$",
        stripped,
        re.IGNORECASE,
    ):
        return True

    for heading in ED_INDEXES:

        if stripped.lower() == heading.lower():
            return True

        if re.match(
            rf"^{re.escape(heading)} No\. \d+$",
            stripped,
            re.IGNORECASE,
        ):
            return True

    return False


def export_pdf(text):

    pdf = FPDF()

    pdf.set_auto_page_break(
        auto=True,
        margin=15,
    )

    pdf.set_margins(
        left=15,
        top=15,
        right=15,
    )

    pdf.add_page()

    usable_width = (
        pdf.w
        - pdf.l_margin
        - pdf.r_margin
    )

    for paragraph in text.splitlines():

        paragraph = clean_pdf_text(
            paragraph
        )

        if not paragraph.strip():

            pdf.ln(4)
            continue

        heading = is_heading_line(
            paragraph
        )

        if heading:

            pdf.set_font(
                "Helvetica",
                "BU",
                11,
            )

        else:

            pdf.set_font(
                "Helvetica",
                "",
                11,
            )

        lines = wrap_pdf_line(
            pdf,
            paragraph,
            usable_width,
        )

        for line in lines:

            if not line:
                pdf.ln(7)
                continue

            pdf.multi_cell(
                usable_width,
                7,
                line,
            )

    return bytes(pdf.output())


# ============================================================
# DOCX EXPORT
# ============================================================

def add_docx_line(document, line):

    line = line.strip()

    if not line:

        document.add_paragraph()
        return

    paragraph = document.add_paragraph()

    heading = is_heading_line(
        clean_markup(line)
    )

    if heading:

        run = paragraph.add_run(
            clean_markup(line)
        )

        run.bold = True
        run.underline = True

    else:

        clean_line = clean_markup(line)

        run = paragraph.add_run(
            clean_line
        )

    run.font.size = Pt(11)


def export_docx(text):

    document = Document()

    for line in text.splitlines():

        add_docx_line(
            document,
            line,
        )

    buffer = io.BytesIO()

    document.save(buffer)

    buffer.seek(0)

    return buffer.getvalue()


# ============================================================
# TXT EXPORT
# ============================================================

def export_txt(text):

    return clean_markup(text).encode(
        "utf-8"
    )


# ============================================================
# PNG EXPORT
# ============================================================

def export_png(text):

    clean_text = clean_markup(text)

    lines = []

    for paragraph in clean_text.splitlines():

        if not paragraph:
            lines.append("")
            continue

        words = paragraph.split()

        current = ""

        for word in words:

            test = (
                word
                if not current
                else current + " " + word
            )

            if len(test) <= 95:
                current = test
            else:
                if current:
                    lines.append(current)
                current = word

        if current:
            lines.append(current)

    width = 1600
    line_height = 32
    padding = 50

    height = max(
        500,
        padding * 2
        + line_height * len(lines),
    )

    image = Image.new(
        "RGB",
        (width, height),
        "white",
    )

    draw = ImageDraw.Draw(image)

    try:
        font = ImageFont.truetype(
            "DejaVuSans.ttf",
            22,
        )

        bold_font = ImageFont.truetype(
            "DejaVuSans-Bold.ttf",
            22,
        )

    except Exception:

        font = ImageFont.load_default()
        bold_font = font

    y = padding

    for line in lines:

        if is_heading_line(line):

            draw.text(
                (padding, y),
                line,
                fill="black",
                font=bold_font,
            )

            bbox = draw.textbbox(
                (padding, y),
                line,
                font=bold_font,
            )

            draw.line(
                (
                    bbox[0],
                    bbox[3] + 2,
                    bbox[2],
                    bbox[3] + 2,
                ),
                fill="black",
                width=1,
            )

        else:

            draw.text(
                (padding, y),
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

    return output.getvalue()


# ============================================================
# HISTORY
# ============================================================

def save_history(document_type, draft):

    if not draft:
        return

    st.session_state.history.insert(
        0,
        {
            "type": document_type,
            "date": datetime.now().strftime(
                "%d %B %Y %H:%M"
            ),
            "draft": draft,
        },
    )

    st.session_state.history = (
        st.session_state.history[:10]
    )


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.markdown(
        "## 📝 DraftForge"
    )

    st.caption(
        "AI Document Composer"
    )

    st.divider()

    st.markdown(
        "### Quick Guide"
    )

    st.markdown(
        """
**1. Choose a document type**

✉️ Email  
📄 Letter  
🔎 Inquiry

**2. Enter information**

Type naturally or use the 🎙️ microphone.

**3. Generate**

DraftForge corrects language while
preserving your facts.

**4. Export**

Download PDF, DOCX, TXT or PNG.
"""
    )

    st.divider()

    if st.button(
        "🗑️ Clear Current Draft",
        use_container_width=True,
    ):
        st.session_state.generated_draft = ""
        st.rerun()

    if st.session_state.history:

        st.divider()

        st.markdown(
            "### Recent Drafts"
        )

        for item in st.session_state.history[:5]:

            with st.expander(
                f"{item['type']} — {item['date']}"
            ):

                st.text(
                    clean_markup(
                        item["draft"]
                    )[:1000]
                )


# ============================================================
# MAIN HEADER
# ============================================================

st.markdown(
    '<div class="main-title">📝 DraftForge</div>',
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="subtitle">'
    "AI Document Composer — create professional "
    "official documents using text or voice"
    "</div>",
    unsafe_allow_html=True,
)


# ============================================================
# PROMINENT DOCUMENT TYPE SELECTOR
# ============================================================

st.markdown(
    "### 📌 What would you like to create?"
)

document_type = st.segmented_control(
    "Document Type",
    options=[
        "✉️ Email",
        "📄 Letter",
        "🔎 Inquiry",
    ],
    default="✉️ Email",
    key="document_type_selector",
    selection_mode="single",
    width="stretch",
    label_visibility="collapsed",
)

if document_type is None:
    document_type = "✉️ Email"

if "Email" in document_type:
    st.session_state.document_type = "Email"

elif "Letter" in document_type:
    st.session_state.document_type = "Letter"

elif "Inquiry" in document_type:
    st.session_state.document_type = "Inquiry"


current_document_type = st.session_state.document_type

st.divider()


# ============================================================
# EMAIL
# ============================================================

if current_document_type == "Email":

    st.markdown(
        '<div class="section-title">✉️ Create Email</div>',
        unsafe_allow_html=True,
    )

    recipient = st.text_input(
        "Recipient",
        placeholder="Enter recipient or office",
        key="email_recipient",
    )

    subject = st.text_input(
        "Subject",
        placeholder="Enter email subject",
        key="email_subject",
    )

    st.markdown(
        "#### Email Instructions"
    )

    st.caption(
        "Type your instructions below. "
        "For voice, press the microphone button "
        "inside the input."
    )

    email_input = st.chat_input(
        "Type your email instructions or 🎙️ record your voice...",
        key="email_composer",
        accept_audio=True,
        audio_sample_rate=16000,
        width="stretch",
    )

    if email_input:

        typed_text = email_input.text or ""

        voice_text = ""

        if email_input.audio is not None:

            with st.spinner(
                "🎙️ Transcribing your voice..."
            ):

                try:

                    voice_text = transcribe_audio(
                        email_input.audio
                    )

                except Exception as exc:

                    st.error(str(exc))

        combined = append_text(
            typed_text,
            voice_text,
        )

        if combined:

            st.session_state.email_instruction_text = (
                append_text(
                    st.session_state.get(
                        "email_instruction_text",
                        "",
                    ),
                    combined,
                )
            )

    existing_email_text = st.session_state.get(
        "email_instruction_text",
        "",
    )

    if existing_email_text:

        st.markdown(
            "#### 📝 Collected Instructions"
        )

        st.text_area(
            "Collected email instructions",
            value=existing_email_text,
            height=180,
            disabled=True,
            label_visibility="collapsed",
        )

    if st.button(
        "✨ Generate Email",
        type="primary",
        use_container_width=True,
    ):

        instructions = st.session_state.get(
            "email_instruction_text",
            "",
        )

        if not instructions.strip():

            st.warning(
                "Please enter or record your email instructions."
            )

        else:

            with st.spinner(
                "DraftForge is preparing your email..."
            ):

                draft = generate_normal_document(
                    "Email",
                    recipient,
                    subject,
                    instructions,
                )

            if draft:

                st.session_state.generated_draft = draft

                save_history(
                    "Email",
                    draft,
                )

                st.rerun()


# ============================================================
# LETTER
# ============================================================

elif current_document_type == "Letter":

    st.markdown(
        '<div class="section-title">📄 Create Official Letter</div>',
        unsafe_allow_html=True,
    )

    recipient = st.text_input(
        "Recipient / Office",
        placeholder="Enter recipient, office or designation",
        key="letter_recipient",
    )

    subject = st.text_input(
        "Subject",
        placeholder="Enter letter subject",
        key="letter_subject",
    )

    st.markdown(
        "#### Letter Instructions"
    )

    st.caption(
        "Type naturally or press 🎙️ to record."
    )

    letter_input = st.chat_input(
        "Type your letter instructions or 🎙️ record your voice...",
        key="letter_composer",
        accept_audio=True,
        audio_sample_rate=16000,
        width="stretch",
    )

    if letter_input:

        typed_text = letter_input.text or ""

        voice_text = ""

        if letter_input.audio is not None:

            with st.spinner(
                "🎙️ Transcribing your voice..."
            ):

                try:

                    voice_text = transcribe_audio(
                        letter_input.audio
                    )

                except Exception as exc:

                    st.error(str(exc))

        combined = append_text(
            typed_text,
            voice_text,
        )

        if combined:

            st.session_state.letter_instruction_text = (
                append_text(
                    st.session_state.get(
                        "letter_instruction_text",
                        "",
                    ),
                    combined,
                )
            )

    existing_letter_text = st.session_state.get(
        "letter_instruction_text",
        "",
    )

    if existing_letter_text:

        st.markdown(
            "#### 📝 Collected Instructions"
        )

        st.text_area(
            "Collected letter instructions",
            value=existing_letter_text,
            height=180,
            disabled=True,
            label_visibility="collapsed",
        )

    if st.button(
        "✨ Generate Letter",
        type="primary",
        use_container_width=True,
    ):

        instructions = st.session_state.get(
            "letter_instruction_text",
            "",
        )

        if not instructions.strip():

            st.warning(
                "Please enter or record your letter instructions."
            )

        else:

            with st.spinner(
                "DraftForge is preparing your letter..."
            ):

                draft = generate_normal_document(
                    "Letter",
                    recipient,
                    subject,
                    instructions,
                )

            if draft:

                st.session_state.generated_draft = draft

                save_history(
                    "Letter",
                    draft,
                )

                st.rerun()


# ============================================================
# INQUIRY
# ============================================================

elif current_document_type == "Inquiry":

    st.markdown(
        '<div class="section-title">🔎 Departmental Inquiry</div>',
        unsafe_allow_html=True,
    )

    inquiry_type = st.segmented_control(
        "Inquiry Type",
        options=[
            "⚖️ E&D Inquiry",
            "🔍 FFI Inquiry",
        ],
        default="⚖️ E&D Inquiry",
        key="inquiry_type_selector",
        selection_mode="single",
        width="stretch",
    )

    if inquiry_type is None:
        inquiry_type = "⚖️ E&D Inquiry"

    if "FFI" in inquiry_type:

        st.session_state.inquiry_type = "FFI Inquiry"

        st.markdown(
            """
            <div class="warning-box">
            <h3>🔍 FFI Inquiry</h3>
            <b>Under Construction</b><br><br>
            The FFI inquiry module is currently under process.
            The E&D Inquiry module is available for use.
            </div>
            """,
            unsafe_allow_html=True,
        )

    else:

        st.session_state.inquiry_type = "E&D Inquiry"

        # ----------------------------------------------------
        # REFERENCE NUMBER
        # ----------------------------------------------------

        reference_no = st.text_input(
            "Inquiry Reference No.",
            placeholder="e.g. ABC/123",
            key="inquiry_reference_no",
        )

        st.markdown(
            """
            <div class="help-box">
            <b>How to build your inquiry:</b><br>
            Select an index below, then type naturally or use the
            microphone. You can select the same index multiple times.
            Each selection remains as a separate section.
            </div>
            """,
            unsafe_allow_html=True,
        )

        # ----------------------------------------------------
        # INDEX SELECTOR
        # ----------------------------------------------------

        selected_index = st.selectbox(
            "➕ Add Inquiry Index",
            options=ED_INDEXES,
            key="ed_index_selector",
        )

        if st.button(
            "➕ Add Selected Index",
            use_container_width=True,
        ):

            select_ed_index(
                selected_index
            )

            st.rerun()

        # ----------------------------------------------------
        # CURRENT INDEXES
        # ----------------------------------------------------

        if st.session_state.ed_instances:

            st.markdown(
                "### 📋 Inquiry Information"
            )

            for item in list(
                st.session_state.ed_instances
            ):

                instance_id = item["id"]
                index_name = item["index"]

                occurrence = occurrence_number_for(
                    index_name,
                    instance_id,
                )

                heading = display_index_name(
                    index_name,
                    occurrence,
                )

                with st.container(
                    border=True
                ):

                    col1, col2 = st.columns(
                        [6, 1]
                    )

                    with col1:

                        st.markdown(
                            f"#### {heading}"
                        )

                    with col2:

                        if st.button(
                            "🗑️",
                            key=f"remove_{instance_id}",
                            help="Remove this section",
                        ):

                            remove_ed_instance(
                                instance_id
                            )

                            st.rerun()

                    st.caption(
                        "Type information or press 🎙️ "
                        "to record it. You may submit several "
                        "text/voice entries for the same index."
                    )

                    # ----------------------------------------
                    # Integrated text + voice input
                    # ----------------------------------------

                    composer_key = (
                        f"ed_composer_{instance_id}"
                    )

                    submission = st.chat_input(
                        f"Enter information for {heading} or 🎙️ record your voice...",
                        key=composer_key,
                        accept_audio=True,
                        audio_sample_rate=16000,
                        width="stretch",
                    )

                    if submission:

                        typed_text = (
                            submission.text
                            or ""
                        )

                        voice_text = ""

                        if submission.audio is not None:

                            with st.spinner(
                                "🎙️ Transcribing..."
                            ):

                                try:

                                    voice_text = (
                                        transcribe_audio(
                                            submission.audio
                                        )
                                    )

                                except Exception as exc:

                                    st.error(
                                        str(exc)
                                    )

                        combined = append_text(
                            typed_text,
                            voice_text,
                        )

                        if combined:

                            for stored in (
                                st.session_state.ed_instances
                            ):

                                if (
                                    stored["id"]
                                    == instance_id
                                ):

                                    stored["text"] = (
                                        append_text(
                                            stored.get(
                                                "text",
                                                "",
                                            ),
                                            combined,
                                        )
                                    )

                            st.rerun()

                    # ----------------------------------------
                    # Display accumulated input
                    # ----------------------------------------

                    current_text = item.get(
                        "text",
                        "",
                    )

                    if current_text:

                        st.text_area(
                            "Collected information",
                            value=current_text,
                            height=180,
                            disabled=True,
                            key=(
                                f"preview_{instance_id}"
                            ),
                            label_visibility="collapsed",
                        )

        # ----------------------------------------------------
        # DOCUMENTS RECORDED
        # ----------------------------------------------------

        documents_recorded = []

        if any(
            item["index"]
            == "Documents Recorded"
            for item in st.session_state.ed_instances
        ):

            documents_recorded = (
                render_documents_recorded()
            )

        # ----------------------------------------------------
        # INQUIRY COMMITTEE
        # ----------------------------------------------------

        committee = []

        if any(
            item["index"]
            == "Inquiry Committee"
            for item in st.session_state.ed_instances
        ):

            committee = render_committee()

        # ----------------------------------------------------
        # ACTION BUTTONS
        # ----------------------------------------------------

        st.divider()

        col1, col2 = st.columns(2)

        with col1:

            if st.button(
                "✨ Generate Inquiry Report",
                type="primary",
                use_container_width=True,
            ):

                if not reference_no.strip():

                    st.warning(
                        "Please enter the Inquiry Reference No."
                    )

                elif not st.session_state.ed_instances:

                    st.warning(
                        "Please add at least one inquiry index."
                    )

                else:

                    has_information = any(
                        item.get("text", "").strip()
                        for item
                        in st.session_state.ed_instances
                    )

                    if not has_information:

                        st.warning(
                            "Please provide information in at least "
                            "one inquiry index."
                        )

                    else:

                        with st.spinner(
                            "DraftForge is preparing the "
                            "Departmental Inquiry Report..."
                        ):

                            draft = generate_ed_report(
                                reference_no,
                                documents_recorded,
                                committee,
                            )

                        if draft:

                            st.session_state.generated_draft = (
                                draft
                            )

                            save_history(
                                "E&D Inquiry",
                                draft,
                            )

                            st.rerun()

        with col2:

            if st.button(
                "🔄 Start New Inquiry",
                use_container_width=True,
            ):

                reset_inquiry()
                st.rerun()


# ============================================================
# GENERATED DOCUMENT
# ============================================================

if st.session_state.generated_draft:

    st.divider()

    st.markdown(
        '<div class="section-title">'
        "📄 Generated Document"
        "</div>",
        unsafe_allow_html=True,
    )

    draft = st.session_state.generated_draft

    st.markdown(
        '<div class="generated-box">',
        unsafe_allow_html=True,
    )

    # Render markdown/underline formatting
    st.markdown(
        draft,
        unsafe_allow_html=True,
    )

    st.markdown(
        "</div>",
        unsafe_allow_html=True,
    )

    st.divider()

    st.markdown(
        "### 📥 Download"

    )

    pdf_data = export_pdf(
        draft
    )

    docx_data = export_docx(
        draft
    )

    txt_data = export_txt(
        draft
    )

    png_data = export_png(
        draft
    )

    c1, c2, c3, c4 = st.columns(4)

    with c1:

        st.download_button(
            "📕 PDF",
            data=pdf_data,
            file_name="DraftForge_Document.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    with c2:

        st.download_button(
            "📘 DOCX",
            data=docx_data,
            file_name="DraftForge_Document.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument"
                ".wordprocessingml.document"
            ),
            use_container_width=True,
        )

    with c3:

        st.download_button(
            "📄 TXT",
            data=txt_data,
            file_name="DraftForge_Document.txt",
            mime="text/plain",
            use_container_width=True,
        )

    with c4:

        st.download_button(
            "🖼️ PNG",
            data=png_data,
            file_name="DraftForge_Document.png",
            mime="image/png",
            use_container_width=True,
        )
