import io
import os
import re
import sqlite3
import uuid
from datetime import datetime

import requests
import streamlit as st
from docx import Document
from docx.shared import Pt
from fpdf import FPDF
from PIL import Image
from groq import Groq


# ============================================================
# APP CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="DraftForge — AI Document Composer",
    page_icon="📝",
    layout="wide",
)

APP_TITLE = "DraftForge"
APP_SUBTITLE = "AI Document Composer"

GROQ_MODEL = "openai/gpt-oss-120b"
WHISPER_MODEL = "whisper-large-v3"

GEMINI_MODEL = "gemini-2.0-flash"

DOC_EMAIL = "Email"
DOC_LETTER = "Letter"
DOC_INQUIRY = "Inquiry"

INQUIRY_FFI = "FFI Inquiry"
INQUIRY_ED = "E&D Inquiry"

LANGUAGES = [
    "English",
    "Urdu",
]

TONES = [
    "Professional",
    "Formal",
    "Concise",
    "Detailed",
]


# ============================================================
# E&D INQUIRY INDEXES
# ============================================================

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

NUMBERABLE_INDEXES = [
    "Statement of the Accused",
    "Questions / Answers with the Accused",
]


# ============================================================
# DATABASE
# ============================================================

DB_FILE = "draftforge.db"


def init_db():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT,
            document_type TEXT,
            inquiry_type TEXT,
            subject TEXT,
            content TEXT
        )
        """
    )

    conn.commit()
    conn.close()


init_db()


def save_history(document_type, inquiry_type, subject, content):
    try:
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()

        cur.execute(
            """
            INSERT INTO history
            (created_at, document_type, inquiry_type, subject, content)
            VALUES (?, ?, ?, ?, ?)
            """,
            (
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                document_type,
                inquiry_type,
                subject,
                content,
            ),
        )

        conn.commit()
        conn.close()
    except Exception:
        pass


def load_history(limit=30):
    try:
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()

        cur.execute(
            """
            SELECT id, created_at, document_type, inquiry_type,
                   subject, content
            FROM history
            ORDER BY id DESC
            LIMIT ?
            """,
            (limit,),
        )

        rows = cur.fetchall()
        conn.close()
        return rows
    except Exception:
        return []


# ============================================================
# SESSION STATE
# ============================================================

DEFAULT_STATE = {
    "generated_draft": "",
    "normal_information": "",
    "normal_voice_status": "",
    "ed_index_instances": [],
    "ed_index_selector": "Select an index...",
    "selected_document_type": DOC_EMAIL,
    "selected_inquiry_type": INQUIRY_ED,
    "selected_language": "English",
    "selected_tone": "Professional",
}

for key, value in DEFAULT_STATE.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# API KEY HELPERS
# ============================================================

def get_secret(name):
    """
    Reads API keys safely from Streamlit secrets first,
    then environment variables.
    """

    try:
        value = st.secrets.get(name)
        if value:
            return value
    except Exception:
        pass

    return os.getenv(name)


def get_groq_client():
    key = get_secret("GROQ_API_KEY")

    if not key:
        raise ValueError(
            "GROQ_API_KEY is not configured. "
            "Please add it to Streamlit Secrets."
        )

    return Groq(api_key=key)


# ============================================================
# VOICE TRANSCRIPTION
# ============================================================

def transcribe_audio(audio_bytes, filename="audio.wav"):
    """
    Transcribe uploaded/recorded audio through Groq Whisper.
    """

    client = get_groq_client()

    audio_file = io.BytesIO(audio_bytes)
    audio_file.name = filename

    result = client.audio.transcriptions.create(
        file=audio_file,
        model=WHISPER_MODEL,
        response_format="text",
    )

    if isinstance(result, str):
        return result.strip()

    return str(result).strip()


def append_voice_to_widget(widget_key, transcript):
    """
    Safely append transcription to an existing Streamlit widget
    BEFORE the widget is rendered again.
    """

    transcript = (transcript or "").strip()

    if not transcript:
        return

    existing = st.session_state.get(widget_key, "")

    if existing.strip():
        st.session_state[widget_key] = (
            existing.rstrip() + "\n" + transcript
        )
    else:
        st.session_state[widget_key] = transcript


def add_normal_voice():
    uploaded = st.session_state.get("normal_voice_upload")

    if not uploaded:
        st.session_state.normal_voice_status = (
            "Please record or upload an audio file first."
        )
        return

    try:
        transcript = transcribe_audio(
            uploaded.getvalue(),
            uploaded.name,
        )

        append_voice_to_widget(
            "normal_information",
            transcript,
        )

        st.session_state.normal_voice_status = (
            "Voice transcription added to the text box."
        )

    except Exception as e:
        st.session_state.normal_voice_status = (
            f"Voice transcription failed: {e}"
        )


def add_inquiry_voice(instance_id):
    upload_key = f"voice_upload_{instance_id}"
    text_key = f"inquiry_text_{instance_id}"
    status_key = f"voice_status_{instance_id}"

    uploaded = st.session_state.get(upload_key)

    if not uploaded:
        st.session_state[status_key] = (
            "Please record or upload an audio file first."
        )
        return

    try:
        transcript = transcribe_audio(
            uploaded.getvalue(),
            uploaded.name,
        )

        append_voice_to_widget(
            text_key,
            transcript,
        )

        st.session_state[status_key] = (
            "Voice transcription added to the text box."
        )

    except Exception as e:
        st.session_state[status_key] = (
            f"Voice transcription failed: {e}"
        )


# ============================================================
# E&D INDEX MANAGEMENT
# ============================================================

def get_display_index_name(instances, current_instance):
    """
    Automatically numbers only:
      Statement of the Accused
      Questions / Answers with the Accused

    If one occurrence:
      Statement of the Accused

    If repeated:
      Statement of the Accused No. 1
      Statement of the Accused No. 2
    """

    index_name = current_instance["index_name"]

    if index_name not in NUMBERABLE_INDEXES:
        return index_name

    same_instances = [
        x
        for x in instances
        if x["index_name"] == index_name
    ]

    if len(same_instances) <= 1:
        return index_name

    number = same_instances.index(current_instance) + 1

    return f"{index_name} No. {number}"


def select_ed_index():
    selected = st.session_state.get(
        "ed_index_selector",
        "Select an index...",
    )

    if (
        not selected
        or selected == "Select an index..."
        or selected not in ED_INDEXES
    ):
        return

    instance_id = uuid.uuid4().hex[:12]

    st.session_state.ed_index_instances.append(
        {
            "id": instance_id,
            "index_name": selected,
        }
    )

    # Important:
    # Reset selector after callback, not after widget creation.
    st.session_state.ed_index_selector = "Select an index..."


def remove_ed_instance(instance_id):
    remaining = [
        x
        for x in st.session_state.ed_index_instances
        if x["id"] != instance_id
    ]

    st.session_state.ed_index_instances = remaining

    text_key = f"inquiry_text_{instance_id}"
    voice_key = f"voice_upload_{instance_id}"
    status_key = f"voice_status_{instance_id}"

    for key in [text_key, voice_key, status_key]:
        st.session_state.pop(key, None)


def reset_inquiry():
    st.session_state.ed_index_instances = []
    st.session_state.ed_index_selector = "Select an index..."
    st.session_state.generated_draft = ""

    keys_to_remove = []

    for key in list(st.session_state.keys()):
        if (
            key.startswith("inquiry_text_")
            or key.startswith("voice_upload_")
            or key.startswith("voice_status_")
            or key.startswith("committee_")
            or key.startswith("doc_record_")
        ):
            keys_to_remove.append(key)

    for key in keys_to_remove:
        st.session_state.pop(key, None)


# ============================================================
# E&D INFORMATION BUILDERS
# ============================================================

def build_selected_index_manifest():
    instances = st.session_state.ed_index_instances

    lines = []

    for instance in instances:
        display_name = get_display_index_name(
            instances,
            instance,
        )

        lines.append(
            f"- {display_name}"
        )

    return "\n".join(lines)


def build_ed_information():
    instances = st.session_state.ed_index_instances

    blocks = []

    for instance in instances:
        instance_id = instance["id"]

        display_name = get_display_index_name(
            instances,
            instance,
        )

        text_key = f"inquiry_text_{instance_id}"

        value = st.session_state.get(
            text_key,
            "",
        ).strip()

        blocks.append(
            f"""
INDEX: {display_name}

USER INPUT:
{value}
""".strip()
        )

    return "\n\n".join(blocks)


def get_first_index_value(index_name):
    instances = st.session_state.ed_index_instances

    for instance in instances:
        if instance["index_name"] == index_name:

            key = f"inquiry_text_{instance['id']}"

            return st.session_state.get(
                key,
                "",
            ).strip()

    return ""


# ============================================================
# STRUCTURED E&D INFORMATION
# ============================================================

DOCUMENT_RECORD_OPTIONS = [
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


def build_documents_recorded():
    selected = []

    for item in DOCUMENT_RECORD_OPTIONS:
        key = f"doc_record_{item}"

        if st.session_state.get(key, False):
            selected.append(item)

    return selected


def build_inquiry_committee():
    committee = []

    for role in COMMITTEE_ROLES:

        erp = st.session_state.get(
            f"committee_erp_{role}",
            "",
        ).strip()

        name = st.session_state.get(
            f"committee_name_{role}",
            "",
        ).strip()

        designation = st.session_state.get(
            f"committee_designation_{role}",
            "",
        ).strip()

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
# AI PROMPTS
# ============================================================

LANGUAGE_CORRECTION_RULES = """
LANGUAGE QUALITY RULES:

1. Correct spelling mistakes in the user's input.
2. Correct grammatical mistakes.
3. Correct punctuation and sentence structure.
4. Correct obvious typing mistakes.
5. Correct obvious speech-to-text/transcription mistakes where the intended
   meaning is reasonably clear.
6. Convert informal wording into professional official language where
   appropriate.
7. Preserve the user's original meaning and factual substance.
8. NEVER invent facts, names, dates, events, allegations, evidence,
   statements, findings, or conclusions.
9. NEVER assume information that the user has not supplied.
10. Do not change the substance of a statement merely to make it sound
    more professional.
11. For statements of an accused, witnesses, or officials, preserve the
    actual position/meaning expressed by the person.
12. Preserve identifiers such as ERP numbers, inquiry reference numbers,
    case numbers, CNIC-related numbers, names, and official designations
    unless there is an obvious spelling/transcription error.
"""


ED_SYSTEM_PROMPT = f"""
You are an expert departmental inquiry report drafting assistant.

Your task is to convert the user's raw natural-language information into
a professional, grammatically correct and properly structured
Departmental Inquiry Report.

{LANGUAGE_CORRECTION_RULES}

STRICT E&D INDEX RULES:

1. The SELECTED INDEX LIST is authoritative.
2. Use ONLY the indexes selected by the user.
3. NEVER add an unselected index.
4. NEVER mention an unselected index.
5. NEVER create a heading such as "Not Recorded" for an unselected index.
6. Preserve the exact order in which the user selected the indexes.
7. If an index has been selected more than once, keep every occurrence.
8. For repeated "Statement of the Accused", preserve:
   Statement of the Accused No. 1
   Statement of the Accused No. 2
   etc.
9. For repeated "Questions / Answers with the Accused", preserve:
   Questions / Answers with the Accused No. 1
   Questions / Answers with the Accused No. 2
   etc.
10. If these indexes occur only once, use their original heading without
    a number.
11. Do NOT introduce a separate "Accused No. 1", "Add Accused No. 2",
    or similar workflow.
12. Do not merge two separately selected instances into one section.
13. Preserve every selected section even when the user's input is short.
14. Do not fabricate content to fill a section.

CONTENT RULES:

- Use only facts supplied by the user.
- Correct grammar and spelling intelligently.
- Improve official wording.
- Keep the report objective and professional.
- Do not change the meaning of evidence or statements.
- Do not reach a finding that is unsupported by the user's information.
- Do not manufacture legal provisions, rules, dates, evidence or testimony.

OUTPUT:

Return only the final departmental inquiry report.
Do not explain your corrections.
Do not include analysis outside the report.
"""


NORMAL_SYSTEM_PROMPT = f"""
You are an expert official correspondence drafting assistant.

You prepare professional:
- Emails
- Official letters

{LANGUAGE_CORRECTION_RULES}

The user's raw text may contain spelling mistakes, grammar mistakes,
poor sentence structure, informal wording, abbreviations, or voice
transcription errors.

You must intelligently correct these problems while preserving the
user's intended meaning and facts.

Do not invent facts or add information that was not supplied.

Return only the finished document.
"""


# ============================================================
# GROQ GENERATION
# ============================================================

def generate_with_groq(system_prompt, user_prompt):
    client = get_groq_client()

    response = client.chat.completions.create(
        model=GROQ_MODEL,
        temperature=0.2,
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
# GEMINI GENERATION
# ============================================================

def generate_with_gemini(system_prompt, user_prompt):
    api_key = get_secret("GEMINI_API_KEY")

    if not api_key:
        raise ValueError(
            "GEMINI_API_KEY is not configured."
        )

    url = (
        f"https://generativelanguage.googleapis.com/v1beta/models/"
        f"{GEMINI_MODEL}:generateContent?key={api_key}"
    )

    payload = {
        "systemInstruction": {
            "parts": [
                {
                    "text": system_prompt,
                }
            ]
        },
        "contents": [
            {
                "role": "user",
                "parts": [
                    {
                        "text": user_prompt,
                    }
                ],
            }
        ],
        "generationConfig": {
            "temperature": 0.2,
        },
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
            data["candidates"][0]["content"]["parts"][0]["text"]
            .strip()
        )
    except Exception:
        raise ValueError(
            f"Unexpected Gemini response: {data}"
        )


# ============================================================
# GENERAL GENERATION ROUTER
# ============================================================

def generate_ai(
    system_prompt,
    user_prompt,
    provider="Groq",
):
    if provider == "Gemini":
        return generate_with_gemini(
            system_prompt,
            user_prompt,
        )

    return generate_with_groq(
        system_prompt,
        user_prompt,
    )


# ============================================================
# E&D HEADER
# ============================================================

def add_inquiry_header(
    draft,
    reference_number,
):
    date_string = datetime.now().strftime(
        "%d %B %Y"
    )

    reference_number = (
        reference_number.strip()
        if reference_number
        else ""
    )

    header = (
        "DEPARTMENTAL INQUIRY REPORT\n\n"
        f"Inquiry Reference No.: {reference_number}\n"
        f"Date: {date_string}\n\n"
    )

    return header + draft.strip()


# ============================================================
# MARKDOWN / HTML CLEANING
# ============================================================

def clean_markup(text):
    """
    Converts simple Markdown/HTML formatting into clean plain text.

    Examples:
       **<u>Subject</u>**
       becomes:
       Subject
    """

    if not text:
        return ""

    text = text.replace("\r\n", "\n")
    text = text.replace("\r", "\n")

    text = re.sub(
        r"<u>(.*?)</u>",
        r"\1",
        text,
        flags=re.IGNORECASE | re.DOTALL,
    )

    text = text.replace("**", "")
    text = text.replace("__", "")

    text = re.sub(
        r"<br\s*/?>",
        "\n",
        text,
        flags=re.IGNORECASE,
    )

    text = re.sub(
        r"<[^>]+>",
        "",
        text,
    )

    return text


# ============================================================
# PDF HELPERS
# ============================================================

def clean_pdf_text(text):
    """
    Make text safe for standard Helvetica/Latin-1 PDF output.
    """

    text = clean_markup(text)

    replacements = {
        "\u2013": "-",
        "\u2014": "-",
        "\u2018": "'",
        "\u2019": "'",
        "\u201c": '"',
        "\u201d": '"',
        "\u2022": "-",
        "\u2026": "...",
        "\u2192": "->",
        "\u00a0": " ",
        "\u2011": "-",
    }

    for old, new in replacements.items():
        text = text.replace(old, new)

    # Remove any remaining characters that Helvetica/Latin-1
    # cannot safely render.
    text = (
        text.encode(
            "latin-1",
            "ignore",
        )
        .decode("latin-1")
    )

    return text


def pdf_safe_chunks(pdf, text, usable_width):
    """
    Safely wrap text.

    Critical fix:
    A single extremely long word/token can cause FPDF2's
    multi_cell() to throw:

        Not enough horizontal space to render a single character

    Therefore long tokens are split character-by-character
    when necessary.
    """

    if not text:
        return [""]

    words = text.split(" ")

    lines = []
    current = ""

    for word in words:

        if not word:
            candidate = current + " "
        else:
            candidate = (
                word
                if not current
                else current + " " + word
            )

        if pdf.get_string_width(candidate) <= usable_width:
            current = candidate
            continue

        # Current line exists; save it.
        if current.strip():
            lines.append(current.rstrip())

        # Handle a word that itself is wider than the page.
        if pdf.get_string_width(word) > usable_width:

            chunk = ""

            for char in word:
                test = chunk + char

                if (
                    pdf.get_string_width(test)
                    <= usable_width
                ):
                    chunk = test
                else:
                    if chunk:
                        lines.append(chunk)

                    chunk = char

            current = chunk

        else:
            current = word

    if current.strip():
        lines.append(current.rstrip())

    return lines or [""]


def export_pdf(text):
    """
    Robust FPDF2 PDF exporter.
    """

    pdf = FPDF(
        orientation="P",
        unit="mm",
        format="A4",
    )

    pdf.set_margins(
        left=15,
        top=15,
        right=15,
    )

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

    safe_text = clean_pdf_text(text)

    paragraphs = safe_text.split("\n")

    for paragraph in paragraphs:

        if not paragraph.strip():
            pdf.ln(4)
            continue

        lines = pdf_safe_chunks(
            pdf,
            paragraph,
            usable_width,
        )

        for line in lines:

            if not line:
                pdf.ln(4)
                continue

            pdf.multi_cell(
                w=usable_width,
                h=7,
                text=line,
                border=0,
                align="L",
            )

    output = pdf.output()

    if isinstance(output, bytes):
        return output

    return bytes(output)


# ============================================================
# DOCX EXPORT
# ============================================================

def export_docx(text):
    doc = Document()

    section = doc.sections[0]

    section.top_margin = Pt(40)
    section.bottom_margin = Pt(40)
    section.left_margin = Pt(50)
    section.right_margin = Pt(50)

    lines = clean_markup(text).split("\n")

    for line in lines:

        p = doc.add_paragraph()

        if not line.strip():
            continue

        run = p.add_run(line)
        run.font.name = "Arial"
        run.font.size = Pt(11)

    buffer = io.BytesIO()

    doc.save(buffer)

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
    """
    Simple image export for sharing/preview.
    """

    from PIL import ImageDraw, ImageFont

    text = clean_markup(text)

    try:
        font = ImageFont.truetype(
            "DejaVuSans.ttf",
            22,
        )
    except Exception:
        font = ImageFont.load_default()

    lines = text.split("\n")

    width = 1600
    line_height = 35

    height = max(
        400,
        len(lines) * line_height + 80,
    )

    image = Image.new(
        "RGB",
        (width, height),
        "white",
    )

    draw = ImageDraw.Draw(image)

    y = 40

    for line in lines:
        draw.text(
            (50, y),
            line,
            fill="black",
            font=font,
        )
        y += line_height

    buffer = io.BytesIO()

    image.save(
        buffer,
        format="PNG",
    )

    return buffer.getvalue()


# ============================================================
# UI HEADER
# ============================================================

st.title("📝 DraftForge")
st.caption(APP_SUBTITLE)


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.header("⚙️ Settings")

    document_type = st.selectbox(
        "Document Type",
        [
            DOC_EMAIL,
            DOC_LETTER,
            DOC_INQUIRY,
        ],
        key="selected_document_type",
    )

    language = st.selectbox(
        "Output Language",
        LANGUAGES,
        key="selected_language",
    )

    tone = st.selectbox(
        "Tone",
        TONES,
        key="selected_tone",
    )

    provider = st.selectbox(
        "AI Provider",
        [
            "Groq",
            "Gemini",
        ],
    )

    st.divider()

    st.subheader("📚 History")

    history_rows = load_history()

    if not history_rows:
        st.caption("No drafts saved yet.")

    else:

        for row in history_rows:

            record_id, created_at, dtype, itype, subject, content = row

            label = (
                f"{dtype} | "
                f"{subject[:35] if subject else 'Untitled'}"
            )

            with st.expander(label):

                st.caption(created_at)

                if itype:
                    st.caption(itype)

                st.text_area(
                    "Draft",
                    content,
                    height=180,
                    key=f"history_{record_id}",
                )


# ============================================================
# NORMAL DOCUMENTS
# ============================================================

if document_type in [DOC_EMAIL, DOC_LETTER]:

    st.subheader(f"Create {document_type}")

    if document_type == DOC_EMAIL:

        sender = st.text_input(
            "Sender / From",
            placeholder="e.g. Assistant Director Technical",
        )

        recipient = st.text_input(
            "Recipient / To",
            placeholder="e.g. Regional Head",
        )

        subject = st.text_input(
            "Subject",
            placeholder="Enter subject",
        )

    else:

        recipient = st.text_input(
            "Recipient",
            placeholder="e.g. Regional Head",
        )

        subject = st.text_input(
            "Subject",
            placeholder="Enter subject",
        )

        sender = st.text_input(
            "Sender",
            placeholder="e.g. Assistant Director Technical",
        )

    st.markdown("### Information")

    st.text_area(
        "Enter information in natural language",
        key="normal_information",
        height=220,
        placeholder=(
            "Type your information here.\n\n"
            "You can write naturally and make spelling or grammar "
            "mistakes. DraftForge will correct them."
        ),
    )

    st.file_uploader(
        "🎙️ Record / Upload Voice",
        type=[
            "wav",
            "mp3",
            "m4a",
            "ogg",
            "webm",
        ],
        key="normal_voice_upload",
    )

    st.button(
        "➕ Add Voice to Text",
        on_click=add_normal_voice,
    )

    if st.session_state.normal_voice_status:
        st.info(
            st.session_state.normal_voice_status
        )

    st.divider()

    if st.button(
        "✨ Generate Draft",
        type="primary",
        use_container_width=True,
    ):

        if not st.session_state.normal_information.strip():

            st.warning(
                "Please enter some information first."
            )

        else:

            user_prompt = f"""
DOCUMENT TYPE:
{document_type}

LANGUAGE:
{language}

TONE:
{tone}

SENDER:
{sender}

RECIPIENT:
{recipient}

SUBJECT:
{subject}

RAW USER INFORMATION:
{st.session_state.normal_information}
"""

            try:

                with st.spinner(
                    "DraftForge is preparing your document..."
                ):

                    draft = generate_ai(
                        NORMAL_SYSTEM_PROMPT,
                        user_prompt,
                        provider,
                    )

                st.session_state.generated_draft = draft

                save_history(
                    document_type,
                    "",
                    subject,
                    draft,
                )

                st.success(
                    "Draft generated successfully."
                )

            except Exception as e:

                st.error(
                    f"Generation failed: {e}"
                )


# ============================================================
# INQUIRY
# ============================================================

else:

    st.subheader("Departmental Inquiry")

    inquiry_type = st.selectbox(
        "Inquiry Type",
        [
            INQUIRY_ED,
            INQUIRY_FFI,
        ],
        key="selected_inquiry_type",
    )

    # --------------------------------------------------------
    # FFI
    # --------------------------------------------------------

    if inquiry_type == INQUIRY_FFI:

        st.info(
            "🚧 FFI Inquiry module is currently under construction / "
            "under process."
        )

    # --------------------------------------------------------
    # E&D
    # --------------------------------------------------------

    else:

        st.markdown(
            "### Select E&D Inquiry Indexes"
        )

        st.caption(
            "You can select the same index multiple times. "
            "Each selection creates an independent input section."
        )

        st.selectbox(
            "E&D Index",
            [
                "Select an index..."
            ] + ED_INDEXES,
            key="ed_index_selector",
            on_change=select_ed_index,
        )

        # ----------------------------------------------------
        # SELECTED INDEX SECTIONS
        # ----------------------------------------------------

        if st.session_state.ed_index_instances:

            st.markdown("### Selected Indexes")

            for number, instance in enumerate(
                st.session_state.ed_index_instances,
                start=1,
            ):

                instance_id = instance["id"]

                display_name = get_display_index_name(
                    st.session_state.ed_index_instances,
                    instance,
                )

                st.markdown(
                    f"**{number}. {display_name}**"
                )

                if instance["index_name"] == "Documents Recorded":

                    st.markdown(
                        "**Documents Recorded**"
                    )

                    for item in DOCUMENT_RECORD_OPTIONS:

                        st.checkbox(
                            item,
                            key=f"doc_record_{item}",
                        )

                elif instance["index_name"] == "Inquiry Committee":

                    st.markdown(
                        "**Inquiry Committee**"
                    )

                    for role in COMMITTEE_ROLES:

                        st.markdown(
                            f"#### {role}"
                        )

                        st.text_input(
                            "ERP#",
                            key=f"committee_erp_{role}",
                        )

                        st.text_input(
                            "Name",
                            key=f"committee_name_{role}",
                        )

                        st.text_input(
                            "Designation",
                            key=(
                                f"committee_designation_"
                                f"{role}"
                            ),
                        )

                else:

                    text_key = f"inquiry_text_{instance_id}"

                    st.text_area(
                        display_name,
                        key=text_key,
                        height=180,
                        placeholder=(
                            "Enter information naturally. "
                            "Grammar and spelling will be corrected "
                            "automatically in the final report."
                        ),
                    )

                    st.file_uploader(
                        "🎙️ Record / Upload Voice",
                        type=[
                            "wav",
                            "mp3",
                            "m4a",
                            "ogg",
                            "webm",
                        ],
                        key=f"voice_upload_{instance_id}",
                    )

                    st.button(
                        "➕ Add Voice to Text",
                        key=f"add_voice_{instance_id}",
                        on_click=add_inquiry_voice,
                        args=(instance_id,),
                    )

                    status = st.session_state.get(
                        f"voice_status_{instance_id}",
                        "",
                    )

                    if status:
                        st.info(status)

                st.button(
                    "🗑️ Remove",
                    key=f"remove_{instance_id}",
                    on_click=remove_ed_instance,
                    args=(instance_id,),
                )

                st.divider()

        else:

            st.info(
                "No E&D indexes selected yet."
            )

        # ----------------------------------------------------
        # RESET
        # ----------------------------------------------------

        col1, col2 = st.columns(2)

        with col1:

            if st.button(
                "🔄 Reset Inquiry",
                use_container_width=True,
            ):

                reset_inquiry()
                st.rerun()

        # ----------------------------------------------------
        # GENERATE E&D
        # ----------------------------------------------------

        if st.button(
            "✨ Generate E&D Inquiry Report",
            type="primary",
            use_container_width=True,
        ):

            if not st.session_state.ed_index_instances:

                st.warning(
                    "Please select at least one E&D index."
                )

            else:

                selected_manifest = (
                    build_selected_index_manifest()
                )

                inquiry_information = (
                    build_ed_information()
                )

                reference_number = (
                    get_first_index_value(
                        "Inquiry Reference No."
                    )
                )

                subject = get_first_index_value(
                    "Subject"
                )

                documents_recorded = (
                    build_documents_recorded()
                )

                committee = (
                    build_inquiry_committee()
                )

                user_prompt = f"""
LANGUAGE:
{language}

TONE:
{tone}

SELECTED INDEXES:

{selected_manifest}

USER-PROVIDED INFORMATION:

{inquiry_information}

DOCUMENTS RECORDED:

{documents_recorded}

INQUIRY COMMITTEE:

{committee}

IMPORTANT:
The selected index list above is authoritative.

Generate the report using ONLY those selected indexes,
in exactly that order.

Do not add any other heading.

Correct spelling and grammar intelligently while preserving
all facts and meanings.

Inquiry Reference Number:
{reference_number}

Subject:
{subject}
"""

                try:

                    with st.spinner(
                        "Preparing your departmental inquiry report..."
                    ):

                        draft = generate_ai(
                            ED_SYSTEM_PROMPT,
                            user_prompt,
                            provider,
                        )

                    draft = add_inquiry_header(
                        draft,
                        reference_number,
                    )

                    st.session_state.generated_draft = draft

                    save_history(
                        DOC_INQUIRY,
                        INQUIRY_ED,
                        subject,
                        draft,
                    )

                    st.success(
                        "Departmental Inquiry Report generated successfully."
                    )

                except Exception as e:

                    st.error(
                        f"Generation failed: {e}"
                    )


# ============================================================
# GENERATED DOCUMENT
# ============================================================

if st.session_state.generated_draft:

    st.divider()

    st.subheader("📄 Generated Document")

    st.text_area(
        "Draft",
        st.session_state.generated_draft,
        height=600,
    )

    st.markdown("### Download")

    download_col1, download_col2 = st.columns(2)

    with download_col1:

        try:

            pdf_data = export_pdf(
                st.session_state.generated_draft
            )

            st.download_button(
                "📕 Download PDF",
                data=pdf_data,
                file_name="DraftForge_Inquiry_Report.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

        except Exception as e:

            st.error(
                f"PDF export failed: {e}"
            )

    with download_col2:

        try:

            docx_data = export_docx(
                st.session_state.generated_draft
            )

            st.download_button(
                "📘 Download DOCX",
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
                f"DOCX export failed: {e}"
            )

    download_col3, download_col4 = st.columns(2)

    with download_col3:

        txt_data = export_txt(
            st.session_state.generated_draft
        )

        st.download_button(
            "📄 Download TXT",
            data=txt_data,
            file_name="DraftForge_Document.txt",
            mime="text/plain",
            use_container_width=True,
        )

    with download_col4:

        try:

            png_data = export_png(
                st.session_state.generated_draft
            )

            st.download_button(
                "🖼️ Download PNG",
                data=png_data,
                file_name="DraftForge_Document.png",
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

st.divider()

st.caption(
    "DraftForge — AI-assisted professional document composition"
)
