import io
import os
import sqlite3
from datetime import datetime

import requests
import streamlit as st

from docx import Document
from docx.shared import Inches

from fpdf import FPDF

from PIL import Image, ImageDraw, ImageFont


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="DraftForge - AI Document Composer",
    page_icon="📝",
    layout="wide",
)


# ============================================================
# CONFIGURATION
# ============================================================

GROQ_MODEL = "openai/gpt-oss-120b"
WHISPER_MODEL = "whisper-large-v3"

GEMINI_URL = (
    "https://generativelanguage.googleapis.com/"
    "v1beta/models/gemini-2.0-flash:generateContent"
)

DOCUMENT_TYPES = [
    "Email",
    "Letter",
    "Inquiry",
    "Custom",
]

INQUIRY_TYPES = [
    "FFI Inquiry",
    "E&D Inquiry",
]

OUTPUT_LANGUAGES = [
    "English",
    "Urdu",
    "Pashto",
    "Arabic",
    "Same as input",
]

TONES = [
    "Formal",
    "Professional",
    "Friendly/Casual",
    "Persuasive",
    "Apologetic",
    "Assertive",
    "Neutral",
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


# ============================================================
# DATABASE
# ============================================================

DB_PATH = os.path.join(
    os.path.dirname(__file__),
    "drafts.db",
)


# ============================================================
# SESSION STATE INITIALIZATION
# ============================================================

DEFAULT_STATE = {
    "generated_draft": "",
    "normal_information": "",
    "ed_selected_indexes": [],
    "ed_index_values": {},
    "ed_documents": [],
    "ed_other_document": "",
    "ed_committee": {
        "Convener of Inquiry": "",
        "Member 1": "",
        "Member 2": "",
        "Departmental Representative": "",
    },
}


for key, value in DEFAULT_STATE.items():

    if key not in st.session_state:

        # Copy mutable objects so they are not shared.
        if isinstance(value, list):
            st.session_state[key] = list(value)

        elif isinstance(value, dict):
            st.session_state[key] = dict(value)

        else:
            st.session_state[key] = value


# ============================================================
# DATABASE FUNCTIONS
# ============================================================

def init_db():

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute(
        "SELECT name FROM sqlite_master "
        "WHERE type='table' AND name='drafts'"
    )

    table_exists = cursor.fetchone()

    if table_exists is None:

        cursor.execute(
            """
            CREATE TABLE drafts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                document_type TEXT,
                inquiry_type TEXT,
                tone TEXT,
                recipient TEXT,
                sender TEXT,
                subject TEXT,
                key_points TEXT,
                draft TEXT,
                created_at TEXT
            )
            """
        )

    else:

        cursor.execute(
            "PRAGMA table_info(drafts)"
        )

        existing_columns = {
            row[1]
            for row in cursor.fetchall()
        }

        required_columns = {
            "document_type": "TEXT",
            "inquiry_type": "TEXT",
            "tone": "TEXT",
            "recipient": "TEXT",
            "sender": "TEXT",
            "subject": "TEXT",
            "key_points": "TEXT",
            "draft": "TEXT",
            "created_at": "TEXT",
        }

        for column, data_type in required_columns.items():

            if column not in existing_columns:

                cursor.execute(
                    "ALTER TABLE drafts ADD COLUMN "
                    + column
                    + " "
                    + data_type
                )

    conn.commit()
    conn.close()


def save_draft(
    document_type,
    inquiry_type,
    tone,
    recipient,
    sender,
    subject,
    key_points,
    draft,
):

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute(
        """
        INSERT INTO drafts (
            document_type,
            inquiry_type,
            tone,
            recipient,
            sender,
            subject,
            key_points,
            draft,
            created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            document_type,
            inquiry_type,
            tone,
            recipient,
            sender,
            subject,
            key_points,
            draft,
            datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            ),
        ),
    )

    conn.commit()
    conn.close()


def get_history():

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute(
        """
        SELECT
            id,
            document_type,
            inquiry_type,
            tone,
            recipient,
            sender,
            subject,
            key_points,
            draft,
            created_at
        FROM drafts
        ORDER BY id DESC
        """
    )

    rows = cursor.fetchall()

    conn.close()

    return rows


init_db()


# ============================================================
# API KEY FUNCTIONS
# ============================================================

def get_secret(name):

    try:

        return st.secrets.get(
            name,
            "",
        )

    except Exception:

        return ""


def get_api_key(provider):

    if provider == "Groq":

        return get_secret(
            "GROQ_API_KEY"
        )

    return get_secret(
        "GEMINI_API_KEY"
    )


# ============================================================
# VOICE TRANSCRIPTION
# ============================================================

def transcribe_audio(
    api_key,
    audio_file,
):

    if not api_key:

        raise ValueError(
            "GROQ_API_KEY is not configured "
            "in Streamlit Secrets."
        )

    if audio_file is None:

        raise ValueError(
            "No audio recording was provided."
        )

    try:

        from groq import Groq

        client = Groq(
            api_key=api_key
        )

        audio_bytes = audio_file.getvalue()

        if not audio_bytes:

            raise ValueError(
                "The recording is empty."
            )

        if len(audio_bytes) > (
            25 * 1024 * 1024
        ):

            raise ValueError(
                "The recording is larger than "
                "25 MB. Please make a shorter recording."
            )

        filename = getattr(
            audio_file,
            "name",
            "recording.wav",
        )

        if not filename:

            filename = "recording.wav"

        result = (
            client.audio.transcriptions.create(
                file=(
                    filename,
                    audio_bytes,
                ),
                model=WHISPER_MODEL,
                response_format="json",
                temperature=0.0,
            )
        )

        text = getattr(
            result,
            "text",
            "",
        )

        if not text:

            raise RuntimeError(
                "No speech could be detected."
            )

        return text.strip()

    except Exception as e:

        raise RuntimeError(
            "Speech-to-text error: "
            + str(e)
        ) from e


# ============================================================
# VOICE CALLBACK HELPERS
#
# IMPORTANT:
# These callbacks run before the next page render.
# Therefore it is safe to update the text widget's
# session-state value here.
# ============================================================

def append_voice_to_widget(
    audio_key,
    text_widget_key,
    state_storage_key,
):

    audio = st.session_state.get(
        audio_key
    )

    if audio is None:

        return

    try:

        groq_key = get_secret(
            "GROQ_API_KEY"
        )

        if not groq_key:

            st.session_state[
                "voice_error"
            ] = (
                "GROQ_API_KEY is not configured "
                "in Streamlit Secrets."
            )

            return

        transcript = transcribe_audio(
            groq_key,
            audio,
        )

        existing = st.session_state.get(
            text_widget_key,
            "",
        )

        if existing.strip():

            combined = (
                existing.rstrip()
                + "\n"
                + transcript
            )

        else:

            combined = transcript

        # This is executed by the callback
        # before the widget is instantiated
        # on the next Streamlit run.
        st.session_state[
            text_widget_key
        ] = combined

        # Keep application state synchronized.
        st.session_state[
            state_storage_key
        ] = combined

        st.session_state[
            "voice_success"
        ] = (
            "Voice information added "
            "to the text box."
        )

    except Exception as e:

        st.session_state[
            "voice_error"
        ] = str(e)


def add_normal_voice():

    append_voice_to_widget(
        audio_key="normal_voice_recorder",
        text_widget_key="normal_information_editor",
        state_storage_key="normal_information",
    )


def add_inquiry_voice(
    index_name
):

    append_voice_to_widget(
        audio_key=(
            "ed_voice_"
            + index_name
        ),
        text_widget_key=(
            "ed_text_"
            + index_name
        ),
        state_storage_key=(
            "ed_state_"
            + index_name
        ),
    )


# ============================================================
# AI PROMPT
# ============================================================

def build_prompt(
    document_type,
    inquiry_type,
    tone,
    recipient,
    sender,
    subject,
    information,
    output_language,
):

    if output_language == "Same as input":

        language_instruction = (
            "Write the final document in the same "
            "language used by the user's information. "
            "If multiple languages are present, use "
            "the dominant language unless context "
            "clearly requires another language."
        )

    else:

        language_instruction = (
            "Write the entire final document in "
            + output_language
            + ". Translate and professionally adapt "
              "the supplied information while preserving "
              "its intended meaning."
        )

    inquiry_instruction = ""

    if document_type == "Inquiry":

        if inquiry_type == "E&D Inquiry":

            inquiry_instruction = """
This is an official E&D (Efficiency &
Discipline / departmental disciplinary)
inquiry.

The user has supplied information through
specific inquiry indexes.

Treat the indexes as structured components
of the inquiry report.

The normal E&D structure may include:

1. Inquiry Reference No.
2. Subject
3. Brief of the Inquiry
4. Articles of Charge / Allegations
5. Statement of the Accused
6. Questions / Answers with the Accused
7. Statements of Witnesses / Officials
8. Documentary Evidence / Record Examined
9. Defence / Written Explanation
10. Findings
11. Findings on Each Charge
12. Discussion / Analysis
13. Conclusion
14. Recommendations
15. Documents Recorded
16. Inquiry Committee

Preserve the supplied index identity.

Do not unnecessarily merge separate indexes.

Use formal departmental inquiry language.

Findings must be based only on supplied
facts, statements and evidence.

Do not invent facts, evidence, witnesses,
dates, names, charges, admissions or findings.

Where information is insufficient, do not
guess.

The final document should read as a coherent
professional E&D inquiry report while
preserving the substance of the supplied
information.

The Documents Recorded section should present
the selected documents clearly.

The Inquiry Committee section should preserve
the Convener, Member 1, Member 2 and
Departmental Representative details.
"""

        else:

            inquiry_instruction = """
FFI Inquiry functionality is currently under
construction.

Do not fabricate an FFI inquiry format.
"""

    prompt = f"""
You are an expert professional writer,
official correspondence specialist and
departmental inquiry-document drafting
assistant.

Create a polished professional document.

DOCUMENT TYPE:
{document_type}

INQUIRY TYPE:
{inquiry_type}

TONE:
{tone}

RECIPIENT:
{recipient}

SENDER:
{sender}

SUBJECT:
{subject}

USER INFORMATION:
{information}

DESIRED OUTPUT LANGUAGE:
{output_language}

LANGUAGE INSTRUCTION:
{language_instruction}

{inquiry_instruction}

GENERAL REQUIREMENTS:

1. Preserve the user's intended meaning.

2. Correct grammar, spelling and obvious
speech-recognition errors.

3. Do not invent important facts.

4. Do not fabricate names, dates, evidence,
statements, allegations or findings.

5. Organize information logically.

6. Use professional official language.

7. Make the document ready for practical use.

8. Keep useful placeholders where appropriate.

9. For Email, use an appropriate email structure.

10. For Letter, use an appropriate official
letter structure.

11. For E&D Inquiry, preserve the supplied
inquiry indexes and their meaning.

12. For E&D Inquiry, use appropriate headings
and subheadings.

13. Do not unnecessarily remove information
provided by the user.

14. Do not add explanations before or after
the document.

15. Return only the finished document.

16. Do not use markdown code blocks.

17. Never create unsupported facts.

18. If the user's information is incomplete,
do not guess.

19. Do not turn missing information into
fictional information.

20. Maintain a formal, objective and
professionally neutral inquiry style.
"""

    return prompt.strip()


# ============================================================
# GROQ
# ============================================================

def call_groq(
    api_key,
    prompt,
):

    if not api_key:

        raise ValueError(
            "GROQ_API_KEY is not configured "
            "in Streamlit Secrets."
        )

    try:

        from groq import Groq

        client = Groq(
            api_key=api_key
        )

        response = (
            client.chat.completions.create(
                model=GROQ_MODEL,
                messages=[
                    {
                        "role": "system",
                        "content": (
                            "You are an expert "
                            "professional writer "
                            "and official inquiry "
                            "document drafting "
                            "specialist."
                        ),
                    },
                    {
                        "role": "user",
                        "content": prompt,
                    },
                ],
                temperature=0.4,
            )
        )

        return (
            response
            .choices[0]
            .message
            .content
            .strip()
        )

    except Exception as e:

        raise RuntimeError(
            "Groq API error: "
            + str(e)
        ) from e


# ============================================================
# GEMINI
# ============================================================

def call_gemini(
    api_key,
    prompt,
):

    if not api_key:

        raise ValueError(
            "GEMINI_API_KEY is not configured "
            "in Streamlit Secrets."
        )

    headers = {
        "Content-Type": "application/json"
    }

    payload = {
        "contents": [
            {
                "parts": [
                    {
                        "text": prompt
                    }
                ]
            }
        ]
    }

    url = (
        GEMINI_URL
        + "?key="
        + api_key
    )

    try:

        response = requests.post(
            url,
            headers=headers,
            json=payload,
            timeout=60,
        )

        if response.status_code != 200:

            raise RuntimeError(
                "Gemini API error: "
                + response.text
            )

        data = response.json()

        candidates = data.get(
            "candidates",
            [],
        )

        if not candidates:

            raise RuntimeError(
                "Gemini returned no response."
            )

        parts = (
            candidates[0]
            .get("content", {})
            .get("parts", [])
        )

        if not parts:

            raise RuntimeError(
                "Gemini returned an empty response."
            )

        return (
            parts[0]
            .get("text", "")
            .strip()
        )

    except requests.RequestException as e:

        raise RuntimeError(
            "Gemini connection error: "
            + str(e)
        ) from e


# ============================================================
# PDF
# ============================================================

def clean_pdf_text(text):

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
        "\u2011": "-",
    }

    for old, new in replacements.items():

        text = text.replace(
            old,
            new,
        )

    return (
        text.encode(
            "latin-1",
            "replace",
        )
        .decode("latin-1")
    )


def wrap_pdf_line(
    pdf,
    text,
    max_width,
):

    if not text:

        return [""]

    words = text.split(" ")

    lines = []
    current = ""

    for word in words:

        test_line = (
            word
            if not current
            else current + " " + word
        )

        if (
            pdf.get_string_width(
                test_line
            )
            <= max_width
        ):

            current = test_line

        else:

            if current:

                lines.append(
                    current
                )

            current = word

    if current:

        lines.append(
            current
        )

    return lines


def export_pdf(text):

    pdf = FPDF()

    pdf.set_auto_page_break(
        auto=True,
        margin=15,
    )

    pdf.add_page()

    pdf.set_font(
        "Arial",
        size=11,
    )

    usable_width = (
        pdf.w
        - pdf.l_margin
        - pdf.r_margin
    )

    cleaned_text = clean_pdf_text(
        text
    )

    for paragraph in cleaned_text.split(
        "\n"
    ):

        if paragraph.strip() == "":

            pdf.ln(5)

            continue

        lines = wrap_pdf_line(
            pdf,
            paragraph,
            usable_width,
        )

        for line in lines:

            pdf.multi_cell(
                usable_width,
                7,
                line,
                new_x="LMARGIN",
                new_y="NEXT",
            )

    return bytes(
        pdf.output()
    )


# ============================================================
# DOCX
# ============================================================

def export_docx(text):

    document = Document()

    section = document.sections[0]

    section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.75)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

    for paragraph in text.split("\n"):

        document.add_paragraph(
            paragraph
        )

    buffer = io.BytesIO()

    document.save(
        buffer
    )

    buffer.seek(0)

    return buffer.getvalue()


# ============================================================
# TXT
# ============================================================

def export_txt(text):

    return text.encode(
        "utf-8"
    )


# ============================================================
# PNG
# ============================================================

def export_png(text):

    try:

        font = ImageFont.truetype(
            "DejaVuSans.ttf",
            24,
        )

    except Exception:

        font = ImageFont.load_default()

    margin = 50
    max_width = 1100
    line_height = 48

    dummy = Image.new(
        "RGB",
        (
            max_width,
            100,
        ),
        "white",
    )

    draw = ImageDraw.Draw(
        dummy
    )

    lines = []

    for paragraph in text.split("\n"):

        if paragraph == "":

            lines.append("")

            continue

        words = paragraph.split(" ")

        current = ""

        for word in words:

            test_line = (
                word
                if not current
                else current + " " + word
            )

            bbox = draw.textbbox(
                (0, 0),
                test_line,
                font=font,
            )

            width = (
                bbox[2]
                - bbox[0]
            )

            if width <= (
                max_width
                - margin * 2
            ):

                current = test_line

            else:

                if current:

                    lines.append(
                        current
                    )

                current = word

        if current:

            lines.append(
                current
            )

        lines.append("")

    height = max(
        200,
        margin * 2
        + len(lines)
        * line_height,
    )

    image = Image.new(
        "RGB",
        (
            max_width,
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

    buffer = io.BytesIO()

    image.save(
        buffer,
        format="PNG",
    )

    buffer.seek(0)

    return buffer.getvalue()


# ============================================================
# E&D CALLBACKS
# ============================================================

def select_ed_index():

    selected = st.session_state.get(
        "ed_index_selector",
        "",
    )

    if (
        selected
        and selected != "-- Select an index --"
    ):

        if selected not in st.session_state.ed_selected_indexes:

            st.session_state.ed_selected_indexes.append(
                selected
            )


def remove_ed_index(index_name):

    if index_name in st.session_state.ed_selected_indexes:

        st.session_state.ed_selected_indexes.remove(
            index_name
        )


def reset_inquiry():

    st.session_state.ed_selected_indexes = []
    st.session_state.ed_index_values = {}
    st.session_state.ed_documents = []
    st.session_state.ed_other_document = ""

    st.session_state.ed_committee = {
        "Convener of Inquiry": "",
        "Member 1": "",
        "Member 2": "",
        "Departmental Representative": "",
    }


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.title(
        "📝 DraftForge"
    )

    st.caption(
        "AI Document Composer"
    )

    st.divider()

    st.subheader(
        "⚙️ AI Provider"
    )

    provider = st.selectbox(
        "Choose AI provider",
        [
            "Groq",
            "Google Gemini",
        ],
    )

    api_key = get_api_key(
        provider
    )

    if api_key:

        st.success(
            "🔐 API key loaded securely."
        )

    else:

        if provider == "Groq":

            st.warning(
                "GROQ_API_KEY not found "
                "in Streamlit Secrets."
            )

        else:

            st.warning(
                "GEMINI_API_KEY not found "
                "in Streamlit Secrets."
            )

    st.divider()

    st.subheader(
        "👤 Profile"
    )

    profile_name = st.text_input(
        "Your name",
        value="Raees Khan",
    )

    st.divider()

    st.subheader(
        "📚 Draft History"
    )

    history = get_history()

    if history:

        st.caption(
            str(len(history))
            + " saved draft(s)"
        )

        for row in history[:10]:

            draft_id = row[0]
            document_type_history = row[1]
            inquiry_type_history = row[2]
            subject_history = row[6]
            created_at = row[9]

            title = (
                subject_history
                if subject_history
                else document_type_history
            )

            with st.expander(
                f"{title} - {created_at}"
            ):

                st.write(
                    "**Type:** "
                    + str(
                        document_type_history
                    )
                )

                if inquiry_type_history:

                    st.write(
                        "**Inquiry:** "
                        + str(
                            inquiry_type_history
                        )
                    )

                if subject_history:

                    st.write(
                        "**Subject:** "
                        + str(
                            subject_history
                        )
                    )

                if st.button(
                    "Load Draft",
                    key=f"load_{draft_id}",
                    use_container_width=True,
                ):

                    st.session_state.generated_draft = (
                        row[8] or ""
                    )

                    st.rerun()


# ============================================================
# MAIN HEADER
# ============================================================

st.title(
    "📝 DraftForge — AI Document Composer"
)

st.write(
    "Create professional emails, letters and "
    "inquiries using natural-language text or "
    "multilingual voice input."
)

st.divider()


# ============================================================
# DOCUMENT TYPE
# ============================================================

st.subheader(
    "📄 Document Settings"
)

document_type = st.selectbox(
    "Document Type",
    DOCUMENT_TYPES,
)


# ============================================================
# INQUIRY TYPE
# ============================================================

inquiry_type = ""

if document_type == "Inquiry":

    inquiry_type = st.selectbox(
        "🔎 Inquiry Type",
        INQUIRY_TYPES,
    )

    if inquiry_type == "FFI Inquiry":

        st.warning(
            "🚧 FFI Inquiry Module — Under Construction"
        )

        st.info(
            "The FFI inquiry format and indexes "
            "are currently under process. "
            "Please use E&D Inquiry for the "
            "currently available inquiry workflow."
        )

    else:

        st.success(
            "✅ E&D Inquiry Module"
        )

        st.caption(
            "Select the required inquiry indexes "
            "from the dropdown below. Each selected "
            "index accepts natural-language typing "
            "and voice input."
        )


# ============================================================
# GENERAL SETTINGS
# ============================================================

tone = st.selectbox(
    "🎯 Tone",
    TONES,
)


# ============================================================
# RECIPIENT / SENDER / SUBJECT
# ============================================================

# For E&D inquiry, Subject is an index and
# therefore the normal Subject field is not shown.

if (
    document_type == "Inquiry"
    and inquiry_type == "E&D Inquiry"
):

    recipient = ""
    sender = profile_name
    subject = ""

else:

    col1, col2 = st.columns(2)

    with col1:

        recipient = st.text_input(
            "👤 Recipient",
            placeholder=(
                "e.g. Manager / Director / Officer"
            ),
        )

    with col2:

        sender = st.text_input(
            "✍️ Sender",
            value=profile_name,
        )

    subject = st.text_input(
        "📌 Subject",
        placeholder="Enter subject",
    )


# ============================================================
# OUTPUT LANGUAGE
# ============================================================

output_language = st.selectbox(
    "🌐 Desired Output Language",
    OUTPUT_LANGUAGES,
)


# ============================================================
# STATUS MESSAGES
# ============================================================

if "voice_error" in st.session_state:

    st.error(
        st.session_state.voice_error
    )

    del st.session_state.voice_error


if "voice_success" in st.session_state:

    st.success(
        st.session_state.voice_success
    )

    del st.session_state.voice_success


# ============================================================
# NORMAL INFORMATION BOX
# ============================================================

def normal_information_box():

    st.markdown(
        "### 📝 Information"
    )

    st.caption(
        "Type naturally or use the microphone. "
        "Voice transcription is inserted directly "
        "into the same text box."
    )

    text_widget_key = (
        "normal_information_editor"
    )

    if text_widget_key not in st.session_state:

        st.session_state[
            text_widget_key
        ] = st.session_state.get(
            "normal_information",
            "",
        )

    st.text_area(
        "Type or edit your information",
        height=250,
        key=text_widget_key,
        placeholder=(
            "Type your request here..."
        ),
    )

    # Synchronize application state.
    # This does NOT modify the widget key.
    st.session_state.normal_information = (
        st.session_state.get(
            text_widget_key,
            "",
        )
    )

    st.markdown(
        "#### 🎙️ Voice Input"
    )

    st.audio_input(
        "Record your voice",
        sample_rate=16000,
        key="normal_voice_recorder",
    )

    st.button(
        "🎙️ Add Voice to Text Box",
        key="normal_transcribe_button",
        use_container_width=True,
        on_click=add_normal_voice,
    )

    return st.session_state.get(
        text_widget_key,
        "",
    )


# ============================================================
# E&D DOCUMENTS RECORDED
# ============================================================

def render_documents_recorded():

    st.markdown(
        "### 📂 Documents Recorded"
    )

    st.caption(
        "Select the documents that were available "
        "or examined during the inquiry."
    )

    documents = [
        "CNICF",
        "BC",
        "Marriage Certificate",
        "CNICs",
        "Domicile",
        "Affidavit",
        "Complaint / Application",
        "Written Explanation",
        "Official Record",
        "Other",
    ]

    selected_documents = []

    columns = st.columns(2)

    for number, document in enumerate(
        documents
    ):

        with columns[number % 2]:

            checked = st.checkbox(
                document,
                key=(
                    "ed_document_"
                    + str(number)
                ),
            )

            if checked:

                selected_documents.append(
                    document
                )

    if "Other" in selected_documents:

        st.text_input(
            "Specify other document",
            key="ed_other_document",
            placeholder="Enter document name",
        )

        if st.session_state.ed_other_document.strip():

            selected_documents.append(
                "Other: "
                + st.session_state.ed_other_document.strip()
            )

    st.session_state.ed_documents = (
        selected_documents
    )

    return selected_documents


# ============================================================
# E&D INQUIRY COMMITTEE
# ============================================================

def render_inquiry_committee():

    st.markdown(
        "### 👥 Inquiry Committee"
    )

    st.caption(
        "Enter ERP#, Name and Designation for "
        "each committee member."
    )

    committee_fields = [
        "Convener of Inquiry",
        "Member 1",
        "Member 2",
        "Departmental Representative",
    ]

    for field in committee_fields:

        st.markdown(
            f"**{field}**"
        )

        value = st.text_input(
            f"{field} details",
            key=(
                "ed_committee_"
                + field
                .lower()
                .replace(" ", "_")
            ),
            placeholder=(
                "ERP# | Name | Designation"
            ),
        )

        st.session_state.ed_committee[
            field
        ] = value


# ============================================================
# E&D INQUIRY INPUT
# ============================================================

def inquiry_information_boxes():

    st.markdown(
        "### 📑 E&D Inquiry Indexes"
    )

    st.caption(
        "Select an index from the dropdown. "
        "Once selected, it will no longer appear "
        "in the dropdown."
    )

    selected_indexes = (
        st.session_state.ed_selected_indexes
    )

    available_indexes = [
        index
        for index in ED_INDEXES
        if index not in selected_indexes
    ]

    selector_options = [
        "-- Select an index --"
    ] + available_indexes

    st.selectbox(
        "➕ Select Inquiry Index",
        selector_options,
        key="ed_index_selector",
        on_change=select_ed_index,
    )

    st.divider()

    # --------------------------------------------------------
    # RENDER SELECTED INDEXES
    # --------------------------------------------------------

    for position, index_name in enumerate(
        list(selected_indexes)
    ):

        safe_name = (
            index_name
            .lower()
            .replace(" ", "_")
            .replace("/", "_")
            .replace("#", "no")
            .replace("&", "and")
            .replace("(", "")
            .replace(")", "")
            .replace("-", "_")
        )

        text_widget_key = (
            "ed_text_"
            + index_name
        )

        state_storage_key = (
            "ed_state_"
            + index_name
        )

        voice_key = (
            "ed_voice_"
            + index_name
        )

        transcribe_key = (
            "ed_transcribe_"
            + safe_name
        )

        if state_storage_key not in st.session_state:

            st.session_state[
                state_storage_key
            ] = st.session_state.ed_index_values.get(
                index_name,
                "",
            )

        if text_widget_key not in st.session_state:

            st.session_state[
                text_widget_key
            ] = st.session_state.get(
                state_storage_key,
                "",
            )

        st.markdown(
            "#### 📌 "
            + index_name
        )

        # ----------------------------------------------------
        # SPECIAL INDEX: DOCUMENTS RECORDED
        # ----------------------------------------------------

        if index_name == "Documents Recorded":

            render_documents_recorded()

        # ----------------------------------------------------
        # SPECIAL INDEX: INQUIRY COMMITTEE
        # ----------------------------------------------------

        elif index_name == "Inquiry Committee":

            render_inquiry_committee()

        # ----------------------------------------------------
        # NORMAL TEXT / VOICE INDEX
        # ----------------------------------------------------

        else:

            st.text_area(
                "Type or edit information",
                height=180,
                key=text_widget_key,
                placeholder=(
                    "Type information naturally "
                    "for this index..."
                ),
            )

            # IMPORTANT:
            # We only copy FROM the widget into
            # separate application state.
            # We NEVER write back to the widget
            # key during this run.
            current_value = st.session_state.get(
                text_widget_key,
                "",
            )

            st.session_state[
                state_storage_key
            ] = current_value

            st.session_state.ed_index_values[
                index_name
            ] = current_value

            st.audio_input(
                "🎙️ Record voice for this index",
                sample_rate=16000,
                key=voice_key,
            )

            st.button(
                "🎙️ Add Voice to This Index",
                key=transcribe_key,
                use_container_width=True,
                on_click=add_inquiry_voice,
                args=(index_name,),
            )

        st.button(
            "🗑️ Remove this index",
            key=(
                "remove_ed_"
                + safe_name
            ),
            on_click=remove_ed_index,
            args=(index_name,),
        )

        st.divider()

    if selected_indexes:

        st.caption(
            "Selected indexes: "
            + str(len(selected_indexes))
            + " / "
            + str(len(ED_INDEXES))
        )

        if st.button(
            "↻ Reset E&D Inquiry Indexes",
            use_container_width=True,
        ):

            reset_inquiry()

            st.rerun()

    # --------------------------------------------------------
    # BUILD INFORMATION FOR AI
    # --------------------------------------------------------

    information_parts = []

    for index_name in st.session_state.ed_selected_indexes:

        if index_name == "Documents Recorded":

            documents = (
                st.session_state.ed_documents
            )

            if documents:

                information_parts.append(
                    "INDEX: Documents Recorded\n"
                    + "\n".join(
                        "- " + item
                        for item in documents
                    )
                )

        elif index_name == "Inquiry Committee":

            committee_parts = []

            for role, details in (
                st.session_state.ed_committee.items()
            ):

                if details.strip():

                    committee_parts.append(
                        role
                        + ": "
                        + details
                    )

            if committee_parts:

                information_parts.append(
                    "INDEX: Inquiry Committee\n"
                    + "\n".join(
                        committee_parts
                    )
                )

        else:

            text = st.session_state.ed_index_values.get(
                index_name,
                "",
            )

            if text.strip():

                information_parts.append(
                    "INDEX: "
                    + index_name
                    + "\n"
                    + text
                )

    return "\n\n".join(
        information_parts
    )


# ============================================================
# DISPLAY INPUT AREA
# ============================================================

information = ""

if (
    document_type == "Inquiry"
    and inquiry_type == "E&D Inquiry"
):

    information = (
        inquiry_information_boxes()
    )

elif (
    document_type == "Inquiry"
    and inquiry_type == "FFI Inquiry"
):

    st.info(
        "FFI input fields will be added when "
        "the FFI inquiry format is finalized."
    )

else:

    information = (
        normal_information_box()
    )


# ============================================================
# GENERATE
# ============================================================

st.divider()

generate_allowed = True

if (
    document_type == "Inquiry"
    and inquiry_type == "FFI Inquiry"
):

    generate_allowed = False


if generate_allowed:

    if st.button(
        "✨ Generate Professional Document",
        type="primary",
        use_container_width=True,
    ):

        if not information.strip():

            st.warning(
                "Please provide information "
                "by typing or speaking."
            )

        elif not api_key:

            st.error(
                "Please configure the selected "
                "AI API key in Streamlit Secrets."
            )

        else:

            # For E&D the subject is obtained
            # from the inquiry index.
            inquiry_subject = ""

            if (
                document_type == "Inquiry"
                and inquiry_type == "E&D Inquiry"
            ):

                inquiry_subject = (
                    st.session_state.ed_index_values.get(
                        "Subject",
                        "",
                    )
                )

            prompt = build_prompt(
                document_type=document_type,
                inquiry_type=inquiry_type,
                tone=tone,
                recipient=recipient,
                sender=sender,
                subject=(
                    inquiry_subject
                    if inquiry_subject
                    else subject
                ),
                information=information,
                output_language=output_language,
            )

            with st.spinner(
                "✨ Preparing your professional document..."
            ):

                try:

                    if provider == "Groq":

                        draft = call_groq(
                            api_key,
                            prompt,
                        )

                    else:

                        draft = call_gemini(
                            api_key,
                            prompt,
                        )

                    st.session_state.generated_draft = (
                        draft
                    )

                    save_draft(
                        document_type,
                        inquiry_type,
                        tone,
                        recipient,
                        sender,
                        (
                            inquiry_subject
                            if inquiry_subject
                            else subject
                        ),
                        information,
                        draft,
                    )

                    st.success(
                        "✅ Document generated and "
                        "saved successfully!"
                    )

                except Exception as e:

                    st.error(
                        str(e)
                    )


# ============================================================
# GENERATED DOCUMENT
# ============================================================

if st.session_state.generated_draft:

    st.divider()

    st.subheader(
        "✏️ Generated Document"
    )

    edited_draft = st.text_area(
        "Review and edit your document",
        value=st.session_state.generated_draft,
        height=600,
        key="generated_document_editor",
    )

    # As with the other widget, only copy
    # the widget value after rendering.
    st.session_state.generated_draft = (
        edited_draft
    )

    st.divider()

    st.subheader(
        "📤 Export"
    )

    col1, col2 = st.columns(2)

    with col1:

        st.download_button(
            "📕 Download PDF",
            data=export_pdf(
                edited_draft
            ),
            file_name=(
                "draftforge_document.pdf"
            ),
            mime="application/pdf",
            use_container_width=True,
        )

        st.download_button(
            "📘 Download DOCX",
            data=export_docx(
                edited_draft
            ),
            file_name=(
                "draftforge_document.docx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            use_container_width=True,
        )

    with col2:

        st.download_button(
            "📄 Download TXT",
            data=export_txt(
                edited_draft
            ),
            file_name=(
                "draftforge_document.txt"
            ),
            mime="text/plain",
            use_container_width=True,
        )

        st.download_button(
            "🖼️ Download PNG",
            data=export_png(
                edited_draft
            ),
            file_name=(
                "draftforge_document.png"
            ),
            mime="image/png",
            use_container_width=True,
        )


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "DraftForge — AI-powered professional "
    "document creation with multilingual "
    "text and voice input."
)
