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
# APPLICATION CONFIGURATION
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

DB_PATH = os.path.join(
    os.path.dirname(__file__),
    "drafts.db",
)


# ============================================================
# SESSION STATE
# ============================================================

if "generated_draft" not in st.session_state:
    st.session_state.generated_draft = ""

if "inquiry_indexes" not in st.session_state:
    st.session_state.inquiry_indexes = [
        {
            "title": "Index 1",
            "text": "",
        }
    ]

if "voice_counter" not in st.session_state:
    st.session_state.voice_counter = 0


# ============================================================
# DATABASE
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

        existing_columns = set()

        for row in cursor.fetchall():
            existing_columns.add(row[1])

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

                sql = (
                    "ALTER TABLE drafts ADD COLUMN "
                    + column
                    + " "
                    + data_type
                )

                cursor.execute(sql)

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

        audio_bytes = (
            audio_file.getvalue()
        )

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
            "language used by the user. If multiple "
            "languages are used, use the dominant "
            "language unless the context clearly "
            "requires otherwise."
        )

    else:

        language_instruction = (
            "Write the entire final document in "
            + output_language
            + ". Translate and professionally adapt "
              "the supplied information while preserving "
              "its exact intended meaning."
        )

    inquiry_instruction = ""

    if document_type == "Inquiry":

        inquiry_instruction = f"""
This is an official {inquiry_type}.

The user's information is organized into multiple
inquiry indexes.

Treat every index as an important part of the inquiry.

Do not unnecessarily remove or merge indexes.

Maintain the logical sequence of the indexes.

Use appropriate official and professional inquiry
language.

Clearly distinguish allegations, statements,
facts, evidence, observations, findings and
recommendations when the supplied information
supports such distinctions.

Do not invent evidence, statements, dates, names,
findings or conclusions.

If the user has not supplied a fact, do not create it.

For E&D inquiries, maintain a professional
departmental disciplinary inquiry style.

For FFI inquiries, maintain a professional
fact-finding inquiry style.

Where appropriate, create headings and subheadings
based on the supplied indexes.
"""

    prompt = f"""
You are an expert professional writer,
official correspondence specialist and
inquiry-document drafting assistant.

Create a polished and professional {document_type}.

Document type:
{document_type}

Inquiry type:
{inquiry_type}

Tone:
{tone}

Recipient:
{recipient}

Sender:
{sender}

Subject:
{subject}

User information:
{information}

Desired output language:
{output_language}

Language instruction:
{language_instruction}

{inquiry_instruction}

General requirements:

1. Preserve the user's intended meaning.

2. Correct grammar, spelling and obvious
speech-recognition errors.

3. Do not invent important facts.

4. Do not fabricate names, dates, evidence,
statements, allegations or findings.

5. Organize the information logically.

6. Use professional official language.

7. Make the document ready for practical use.

8. Keep useful placeholders such as
[Date], [Reference], [Office Name] or
[Designation] when information is missing.

9. For Email, include an appropriate greeting
and closing.

10. For Letter, use appropriate official
letter structure.

11. For Inquiry, preserve the supplied
index structure and use appropriate
professional inquiry headings.

12. Return only the completed document.

13. Do not add explanations before or after
the document.

14. Do not use markdown code blocks.

15. Do not claim facts that are not contained
in the user's information.

16. If information is incomplete, use a suitable
placeholder instead of inventing information.
"""

    return prompt.strip()


# ============================================================
# GROQ GENERATION
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
                            "and official document "
                            "drafting specialist."
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
# GEMINI GENERATION
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
# PDF HELPERS
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

    words = text.split(
        " "
    )

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

    paragraphs = cleaned_text.split(
        "\n"
    )

    for paragraph in paragraphs:

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
# DOCX EXPORT
# ============================================================

def export_docx(text):

    document = Document()

    section = document.sections[0]

    section.top_margin = Inches(
        0.75
    )

    section.bottom_margin = Inches(
        0.75
    )

    section.left_margin = Inches(
        0.8
    )

    section.right_margin = Inches(
        0.8
    )

    for paragraph in text.split(
        "\n"
    ):

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
# TXT EXPORT
# ============================================================

def export_txt(text):

    return text.encode(
        "utf-8"
    )


# ============================================================
# PNG EXPORT
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
    line_spacing = 12
    max_width = 1100

    dummy_image = Image.new(
        "RGB",
        (
            max_width,
            100,
        ),
        "white",
    )

    draw = ImageDraw.Draw(
        dummy_image
    )

    lines = []

    for paragraph in text.split(
        "\n"
    ):

        if paragraph == "":

            lines.append("")

            continue

        words = paragraph.split(
            " "
        )

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

    line_height = (
        36
        + line_spacing
    )

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
            document_type = row[1]
            inquiry_type = row[2]
            subject = row[6]
            created_at = row[9]

            title = (
                subject
                if subject
                else document_type
            )

            with st.expander(
                f"{title} - {created_at}"
            ):

                st.write(
                    "**Type:** "
                    + str(document_type)
                )

                if inquiry_type:

                    st.write(
                        "**Inquiry:** "
                        + str(inquiry_type)
                    )

                if subject:

                    st.write(
                        "**Subject:** "
                        + str(subject)
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

    else:

        st.caption(
            "No drafts saved yet."
        )

    st.divider()

    st.caption(
        "DraftForge • AI-powered "
        "professional document creation"
    )


# ============================================================
# MAIN HEADER
# ============================================================

st.title(
    "📝 DraftForge — AI Document Composer"
)

st.write(
    "Create professional emails, letters, "
    "inquiries and custom documents using "
    "text or multilingual voice input."
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

        st.info(
            "FFI Inquiry: Add as many inquiry "
            "indexes as required for the "
            "fact-finding inquiry."
        )

    else:

        st.info(
            "E&D Inquiry: Add as many inquiry "
            "indexes as required for the "
            "departmental disciplinary inquiry."
        )


# ============================================================
# TONE
# ============================================================

tone = st.selectbox(
    "🎯 Tone",
    TONES,
)


# ============================================================
# BASIC DETAILS
# ============================================================

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
# NORMAL DOCUMENT INPUT
# ============================================================

def normal_information_box():

    st.markdown(
        "### 📝 Information"
    )

    st.caption(
        "Type your information or use the microphone. "
        "Voice transcription will be added to this "
        "same text box."
    )

    text_key = "normal_information"

    if text_key not in st.session_state:

        st.session_state[text_key] = ""

    typed_text = st.text_area(
        "Type or edit your information",
        value=st.session_state[text_key],
        height=250,
        key="normal_information_editor",
        placeholder=(
            "Type your request here..."
        ),
    )

    st.session_state[text_key] = typed_text

    st.markdown(
        "#### 🎙️ Add information by voice"
    )

    audio = st.audio_input(
        "Record your voice",
        sample_rate=16000,
        key="normal_voice_recorder",
    )

    if audio is not None:

        if st.button(
            "🎙️ Add Voice to Text Box",
            key="normal_transcribe",
            use_container_width=True,
        ):

            groq_key = get_secret(
                "GROQ_API_KEY"
            )

            if not groq_key:

                st.error(
                    "GROQ_API_KEY is not configured "
                    "in Streamlit Secrets."
                )

            else:

                with st.spinner(
                    "🎙️ Transcribing voice..."
                ):

                    try:

                        transcript = (
                            transcribe_audio(
                                groq_key,
                                audio,
                            )
                        )

                        current_text = (
                            st.session_state.get(
                                "normal_information",
                                "",
                            )
                        )

                        if current_text.strip():

                            combined_text = (
                                current_text
                                + "\n"
                                + transcript
                            )

                        else:

                            combined_text = (
                                transcript
                            )

                        st.session_state[
                            "normal_information"
                        ] = combined_text

                        st.success(
                            "Voice added to the text box."
                        )

                        st.rerun()

                    except Exception as e:

                        st.error(
                            str(e)
                        )

    return st.session_state.get(
        "normal_information",
        "",
    )


# ============================================================
# INQUIRY INPUT
# ============================================================

def inquiry_information_boxes():

    st.markdown(
        "### 📑 Inquiry Indexes"
    )

    st.caption(
        "Each index can contain information entered "
        "by typing, voice, or a combination of both."
    )

    remove_index = None

    for index_number in range(
        len(
            st.session_state.inquiry_indexes
        )
    ):

        index_data = (
            st.session_state.inquiry_indexes[
                index_number
            ]
        )

        st.markdown(
            f"#### 📌 Index {index_number + 1}"
        )

        title_key = (
            "inquiry_title_"
            + str(index_number)
        )

        text_key = (
            "inquiry_text_"
            + str(index_number)
        )

        if title_key not in st.session_state:

            st.session_state[
                title_key
            ] = index_data["title"]

        if text_key not in st.session_state:

            st.session_state[
                text_key
            ] = index_data["text"]

        index_title = st.text_input(
            "Index title",
            value=st.session_state[
                title_key
            ],
            key=title_key,
            placeholder=(
                "e.g. Allegation, Statement "
                "of accused, Witness statement, "
                "Documentary evidence"
            ),
        )

        index_text = st.text_area(
            "Type or edit information",
            value=st.session_state[
                text_key
            ],
            height=180,
            key=text_key,
            placeholder=(
                "Type information for this index..."
            ),
        )

        st.session_state.inquiry_indexes[
            index_number
        ]["title"] = index_title

        st.session_state.inquiry_indexes[
            index_number
        ]["text"] = index_text

        voice_key = (
            "inquiry_voice_"
            + str(index_number)
        )

        transcribe_key = (
            "inquiry_transcribe_"
            + str(index_number)
        )

        audio = st.audio_input(
            "🎙️ Record voice for this index",
            sample_rate=16000,
            key=voice_key,
        )

        if audio is not None:

            if st.button(
                "🎙️ Add Voice to This Index",
                key=transcribe_key,
                use_container_width=True,
            ):

                groq_key = get_secret(
                    "GROQ_API_KEY"
                )

                if not groq_key:

                    st.error(
                        "GROQ_API_KEY is not configured "
                        "in Streamlit Secrets."
                    )

                else:

                    with st.spinner(
                        "🎙️ Transcribing voice..."
                    ):

                        try:

                            transcript = (
                                transcribe_audio(
                                    groq_key,
                                    audio,
                                )
                            )

                            current_text = (
                                st.session_state.get(
                                    text_key,
                                    "",
                                )
                            )

                            if current_text.strip():

                                combined_text = (
                                    current_text
                                    + "\n"
                                    + transcript
                                )

                            else:

                                combined_text = (
                                    transcript
                                )

                            st.session_state[
                                text_key
                            ] = combined_text

                            st.session_state.inquiry_indexes[
                                index_number
                            ]["text"] = (
                                combined_text
                            )

                            st.success(
                                "Voice added to Index "
                                + str(index_number + 1)
                                + "."
                            )

                            st.rerun()

                        except Exception as e:

                            st.error(
                                str(e)
                            )

        if len(
            st.session_state.inquiry_indexes
        ) > 1:

            if st.button(
                "🗑️ Remove This Index",
                key=(
                    "remove_index_"
                    + str(index_number)
                ),
            ):

                remove_index = (
                    index_number
                )

        st.divider()

    if remove_index is not None:

        st.session_state.inquiry_indexes.pop(
            remove_index
        )

        st.rerun()

    if st.button(
        "➕ Add Another Index",
        use_container_width=True,
    ):

        new_number = (
            len(
                st.session_state.inquiry_indexes
            )
            + 1
        )

        st.session_state.inquiry_indexes.append(
            {
                "title": (
                    "Index "
                    + str(new_number)
                ),
                "text": "",
            }
        )

        st.rerun()

    information_parts = []

    for index_number, index_data in enumerate(
        st.session_state.inquiry_indexes
    ):

        title = index_data.get(
            "title",
            "Index "
            + str(index_number + 1),
        )

        text = index_data.get(
            "text",
            "",
        )

        if text.strip():

            information_parts.append(
                "INDEX "
                + str(index_number + 1)
                + ": "
                + title
                + "\n"
                + text
            )

    return "\n\n".join(
        information_parts
    )


# ============================================================
# SELECT INPUT FORM
# ============================================================

if document_type == "Inquiry":

    information = inquiry_information_boxes()

else:

    information = normal_information_box()


# ============================================================
# GENERATE DOCUMENT
# ============================================================

st.divider()

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

        prompt = build_prompt(
            document_type=document_type,
            inquiry_type=inquiry_type,
            tone=tone,
            recipient=recipient,
            sender=sender,
            subject=subject,
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
                    subject,
                    information,
                    draft,
                )

                st.success(
                    "✅ Document generated and saved successfully!"
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

    st.session_state.generated_draft = (
        edited_draft
    )

    st.divider()

    st.subheader(
        "📤 Export"
    )

    export_col1, export_col2 = st.columns(
        2
    )

    with export_col1:

        pdf_data = export_pdf(
            edited_draft
        )

        st.download_button(
            "📕 Download PDF",
            data=pdf_data,
            file_name=(
                "draftforge_document.pdf"
            ),
            mime="application/pdf",
            use_container_width=True,
        )

        docx_data = export_docx(
            edited_draft
        )

        st.download_button(
            "📘 Download DOCX",
            data=docx_data,
            file_name=(
                "draftforge_document.docx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            use_container_width=True,
        )

    with export_col2:

        txt_data = export_txt(
            edited_draft
        )

        st.download_button(
            "📄 Download TXT",
            data=txt_data,
            file_name=(
                "draftforge_document.txt"
            ),
            mime="text/plain",
            use_container_width=True,
        )

        png_data = export_png(
            edited_draft
        )

        st.download_button(
            "🖼️ Download PNG",
            data=png_data,
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
