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

DB_PATH = os.path.join(
    os.path.dirname(__file__),
    "drafts.db",
)


# ============================================================
# SESSION STATE
# ============================================================

if "generated_draft" not in st.session_state:
    st.session_state.generated_draft = ""

if "normal_information" not in st.session_state:
    st.session_state.normal_information = ""

if "inquiry_indexes" not in st.session_state:
    st.session_state.inquiry_indexes = [
        {
            "title": "Index 1",
            "text": "",
        }
    ]


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

        inquiry_instruction = f"""
This is an official {inquiry_type}.

The information is organized into multiple
inquiry indexes.

Every supplied index is important.

Preserve the sequence and identity of the indexes.

Do not unnecessarily merge indexes.

Use professional official inquiry language.

For FFI Inquiry:
Use a fact-finding inquiry approach.

For E&D Inquiry:
Use a departmental disciplinary inquiry approach.

Clearly distinguish between allegations,
statements, facts, evidence, observations,
analysis, findings, conclusions and recommendations
when the supplied information supports such
distinctions.

Do not invent facts, evidence, statements,
dates, findings or conclusions.

If information is missing, use a suitable
placeholder or omit the unsupported fact.
"""

    prompt = f"""
You are an expert professional writer,
official correspondence specialist and
inquiry-document drafting assistant.

Create a polished professional {document_type}.

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

8. Keep useful placeholders such as [Date],
[Reference], [Office Name] or [Designation]
where appropriate.

9. For Email, use appropriate email structure.

10. For Letter, use appropriate official
letter structure.

11. For Inquiry, preserve all supplied indexes.

12. For Inquiry, create appropriate headings
and subheadings where useful.

13. Do not unnecessarily remove information
provided by the user.

14. Do not add explanations before or after
the document.

15. Return only the finished document.

16. Do not use markdown code blocks.

17. Never create unsupported facts.

18. If the user's information is incomplete,
do not guess.
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
            "FFI Inquiry — Fact Finding Inquiry"
        )

    else:

        st.info(
            "E&D Inquiry — Departmental "
            "disciplinary inquiry"
        )


# ============================================================
# TONE
# ============================================================

tone = st.selectbox(
    "🎯 Tone",
    TONES,
)


# ============================================================
# RECIPIENT / SENDER
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


# ============================================================
# SUBJECT
# ============================================================

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
# NORMAL INPUT
# ============================================================

def normal_information_box():

    st.markdown(
        "### 📝 Information"
    )

    st.caption(
        "You can type information and/or add "
        "voice information. Voice transcription "
        "will be added to the same text box."
    )

    text_key = (
        "normal_information"
    )

    editor_value = st.session_state.get(
        text_key,
        "",
    )

    st.text_area(
        "Type or edit your information",
        value=editor_value,
        height=250,
        key="normal_information_editor",
        placeholder=(
            "Type your request here..."
        ),
    )

    current_text = st.session_state.get(
        "normal_information_editor",
        "",
    )

    # Keep application state separate from
    # the widget's own state.
    st.session_state[
        text_key
    ] = current_text

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
            key="normal_transcribe_button",
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

                        existing = (
                            st.session_state.get(
                                text_key,
                                "",
                            )
                        )

                        if existing.strip():

                            combined = (
                                existing
                                + "\n"
                                + transcript
                            )

                        else:

                            combined = (
                                transcript
                            )

                        # Store result in the
                        # application state.
                        st.session_state[
                            text_key
                        ] = combined

                        st.session_state[
                            "normal_information_editor"
                        ] = combined

                        st.success(
                            "Voice added to the "
                            "information."
                        )

                        st.rerun()

                    except Exception as e:

                        st.error(
                            str(e)
                        )

    return st.session_state.get(
        text_key,
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
        "Add as many indexes as required. "
        "Each index accepts both typed and "
        "voice information."
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

        title_state_key = (
            "inquiry_title_state_"
            + str(index_number)
        )

        text_state_key = (
            "inquiry_text_state_"
            + str(index_number)
        )

        title_widget_key = (
            "inquiry_title_widget_"
            + str(index_number)
        )

        text_widget_key = (
            "inquiry_text_widget_"
            + str(index_number)
        )

        if title_state_key not in st.session_state:

            st.session_state[
                title_state_key
            ] = index_data.get(
                "title",
                "Index "
                + str(index_number + 1),
            )

        if text_state_key not in st.session_state:

            st.session_state[
                text_state_key
            ] = index_data.get(
                "text",
                "",
            )

        st.markdown(
            "#### 📌 Index "
            + str(index_number + 1)
        )

        title_value = st.text_input(
            "Index title",
            value=st.session_state[
                title_state_key
            ],
            key=title_widget_key,
            placeholder=(
                "e.g. Allegation, Statement "
                "of accused, Witness statement, "
                "Documentary evidence"
            ),
        )

        text_value = st.text_area(
            "Type or edit information",
            value=st.session_state[
                text_state_key
            ],
            height=180,
            key=text_widget_key,
            placeholder=(
                "Type information for this index..."
            ),
        )

        # IMPORTANT:
        # Do not modify text_widget_key after
        # the widget has been instantiated.
        # Instead, copy the widget value into
        # our separate application state.
        st.session_state[
            title_state_key
        ] = title_value

        st.session_state[
            text_state_key
        ] = text_value

        st.session_state.inquiry_indexes[
            index_number
        ]["title"] = title_value

        st.session_state.inquiry_indexes[
            index_number
        ]["text"] = text_value

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

                            existing = (
                                st.session_state.get(
                                    text_state_key,
                                    "",
                                )
                            )

                            if existing.strip():

                                combined = (
                                    existing
                                    + "\n"
                                    + transcript
                                )

                            else:

                                combined = (
                                    transcript
                                )

                            # Update application state,
                            # NOT the instantiated widget.
                            st.session_state[
                                text_state_key
                            ] = combined

                            st.session_state.inquiry_indexes[
                                index_number
                            ]["text"] = combined

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
# DISPLAY INPUT AREA
# ============================================================

if document_type == "Inquiry":

    information = (
        inquiry_information_boxes()
    )

else:

    information = (
        normal_information_box()
    )


# ============================================================
# GENERATE
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
