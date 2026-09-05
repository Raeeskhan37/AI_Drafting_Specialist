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
# CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="DraftForge - AI Document Composer",
    page_icon="📝",
    layout="wide",
)

GROQ_MODEL = "openai/gpt-oss-120b"
WHISPER_MODEL = "whisper-large-v3"

GEMINI_URL = (
    "https://generativelanguage.googleapis.com/"
    "v1beta/models/gemini-2.0-flash:generateContent"
)

DOC_TYPES = [
    "Email",
    "Formal Letter",
    "Business Report",
    "Cover Letter",
    "Custom",
]

TONES = [
    "Formal",
    "Friendly/Casual",
    "Persuasive",
    "Apologetic",
    "Assertive",
    "Neutral",
]

OUTPUT_LANGUAGES = [
    "Same as spoken language",
    "English",
    "Urdu",
    "Pashto",
    "Arabic",
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

if "voice_transcript" not in st.session_state:
    st.session_state.voice_transcript = ""


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
        cursor.execute("PRAGMA table_info(drafts)")

        existing_columns = set()

        for row in cursor.fetchall():
            existing_columns.add(row[1])

        required_columns = {
            "document_type": "TEXT",
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
            tone,
            recipient,
            sender,
            subject,
            key_points,
            draft,
            created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            document_type,
            tone,
            recipient,
            sender,
            subject,
            key_points,
            draft,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
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


# Initialize database before anything uses it.
init_db()


# ============================================================
# API KEY FUNCTIONS
# ============================================================

def get_secret(name):
    try:
        return st.secrets.get(name, "")
    except Exception:
        return ""


def get_api_key(provider):
    if provider == "Groq (Llama 3.3)":
        return get_secret("GROQ_API_KEY")

    return get_secret("GEMINI_API_KEY")


# ============================================================
# VOICE TO TEXT
# ============================================================

def transcribe_audio(api_key, audio_file):
    if not api_key:
        raise ValueError(
            "GROQ_API_KEY is not configured in Streamlit Secrets."
        )

    if audio_file is None:
        raise ValueError(
            "No audio recording was provided."
        )

    try:
        from groq import Groq

        client = Groq(api_key=api_key)

        audio_bytes = audio_file.getvalue()

        if not audio_bytes:
            raise ValueError(
                "The recording is empty."
            )

        if len(audio_bytes) > 25 * 1024 * 1024:
            raise ValueError(
                "The recording is larger than 25 MB. "
                "Please make a shorter recording."
            )

        filename = getattr(
            audio_file,
            "name",
            "recording.wav",
        )

        if not filename:
            filename = "recording.wav"

        result = client.audio.transcriptions.create(
            file=(filename, audio_bytes),
            model=WHISPER_MODEL,
            response_format="json",
            temperature=0.0,
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
            "Speech-to-text error: " + str(e)
        ) from e


# ============================================================
# PROMPT BUILDER
# ============================================================

def build_prompt(
    document_type,
    tone,
    recipient,
    sender,
    subject,
    key_points,
    output_language,
):
    if output_language == "Same as spoken language":
        language_instruction = (
            "Write the final document in the same language "
            "as the user's content. Preserve the intended meaning."
        )
    else:
        language_instruction = (
            "Write the final document entirely in "
            + output_language
            + ". Translate and professionally adapt the "
              "user's content while preserving its meaning."
        )

    prompt = f"""
You are an expert professional writer and editor.

Create a polished, professional {document_type}.

Tone:
{tone}

Recipient:
{recipient}

Sender:
{sender}

Subject:
{subject}

User information:
{key_points}

Desired output language:
{output_language}

Language instruction:
{language_instruction}

Requirements:

1. Maintain the requested tone.
2. Preserve the user's intended meaning.
3. Do not invent important facts.
4. Organize the document logically.
5. Correct grammar and spelling.
6. Make the document ready to use.
7. For letters and emails, include an appropriate greeting and closing.
8. Do not add explanations before or after the document.
9. Do not use markdown code blocks.
10. Keep useful placeholders such as [Date] or [Company Name].
11. Correct obvious speech-recognition errors when the intended meaning is clear.
12. Return only the finished document.
"""

    return prompt.strip()


# ============================================================
# GROQ TEXT GENERATION
# ============================================================

def call_groq(api_key, prompt):
    if not api_key:
        raise ValueError(
            "GROQ_API_KEY is not configured in Streamlit Secrets."
        )

    try:
        from groq import Groq

        client = Groq(
            api_key=api_key
        )

        response = client.chat.completions.create(
            model=GROQ_MODEL,
            messages=[
                {
                    "role": "system",
                    "content": (
                        "You are an expert professional "
                        "writer and editor."
                    ),
                },
                {
                    "role": "user",
                    "content": prompt,
                },
            ],
            temperature=0.6,
        )

        return response.choices[0].message.content.strip()

    except Exception as e:
        raise RuntimeError(
            "Groq API error: " + str(e)
        ) from e


# ============================================================
# GEMINI TEXT GENERATION
# ============================================================

def call_gemini(api_key, prompt):
    if not api_key:
        raise ValueError(
            "GEMINI_API_KEY is not configured in Streamlit Secrets."
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

    url = GEMINI_URL + "?key=" + api_key

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

        parts = candidates[0].get(
            "content",
            {},
        ).get(
            "parts",
            [],
        )

        if not parts:
            raise RuntimeError(
                "Gemini returned an empty response."
            )

        return parts[0].get(
            "text",
            "",
        ).strip()

    except requests.RequestException as e:
        raise RuntimeError(
            "Gemini connection error: " + str(e)
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

    return text.encode(
        "latin-1",
        "replace",
    ).decode(
        "latin-1"
    )


def wrap_pdf_line(pdf, text, max_width):
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

        if pdf.get_string_width(
            test_line
        ) <= max_width:
            current = test_line
        else:
            if current:
                lines.append(current)

            current = word

    if current:
        lines.append(current)

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
        pdf.w - pdf.l_margin - pdf.r_margin
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

    section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.75)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

    paragraphs = text.split(
        "\n"
    )

    for paragraph in paragraphs:
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

    lines = []

    dummy_image = Image.new(
        "RGB",
        (max_width, 100),
        "white",
    )

    draw = ImageDraw.Draw(
        dummy_image
    )

    for paragraph in text.split(
        "\n"
    ):
        words = paragraph.split(" ")

        current = ""

        if not words:
            lines.append("")
            continue

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

            width = bbox[2] - bbox[0]

            if width <= max_width - (
                margin * 2
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

    line_height = 36

    height = max(
        200,
        margin * 2
        + len(lines) * line_height,
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

    st.title("📝 DraftForge")

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
            "Groq (Llama 3.3)",
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
        if provider == "Groq (Llama 3.3)":
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
            f"{len(history)} saved draft(s)"
        )

        for row in history[:10]:

            draft_id = row[0]
            document_type = row[1]
            subject = row[5]
            created_at = row[8]

            title = (
                subject
                if subject
                else document_type
            )

            with st.expander(
                f"{title} - {created_at}"
            ):

                st.write(
                    f"**Type:** {document_type}"
                )

                st.write(
                    f"**Subject:** "
                    f"{subject or 'N/A'}"
                )

                if st.button(
                    "Load Draft",
                    key=f"load_{draft_id}",
                    use_container_width=True,
                ):
                    st.session_state.generated_draft = (
                        row[7] or ""
                    )

                    st.rerun()

    else:
        st.caption(
            "No drafts saved yet."
        )

    st.divider()

    st.caption(
        "DraftForge • AI-powered document creation"
    )


# ============================================================
# MAIN HEADER
# ============================================================

st.title(
    "📝 DraftForge — AI Document Composer"
)

st.write(
    "Create professional emails, letters, reports "
    "and other documents using AI."
)

st.divider()


# ============================================================
# DOCUMENT SETTINGS
# ============================================================

col1, col2 = st.columns(2)

with col1:

    document_type = st.selectbox(
        "📄 Document Type",
        DOC_TYPES,
    )

with col2:

    tone = st.selectbox(
        "🎯 Tone",
        TONES,
    )


col3, col4 = st.columns(2)

with col3:

    recipient = st.text_input(
        "👤 Recipient",
        placeholder="e.g. Manager",
    )

with col4:

    sender = st.text_input(
        "✍️ Sender",
        value=profile_name,
    )


subject = st.text_input(
    "📌 Subject",
    placeholder="Enter document subject",
)


# ============================================================
# INPUT MODE
# ============================================================

st.subheader(
    "💬 Provide Your Information"
)

input_mode = st.radio(
    "Choose how you want to provide the information",
    [
        "⌨️ Type",
        "🎙️ Speak",
    ],
    horizontal=True,
)


# ============================================================
# TYPE MODE
# ============================================================

if input_mode == "⌨️ Type":

    key_points = st.text_area(
        "📝 Key Points / Details",
        height=220,
        placeholder=(
            "Describe what you want the document to say..."
        ),
    )


# ============================================================
# VOICE MODE
# ============================================================

else:

    selected_output_language = st.selectbox(
        "🌐 Desired output language",
        OUTPUT_LANGUAGES,
    )

    st.info(
        "Speak naturally in your preferred language. "
        "DraftForge will convert your speech to text "
        "and then generate the requested document."
    )

    voice_audio = st.audio_input(
        "🎙️ Record your request",
        sample_rate=16000,
        key="voice_recorder",
    )

    if voice_audio is not None:

        st.audio(
            voice_audio
        )

        if st.button(
            "📝 Transcribe Voice",
            type="secondary",
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
                    "🎙️ Converting speech to text..."
                ):

                    try:

                        transcript = transcribe_audio(
                            groq_key,
                            voice_audio,
                        )

                        st.session_state.voice_transcript = (
                            transcript
                        )

                        st.success(
                            "Voice successfully converted to text!"
                        )

                    except Exception as e:

                        st.error(
                            str(e)
                        )

    key_points = st.text_area(
        "📝 Transcribed request - edit if necessary",
        value=st.session_state.voice_transcript,
        height=220,
        key="voice_transcript_editor",
        placeholder=(
            "Your spoken request will appear here "
            "after transcription..."
        ),
    )

    st.session_state.voice_transcript = (
        key_points
    )


# ============================================================
# GENERATE BUTTON
# ============================================================

st.divider()

generate_button = st.button(
    "✨ Generate Draft",
    type="primary",
    use_container_width=True,
)


if generate_button:

    if not key_points.strip():

        st.warning(
            "Please provide some information "
            "by typing or speaking."
        )

    elif not api_key:

        st.error(
            "Please configure the selected API key "
            "in Streamlit Secrets."
        )

    else:

        if input_mode == "🎙️ Speak":

            output_language = (
                selected_output_language
            )

        else:

            output_language = "English"

        prompt = build_prompt(
            document_type=document_type,
            tone=tone,
            recipient=recipient,
            sender=sender,
            subject=subject,
            key_points=key_points,
            output_language=output_language,
        )

        with st.spinner(
            "✨ Creating your professional document..."
        ):

            try:

                if provider == "Groq (Llama 3.3)":

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
                    tone,
                    recipient,
                    sender,
                    subject,
                    key_points,
                    draft,
                )

                st.success(
                    "✅ Draft generated and saved successfully!"
                )

            except Exception as e:

                st.error(
                    str(e)
                )


# ============================================================
# GENERATED DRAFT
# ============================================================

if st.session_state.generated_draft:

    st.divider()

    st.subheader(
        "✏️ Generated Draft"
    )

    edited_draft = st.text_area(
        "Edit your document before exporting",
        value=st.session_state.generated_draft,
        height=500,
        key="draft_editor",
    )

    st.session_state.generated_draft = (
        edited_draft
    )

    st.divider()

    st.subheader(
        "📤 Export Document"
    )

    export_col1, export_col2 = st.columns(2)

    with export_col1:

        pdf_data = export_pdf(
            edited_draft
        )

        st.download_button(
            "📕 Download PDF",
            data=pdf_data,
            file_name="draftforge_document.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

        docx_data = export_docx(
            edited_draft
        )

        st.download_button(
            "📘 Download DOCX",
            data=docx_data,
            file_name="draftforge_document.docx",
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
            file_name="draftforge_document.txt",
            mime="text/plain",
            use_container_width=True,
        )

        png_data = export_png(
            edited_draft
        )

        st.download_button(
            "🖼️ Download PNG",
            data=png_data,
            file_name="draftforge_document.png",
            mime="image/png",
            use_container_width=True,
        )


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "DraftForge — Professional AI document creation "
    "with text and multilingual voice input."
)
