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
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="DraftForge - AI Document Composer",
    page_icon="✍️",
    layout="wide",
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

    conn.execute(
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

    try:

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

    finally:

        conn.close()

    return rows


def delete_draft(draft_id):

    conn = sqlite3.connect(DB_PATH)

    conn.execute(
        "DELETE FROM drafts WHERE id = ?",
        (draft_id,),
    )

    conn.commit()
    conn.close()


init_db()


# ============================================================
# API KEYS
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
# VOICE TRANSCRIPTION
# ============================================================

def transcribe_audio(api_key, audio_file):

    if not api_key:
        raise ValueError(
            "GROQ_API_KEY is not configured in Streamlit Secrets."
        )

    if audio_file is None:
        raise ValueError("No audio recording was provided.")

    try:

        from groq import Groq

        client = Groq(api_key=api_key)

        audio_bytes = audio_file.getvalue()

        if not audio_bytes:
            raise ValueError("The recording is empty.")

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

        text = getattr(result, "text", "")

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
# PROMPT
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

        client = Groq(api_key=api_key)

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
# GEMINI
# ============================================================

def call_gemini(api_key, prompt):

    if not api_key:
        raise ValueError(
            "GEMINI_API_KEY is not configured in Streamlit Secrets."
        )

    try:

        url = GEMINI_URL + "?key=" + api_key

        payload = {
            "contents": [
                {
                    "parts": [
                        {
                            "text": prompt,
                        }
                    ]
                }
            ]
        }

        response = requests.post(
            url,
            json=payload,
            timeout=60,
        )

        response.raise_for_status()

        data = response.json()

        return (
            data["candidates"][0]
            ["content"]["parts"][0]
            ["text"]
            .strip()
        )

    except Exception as e:

        raise RuntimeError(
            "Gemini API error: " + str(e)
        ) from e


# ============================================================
# PDF TEXT CLEANING
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
        "\u2011": "-",
        "\u2010": "-",
    }

    for old, new in replacements.items():
        text = text.replace(old, new)

    return text.encode(
        "latin-1",
        "replace",
    ).decode("latin-1")


# ============================================================
# PDF WRAPPING
# ============================================================

def wrap_pdf_line(pdf, text, max_width):

    text = clean_pdf_text(text)

    if not text:
        return [""]

    words = text.split()

    lines = []
    current = ""

    for word in words:

        if pdf.get_string_width(word) > max_width:

            if current:
                lines.append(current)
                current = ""

            remaining = word

            while remaining:

                piece = ""

                for i in range(
                    1,
                    len(remaining) + 1,
                ):

                    candidate = remaining[:i]

                    if (
                        pdf.get_string_width(candidate)
                        <= max_width
                    ):
                        piece = candidate
                    else:
                        break

                if not piece:
                    piece = remaining[0]

                lines.append(piece)
                remaining = remaining[len(piece):]

            continue

        if current:
            candidate = current + " " + word
        else:
            candidate = word

        if pdf.get_string_width(candidate) <= max_width:

            current = candidate

        else:

            if current:
                lines.append(current)

            current = word

    if current:
        lines.append(current)

    return lines


# ============================================================
# PDF EXPORT
# ============================================================

def export_pdf(
    text,
    sender="",
    recipient="",
    subject="",
    logo=None,
):

    pdf = FPDF()

    pdf.set_auto_page_break(
        auto=True,
        margin=15,
    )

    pdf.add_page()

    if logo is not None:

        try:

            logo_stream = io.BytesIO(logo)

            pdf.image(
                logo_stream,
                x=10,
                y=10,
                w=35,
            )

            pdf.ln(30)

        except Exception:
            pass

    if sender:

        pdf.set_font(
            "Helvetica",
            "B",
            12,
        )

        pdf.cell(
            0,
            7,
            clean_pdf_text(sender),
            new_x="LMARGIN",
            new_y="NEXT",
        )

    if recipient:

        pdf.set_font(
            "Helvetica",
            "",
            11,
        )

        pdf.cell(
            0,
            7,
            clean_pdf_text(recipient),
            new_x="LMARGIN",
            new_y="NEXT",
        )

    if subject:

        pdf.ln(3)

        pdf.set_font(
            "Helvetica",
            "B",
            11,
        )

        pdf.cell(
            0,
            7,
            clean_pdf_text(
                "Subject: " + subject
            ),
            new_x="LMARGIN",
            new_y="NEXT",
        )

    pdf.ln(5)

    pdf.set_font(
        "Helvetica",
        "",
        11,
    )

    usable_width = (
        pdf.w
        - pdf.l_margin
        - pdf.r_margin
    )

    for raw_line in text.splitlines():

        if not raw_line.strip():

            pdf.ln(5)
            continue

        wrapped = wrap_pdf_line(
            pdf,
            raw_line,
            usable_width,
        )

        for line in wrapped:

            if not line:
                pdf.ln(5)
                continue

            line = clean_pdf_text(line)

            if not line:
                continue

            pdf.multi_cell(
                usable_width,
                7,
                line,
                new_x="LMARGIN
