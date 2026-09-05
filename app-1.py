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

# Groq multilingual speech-to-text model
WHISPER_MODEL = "whisper-large-v3"

GROQ_URL = "https://api.groq.com/openai/v1/chat/completions"

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
    "drafts.db"
)


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="DraftForge — AI Document Composer",
    page_icon="✍️",
    layout="wide",
)


# ============================================================
# SESSION STATE
# ============================================================

if "generated_draft" not in st.session_state:
    st.session_state["generated_draft"] = ""

if "voice_transcript" not in st.session_state:
    st.session_state["voice_transcript"] = ""


# ============================================================
# DATABASE
# ============================================================

def init_db():

    conn = sqlite3.connect(DB_PATH)

    cursor = conn.cursor()

    cursor.execute(
        """
        SELECT name
        FROM sqlite_master
        WHERE type='table'
        AND name='drafts'
        """
    )

    table_exists = cursor.fetchone()

    if not table_exists:

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

        cursor.execute(
            "PRAGMA table_info(drafts)"
        )

        existing_columns = {
            row[1]
            for row in cursor.fetchall()
        }

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

                cursor.execute(
                    f"""
                    ALTER TABLE drafts
                    ADD COLUMN {column} {data_type}
                    """
                )

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
        INSERT INTO drafts
        (
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
            datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            ),
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
        (draft_id,)
    )

    conn.commit()
    conn.close()


# Initialize / repair database before using history
init_db()


# ============================================================
# API KEY FUNCTIONS
# ============================================================

def get_secret(name):

    try:

        return st.secrets.get(
            name,
            ""
        )

    except Exception:

        return ""


def get_api_key(provider):

    if provider == "Groq (Llama 3.3)":

        return get_secret(
            "GROQ_API_KEY"
        )

    return get_secret(
        "GEMINI_API_KEY"
    )


# ============================================================
# VOICE TO TEXT — GROQ WHISPER
# ============================================================

def transcribe_audio(
    api_key,
    audio_file
):

    if not api_key:

        raise ValueError(
            "GROQ_API_KEY is not configured in "
            "Streamlit Secrets."
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
                "The audio recording is empty."
            )

        # Groq currently has a file-size limit.
        # This prevents unnecessarily large uploads.
        max_size = 25 * 1024 * 1024

        if len(audio_bytes) > max_size:

            raise ValueError(
                "The recording is larger than 25 MB. "
                "Please make a shorter recording."
            )

        filename = getattr(
            audio_file,
            "name",
            "recording.wav"
        )

        if not filename:

            filename = "recording.wav"

        transcription = (
            client.audio.transcriptions.create(
                file=(
                    filename,
                    audio_bytes
                ),
                model=WHISPER_MODEL,
                response_format="json",
                temperature=0.0,
            )
        )

        text = getattr(
            transcription,
            "text",
            ""
        )

        if not text:

            raise RuntimeError(
                "No speech could be detected in the recording."
            )

        return text.strip()

    except Exception as e:

        raise RuntimeError(
            f"Speech-to-text error: {str(e)}"
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
    output_language="Same as spoken language",
):

    if output_language == "Same as spoken language":

        language_instruction = """
Write the final document in the same language as
the user's provided content/request.

If the user's content is in Urdu, write in Urdu.
If it is in Pashto, write in Pashto.
If it is in English, write in English.
If it is in Arabic, write in Arabic.

Do not translate unless necessary for clarity.
""".strip()

    else:

        language_instruction = f"""
Write the final document entirely in {output_language}.

The user's input may be in another language.
Translate and professionally adapt the content into
{output_language} while preserving the intended meaning.
""".strip()

    return f"""
You are an expert professional writer and editor.

Create a polished, professional {document_type}.

Writing tone:
{tone}

Recipient:
{recipient}

Sender:
{sender}

Subject:
{subject}

Key points provided by the user:
{key_points}

OUTPUT LANGUAGE:
{output_language}

Language instructions:
{language_instruction}

Requirements:

1. Maintain the requested tone.
2. Preserve the user's intended meaning.
3. Do not invent important facts.
4. Organize the document logically.
5. Correct grammar and spelling.
6. Make the document ready to use.
7. Do not add explanations before or after the document.
8. Do not use markdown code blocks.
9. For formal letters and emails, include an appropriate greeting and closing.
10. Keep useful placeholders such as [Date], [Company Name], etc.
11. If the input is spoken language, clean up speech-recognition errors where the intended meaning is obvious.
12. Do not mention that the content came from speech recognition.
13. Return only the finished document.

Return only the finished document.
""".strip()


# ============================================================
# GROQ
# ============================================================

def call_groq(api_key, prompt):

    if not api_key:

        raise ValueError(
            "GROQ_API_KEY is not configured in "
            "Streamlit Secrets."
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

        return (
            response
            .choices[0]
            .message
            .content
            .strip()
        )

    except Exception as e:

        raise RuntimeError(
            f"Groq API error: {str(e)}"
        ) from e


# ============================================================
# GEMINI
# ============================================================

def call_gemini(api_key, prompt):

    if not api_key:

        raise ValueError(
            "GEMINI_API_KEY is not configured in "
            "Streamlit Secrets."
        )

    try:

        url = (
            f"{GEMINI_URL}?key={api_key}"
        )

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

        response = requests.post(
            url,
            json=payload,
            timeout=60
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
            f"Gemini API error: {str(e)}"
        ) from e


# ============================================================
# TEXT CLEANING FOR PDF
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

        text = text.replace(
            old,
            new
        )

    return (
        text
        .encode(
            "latin-1",
            "replace"
        )
        .decode("latin-1")
    )


# ============================================================
# PDF LINE WRAPPING
# ============================================================

def wrap_pdf_line(
    pdf,
    text,
    max_width
):

    text = clean_pdf_text(
