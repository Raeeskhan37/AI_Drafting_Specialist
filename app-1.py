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
    "drafts.db",
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
        cursor.execute("PRAGMA table_info(drafts)")

        existing_columns = {
            row[1]
            for row in cursor.fetchall()
        }

        required_columns = {
            "document_type": "
