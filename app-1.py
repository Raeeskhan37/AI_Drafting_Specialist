"""
DraftForge — AI Document Composer
-----------------------------------
A Streamlit app that generates professional emails, letters, reports, and
other documents using a free LLM API (Groq or Gemini), lets the user edit
the draft, saves a history of past drafts, and exports the result as a
letterhead-styled PDF, DOCX, TXT, or image (PNG/JPG).

Run locally:
    streamlit run app.py

See README.md for Colab / Streamlit Cloud / GitHub deployment instructions.
"""

import io
import os
import sqlite3
import textwrap
from datetime import datetime

import requests
import streamlit as st
from docx import Document
from docx.shared import Inches
from fpdf import FPDF
from PIL import Image, ImageDraw, ImageFont

# --------------------------------------------------------------------------
# CONFIG
# --------------------------------------------------------------------------
st.set_page_config(page_title="DraftForge — AI Document Composer", page_icon="📝", layout="wide")

GROQ_MODEL = "llama-3.3-70b-versatile"
GROQ_URL = "https://api.groq.com/openai/v1/chat/completions"
GEMINI_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent"

DOC_TYPES = ["Email", "Formal Letter", "Business Report", "Cover Letter", "Custom"]
TONES = ["Formal", "Friendly/Casual", "Persuasive", "Apologetic", "Assertive", "Neutral"]

DB_PATH = os.path.join(os.path.dirname(__file__), "drafts.db")


# --------------------------------------------------------------------------
# DATABASE (lightweight local history — see README for persistence notes)
# --------------------------------------------------------------------------
def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        """CREATE TABLE IF NOT EXISTS drafts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user TEXT,
            created_at TEXT,
            doc_type TEXT,
            subject TEXT,
            content TEXT
        )"""
    )
    conn.commit()
    return conn


def save_draft(user, doc_type, subject, content):
    conn = init_db()
    conn.execute(
        "INSERT INTO drafts (user, created_at, doc_type, subject, content) VALUES (?, ?, ?, ?, ?)",
        (user, datetime.now().strftime("%Y-%m-%d %H:%M"), doc_type, subject, content),
    )
    conn.commit()
    conn.close()


def get_history(user, limit=20):
    conn = init_db()
    rows = conn.execute(
        "SELECT id, created_at, doc_type, subject, content FROM drafts "
        "WHERE user = ? ORDER BY id DESC LIMIT ?",
        (user, limit),
    ).fetchall()
    conn.close()
    return rows


def delete_draft(draft_id):
    conn = init_db()
    conn.execute("DELETE FROM drafts WHERE id = ?", (draft_id,))
    conn.commit()
    conn.close()


# --------------------------------------------------------------------------
# LLM CALLS
# --------------------------------------------------------------------------
def build_prompt(doc_type, tone, recipient, sender, subject, key_points, extra):
    parts = [f"Write a professional {doc_type.lower()} in a {tone.lower()} tone."]
    if recipient:
        parts.append(f"Recipient / audience: {recipient}.")
    if sender:
        parts.append(f"Sender: {sender}.")
    if subject:
        parts.append(f"Subject / purpose: {subject}.")
    if key_points:
        parts.append(f"Key points to include:\n{key_points}")
    if extra:
        parts.append(f"Additional instructions: {extra}")
    parts.append(
        "Return only the finished document text (with a greeting/sign-off if "
        "appropriate for the format). Do not add explanations, notes, or "
        "markdown formatting symbols like ** or #."
    )
    return "\n".join(parts)


def call_groq(api_key, prompt):
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    payload = {
        "model": GROQ_MODEL,
        "messages": [
            {"role": "system", "content": "You are an expert professional writer and editor."},
            {"role": "user", "content": prompt},
        ],
        "temperature": 0.6,
    }
    r = requests.post(GROQ_URL, headers=headers, json=payload, timeout=60)
    r.raise_for_status()
    return r.json()["choices"][0]["message"]["content"].strip()


def call_gemini(api_key, prompt):
    headers = {"Content-Type": "application/json"}
    payload = {"contents": [{"parts": [{"text": prompt}]}]}
    r = requests.post(f"{GEMINI_URL}?key={api_key}", headers=headers, json=payload, timeout=60)
    r.raise_for_status()
    return r.json()["candidates"][0]["content"]["parts"][0]["text"].strip()


def generate_document(provider, api_key, prompt):
    if provider == "Groq (Llama 3.3)":
        return call_groq(api_key, prompt)
    return call_gemini(api_key, prompt)


# --------------------------------------------------------------------------
# LETTERHEAD HELPERS
# --------------------------------------------------------------------------
def get_font(size, bold=False):
    candidates = ["DejaVuSans-Bold.ttf", "DejaVuSans.ttf"] if bold else ["DejaVuSans.ttf"]
    for name in candidates:
        try:
            return ImageFont.truetype(name, size)
        except Exception:
            continue
    return ImageFont.load_default()


# --------------------------------------------------------------------------
# EXPORT HELPERS
# --------------------------------------------------------------------------
def export_txt(text):
    return text.encode("utf-8")


def export_docx(text, company, address, logo_bytes):
    doc = Document()
    if logo_bytes:
        try:
            doc.add_picture(io.BytesIO(logo_bytes), width=Inches(1.2))
        except Exception:
            pass
    if company:
        doc.add_heading(company, level=1)
    if address:
        doc.add_paragraph(address)
    if company or address:
        doc.add_paragraph("_" * 60)
    for para in text.split("\n"):
        doc.add_paragraph(para)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def export_pdf(text, company, address, logo_bytes):
    pdf = FPDF()
    pdf.add_page()

    if logo_bytes:
        try:
            logo_path = "/tmp/_draftforge_logo.png"
            with open(logo_path, "wb") as f:
                f.write(logo_bytes)
            pdf.image(logo_path, x=10, y=8, w=25)
            pdf.set_xy(40, 10)
        except Exception:
            pdf.set_xy(10, 10)
    else:
        pdf.set_xy(10, 10)

    if company:
        pdf.set_font("Helvetica", "B", 16)
        pdf.cell(0, 8, company, ln=True)
        pdf.set_x(pdf.l_margin if not logo_bytes else 40)
    if address:
        pdf.set_font("Helvetica", size=10)
        pdf.set_x(pdf.l_margin if not logo_bytes else 40)
        pdf.cell(0, 6, address, ln=True)

    if company or address:
        pdf.ln(4)
        pdf.set_draw_color(150, 150, 150)
        pdf.line(10, pdf.get_y(), 200, pdf.get_y())
        pdf.ln(8)
    else:
        pdf.set_y(20)

    pdf.set_font("Helvetica", size=12)
    for line in text.split("\n"):
        wrapped = textwrap.wrap(line, 100) or [""]
        for w in wrapped:
            pdf.multi_cell(0, 8, w)
    return bytes(pdf.output(dest="S"))


def export_image(text, company, address, logo_bytes, fmt="PNG"):
    width = 1000
    body_font = get_font(20)
    header_font = get_font(26, bold=True)
    sub_font = get_font(16)

    wrapped_lines = []
    for line in text.split("\n"):
        wrapped_lines.extend(textwrap.wrap(line, 80) or [""])

    header_height = 0
    if company:
        header_height += 40
    if address:
        header_height += 26
    if company or address:
        header_height += 20

    line_height = 28
    height = max(400, header_height + line_height * len(wrapped_lines) + 100)
    img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(img)

    y = 30
    x_text = 40
    if logo_bytes:
        try:
            logo = Image.open(io.BytesIO(logo_bytes))
            logo.thumbnail((80, 80))
            img.paste(logo, (40, y))
            x_text = 140
        except Exception:
            pass

    if company:
        draw.text((x_text, y), company, fill="black", font=header_font)
        y += 40
    if address:
        draw.text((x_text, y), address, fill="gray", font=sub_font)
        y += 30

    if company or address:
        y += 10
        draw.line((40, y, width - 40, y), fill=(180, 180, 180), width=2)
        y += 25
    else:
        y = 40

    for line in wrapped_lines:
        draw.text((40, y), line, fill="black", font=body_font)
        y += line_height

    buf = io.BytesIO()
    save_fmt = "JPEG" if fmt.upper() == "JPG" else fmt.upper()
    if save_fmt == "JPEG":
        img = img.convert("RGB")
    img.save(buf, format=save_fmt)
    return buf.getvalue()


# --------------------------------------------------------------------------
# UI
# --------------------------------------------------------------------------
st.title("📝 DraftForge — AI Document Composer")
st.caption("Draft professional emails, letters, and reports — then edit and export.")

with st.sidebar:
    st.header("AI Provider")
    provider = st.selectbox("Provider", ["Groq (Llama 3.3)", "Gemini"])
    api_key = st.text_input(
        f"{provider.split()[0]} API key",
        type="password",
        value=os.environ.get("GROQ_API_KEY" if "Groq" in provider else "GEMINI_API_KEY", ""),
        help="Get a free key at console.groq.com or aistudio.google.com/apikey.",
    )

    st.divider()
    st.header("Your profile")
    username = st.text_input("Your name", value=st.session_state.get("username", "Guest"))
    st.session_state.username = username or "Guest"
    st.caption("Used only to tag and filter your saved draft history on this device.")

    st.divider()
    with st.expander("🖋️ Letterhead (for PDF / image / Word export)"):
        company = st.text_input("Company / your name")
        address = st.text_input("Address / contact line")
        logo_file = st.file_uploader("Logo (optional)", type=["png", "jpg", "jpeg"])
        logo_bytes = logo_file.read() if logo_file else None

    st.divider()
    st.subheader("📜 History")
    history = get_history(st.session_state.username)
    if not history:
        st.caption("No saved drafts yet.")
    for draft_id, created_at, d_type, subj, content in history:
        with st.expander(f"{created_at} — {d_type}: {subj or '(no subject)'}"):
            st.text(content[:200] + ("..." if len(content) > 200 else ""))
            c1, c2 = st.columns(2)
            if c1.button("Load", key=f"load_{draft_id}"):
                st.session_state.draft = content
                st.rerun()
            if c2.button("Delete", key=f"del_{draft_id}"):
                delete_draft(draft_id)
                st.rerun()

col1, col2 = st.columns(2)
with col1:
    doc_type = st.selectbox("Document type", DOC_TYPES)
    tone = st.selectbox("Tone", TONES)
    recipient = st.text_input("Recipient / audience (optional)")
    sender = st.text_input("Your name / sender (optional)")
with col2:
    subject = st.text_input("Subject / purpose")
    key_points = st.text_area("Key points to include", height=100)
    extra = st.text_area("Any other instructions (optional)", height=68)

generate = st.button("✨ Generate draft", type="primary", use_container_width=True)

if "draft" not in st.session_state:
    st.session_state.draft = ""

if generate:
    if not api_key:
        st.error("Please enter an API key in the sidebar first.")
    elif not subject and not key_points:
        st.error("Please provide at least a subject or some key points.")
    else:
        prompt = build_prompt(doc_type, tone, recipient, sender, subject, key_points, extra)
        with st.spinner("Composing your document..."):
            try:
                result = generate_document(provider, api_key, prompt)
                st.session_state.draft = result
                save_draft(st.session_state.username, doc_type, subject, result)
            except requests.HTTPError as e:
                st.error(f"API error: {e.response.status_code} — check your API key and quota.")
            except Exception as e:
                st.error(f"Something went wrong: {e}")

if st.session_state.draft:
    st.subheader("Your draft — edit freely below")
    edited = st.text_area("Draft", value=st.session_state.draft, height=350, label_visibility="collapsed")
    st.session_state.draft = edited

    st.subheader("Export")
    filename = st.text_input(
        "File name (no extension)",
        value=f"{doc_type.replace(' ', '_').lower()}_{datetime.now().strftime('%Y%m%d')}",
    )
    fmt = st.radio("Format", ["PDF", "DOCX", "TXT", "PNG", "JPG"], horizontal=True)

    if edited.strip():
        if fmt == "PDF":
            data, mime, ext = export_pdf(edited, company, address, logo_bytes), "application/pdf", "pdf"
        elif fmt == "DOCX":
            data = export_docx(edited, company, address, logo_bytes)
            mime, ext = "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "docx"
        elif fmt == "TXT":
            data, mime, ext = export_txt(edited), "text/plain", "txt"
        elif fmt == "PNG":
            data, mime, ext = export_image(edited, company, address, logo_bytes, "PNG"), "image/png", "png"
        else:
            data, mime, ext = export_image(edited, company, address, logo_bytes, "JPG"), "image/jpeg", "jpg"

        st.download_button(
            f"⬇️ Download as {fmt}",
            data=data,
            file_name=f"{filename}.{ext}",
            mime=mime,
            use_container_width=True,
        )

    if st.button("💾 Save current edits to history"):
        save_draft(st.session_state.username, doc_type, subject, edited)
        st.success("Saved!")
        st.rerun()
