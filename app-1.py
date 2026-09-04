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
from PIL import Image


# ============================================================
# CONFIGURATION
# ============================================================

GROQ_MODEL = "openai/gpt-oss-120b"

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

DB_PATH = os.path.join(os.path.dirname(__file__), "drafts.db")


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="DraftForge — AI Document Composer",
    page_icon="✍️",
    layout="wide",
)


# ============================================================
# DATABASE
# ============================================================

def init_db():
    conn = sqlite3.connect(DB_PATH)

    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS drafts (
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
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        ),
    )

    conn.commit()
    conn.close()


def get_history():
    conn = sqlite3.connect(DB_PATH)

    rows = conn.execute(
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
    ).fetchall()

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
# SECRETS / API KEYS
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
# PROMPT BUILDER
# ============================================================

def build_prompt(
    document_type,
    tone,
    recipient,
    sender,
    subject,
    key_points,
):

    prompt = f"""
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

Requirements:

1. Write clear and professional English.
2. Maintain the requested tone.
3. Do not invent important facts.
4. Organize the document logically.
5. Correct grammar and spelling.
6. Make the document ready to use.
7. Do not add explanations before or after the document.
8. Do not use markdown code blocks.
9. For formal letters and emails, include an appropriate greeting and closing.
10. Keep placeholders such as [Date], [Company Name], etc. when useful.

Return only the finished document.
"""

    return prompt.strip()


# ============================================================
# GROQ API
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
                        "You are an expert professional writer and editor."
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
            f"Groq API error: {str(e)}"
        ) from e


# ============================================================
# GEMINI API
# ============================================================

def call_gemini(api_key, prompt):

    if not api_key:
        raise ValueError(
            "GEMINI_API_KEY is not configured in Streamlit Secrets."
        )

    try:

        url = f"{GEMINI_URL}?key={api_key}"

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
            timeout=60,
        )

        response.raise_for_status()

        data = response.json()

        return (
            data["candidates"][0]["content"]["parts"][0]["text"]
            .strip()
        )

    except Exception as e:

        raise RuntimeError(
            f"Gemini API error: {str(e)}"
        ) from e


# ============================================================
# DOCUMENT EXPORT
# ============================================================

def export_docx(
    text,
    sender="",
    recipient="",
    subject="",
    logo=None,
):

    document = Document()

    # Logo
    if logo is not None:
        try:
            document.add_picture(
                logo,
                width=Inches(1.5),
            )
        except Exception:
            pass

    # Letterhead
    if sender:
        p = document.add_paragraph()
        run = p.add_run(sender)
        run.bold = True

    if recipient:
        document.add_paragraph(recipient)

    if subject:
        p = document.add_paragraph()
        run = p.add_run(f"Subject: {subject}")
        run.bold = True

    document.add_paragraph("")

    # Main text
    for line in text.splitlines():

        document.add_paragraph(line)

    buffer = io.BytesIO()

    document.save(buffer)

    buffer.seek(0)

    return buffer.getvalue()


# ============================================================
# SAFE TEXT FOR PDF
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

    # Helvetica in fpdf2 is Latin-1 based.
    text = text.encode(
        "latin-1",
        "replace",
    ).decode("latin-1")

    return text


# ============================================================
# SAFE PDF LINE WRAPPING
# ============================================================

def wrap_pdf_line(pdf, text, max_width):

    text = clean_pdf_text(text)

    if not text:
        return [""]

    words = text.split(" ")

    lines = []

    current = ""

    for word in words:

        # Handle extremely long words.
        if pdf.get_string_width(word) > max_width:

            if current:
                lines.append(current)
                current = ""

            remaining = word

            while remaining:

                best = ""

                for i in range(1, len(remaining) + 1):

                    candidate = remaining[:i]

                    if pdf.get_string_width(candidate) <= max_width:
                        best = candidate
                    else:
                        break

                if not best:
                    best = remaining[0]

                lines.append(best)

                remaining = remaining[len(best):]

            continue

        candidate = (
            word
            if not current
            else current + " " + word
        )

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

    # Page settings
    pdf.set_auto_page_break(
        auto=True,
        margin=15,
    )

    pdf.add_page()

    # --------------------------------------------------------
    # LOGO
    # --------------------------------------------------------

    if logo is not None:

        try:

            if isinstance(logo, bytes):

                logo_stream = io.BytesIO(logo)

                pdf.image(
                    logo_stream,
                    x=10,
                    y=10,
                    w=35,
                )

            elif isinstance(logo, str) and os.path.exists(logo):

                pdf.image(
                    logo,
                    x=10,
                    y=10,
                    w=35,
                )

            pdf.ln(30)

        except Exception:
            pass

    # --------------------------------------------------------
    # LETTERHEAD
    # --------------------------------------------------------

    pdf.set_font(
        "Helvetica",
        "B",
        12,
    )

    if sender:

        sender_clean = clean_pdf_text(sender)

        pdf.cell(
            0,
            7,
            sender_clean,
            new_x="LMARGIN",
            new_y="NEXT",
        )

    pdf.set_font(
        "Helvetica",
        "",
        11,
    )

    if recipient:

        recipient_clean = clean_pdf_text(recipient)

        pdf.cell(
            0,
            7,
            recipient_clean,
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

        subject_clean = clean_pdf_text(
            f"Subject: {subject}"
        )

        pdf.cell(
            0,
            7,
            subject_clean,
            new_x="LMARGIN",
            new_y="NEXT",
        )

    pdf.ln(5)

    # --------------------------------------------------------
    # BODY
    # --------------------------------------------------------

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

    # Safety margin so fpdf2 never receives an
    # unusably small width.
    usable_width = max(
        20,
        usable_width,
    )

    for raw_line in text.splitlines():

        raw_line = raw_line.rstrip()

        # Blank line
        if not raw_line.strip():

            pdf.ln(5)

            continue

        wrapped_lines = wrap_pdf_line(
            pdf,
            raw_line,
            usable_width,
        )

        for line in wrapped_lines:

            if not line:
                pdf.ln(5)
                continue

            # Extra safety check
            line = clean_pdf_text(line)

            if not line:
                pdf.ln(5)
                continue

            # If somehow the line is still too wide,
            # shrink it one character at a time.
            while (
                pdf.get_string_width(line)
                > usable_width
                and len(line) > 1
            ):
                line = line[:-1]

            if not line:
                continue

            pdf.multi_cell(
                usable_width,
                7,
                line,
                new_x="LMARGIN",
                new_y="NEXT",
            )

    # --------------------------------------------------------
    # RETURN PDF BYTES
    # --------------------------------------------------------

    output = pdf.output()

    return bytes(output)


# ============================================================
# TEXT EXPORT
# ============================================================

def export_txt(text):

    return text.encode(
        "utf-8"
    )


# ============================================================
# IMAGE EXPORT
# ============================================================

def export_image(text):

    try:

        from PIL import ImageDraw, ImageFont

        lines = text.splitlines()

        font = ImageFont.load_default()

        line_height = 18

        width = 1200

        height = max(
            200,
            (len(lines) + 4) * line_height,
        )

        image = Image.new(
            "RGB",
            (width, height),
            "white",
        )

        draw = ImageDraw.Draw(image)

        y = 20

        for line in lines:

            draw.text(
                (30, y),
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

    except Exception as e:

        raise RuntimeError(
            f"Image export error: {str(e)}"
        )


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.title("✍️ DraftForge")

    st.caption(
        "AI Document Composer"
    )

    st.divider()

    # --------------------------------------------------------
    # AI PROVIDER
    # --------------------------------------------------------

    st.subheader("🤖 AI Provider")

    provider = st.selectbox(
        "Choose AI provider",
        [
            "Groq (Llama 3.3)",
            "Gemini",
        ],
    )

    api_key = get_api_key(provider)

    if api_key:

        st.success(
            "🔐 API key loaded securely."
        )

    else:

        if provider == "Groq (Llama 3.3)":

            st.warning(
                "GROQ_API_KEY not found in Streamlit Secrets."
            )

        else:

            st.warning(
                "GEMINI_API_KEY not found in Streamlit Secrets."
            )

    st.divider()

    # --------------------------------------------------------
    # PROFILE
    # --------------------------------------------------------

    st.subheader("👤 Profile")

    profile_name = st.text_input(
        "Your name",
        value="",
        key="profile_name",
    )

    profile_organization = st.text_input(
        "Organization",
        value="",
        key="profile_organization",
    )

    st.divider()

    # --------------------------------------------------------
    # LETTERHEAD
    # --------------------------------------------------------

    st.subheader("🖼️ Letterhead")

    logo_file = st.file_uploader(
        "Upload logo",
        type=[
            "png",
            "jpg",
            "jpeg",
        ],
    )

    logo_bytes = None

    if logo_file is not None:

        logo_bytes = logo_file.getvalue()

        st.image(
            logo_bytes,
            width=120,
        )

    st.divider()

    # --------------------------------------------------------
    # HISTORY
    # --------------------------------------------------------

    st.subheader("📚 Draft History")

    history = get_history()

    if history:

        for row in history[:10]:

            draft_id = row[0]
            doc_type = row[1]
            subject = row[5]
            created = row[8]

            title = subject or doc_type

            with st.expander(
                f"{title} — {created}"
            ):

                st.caption(
                    f"Type: {doc_type}"
                )

                if st.button(
                    "Delete",
                    key=f"delete_{draft_id}",
                ):

                    delete_draft(draft_id)

                    st.rerun()

                st.text_area(
                    "Draft",
                    row[7],
                    height=150,
                    key=f"history_{draft_id}",
                )

    else:

        st.info(
            "No saved drafts yet."
        )


# ============================================================
# MAIN PAGE
# ============================================================

st.title(
    "✍️ DraftForge — AI Document Composer"
)

st.write(
    "Create professional documents using AI."
)

st.divider()


# ============================================================
# INPUT FORM
# ============================================================

col1, col2 = st.columns(2)


with col1:

    document_type = st.selectbox(
        "Document Type",
        DOC_TYPES,
    )

    tone = st.selectbox(
        "Tone",
        TONES,
    )

    recipient = st.text_input(
        "Recipient",
        placeholder="e.g. Manager",
    )

    sender = st.text_input(
        "Sender",
        value=profile_name,
        placeholder="Your name",
    )


with col2:

    subject = st.text_input(
        "Subject",
        placeholder="e.g. Request for grant of bonus salary",
    )

    key_points = st.text_area(
        "Key Points / Details",
        placeholder=(
            "Enter the important information "
            "you want included..."
        ),
        height=180,
    )


st.divider()


# ============================================================
# GENERATE
# ============================================================

generate = st.button(
    "✨ Generate Draft",
    type="primary",
    use_container_width=True,
)


if generate:

    if not api_key:

        st.error(
            "Please configure the selected API key "
            "in Streamlit Secrets first."
        )

    elif not key_points.strip():

        st.warning(
            "Please enter some key points or details."
        )

    else:

        prompt = build_prompt(
            document_type=document_type,
            tone=tone,
            recipient=recipient,
            sender=sender,
            subject=subject,
            key_points=key_points,
        )

        with st.spinner(
            "Generating your professional draft..."
        ):

            try:

                if provider == "Groq (Llama 3.3)":

                    generated = call_groq(
                        api_key,
                        prompt,
                    )

                else:

                    generated = call_gemini(
                        api_key,
                        prompt,
                    )

                st.session_state[
                    "generated_draft"
                ] = generated

                save_draft(
                    document_type,
                    tone,
                    recipient,
                    sender,
                    subject,
                    key_points,
                    generated,
                )

                st.success(
                    "Draft generated successfully!"
                )

            except Exception as e:

                st.error(
                    str(e)
                )


# ============================================================
# DRAFT EDITOR
# ============================================================

if "generated_draft" in st.session_state:

    st.subheader(
        "📝 Your draft — edit freely below"
    )

    edited = st.text_area(
        "Draft",
        value=st.session_state[
            "generated_draft"
        ],
        height=500,
        key="draft_editor",
    )

    # Keep session state updated
    st.session_state[
        "generated_draft"
    ] = edited

    st.divider()

    # ========================================================
    # EXPORT
    # ========================================================

    st.subheader(
        "📤 Export"
    )

    fmt = st.radio(
        "Choose format",
        [
            "PDF",
            "DOCX",
            "TXT",
            "PNG",
        ],
        horizontal=True,
    )

    # --------------------------------------------------------
    # PDF
    # --------------------------------------------------------

    if fmt == "PDF":

        try:

            pdf_data = export_pdf(
                edited,
                sender=sender,
                recipient=recipient,
                subject=subject,
                logo=logo_bytes,
            )

            st.download_button(
                label="⬇️ Download PDF",
                data=pdf_data,
                file_name="draftforge_document.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

        except Exception as e:

            st.error(
                f"PDF export failed: {str(e)}"
            )

    # --------------------------------------------------------
    # DOCX
    # --------------------------------------------------------

    elif fmt == "DOCX":

        try:

            docx_data = export_docx(
                edited,
                sender=sender,
                recipient=recipient,
                subject=subject,
                logo=logo_bytes,
            )

            st.download_button(
                label="⬇️ Download DOCX",
                data=docx_data,
                file_name="draftforge_document.docx",
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.wordprocessingml.document"
                ),
                use_container_width=True,
            )

        except Exception as e:

            st.error(
                f"DOCX export failed: {str(e)}"
            )

    # --------------------------------------------------------
    # TXT
    # --------------------------------------------------------

    elif fmt == "TXT":

        try:

            txt_data = export_txt(
                edited
            )

            st.download_button(
                label="⬇️ Download TXT",
                data=txt_data,
                file_name="draftforge_document.txt",
                mime="text/plain",
                use_container_width=True,
            )

        except Exception as e:

            st.error(
                f"TXT export failed: {str(e)}"
            )

    # --------------------------------------------------------
    # PNG
    # --------------------------------------------------------

    elif fmt == "PNG":

        try:

            png_data = export_image(
                edited
            )

            st.download_button(
                label="⬇️ Download PNG",
                data=png_data,
                file_name="draftforge_document.png",
                mime="image/png",
                use_container_width=True,
            )

        except Exception as e:

            st.error(
                f"PNG export failed: {str(e)}"
            )


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "DraftForge — AI-powered professional document creation"
            )
