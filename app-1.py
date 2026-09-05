import io
import os
import sqlite3
import uuid
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
    layout="wide"
)


# ============================================================
# CONSTANTS
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
    "Inquiry"
]

INQUIRY_TYPES = [
    "FFI Inquiry",
    "E&D Inquiry"
]

OUTPUT_LANGUAGES = [
    "English",
    "Urdu",
    "Pashto",
    "Arabic",
    "Same as input"
]

TONES = [
    "Formal",
    "Professional",
    "Friendly/Casual",
    "Persuasive",
    "Apologetic",
    "Assertive",
    "Neutral"
]


# ============================================================
# E&D INDEXES
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
# INDEXES THAT GET AUTOMATIC NUMBERING WHEN REPEATED
# ============================================================

NUMBERABLE_INDEXES = [
    "Statement of the Accused",
    "Questions / Answers with the Accused",
]


# ============================================================
# DOCUMENTS RECORDED
# ============================================================

DOCUMENT_OPTIONS = [
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


# ============================================================
# DATABASE
# ============================================================

DB_PATH = os.path.join(
    os.path.dirname(__file__),
    "drafts.db"
)


def init_db():

    conn = sqlite3.connect(DB_PATH)

    cursor = conn.cursor()

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS drafts (
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

    conn.commit()

    cursor.execute("PRAGMA table_info(drafts)")

    columns = [
        row[1]
        for row in cursor.fetchall()
    ]

    required_columns = {
        "inquiry_type": "TEXT",
        "tone": "TEXT",
        "recipient": "TEXT",
        "sender": "TEXT",
        "subject": "TEXT",
        "key_points": "TEXT",
        "draft": "TEXT",
        "created_at": "TEXT",
    }

    for column, column_type in required_columns.items():

        if column not in columns:

            cursor.execute(
                f"ALTER TABLE drafts ADD COLUMN "
                f"{column} {column_type}"
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
    draft
):

    conn = sqlite3.connect(DB_PATH)

    cursor = conn.cursor()

    cursor.execute(
        """
        INSERT INTO drafts
        (
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
        )
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
# SESSION STATE
# ============================================================

DEFAULT_STATE = {
    "generated_draft": "",
    "normal_information": "",
    "normal_voice_status": "",
    "ed_index_instances": [],
    "ed_index_selector": "-- Select an index --",
}

for key, value in DEFAULT_STATE.items():

    if key not in st.session_state:

        if isinstance(value, list):
            st.session_state[key] = []

        elif isinstance(value, dict):
            st.session_state[key] = dict(value)

        else:
            st.session_state[key] = value


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
    audio_file
):

    if not api_key:

        raise ValueError(
            "Groq API key is not configured."
        )

    if audio_file is None:

        raise ValueError(
            "No audio recording was found."
        )

    audio_bytes = audio_file.getvalue()

    if not audio_bytes:

        raise ValueError(
            "The audio recording is empty."
        )

    if len(audio_bytes) > 25 * 1024 * 1024:

        raise ValueError(
            "Audio file is larger than 25 MB."
        )

    try:

        from groq import Groq

        client = Groq(
            api_key=api_key
        )

        filename = getattr(
            audio_file,
            "name",
            "recording.wav"
        )

        result = client.audio.transcriptions.create(
            file=(
                filename,
                audio_bytes
            ),
            model=WHISPER_MODEL,
            response_format="json",
            temperature=0.0,
        )

        text = getattr(
            result,
            "text",
            ""
        )

        return text.strip()

    except Exception as e:

        raise RuntimeError(
            f"Voice transcription failed: {e}"
        )


# ============================================================
# VOICE -> TEXT
# ============================================================

def append_voice_to_widget(
    audio_key,
    text_widget_key,
    status_key
):

    audio_file = st.session_state.get(
        audio_key
    )

    if not audio_file:

        st.session_state[
            status_key
        ] = "Please record your voice first."

        return

    api_key = get_api_key(
        "Groq"
    )

    try:

        transcribed_text = transcribe_audio(
            api_key,
            audio_file
        )

        if not transcribed_text:

            st.session_state[
                status_key
            ] = "No speech was detected."

            return

        existing_text = st.session_state.get(
            text_widget_key,
            ""
        )

        if existing_text.strip():

            combined_text = (
                existing_text.rstrip()
                + "\n"
                + transcribed_text
            )

        else:

            combined_text = transcribed_text

        st.session_state[
            text_widget_key
        ] = combined_text

        st.session_state[
            status_key
        ] = "Voice text added successfully."

    except Exception as e:

        st.session_state[
            status_key
        ] = str(e)


def add_normal_voice():

    append_voice_to_widget(
        "normal_voice",
        "normal_information",
        "normal_voice_status"
    )


def add_inquiry_voice(
    instance_id
):

    append_voice_to_widget(
        f"ed_voice_{instance_id}",
        f"ed_text_{instance_id}",
        f"ed_state_{instance_id}"
    )


# ============================================================
# AUTOMATIC INDEX NUMBERING
# ============================================================

def get_display_index_name(
    instances,
    current_instance
):

    index_name = current_instance[
        "name"
    ]

    current_id = current_instance[
        "id"
    ]

    # Only these two indexes are automatically numbered.
    if index_name not in NUMBERABLE_INDEXES:

        return index_name

    matching_instances = [
        instance
        for instance in instances
        if instance["name"] == index_name
    ]

    # If selected only once, keep original name.
    if len(matching_instances) == 1:

        return index_name

    # If selected more than once, determine
    # this instance's occurrence number.
    for number, instance in enumerate(
        matching_instances,
        start=1
    ):

        if instance["id"] == current_id:

            return (
                f"{index_name} No. {number}"
            )

    return index_name


# ============================================================
# SELECT E&D INDEX
# ============================================================

def select_ed_index():

    selected = st.session_state.get(
        "ed_index_selector",
        ""
    )

    if (
        selected
        and selected != "-- Select an index --"
    ):

        instance_id = uuid.uuid4().hex[:10]

        instance = {

            "id": instance_id,

            "name": selected,

            "text": "",

            "documents": [],

            "other_document": "",

        }

        st.session_state.ed_index_instances.append(
            instance
        )

        # Reset selector.
        # IMPORTANT:
        # We DO NOT remove the selected index.
        st.session_state.ed_index_selector = (
            "-- Select an index --"
        )


# ============================================================
# REMOVE ONE INDEX INSTANCE
# ============================================================

def remove_ed_instance(
    instance_id
):

    st.session_state.ed_index_instances = [
        instance
        for instance
        in st.session_state.ed_index_instances
        if instance["id"] != instance_id
    ]

    keys_to_remove = [
        f"ed_text_{instance_id}",
        f"ed_voice_{instance_id}",
        f"ed_state_{instance_id}",
        f"ed_other_{instance_id}",
    ]

    for document in DOCUMENT_OPTIONS:

        keys_to_remove.append(
            f"ed_doc_{instance_id}_{document}"
        )

    committee_roles = [
        "Convener of Inquiry",
        "Member 1",
        "Member 2",
        "Departmental Representative",
    ]

    for role in committee_roles:

        keys_to_remove.extend([
            (
                f"ed_committee_"
                f"{instance_id}_"
                f"{role}_erp"
            ),
            (
                f"ed_committee_"
                f"{instance_id}_"
                f"{role}_name"
            ),
            (
                f"ed_committee_"
                f"{instance_id}_"
                f"{role}_designation"
            ),
        ])

    for key in keys_to_remove:

        st.session_state.pop(
            key,
            None
        )


# ============================================================
# RESET E&D INQUIRY
# ============================================================

def reset_inquiry():

    st.session_state.ed_index_instances = []

    st.session_state.ed_index_selector = (
        "-- Select an index --"
    )

    st.session_state.generated_draft = ""

    keys_to_remove = []

    for key in list(
        st.session_state.keys()
    ):

        if (
            key.startswith("ed_text_")
            or key.startswith("ed_voice_")
            or key.startswith("ed_state_")
            or key.startswith("ed_doc_")
            or key.startswith("ed_other_")
            or key.startswith("ed_committee_")
        ):

            keys_to_remove.append(
                key
            )

    for key in keys_to_remove:

        st.session_state.pop(
            key,
            None
        )


# ============================================================
# BUILD SELECTED INDEX MANIFEST
# ============================================================

def build_selected_index_manifest():

    instances = st.session_state.get(
        "ed_index_instances",
        []
    )

    if not instances:

        return "(No E&D indexes selected.)"

    manifest = []

    for number, instance in enumerate(
        instances,
        start=1
    ):

        display_name = get_display_index_name(
            instances,
            instance
        )

        manifest.append(
            f"{number}. {display_name}"
        )

    return "\n".join(
        manifest
    )


# ============================================================
# BUILD E&D INFORMATION
# ============================================================

def build_ed_information():

    instances = st.session_state.get(
        "ed_index_instances",
        []
    )

    information_parts = []

    for number, instance in enumerate(
        instances,
        start=1
    ):

        instance_id = instance[
            "id"
        ]

        index_name = instance[
            "name"
        ]

        display_name = get_display_index_name(
            instances,
            instance
        )

        # ----------------------------------------------------
        # DOCUMENTS RECORDED
        # ----------------------------------------------------

        if index_name == "Documents Recorded":

            selected_documents = []

            for document in DOCUMENT_OPTIONS:

                key = (
                    f"ed_doc_"
                    f"{instance_id}_"
                    f"{document}"
                )

                if st.session_state.get(
                    key,
                    False
                ):

                    selected_documents.append(
                        document
                    )

            other_document = st.session_state.get(
                f"ed_other_{instance_id}",
                ""
            ).strip()

            if other_document:

                selected_documents.append(
                    f"Other: {other_document}"
                )

            document_text = ", ".join(
                selected_documents
            )

            information_parts.append(
                f"INDEX INSTANCE {number}\n"
                f"INDEX: {display_name}\n"
                f"CONTENT:\n"
                f"{document_text}"
            )

        # ----------------------------------------------------
        # INQUIRY COMMITTEE
        # ----------------------------------------------------

        elif index_name == "Inquiry Committee":

            roles = [
                "Convener of Inquiry",
                "Member 1",
                "Member 2",
                "Departmental Representative",
            ]

            committee_lines = []

            for role in roles:

                erp = st.session_state.get(
                    (
                        f"ed_committee_"
                        f"{instance_id}_"
                        f"{role}_erp"
                    ),
                    ""
                ).strip()

                name = st.session_state.get(
                    (
                        f"ed_committee_"
                        f"{instance_id}_"
                        f"{role}_name"
                    ),
                    ""
                ).strip()

                designation = st.session_state.get(
                    (
                        f"ed_committee_"
                        f"{instance_id}_"
                        f"{role}_designation"
                    ),
                    ""
                ).strip()

                committee_lines.append(
                    f"{role}: "
                    f"ERP#={erp}; "
                    f"Name={name}; "
                    f"Designation={designation}"
                )

            information_parts.append(
                f"INDEX INSTANCE {number}\n"
                f"INDEX: {display_name}\n"
                f"CONTENT:\n"
                + "\n".join(
                    committee_lines
                )
            )

        # ----------------------------------------------------
        # NORMAL INDEX
        # ----------------------------------------------------

        else:

            text = st.session_state.get(
                f"ed_text_{instance_id}",
                ""
            ).strip()

            information_parts.append(
                f"INDEX INSTANCE {number}\n"
                f"INDEX: {display_name}\n"
                f"CONTENT:\n"
                f"{text}"
            )

    return "\n\n".join(
        information_parts
    )


# ============================================================
# GET FIRST INDEX VALUE
# ============================================================

def get_first_index_value(
    index_name
):

    instances = st.session_state.get(
        "ed_index_instances",
        []
    )

    for instance in instances:

        if instance["name"] == index_name:

            instance_id = instance[
                "id"
            ]

            text = st.session_state.get(
                f"ed_text_{instance_id}",
                ""
            )

            if text.strip():

                return text.strip()

    return ""


# ============================================================
# LANGUAGE
# ============================================================

def language_instruction(
    output_language
):

    if output_language == "Same as input":

        return (
            "Write the final document in the "
            "same language used by the user."
        )

    return (
        f"Write the final document in "
        f"{output_language}."
    )


# ============================================================
# PROMPT
# ============================================================

def build_prompt(
    document_type,
    inquiry_type,
    output_language,
    tone,
    recipient,
    sender,
    subject,
    information,
    selected_index_manifest=""
):

    language_rule = language_instruction(
        output_language
    )

    # ========================================================
    # E&D
    # ========================================================

    if (
        document_type == "Inquiry"
        and inquiry_type == "E&D Inquiry"
    ):

        return f"""
You are an expert professional departmental
inquiry document drafting specialist.

Prepare an official E&D Inquiry Report based
ONLY on the information supplied by the user.

============================================================
ABSOLUTE INDEX CONTROL
============================================================

The following list is the COMPLETE list of indexes
selected by the user:

{selected_index_manifest}

THIS LIST IS AUTHORITATIVE.

YOU MUST NOT ADD ANY INDEX THAT IS NOT IN THIS LIST.

Do NOT automatically add standard inquiry sections.

Do NOT assume that a departmental inquiry must contain
Witness Statements, Documentary Evidence, Findings,
Conclusion, Recommendations, Defence, Committee,
Documents Recorded, or any other section unless that
index appears in the selected list above.

If an index is NOT in the selected list:

- Do not create its heading.
- Do not create its section.
- Do not mention it.
- Do not say it was not provided.
- Do not say it was not recorded.
- Do not write a remark about it.
- Do not create a placeholder for it.

For example, if the user did NOT select:

Statements of Witnesses / Officials

then you MUST NOT write:

"No witness statement was recorded."

You must completely omit that index.

============================================================
REPEATED INDEXES
============================================================

The user may select an index more than once.

The system automatically numbers certain repeated indexes.

For example, if the selected list contains:

Statement of the Accused

only once, use exactly:

Statement of the Accused

If it appears twice, use:

Statement of the Accused No. 1
Statement of the Accused No. 2

If it appears three times, use:

Statement of the Accused No. 1
Statement of the Accused No. 2
Statement of the Accused No. 3

The same rule applies to:

Questions / Answers with the Accused

If selected twice:

Questions / Answers with the Accused No. 1
Questions / Answers with the Accused No. 2

DO NOT invent another numbering scheme.

DO NOT use "Accused No. 1" as a separate index.

DO NOT create an "Add Accused" section.

============================================================
ORDER
============================================================

Preserve the exact order of the selected indexes.

============================================================
HEADINGS
============================================================

Use the supplied selected index names as headings.

Do not rename indexes.

Section headings must be bold and underlined.

============================================================
FACTUAL ACCURACY
============================================================

Use ONLY information supplied by the user.

Do not invent:

- names
- ERP numbers
- dates
- allegations
- statements
- witnesses
- evidence
- findings
- conclusions
- recommendations
- documents
- events

Do not fill gaps using assumptions.

============================================================
IMPORTANT
============================================================

The AI must NOT "complete" the inquiry report
using a standard template.

The user's selected indexes control the structure.

Selected indexes = allowed sections.

Unselected indexes = forbidden sections.

============================================================
LANGUAGE
============================================================

{language_rule}

============================================================
TONE
============================================================

{tone}

============================================================
USER INFORMATION
============================================================

{information}

============================================================
FINAL VALIDATION
============================================================

Before returning the final report:

1. Compare every heading against the selected index list.
2. Remove every unselected section.
3. Do not add remarks about unselected sections.
4. Preserve repeated numbered indexes.
5. Preserve the selected order.
6. Do not invent facts.

Return only the professional E&D Inquiry Report.
""".strip()

    # ========================================================
    # FFI
    # ========================================================

    if (
        document_type == "Inquiry"
        and inquiry_type == "FFI Inquiry"
    ):

        return f"""
Prepare a professional FFI Inquiry document.

{language_rule}

Tone:

{tone}

Information supplied by the user:

{information}

Do not invent facts or information.
""".strip()

    # ========================================================
    # EMAIL / LETTER
    # ========================================================

    prompt = f"""
You are an expert professional document writer.

Create a high-quality {document_type}.

{language_rule}

Tone:
{tone}

Recipient:
{recipient}

Sender:
{sender}

Subject:
{subject}

Information supplied by the user:

{information}

Requirements:

- Preserve the facts supplied by the user.
- Do not invent important facts.
- Correct grammar and spelling.
- Use professional language.
- Make the document clear and logically organized.
- Do not add unnecessary information.
"""

    if document_type == "Email":

        prompt += """
For an email:

- Include an appropriate subject.
- Use an appropriate salutation.
- Write a professional body.
- Use an appropriate closing.
"""

    elif document_type == "Letter":

        prompt += """
For a letter:

- Use a formal letter structure.
- Include an appropriate subject.
- Use appropriate salutation.
- Use an appropriate closing.
"""

    return prompt.strip()


# ============================================================
# GROQ
# ============================================================

def call_groq(
    api_key,
    prompt
):

    if not api_key:

        raise ValueError(
            "Groq API key is not configured."
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
                        "writer and official departmental "
                        "inquiry document drafting specialist."
                    ),
                },
                {
                    "role": "user",
                    "content": prompt,
                },
            ],
            temperature=0.4,
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
            f"Groq request failed: {e}"
        )


# ============================================================
# GEMINI
# ============================================================

def call_gemini(
    api_key,
    prompt
):

    if not api_key:

        raise ValueError(
            "Gemini API key is not configured."
        )

    url = (
        GEMINI_URL
        + "?key="
        + api_key
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

    try:

        response = requests.post(
            url,
            json=payload,
            timeout=60
        )

        response.raise_for_status()

        data = response.json()

        candidates = data.get(
            "candidates",
            []
        )

        if not candidates:

            raise RuntimeError(
                "Gemini returned no candidates."
            )

        parts = (
            candidates[0]
            .get("content", {})
            .get("parts", [])
        )

        if not parts:

            raise RuntimeError(
                "Gemini returned no text."
            )

        return parts[0].get(
            "text",
            ""
        ).strip()

    except requests.RequestException as e:

        raise RuntimeError(
            f"Gemini request failed: {e}"
        )


# ============================================================
# E&D HEADER
# ============================================================

def add_inquiry_header(
    draft,
    reference_number
):

    date_text = datetime.now().strftime(
        "%d %B %Y"
    )

    lines = draft.splitlines()

    cleaned_lines = []

    for line in lines:

        stripped = line.strip()

        lower = stripped.lower()

        if lower.startswith(
            "inquiry reference no:"
        ):

            continue

        if lower.startswith(
            "date:"
        ):

            continue

        cleaned_lines.append(
            line
        )

    body = "\n".join(
        cleaned_lines
    ).strip()

    header = (
        "DEPARTMENTAL INQUIRY REPORT\n\n"
        f"Inquiry Reference No.: "
        f"{reference_number}\n"
        f"Date: {date_text}\n\n"
    )

    return (
        header
        + body
    )


# ============================================================
# PDF
# ============================================================

def clean_pdf_text(
    text
):

    replacements = {
        "–": "-",
        "—": "-",
        "’": "'",
        "‘": "'",
        "“": '"',
        "”": '"',
        "•": "-",
        "…": "...",
        "→": "->",
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
            "ignore"
        )
        .decode(
            "latin-1"
        )
    )


def wrap_pdf_line(
    pdf,
    text,
    max_width=180
):

    words = text.split()

    lines = []

    current = ""

    for word in words:

        test = (
            word
            if not current
            else current + " " + word
        )

        if pdf.get_string_width(
            test
        ) <= max_width:

            current = test

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


def export_pdf(
    text
):

    pdf = FPDF()

    pdf.set_auto_page_break(
        auto=True,
        margin=15
    )

    pdf.add_page()

    pdf.set_font(
        "Arial",
        size=11
    )

    text = clean_pdf_text(
        text
    )

    for paragraph in text.split("\n"):

        if not paragraph.strip():

            pdf.ln(5)

            continue

        lines = wrap_pdf_line(
            pdf,
            paragraph
        )

        for line in lines:

            pdf.multi_cell(
                0,
                7,
                line
            )

    return bytes(
        pdf.output()
    )


# ============================================================
# DOCX
# ============================================================

def export_docx(
    text
):

    document = Document()

    section = document.sections[0]

    section.top_margin = Inches(
        0.7
    )

    section.bottom_margin = Inches(
        0.7
    )

    section.left_margin = Inches(
        0.8
    )

    section.right_margin = Inches(
        0.8
    )

    for line in text.splitlines():

        paragraph = document.add_paragraph(
            line
        )

        paragraph.paragraph_format.space_after = (
            Inches(0.05)
        )

    output = io.BytesIO()

    document.save(
        output
    )

    return output.getvalue()


# ============================================================
# TXT
# ============================================================

def export_txt(
    text
):

    return text.encode(
        "utf-8"
    )


# ============================================================
# PNG
# ============================================================

def export_png(
    text
):

    try:

        font_path = (
            "/usr/share/fonts/truetype/"
            "dejavu/DejaVuSans.ttf"
        )

        font = ImageFont.truetype(
            font_path,
            24
        )

    except Exception:

        font = ImageFont.load_default()

    lines = text.splitlines()

    if not lines:

        lines = [""]

    width = 1600

    line_height = 38

    height = max(
        400,
        len(lines) * line_height + 80
    )

    image = Image.new(
        "RGB",
        (width, height),
        "white"
    )

    draw = ImageDraw.Draw(
        image
    )

    y = 30

    for line in lines:

        draw.text(
            (30, y),
            line,
            fill="black",
            font=font
        )

        y += line_height

    output = io.BytesIO()

    image.save(
        output,
        format="PNG"
    )

    return output.getvalue()


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.title(
        "📝 DraftForge"
    )

    st.markdown(
        "### AI Document Composer"
    )

    provider = st.selectbox(
        "AI Provider",
        ["Groq", "Gemini"],
        key="provider"
    )

    output_language = st.selectbox(
        "Output Language",
        OUTPUT_LANGUAGES
    )

    tone = st.selectbox(
        "Tone",
        TONES
    )

    st.divider()

    st.subheader(
        "Draft History"
    )

    history = get_history()

    if not history:

        st.caption(
            "No saved drafts yet."
        )

    else:

        for row in history[:15]:

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
                f"{title} — {created_at}"
            ):

                st.caption(
                    f"ID: {draft_id}"
                )

                if inquiry_type_history:

                    st.caption(
                        inquiry_type_history
                    )

                st.text_area(
                    "Draft",
                    row[8],
                    height=180,
                    key=f"history_{draft_id}"
                )


# ============================================================
# MAIN APPLICATION
# ============================================================

st.title(
    "📝 DraftForge — AI Document Composer"
)

st.caption(
    "Create professional emails, letters and "
    "departmental inquiry documents using text "
    "or voice input."
)


# ============================================================
# DOCUMENT TYPE
# ============================================================

document_type = st.selectbox(
    "Document Type",
    DOCUMENT_TYPES
)


# ============================================================
# EMAIL
# ============================================================

if document_type == "Email":

    st.subheader(
        "Email Details"
    )

    recipient = st.text_input(
        "Recipient",
        key="email_recipient"
    )

    sender = st.text_input(
        "Sender",
        key="email_sender"
    )

    subject = st.text_input(
        "Subject",
        key="email_subject"
    )

    st.markdown(
        "**Information / Instructions**"
    )

    st.text_area(
        "Type your information",
        height=220,
        key="normal_information",
        label_visibility="collapsed"
    )

    st.audio_input(
        "🎙️ Record Voice",
        key="normal_voice"
    )

    st.button(
        "➕ Add Voice to Text Box",
        on_click=add_normal_voice
    )

    if st.session_state.normal_voice_status:

        st.info(
            st.session_state.normal_voice_status
        )


# ============================================================
# LETTER
# ============================================================

elif document_type == "Letter":

    st.subheader(
        "Letter Details"
    )

    recipient = st.text_input(
        "Recipient",
        key="letter_recipient"
    )

    sender = st.text_input(
        "Sender",
        key="letter_sender"
    )

    subject = st.text_input(
        "Subject",
        key="letter_subject"
    )

    st.markdown(
        "**Information / Instructions**"
    )

    st.text_area(
        "Type your information",
        height=220,
        key="normal_information",
        label_visibility="collapsed"
    )

    st.audio_input(
        "🎙️ Record Voice",
        key="normal_voice"
    )

    st.button(
        "➕ Add Voice to Text Box",
        on_click=add_normal_voice
    )

    if st.session_state.normal_voice_status:

        st.info(
            st.session_state.normal_voice_status
        )


# ============================================================
# INQUIRY
# ============================================================

else:

    inquiry_type = st.selectbox(
        "Inquiry Type",
        INQUIRY_TYPES
    )

    # ========================================================
    # FFI
    # ========================================================

    if inquiry_type == "FFI Inquiry":

        st.info(
            "🚧 FFI Inquiry format is currently "
            "under construction / under process."
        )

    # ========================================================
    # E&D
    # ========================================================

    else:

        st.subheader(
            "E&D Inquiry"
        )

        st.info(
            "Select the required indexes. "
            "You can select the same index multiple times."
        )

        # ----------------------------------------------------
        # INDEX DROPDOWN
        # ----------------------------------------------------

        st.selectbox(
            "Select E&D Index",
            [
                "-- Select an index --"
            ] + ED_INDEXES,
            key="ed_index_selector",
            on_change=select_ed_index
        )

        if st.session_state.ed_index_instances:

            st.button(
                "🔄 Reset E&D Inquiry",
                on_click=reset_inquiry
            )

        st.divider()

        # ----------------------------------------------------
        # SELECTED INDEXES
        # ----------------------------------------------------

        for instance in list(
            st.session_state.ed_index_instances
        ):

            instance_id = instance[
                "id"
            ]

            index_name = instance[
                "name"
            ]

            display_index_name = (
                get_display_index_name(
                    st.session_state.ed_index_instances,
                    instance
                )
            )

            # ================================================
            # HEADING
            # ================================================

            st.markdown(
                f"**<u>{display_index_name}</u>**",
                unsafe_allow_html=True
            )

            # ================================================
            # REFERENCE NUMBER
            # ================================================

            if index_name == (
                "Inquiry Reference No."
            ):

                st.text_input(
                    "Enter Inquiry Reference Number",
                    key=f"ed_text_{instance_id}",
                    label_visibility="collapsed"
                )

            # ================================================
            # DOCUMENTS RECORDED
            # ================================================

            elif index_name == (
                "Documents Recorded"
            ):

                st.write(
                    "Select the documents recorded:"
                )

                columns = st.columns(2)

                for number, document in enumerate(
                    DOCUMENT_OPTIONS
                ):

                    key = (
                        f"ed_doc_"
                        f"{instance_id}_"
                        f"{document}"
                    )

                    with columns[
                        number % 2
                    ]:

                        st.checkbox(
                            document,
                            key=key
                        )

                st.text_input(
                    "Other Document",
                    key=f"ed_other_{instance_id}"
                )

            # ================================================
            # INQUIRY COMMITTEE
            # ================================================

            elif index_name == (
                "Inquiry Committee"
            ):

                roles = [
                    "Convener of Inquiry",
                    "Member 1",
                    "Member 2",
                    "Departmental Representative",
                ]

                for role in roles:

                    st.markdown(
                        f"**{role}**"
                    )

                    col1, col2, col3 = st.columns(3)

                    with col1:

                        st.text_input(
                            "ERP#",
                            key=(
                                f"ed_committee_"
                                f"{instance_id}_"
                                f"{role}_erp"
                            )
                        )

                    with col2:

                        st.text_input(
                            "Name",
                            key=(
                                f"ed_committee_"
                                f"{instance_id}_"
                                f"{role}_name"
                            )
                        )

                    with col3:

                        st.text_input(
                            "Designation",
                            key=(
                                f"ed_committee_"
                                f"{instance_id}_"
                                f"{role}_designation"
                            )
                        )

            # ================================================
            # NORMAL TEXT / VOICE INDEX
            # ================================================

            else:

                st.text_area(
                    "Enter information",
                    height=180,
                    key=f"ed_text_{instance_id}",
                    label_visibility="collapsed"
                )

                st.audio_input(
                    "🎙️ Record Voice",
                    key=f"ed_voice_{instance_id}"
                )

                st.button(
                    "➕ Add Voice to Text Box",
                    key=f"ed_transcribe_{instance_id}",
                    on_click=add_inquiry_voice,
                    args=(instance_id,)
                )

                status = st.session_state.get(
                    f"ed_state_{instance_id}",
                    ""
                )

                if status:

                    st.info(
                        status
                    )

            # ================================================
            # REMOVE THIS SPECIFIC INSTANCE
            # ================================================

            st.button(
                "🗑️ Remove this index",
                key=f"ed_remove_{instance_id}",
                on_click=remove_ed_instance,
                args=(instance_id,)
            )

            st.divider()


# ============================================================
# GENERATE
# ============================================================

st.subheader(
    "Generate Document"
)

if st.button(
    "✨ Generate Document",
    type="primary",
    use_container_width=True
):

    try:

        # ====================================================
        # E&D
        # ====================================================

        if (
            document_type == "Inquiry"
            and inquiry_type == "E&D Inquiry"
        ):

            instances = (
                st.session_state.ed_index_instances
            )

            if not instances:

                st.error(
                    "Please select at least one "
                    "E&D inquiry index."
                )

                st.stop()

            information = (
                build_ed_information()
            )

            selected_manifest = (
                build_selected_index_manifest()
            )

            reference_number = (
                get_first_index_value(
                    "Inquiry Reference No."
                )
            )

            selected_subject = (
                get_first_index_value(
                    "Subject"
                )
            )

            prompt = build_prompt(
                document_type=document_type,
                inquiry_type=inquiry_type,
                output_language=output_language,
                tone=tone,
                recipient="",
                sender="",
                subject=selected_subject,
                information=information,
                selected_index_manifest=selected_manifest
            )

        # ====================================================
        # FFI
        # ====================================================

        elif (
            document_type == "Inquiry"
            and inquiry_type == "FFI Inquiry"
        ):

            st.warning(
                "FFI Inquiry is currently under construction."
            )

            st.stop()

        # ====================================================
        # EMAIL / LETTER
        # ====================================================

        else:

            if document_type == "Email":

                recipient = st.session_state.get(
                    "email_recipient",
                    ""
                )

                sender = st.session_state.get(
                    "email_sender",
                    ""
                )

                subject = st.session_state.get(
                    "email_subject",
                    ""
                )

            else:

                recipient = st.session_state.get(
                    "letter_recipient",
                    ""
                )

                sender = st.session_state.get(
                    "letter_sender",
                    ""
                )

                subject = st.session_state.get(
                    "letter_subject",
                    ""
                )

            information = st.session_state.get(
                "normal_information",
                ""
            )

            if not information.strip():

                st.warning(
                    "Please enter information "
                    "or add voice input."
                )

                st.stop()

            prompt = build_prompt(
                document_type=document_type,
                inquiry_type="",
                output_language=output_language,
                tone=tone,
                recipient=recipient,
                sender=sender,
                subject=subject,
                information=information
            )

            reference_number = ""

        # ====================================================
        # API KEY
        # ====================================================

        api_key = get_api_key(
            provider
        )

        if not api_key:

            st.error(
                f"{provider} API key is not configured."
            )

            st.stop()

        # ====================================================
        # AI GENERATION
        # ====================================================

        with st.spinner(
            "Generating professional document..."
        ):

            if provider == "Groq":

                draft = call_groq(
                    api_key,
                    prompt
                )

            else:

                draft = call_gemini(
                    api_key,
                    prompt
                )

        # ====================================================
        # E&D HEADER
        # ====================================================

        if (
            document_type == "Inquiry"
            and inquiry_type == "E&D Inquiry"
        ):

            draft = add_inquiry_header(
                draft,
                (
                    reference_number
                    if reference_number
                    else "Not Provided"
                )
            )

        # ====================================================
        # SAVE GENERATED DRAFT
        # ====================================================

        st.session_state.generated_draft = (
            draft
        )

        save_draft(
            document_type=document_type,

            inquiry_type=(
                inquiry_type
                if document_type == "Inquiry"
                else ""
            ),

            tone=tone,

            recipient=(
                recipient
                if document_type != "Inquiry"
                else ""
            ),

            sender=(
                sender
                if document_type != "Inquiry"
                else ""
            ),

            subject=(
                selected_subject
                if (
                    document_type == "Inquiry"
                    and inquiry_type == "E&D Inquiry"
                )
                else (
                    subject
                    if document_type != "Inquiry"
                    else ""
                )
            ),

            key_points=information,

            draft=draft
        )

        st.success(
            "Document generated successfully."
        )

    except Exception as e:

        st.error(
            f"Error: {e}"
        )


# ============================================================
# GENERATED DOCUMENT
# ============================================================

if st.session_state.generated_draft:

    st.divider()

    st.subheader(
        "Generated Document"
    )

    st.text_area(
        "Generated Draft",
        value=st.session_state.generated_draft,
        height=600
    )

    # ========================================================
    # DOWNLOAD BUTTONS
    # ========================================================

    col1, col2, col3, col4 = st.columns(4)

    with col1:

        st.download_button(
            "📄 Download DOCX",
            data=export_docx(
                st.session_state.generated_draft
            ),
            file_name="DraftForge_Document.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            use_container_width=True
        )

    with col2:

        st.download_button(
            "📕 Download PDF",
            data=export_pdf(
                st.session_state.generated_draft
            ),
            file_name="DraftForge_Document.pdf",
            mime="application/pdf",
            use_container_width=True
        )

    with col3:

        st.download_button(
            "📝 Download TXT",
            data=export_txt(
                st.session_state.generated_draft
            ),
            file_name="DraftForge_Document.txt",
            mime="text/plain",
            use_container_width=True
        )

    with col4:

        st.download_button(
            "🖼️ Download PNG",
            data=export_png(
                st.session_state.generated_draft
            ),
            file_name="DraftForge_Document.png",
            mime="image/png",
            use_container_width=True
        )
