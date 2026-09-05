import io
import os
import re
import json
import hashlib
from datetime import datetime

import requests
import streamlit as st
from groq import Groq

from docx import Document
from docx.shared import Pt

from fpdf import FPDF

from PIL import Image, ImageDraw, ImageFont


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="DraftForge — AI Document Composer",
    page_icon="📝",
    layout="wide",
)


# ============================================================
# AI MODELS
# ============================================================

GROQ_MODEL = "openai/gpt-oss-120b"
WHISPER_MODEL = "whisper-large-v3"
GEMINI_MODEL = "gemini-2.0-flash"


# ============================================================
# PROFILE FILE
# ============================================================

PROFILE_FILE = "user_profile.json"


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
# SPECIAL STRUCTURED INDEXES
# ============================================================

SPECIAL_INDEXES = {
    "Documents Recorded",
    "Inquiry Committee",
}


# ============================================================
# Q&A INDEX
# ============================================================

QA_INDEX_PREFIX = "Questions / Answers with the Accused"


# ============================================================
# DOCUMENTS RECORDED
# ============================================================

DOCUMENTS_RECORDED = [
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
# INQUIRY COMMITTEE
# ============================================================

COMMITTEE_ROLES = [
    "Convener of Inquiry",
    "Member 1",
    "Member 2",
    "Departmental Representative",
]


# ============================================================
# CSS
# ============================================================

st.markdown(
    """
    <style>

    .main-title {
        font-size: 2.2rem;
        font-weight: 700;
        margin-bottom: 0;
    }

    .subtitle {
        font-size: 1rem;
        color: #666;
        margin-top: 0;
    }

    .composer-box {
        padding: 18px;
        border-radius: 12px;
        border: 1px solid #ddd;
        background-color: #fafafa;
        margin-bottom: 18px;
    }

    .voice-heading {
        font-weight: 600;
        margin-bottom: 5px;
    }

    .voice-help {
        font-size: 0.85rem;
        color: #666;
        margin-bottom: 8px;
    }

    .index-card {
        padding: 15px;
        border-radius: 10px;
        border: 1px solid #ddd;
        background-color: #fafafa;
        margin-bottom: 15px;
    }

    .warning-box {
        padding: 12px;
        border-radius: 8px;
        border: 1px solid #e0b000;
        background-color: #fff8d8;
    }

    .qa-table {
        width: 100%;
        border-collapse: collapse;
        margin-top: 10px;
        margin-bottom: 20px;
    }

    .qa-table th,
    .qa-table td {
        border: 1px solid #999;
        padding: 10px;
        vertical-align: top;
        text-align: left;
    }

    .qa-table th {
        font-weight: 700;
    }

    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# USER PROFILE
# ============================================================

def load_user_profile():

    default = {
        "name": "",
        "designation": "",
        "contact_no": "",
        "current_station": "",
    }

    try:

        if os.path.exists(PROFILE_FILE):

            with open(
                PROFILE_FILE,
                "r",
                encoding="utf-8",
            ) as f:

                data = json.load(f)

            if isinstance(data, dict):

                for key in default:

                    if key in data:

                        default[key] = data.get(
                            key,
                            "",
                        )

    except Exception:

        pass

    return default


def save_user_profile(profile):

    try:

        with open(
            PROFILE_FILE,
            "w",
            encoding="utf-8",
        ) as f:

            json.dump(
                profile,
                f,
                ensure_ascii=False,
                indent=2,
            )

        return True

    except Exception:

        return False


def build_sender_signature():

    profile = st.session_state.get(
        "user_profile",
        {},
    )

    lines = []

    if profile.get("name"):
        lines.append(profile["name"])

    if profile.get("designation"):
        lines.append(profile["designation"])

    if profile.get("contact_no"):
        lines.append(
            f"Contact No.: {profile['contact_no']}"
        )

    if profile.get("current_station"):
        lines.append(
            f"Current Station: {profile['current_station']}"
        )

    return "\n".join(lines)


def append_profile_signature(draft):

    signature = build_sender_signature()

    if not signature:
        return draft

    return (
        draft.rstrip()
        + "\n\n"
        + signature.strip()
    )


def profile_is_complete():

    profile = st.session_state.get(
        "user_profile",
        {},
    )

    required = [
        "name",
        "designation",
        "contact_no",
        "current_station",
    ]

    return all(
        str(
            profile.get(
                item,
                "",
            )
        ).strip()
        for item in required
    )


# ============================================================
# SESSION STATE
# ============================================================

def initialize_state():

    defaults = {

        "document_type": "Email",

        "inquiry_type": "E&D Inquiry",

        "generated_draft": "",

        "editable_draft": "",

        "email_input": "",

        "letter_input": "",

        "email_audio_seen": "",

        "letter_audio_seen": "",

        "ed_instances": [],

        "history": [],

        "user_profile": load_user_profile(),

        "edit_instruction": "",

        "documents_recorded_selected": [],

        "inquiry_committee_data": {},

    }

    for key, value in defaults.items():

        if key not in st.session_state:

            st.session_state[key] = value

    profile = st.session_state.user_profile

    if "profile_name" not in st.session_state:

        st.session_state.profile_name = profile.get(
            "name",
            "",
        )

    if "profile_designation" not in st.session_state:

        st.session_state.profile_designation = profile.get(
            "designation",
            "",
        )

    if "profile_contact" not in st.session_state:

        st.session_state.profile_contact = profile.get(
            "contact_no",
            "",
        )

    if "profile_station" not in st.session_state:

        st.session_state.profile_station = profile.get(
            "current_station",
            "",
        )


initialize_state()


# ============================================================
# API KEY
# ============================================================

def get_secret(
    name,
    default=None,
):

    try:

        value = st.secrets.get(name)

        if value:
            return value

    except Exception:

        pass

    return os.getenv(
        name,
        default,
    )


# ============================================================
# GROQ CLIENT
# ============================================================

@st.cache_resource
def get_groq_client():

    api_key = get_secret(
        "GROQ_API_KEY"
    )

    if not api_key:
        return None

    try:

        return Groq(
            api_key=api_key
        )

    except Exception:

        return None


# ============================================================
# TEXT HELPERS
# ============================================================

def append_text(
    old_text,
    new_text,
):

    old_text = (
        old_text or ""
    ).strip()

    new_text = (
        new_text or ""
    ).strip()

    if not new_text:
        return old_text

    if not old_text:
        return new_text

    return (
        old_text
        + "\n"
        + new_text
    )


def audio_signature(audio):

    if audio is None:
        return ""

    try:

        audio_bytes = audio.getvalue()

        return hashlib.md5(
            audio_bytes
        ).hexdigest()

    except Exception:

        return ""


def clean_markup(text):

    if not text:
        return ""

    text = re.sub(
        r"```(?:text|markdown)?",
        "",
        text,
        flags=re.IGNORECASE,
    )

    text = text.replace(
        "```",
        "",
    )

    return text.strip()


# ============================================================
# VOICE TRANSCRIPTION
# ============================================================

def transcribe_audio(audio_file):

    client = get_groq_client()

    if client is None:

        return (
            None,
            "Groq API key is not configured.",
        )

    try:

        audio_bytes = audio_file.getvalue()

        if not audio_bytes:

            return (
                None,
                "No audio was recorded.",
            )

        audio_buffer = io.BytesIO(
            audio_bytes
        )

        audio_buffer.name = (
            "voice_input.wav"
        )

        transcription = (
            client.audio.transcriptions.create(
                file=audio_buffer,
                model=WHISPER_MODEL,
                response_format="text",
            )
        )

        if hasattr(
            transcription,
            "text",
        ):

            text = transcription.text

        else:

            text = str(
                transcription
            )

        return (
            text.strip(),
            None,
        )

    except Exception as e:

        return (
            None,
            f"Voice transcription failed: {e}",
        )


# ============================================================
# AI GENERATION
# ============================================================

def generate_ai(
    system_prompt,
    user_prompt,
):

    groq_client = get_groq_client()

    # --------------------------------------------------------
    # GROQ
    # --------------------------------------------------------

    if groq_client:

        try:

            response = (
                groq_client.chat.completions.create(
                    model=GROQ_MODEL,
                    messages=[
                        {
                            "role": "system",
                            "content": system_prompt,
                        },
                        {
                            "role": "user",
                            "content": user_prompt,
                        },
                    ],
                    temperature=0.2,
                )
            )

            return clean_markup(
                response.choices[
                    0
                ].message.content
            )

        except Exception as groq_error:

            groq_error_text = str(
                groq_error
            )

    else:

        groq_error_text = (
            "Groq API key not configured."
        )

    # --------------------------------------------------------
    # GEMINI FALLBACK
    # --------------------------------------------------------

    gemini_key = get_secret(
        "GEMINI_API_KEY"
    )

    if gemini_key:

        try:

            url = (
                "https://generativelanguage.googleapis.com/"
                "v1beta/models/"
                f"{GEMINI_MODEL}:generateContent"
                f"?key={gemini_key}"
            )

            payload = {

                "system_instruction": {
                    "parts": [
                        {
                            "text": system_prompt
                        }
                    ]
                },

                "contents": [
                    {
                        "parts": [
                            {
                                "text": user_prompt
                            }
                        ]
                    }
                ],

                "generationConfig": {
                    "temperature": 0.2
                },
            }

            response = requests.post(
                url,
                json=payload,
                timeout=120,
            )

            response.raise_for_status()

            data = response.json()

            text = (
                data[
                    "candidates"
                ][0][
                    "content"
                ][
                    "parts"
                ][0]["text"]
            )

            return clean_markup(
                text
            )

        except Exception as gemini_error:

            raise RuntimeError(
                "Both AI services failed.\n\n"
                f"Groq: {groq_error_text}\n"
                f"Gemini: {gemini_error}"
            )

    raise RuntimeError(
        "No AI API is available. "
        "Please configure GROQ_API_KEY "
        "or GEMINI_API_KEY."
    )


# ============================================================
# AI RULES
# ============================================================

LANGUAGE_RULES = """
Write in professional, formal and clear official English.

Correct:
- spelling
- grammar
- punctuation
- obvious transcription mistakes
- sentence structure

Preserve the user's intended meaning.

DO NOT invent:
- names
- dates
- allegations
- evidence
- witnesses
- findings
- recommendations
- reference numbers
- events
- facts

If information is missing, do not fabricate it.

For statements of accused persons or witnesses,
preserve the actual meaning of their statements.

Do not introduce unsupported facts.
"""


NORMAL_SYSTEM_PROMPT = f"""
You are an AI assistant for drafting official correspondence.

{LANGUAGE_RULES}

Prepare professional official documents suitable
for government/organizational communication.

Do not add a sender signature because the application
will automatically add the user's official information.
"""


ED_SYSTEM_PROMPT = f"""
You are an expert assistant for preparing
departmental inquiry reports.

{LANGUAGE_RULES}

Organize the supplied information into a professional
departmental inquiry report.

Do not invent missing evidence or findings.

Where information is provided as a short note,
convert it into proper official language without
changing its meaning.

Do not add an inquiry committee or documents recorded
unless they are actually supplied.

For every "Questions / Answers with the Accused"
section, format the questions and answers as a
Markdown table using exactly two columns:

| Questions | Answers |
|---|---|
| 1. Question | Answer |
| 2. Question | Answer |

Every question and its corresponding answer must
occupy the same row.

If an answer is not supplied, do not invent one.
"""


# ============================================================
# E&D INDEX MANAGEMENT
# ============================================================

def add_ed_index(
    index_name,
):

    instance_id = (
        len(
            st.session_state.ed_instances
        )
        + 1
    )

    st.session_state.ed_instances.append(
        {
            "id": instance_id,
            "index": index_name,
            "text": "",
            "audio_seen": "",
        }
    )


def remove_ed_index(
    instance_id,
):

    st.session_state.ed_instances = [
        item
        for item in st.session_state.ed_instances
        if item["id"] != instance_id
    ]


def get_occurrence(
    index_name,
    instance_id,
):

    occurrence = 0

    for item in (
        st.session_state.ed_instances
    ):

        if item["index"] == index_name:

            occurrence += 1

            if item["id"] == instance_id:

                return occurrence

    return 1


def get_display_heading(
    index_name,
    occurrence,
):

    if occurrence <= 1:

        return index_name

    return (
        f"{index_name} No. "
        f"{occurrence}"
    )


# ============================================================
# NORMAL COMPOSER
# ============================================================

def render_composer(
    prefix,
    title,
):

    st.markdown(
        f"#### {title}"
    )

    st.caption(
        "Type naturally or tap 🎙️ to record. "
        "Voice transcription is added to the same input."
    )

    audio = st.audio_input(
        "🎙️ Record Voice",
        key=f"{prefix}_audio",
    )

    if audio is not None:

        signature = audio_signature(
            audio
        )

        seen_key = (
            f"{prefix}_audio_seen"
        )

        if (
            signature
            and signature
            != st.session_state.get(
                seen_key,
                "",
            )
        ):

            with st.spinner(
                "Transcribing voice..."
            ):

                transcript, error = (
                    transcribe_audio(
                        audio
                    )
                )

            if error:

                st.error(error)

            elif transcript:

                current = (
                    st.session_state.get(
                        f"{prefix}_input",
                        "",
                    )
                )

                st.session_state[
                    f"{prefix}_input"
                ] = append_text(
                    current,
                    transcript,
                )

                st.session_state[
                    seen_key
                ] = signature

    return st.text_area(
        "Instructions / Information",
        key=f"{prefix}_input",
        height=160,
        placeholder=(
            "Type information naturally here, "
            "or use the microphone above."
        ),
        label_visibility="collapsed",
    )


# ============================================================
# E&D NORMAL INPUT
# ============================================================

def render_ed_input(
    item,
):

    instance_id = item["id"]

    index_name = item["index"]

    occurrence = get_occurrence(
        index_name,
        instance_id,
    )

    heading = get_display_heading(
        index_name,
        occurrence,
    )

    st.markdown(
        f"""
        <div class="index-card">
        <strong>{heading}</strong>
        </div>
        """,
        unsafe_allow_html=True,
    )

    audio_key = (
        f"ed_audio_{instance_id}"
    )

    text_key = (
        f"ed_text_{instance_id}"
    )

    seen_key = (
        f"ed_audio_seen_{instance_id}"
    )

    audio = st.audio_input(
        "🎙️ Record Voice",
        key=audio_key,
    )

    if audio is not None:

        signature = audio_signature(
            audio
        )

        if (
            signature
            and signature
            != st.session_state.get(
                seen_key,
                "",
            )
        ):

            with st.spinner(
                "Transcribing voice..."
            ):

                transcript, error = (
                    transcribe_audio(
                        audio
                    )
                )

            if error:

                st.error(error)

            elif transcript:

                current = (
                    st.session_state.get(
                        text_key,
                        item.get(
                            "text",
                            "",
                        ),
                    )
                )

                st.session_state[
                    text_key
                ] = append_text(
                    current,
                    transcript,
                )

                st.session_state[
                    seen_key
                ] = signature

    text = st.text_area(
        "Information",
        key=text_key,
        height=150,
        placeholder=(
            "Type information naturally here, "
            "or use the microphone above."
        ),
    )

    item["text"] = text


# ============================================================
# DOCUMENTS RECORDED
# ============================================================

def render_documents_recorded():

    st.markdown(
        "### 📋 Documents Recorded"
    )

    st.caption(
        "Select only the documents or records "
        "actually examined or recorded."
    )

    selected = []

    columns = st.columns(2)

    for i, document in enumerate(
        DOCUMENTS_RECORDED
    ):

        with columns[
            i % 2
        ]:

            checked = st.checkbox(
                document,
                key=f"doc_recorded_{i}",
            )

            if checked:

                selected.append(
                    document
                )

    return selected


# ============================================================
# INQUIRY COMMITTEE
# ============================================================

def render_committee():

    st.markdown(
        "### 👥 Inquiry Committee"
    )

    st.caption(
        "Enter the details of the officers/members "
        "constituting the inquiry committee."
    )

    committee = {}

    for index, role in enumerate(
        COMMITTEE_ROLES
    ):

        st.markdown(
            f"#### {role}"
        )

        col1, col2, col3 = st.columns(
            3
        )

        with col1:

            erp = st.text_input(
                "ERP#",
                key=f"committee_erp_{index}",
            )

        with col2:

            name = st.text_input(
                "Name",
                key=f"committee_name_{index}",
            )

        with col3:

            designation = st.text_input(
                "Designation",
                key=f"committee_designation_{index}",
            )

        committee[role] = {
            "erp": erp.strip(),
            "name": name.strip(),
            "designation": designation.strip(),
        }

    return committee


# ============================================================
# Q&A DETECTION
# ============================================================

def is_qa_heading(
    line,
):

    normalized = (
        line.strip()
        .lower()
    )

    return (
        normalized.startswith(
            QA_INDEX_PREFIX.lower()
        )
    )


# ============================================================
# PARSE MARKDOWN TABLE
# ============================================================

def parse_markdown_table(
    lines,
    start_index,
):

    table_rows = []

    i = start_index

    while i < len(lines):

        line = lines[i].strip()

        if not line.startswith("|"):

            break

        cells = [
            cell.strip()
            for cell in line.strip(
                "|"
            ).split("|")
        ]

        # Header separator row.
        if all(
            re.fullmatch(
                r":?-+:?",
                cell.replace(
                    " ",
                    "",
                ),
            )
            for cell in cells
        ):

            i += 1
            continue

        if len(cells) >= 2:

            table_rows.append(
                (
                    cells[0],
                    cells[1],
                )
            )

        i += 1

    return (
        table_rows,
        i,
    )


# ============================================================
# RENDER DOCUMENT FOR SCREEN
# ============================================================

def render_document_markdown(
    text,
):

    if not text:

        return

    lines = text.splitlines()

    i = 0

    while i < len(lines):

        line = lines[i].strip()

        if (
            is_qa_heading(line)
            and i + 1 < len(lines)
            and lines[i + 1].strip().startswith("|")
        ):

            st.markdown(
                f"### {line}"
            )

            rows, next_index = (
                parse_markdown_table(
                    lines,
                    i + 1,
                )
            )

            if rows:

                table_markdown = (
                    "| Questions | Answers |\n"
                    "|---|---|\n"
                )

                for question, answer in rows:

                    table_markdown += (
                        f"| {question} | {answer} |\n"
                    )

                st.markdown(
                    table_markdown
                )

                i = next_index
                continue

        # Normal paragraph/heading.
        if line:

            st.markdown(
                line
            )

        else:

            st.write("")

        i += 1


# ============================================================
# GENERATE EMAIL
# ============================================================

def generate_email():

    recipient = st.session_state.get(
        "email_recipient",
        "",
    )

    subject = st.session_state.get(
        "email_subject",
        "",
    )

    instructions = st.session_state.get(
        "email_input",
        "",
    )

    prompt = f"""
Prepare a professional official email.

Recipient:
{recipient}

Subject:
{subject}

User's instructions/information:
{instructions}

Requirements:

- Use an appropriate professional email structure.
- Correct grammar and spelling.
- Preserve all facts.
- Do not invent information.
- Do not include the sender's name,
  designation, contact number or current station.
- The application will automatically append
  the sender's official information at the end.
"""

    draft = generate_ai(
        NORMAL_SYSTEM_PROMPT,
        prompt,
    )

    return append_profile_signature(
        draft
    )


# ============================================================
# GENERATE LETTER
# ============================================================

def generate_letter():

    recipient = st.session_state.get(
        "letter_recipient",
        "",
    )

    subject = st.session_state.get(
        "letter_subject",
        "",
    )

    instructions = st.session_state.get(
        "letter_input",
        "",
    )

    prompt = f"""
Prepare a professional official letter.

Recipient:
{recipient}

Subject:
{subject}

User's instructions/information:
{instructions}

Requirements:

- Use a formal official letter structure.
- Correct grammar and spelling.
- Preserve all facts.
- Do not invent information.
- Do not include the sender's name,
  designation, contact number or current station.
- The application will automatically append
  the sender's official information at the end.
"""

    draft = generate_ai(
        NORMAL_SYSTEM_PROMPT,
        prompt,
    )

    return append_profile_signature(
        draft
    )


# ============================================================
# GENERATE E&D INQUIRY
# ============================================================

def generate_ed_report(
    reference_no,
    documents_recorded,
    committee,
):

    # --------------------------------------------------------
    # NORMAL INDEXES
    # --------------------------------------------------------

    information_parts = []

    for item in (
        st.session_state.ed_instances
    ):

        index_name = item["index"]

        # Do NOT send structured indexes
        # as normal text.
        if index_name in SPECIAL_INDEXES:

            continue

        text = item.get(
            "text",
            "",
        ).strip()

        if not text:

            continue

        occurrence = get_occurrence(
            index_name,
            item["id"],
        )

        heading = get_display_heading(
            index_name,
            occurrence,
        )

        information_parts.append(
            f"{heading}:\n{text}"
        )

    inquiry_information = (
        "\n\n".join(
            information_parts
        )
    )

    # --------------------------------------------------------
    # DOCUMENTS
    # --------------------------------------------------------

    documents_text = "\n".join(
        f"- {item}"
        for item in documents_recorded
    )

    # --------------------------------------------------------
    # COMMITTEE
    # --------------------------------------------------------

    committee_parts = []

    for role in COMMITTEE_ROLES:

        details = committee.get(
            role,
            {},
        )

        erp = details.get(
            "erp",
            "",
        )

        name = details.get(
            "name",
            "",
        )

        designation = details.get(
            "designation",
            "",
        )

        if (
            erp
            or name
            or designation
        ):

            committee_parts.append(
                f"{role}: "
                f"ERP# {erp}; "
                f"Name: {name}; "
                f"Designation: {designation}"
            )

    committee_text = "\n".join(
        committee_parts
    )

    # --------------------------------------------------------
    # PROMPT
    # --------------------------------------------------------

    prompt = f"""
Prepare a professional Departmental Inquiry Report.

Inquiry Reference No.:
{reference_no}

Information supplied for the inquiry:
{inquiry_information}

Documents Recorded:
{documents_text if documents_text else "None specified."}

Inquiry Committee:
{committee_text if committee_text else "Not specified."}

IMPORTANT REQUIREMENTS:

1. Use only the information supplied.

2. Do not invent facts.

3. Correct spelling and grammar.

4. Preserve statements of the accused.

5. Preserve the meaning of witness statements.

6. Do not create allegations, evidence, findings,
   recommendations or conclusions that were not supplied.

7. Do not duplicate any section.

8. Documents Recorded must appear ONLY ONCE.

9. Inquiry Committee must appear ONLY ONCE.

10. If multiple instances of an index are supplied,
    retain each instance separately.

11. Use professional official language.

12. Include the inquiry reference number near the beginning.

13. Include the current date.

14. Do not add commentary about how the report was created.

15. QUESTIONS AND ANSWERS:

For every section titled:

Questions / Answers with the Accused
Questions / Answers with the Accused No. 2
Questions / Answers with the Accused No. 3
etc.

you MUST format the content as a Markdown table.

Use exactly:

| Questions | Answers |
|---|---|
| 1. Question | Answer |
| 2. Question | Answer |
| 3. Question | Answer |

Each question and its corresponding answer must
be in the SAME ROW.

Do not put all questions in one cell.

Do not put all answers in one cell.

If there is no answer supplied for a question,
do not invent an answer.

16. Do not use a table for any other section.

17. Return only the final inquiry report.
"""

    draft = generate_ai(
        ED_SYSTEM_PROMPT,
        prompt,
    )

    return draft


# ============================================================
# MODIFY GENERATED DOCUMENT
# ============================================================

def modify_generated_document(
    current_document,
    instruction,
):

    prompt = f"""
You are editing an already generated official document.

CURRENT DOCUMENT:
----------------
{current_document}
----------------

USER'S REQUESTED CHANGE:
------------------------
{instruction}
------------------------

Apply ONLY the requested changes.

Rules:

1. Preserve all existing facts.

2. Do not invent names, dates, evidence,
   allegations, witnesses, findings,
   recommendations or events.

3. Do not remove important information unless
   the user explicitly asks for removal.

4. Correct grammar and spelling where appropriate.

5. Maintain professional official language.

6. Keep the document structure professional.

7. If a Questions / Answers with the Accused
   section exists, preserve its table structure.

8. Every Q&A table must have exactly two columns:

| Questions | Answers |
|---|---|

9. Keep each question and corresponding answer
   in the same row.

10. Do not convert a Q&A table into ordinary paragraphs.

11. Return ONLY the revised document.
"""

    return generate_ai(
        NORMAL_SYSTEM_PROMPT,
        prompt,
    )


# ============================================================
# PDF HELPERS
# ============================================================

def pdf_clean(text):

    replacements = {

        "–": "-",

        "—": "-",

        "’": "'",

        "‘": "'",

        "“": '"',

        "”": '"',

        "•": "-",

        " ": " ",
    }

    for old, new in replacements.items():

        text = text.replace(
            old,
            new,
        )

    return text


def is_heading(
    line,
):

    line = line.strip()

    if not line:

        return False

    headings = [

        "DEPARTMENTAL INQUIRY REPORT",

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

    for heading in headings:

        if line.lower().startswith(
            heading.lower()
        ):

            return True

    return False


def wrap_pdf_line(
    pdf,
    text,
    width,
):

    words = text.split()

    if not words:

        pdf.ln(5)

        return

    current = ""

    for word in words:

        test = (
            current
            + " "
            + word
        ).strip()

        if pdf.get_string_width(
            test
        ) <= width:

            current = test

        else:

            if current:

                pdf.cell(
                    0,
                    6,
                    current,
                    ln=1,
                )

            current = word

    if current:

        pdf.cell(
            0,
            6,
            current,
            ln=1,
        )


# ============================================================
# PDF Q&A TABLE
# ============================================================

def draw_pdf_qa_table(
    pdf,
    rows,
):

    page_width = 180

    question_width = 85

    answer_width = 95

    # Header
    pdf.set_font(
        "Arial",
        "B",
        9,
    )

    pdf.cell(
        question_width,
        8,
        "Questions",
        border=1,
        align="L",
    )

    pdf.cell(
        answer_width,
        8,
        "Answers",
        border=1,
        align="L",
        ln=1,
    )

    pdf.set_font(
        "Arial",
        "",
        9,
    )

    for question, answer in rows:

        question = pdf_clean(
            question
        )

        answer = pdf_clean(
            answer
        )

        # Approximate row height.
        q_words = question.split()
        a_words = answer.split()

        q_lines = max(
            1,
            len(q_words) // 9 + 1,
        )

        a_lines = max(
            1,
            len(a_words) // 11 + 1,
        )

        row_height = max(
            6,
            max(
                q_lines,
                a_lines,
            ) * 6,
        )

        # Page break if needed.
        if (
            pdf.get_y()
            + row_height
            > 270
        ):

            pdf.add_page()

            pdf.set_font(
                "Arial",
                "B",
                9,
            )

            pdf.cell(
                question_width,
                8,
                "Questions",
                border=1,
                align="L",
            )

            pdf.cell(
                answer_width,
                8,
                "Answers",
                border=1,
                align="L",
                ln=1,
            )

            pdf.set_font(
                "Arial",
                "",
                9,
            )

        x = pdf.get_x()
        y = pdf.get_y()

        # Question cell
        pdf.multi_cell(
            question_width,
            6,
            question,
            border=1,
            align="L",
        )

        question_end_y = pdf.get_y()

        # Reset to beginning of row.
        pdf.set_xy(
            x + question_width,
            y,
        )

        pdf.multi_cell(
            answer_width,
            6,
            answer,
            border=1,
            align="L",
        )

        answer_end_y = pdf.get_y()

        final_y = max(
            question_end_y,
            answer_end_y,
        )

        pdf.set_xy(
            x,
            final_y,
        )


# ============================================================
# EXPORT PDF
# ============================================================

def export_pdf(
    text,
):

    pdf = FPDF()

    pdf.set_auto_page_break(
        auto=True,
        margin=15,
    )

    pdf.add_page()

    pdf.set_font(
        "Arial",
        size=10,
    )

    clean_text = pdf_clean(
        text
    )

    lines = clean_text.splitlines()

    i = 0

    while i < len(lines):

        line = lines[i].strip()

        # ----------------------------------------------------
        # Q&A TABLE
        # ----------------------------------------------------

        if (
            is_qa_heading(line)
            and i + 1 < len(lines)
            and lines[i + 1].strip().startswith("|")
        ):

            pdf.set_font(
                "Arial",
                "B",
                11,
            )

            pdf.ln(3)

            wrap_pdf_line(
                pdf,
                line,
                180,
            )

            rows, next_index = (
                parse_markdown_table(
                    lines,
                    i + 1,
                )
            )

            if rows:

                draw_pdf_qa_table(
                    pdf,
                    rows,
                )

            pdf.set_font(
                "Arial",
                "",
                10,
            )

            i = next_index

            continue

        # ----------------------------------------------------
        # NORMAL HEADING
        # ----------------------------------------------------

        if is_heading(line):

            pdf.set_font(
                "Arial",
                "B",
                11,
            )

            pdf.ln(2)

            wrap_pdf_line(
                pdf,
                line,
                180,
            )

            pdf.set_font(
                "Arial",
                "",
                10,
            )

        # ----------------------------------------------------
        # NORMAL TEXT
        # ----------------------------------------------------

        elif line:

            wrap_pdf_line(
                pdf,
                line,
                180,
            )

        else:

            pdf.ln(4)

        i += 1

    return bytes(
        pdf.output()
    )


# ============================================================
# DOCX Q&A TABLE
# ============================================================

def add_docx_qa_table(
    document,
    rows,
):

    table = document.add_table(
        rows=1,
        cols=2,
    )

    table.style = "Table Grid"

    hdr = table.rows[0].cells

    hdr[0].text = "Questions"
    hdr[1].text = "Answers"

    for cell in hdr:

        for paragraph in cell.paragraphs:

            for run in paragraph.runs:

                run.bold = True

    for question, answer in rows:

        cells = table.add_row().cells

        cells[0].text = question

        cells[1].text = answer

    document.add_paragraph()


# ============================================================
# EXPORT DOCX
# ============================================================

def export_docx(
    text,
):

    document = Document()

    normal_style = document.styles[
        "Normal"
    ]

    normal_style.font.name = "Arial"

    normal_style.font.size = Pt(
        11
    )

    lines = text.splitlines()

    i = 0

    while i < len(lines):

        line = lines[i].strip()

        # ----------------------------------------------------
        # Q&A TABLE
        # ----------------------------------------------------

        if (
            is_qa_heading(line)
            and i + 1 < len(lines)
            and lines[i + 1].strip().startswith("|")
        ):

            paragraph = document.add_paragraph()

            run = paragraph.add_run(
                line
            )

            run.bold = True

            rows, next_index = (
                parse_markdown_table(
                    lines,
                    i + 1,
                )
            )

            if rows:

                add_docx_qa_table(
                    document,
                    rows,
                )

            i = next_index

            continue

        # ----------------------------------------------------
        # EMPTY LINE
        # ----------------------------------------------------

        if not line:

            document.add_paragraph()

            i += 1

            continue

        # ----------------------------------------------------
        # HEADING
        # ----------------------------------------------------

        if is_heading(line):

            paragraph = document.add_paragraph()

            run = paragraph.add_run(
                line
            )

            run.bold = True

            run.font.size = Pt(
                11
            )

        # ----------------------------------------------------
        # NORMAL TEXT
        # ----------------------------------------------------

        else:

            document.add_paragraph(
                line
            )

        i += 1

    buffer = io.BytesIO()

    document.save(
        buffer
    )

    buffer.seek(0)

    return buffer.getvalue()


# ============================================================
# TXT EXPORT
# ============================================================

def export_txt(
    text,
):

    return text.encode(
        "utf-8"
    )


# ============================================================
# PNG EXPORT
# ============================================================

def export_png(
    text,
):

    font = None

    possible_fonts = [

        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",

        "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
    ]

    bold_fonts = [

        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",

        "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
    ]

    for path in possible_fonts:

        if os.path.exists(path):

            font = ImageFont.truetype(
                path,
                24,
            )

            break

    if font is None:

        font = ImageFont.load_default()

    bold_font = font

    for path in bold_fonts:

        if os.path.exists(path):

            bold_font = ImageFont.truetype(
                path,
                26,
            )

            break

    lines = text.splitlines()

    line_height = 36

    height = max(
        800,
        (
            len(lines)
            + 4
        )
        * line_height,
    )

    image = Image.new(
        "RGB",
        (
            1600,
            height,
        ),
        "white",
    )

    draw = ImageDraw.Draw(
        image
    )

    y = 50

    for line in lines:

        if is_heading(line):

            draw.text(
                (
                    60,
                    y,
                ),
                line,
                fill="black",
                font=bold_font,
            )

        else:

            draw.text(
                (
                    60,
                    y,
                ),
                line,
                fill="black",
                font=font,
            )

        y += line_height

    image = image.crop(
        (
            0,
            0,
            1600,
            min(
                height,
                y + 50,
            ),
        )
    )

    buffer = io.BytesIO()

    image.save(
        buffer,
        format="PNG",
    )

    return buffer.getvalue()


# ============================================================
# HISTORY
# ============================================================

def save_history(
    document_type,
    draft,
):

    if not draft:

        return

    entry = {

        "type": document_type,

        "date": datetime.now().strftime(
            "%d %b %Y %H:%M"
        ),

        "draft": draft,
    }

    st.session_state.history.insert(
        0,
        entry,
    )

    st.session_state.history = (
        st.session_state.history[:10]
    )


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.markdown(
        "## 📝 DraftForge"
    )

    st.caption(
        "AI Document Composer"
    )

    # --------------------------------------------------------
    # OFFICIAL INFORMATION
    # --------------------------------------------------------

    st.markdown(
        "### 👤 My Official Information"
    )

    st.caption(
        "Save your information once. It will automatically "
        "appear at the end of generated emails and letters."
    )

    st.text_input(
        "Name",
        key="profile_name",
    )

    st.text_input(
        "Designation",
        key="profile_designation",
    )

    st.text_input(
        "Contact No.",
        key="profile_contact",
    )

    st.text_input(
        "Current Station",
        key="profile_station",
    )

    if st.button(
        "💾 Save Official Information",
        use_container_width=True,
    ):

        profile = {

            "name":
                st.session_state.profile_name.strip(),

            "designation":
                st.session_state.profile_designation.strip(),

            "contact_no":
                st.session_state.profile_contact.strip(),

            "current_station":
                st.session_state.profile_station.strip(),
        }

        if save_user_profile(
            profile
        ):

            st.session_state.user_profile = (
                profile
            )

            st.success(
                "Official information saved."
            )

        else:

            st.error(
                "Could not save profile information."
            )

    st.divider()

    # --------------------------------------------------------
    # QUICK GUIDE
    # --------------------------------------------------------

    st.markdown(
        "### Quick Guide"
    )

    st.markdown(
        """
**1. Choose a document**

✉️ Email  
📄 Letter  
🔎 Inquiry

**2. Enter information**

Type normally or use 🎙️.

**3. Generate**

AI corrects language while preserving facts.

**4. Review & Modify**

Edit the generated document or ask AI to make changes.

**5. Export**

PDF • DOCX • TXT • PNG
"""
    )

    st.divider()

    # --------------------------------------------------------
    # SUGGESTIONS
    # --------------------------------------------------------

    with st.expander(
        "💡 Suggestions for Improvement"
    ):

        st.markdown(
            """
The following features can be considered
for future versions:

📋 **Document templates**

✏️ **Edit generated document**

🔄 **Regenerate with changes**

🌐 **Urdu / English support**

📎 **Attachment support**

📚 **Saved document library**

🔐 **User login and secure profiles**

🗂️ **Searchable document history**

📊 **Inquiry progress tracking**

🖨️ **Improved official printing format**
"""
        )

    # --------------------------------------------------------
    # ABOUT DEVELOPER
    # --------------------------------------------------------

    with st.expander(
        "ℹ️ About the Developer"
    ):

        st.markdown(
            """
**Developed by: Raees Khan**

Assistant Director, NADRA
"""
        )

    st.divider()

    # --------------------------------------------------------
    # CLEAR GENERATED DRAFT
    # --------------------------------------------------------

    if st.button(
        "🗑️ Clear Generated Draft",
        use_container_width=True,
    ):

        st.session_state.generated_draft = ""

        st.session_state.editable_draft = ""

        st.success(
            "Generated draft cleared."
        )

    # --------------------------------------------------------
    # RECENT DRAFTS
    # --------------------------------------------------------

    with st.expander(
        "Recent Drafts"
    ):

        if not st.session_state.history:

            st.caption(
                "No recent drafts."
            )

        else:

            for i, item in enumerate(
                st.session_state.history
            ):

                st.markdown(
                    f"**{item['type']}**"
                )

                st.caption(
                    item["date"]
                )

                if st.button(
                    "Open",
                    key=f"history_open_{i}",
                    use_container_width=True,
                ):

                    st.session_state.generated_draft = (
                        item["draft"]
                    )

                    st.session_state.editable_draft = (
                        item["draft"]
                    )

                    st.rerun()


# ============================================================
# MAIN HEADER
# ============================================================

st.markdown(
    '<div class="main-title">📝 DraftForge</div>',
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="subtitle">'
    'AI Document Composer — create professional official '
    'documents using text or voice'
    '</div>',
    unsafe_allow_html=True,
)

st.divider()

st.markdown(
    "### 📌 What would you like to create?"
)


# ============================================================
# DOCUMENT TYPE
# ============================================================

document_choice = st.segmented_control(
    "Document Type",
    [
        "✉️ Email",
        "📄 Letter",
        "🔎 Inquiry",
    ],
    default="✉️ Email",
    key="document_type_selector",
    selection_mode="single",
    width="stretch",
    label_visibility="collapsed",
)


if document_choice == "📄 Letter":

    current_type = "Letter"

elif document_choice == "🔎 Inquiry":

    current_type = "Inquiry"

else:

    current_type = "Email"


st.session_state.document_type = (
    current_type
)


# ============================================================
# EMAIL
# ============================================================

if current_type == "Email":

    st.markdown(
        "### ✉️ Official Email"
    )

    st.text_input(
        "Recipient",
        key="email_recipient",
        placeholder="Enter recipient",
    )

    st.text_input(
        "Subject",
        key="email_subject",
        placeholder="Enter email subject",
    )

    render_composer(
        "email",
        "Email Instructions",
    )

    # Generate remains LAST.

    if st.button(
        "✨ Generate Email",
        type="primary",
        use_container_width=True,
    ):

        if not profile_is_complete():

            st.warning(
                "Please complete and save all four fields "
                "under 'My Official Information' first."
            )

        elif not st.session_state.email_input.strip():

            st.warning(
                "Please provide email information by typing "
                "or using the microphone."
            )

        else:

            with st.spinner(
                "Drafting professional email..."
            ):

                try:

                    draft = generate_email()

                    st.session_state.generated_draft = (
                        draft
                    )

                    st.session_state.editable_draft = (
                        draft
                    )

                    save_history(
                        "Email",
                        draft,
                    )

                    st.rerun()

                except Exception as e:

                    st.error(
                        f"Could not generate email: {e}"
                    )


# ============================================================
# LETTER
# ============================================================

elif current_type == "Letter":

    st.markdown(
        "### 📄 Official Letter"
    )

    st.text_input(
        "Recipient",
        key="letter_recipient",
        placeholder="Enter recipient",
    )

    st.text_input(
        "Subject",
        key="letter_subject",
        placeholder="Enter letter subject",
    )

    render_composer(
        "letter",
        "Letter Instructions",
    )

    # Generate remains LAST.

    if st.button(
        "✨ Generate Letter",
        type="primary",
        use_container_width=True,
    ):

        if not profile_is_complete():

            st.warning(
                "Please complete and save all four fields "
                "under 'My Official Information' first."
            )

        elif not st.session_state.letter_input.strip():

            st.warning(
                "Please provide letter information by typing "
                "or using the microphone."
            )

        else:

            with st.spinner(
                "Drafting professional letter..."
            ):

                try:

                    draft = generate_letter()

                    st.session_state.generated_draft = (
                        draft
                    )

                    st.session_state.editable_draft = (
                        draft
                    )

                    save_history(
                        "Letter",
                        draft,
                    )

                    st.rerun()

                except Exception as e:

                    st.error(
                        f"Could not generate letter: {e}"
                    )


# ============================================================
# INQUIRY
# ============================================================

elif current_type == "Inquiry":

    st.markdown(
        "### 🔎 Departmental Inquiry"
    )

    inquiry_choice = st.segmented_control(
        "Inquiry Type",
        [
            "⚖️ E&D Inquiry",
            "🔍 FFI Inquiry",
        ],
        default="⚖️ E&D Inquiry",
        key="inquiry_type_selector",
        selection_mode="single",
        width="stretch",
    )

    if inquiry_choice == "🔍 FFI Inquiry":

        st.markdown(
            """
            <div class="warning-box">
            <strong>FFI Inquiry</strong><br>
            This module is currently under construction.
            </div>
            """,
            unsafe_allow_html=True,
        )

    else:

        st.session_state.inquiry_type = (
            "E&D Inquiry"
        )

        # ----------------------------------------------------
        # REFERENCE NUMBER
        # ----------------------------------------------------

        reference_no = st.text_input(
            "Inquiry Reference No.",
            placeholder="ABC/678",
            key="inquiry_reference_no",
        )

        st.caption(
            "Select an index, add it, then provide information "
            "by typing or voice. The same index can be selected "
            "multiple times."
        )

        # ----------------------------------------------------
        # INDEX SELECTOR
        # ----------------------------------------------------

        selected_index = st.selectbox(
            "➕ Select E&D Index",
            ED_INDEXES,
            key="selected_ed_index",
        )

        if st.button(
            "➕ Add Selected Index",
            use_container_width=True,
        ):

            add_ed_index(
                selected_index
            )

            st.rerun()

        st.divider()

        # ----------------------------------------------------
        # INQUIRY INFORMATION
        # ----------------------------------------------------

        if st.session_state.ed_instances:

            st.markdown(
                "### 📋 Inquiry Information"
            )

            for item in list(
                st.session_state.ed_instances
            ):

                index_name = item["index"]

                # ============================================
                # DOCUMENTS RECORDED
                # ============================================

                if index_name == "Documents Recorded":

                    documents_recorded = (
                        render_documents_recorded()
                    )

                    st.session_state[
                        "documents_recorded_selected"
                    ] = documents_recorded

                    st.divider()

                    continue

                # ============================================
                # INQUIRY COMMITTEE
                # ============================================

                if index_name == "Inquiry Committee":

                    committee = render_committee()

                    st.session_state[
                        "inquiry_committee_data"
                    ] = committee

                    st.divider()

                    continue

                # ============================================
                # NORMAL INDEX
                # ============================================

                render_ed_input(
                    item
                )

                if st.button(
                    "🗑️ Remove This Index",
                    key=f"remove_ed_{item['id']}",
                ):

                    remove_ed_index(
                        item["id"]
                    )

                    st.rerun()

                st.divider()

        # ----------------------------------------------------
        # GENERATE REPORT
        # ----------------------------------------------------

        if st.button(
            "✨ Generate Inquiry Report",
            type="primary",
            use_container_width=True,
        ):

            if not reference_no.strip():

                st.warning(
                    "Please enter the Inquiry Reference No."
                )

            elif not st.session_state.ed_instances:

                st.warning(
                    "Please add at least one inquiry index."
                )

            else:

                documents_recorded = (
                    st.session_state.get(
                        "documents_recorded_selected",
                        [],
                    )
                )

                committee = (
                    st.session_state.get(
                        "inquiry_committee_data",
                        {},
                    )
                )

                with st.spinner(
                    "Preparing departmental inquiry report..."
                ):

                    try:

                        draft = generate_ed_report(
                            reference_no,
                            documents_recorded,
                            committee,
                        )

                        st.session_state.generated_draft = (
                            draft
                        )

                        st.session_state.editable_draft = (
                            draft
                        )

                        save_history(
                            "E&D Inquiry",
                            draft,
                        )

                        st.rerun()

                    except Exception as e:

                        st.error(
                            f"Could not generate inquiry report: {e}"
                        )

        # ----------------------------------------------------
        # START NEW INQUIRY
        # ----------------------------------------------------

        if st.button(
            "🔄 Start New Inquiry",
            use_container_width=True,
        ):

            st.session_state.ed_instances = []

            st.session_state.generated_draft = ""

            st.session_state.editable_draft = ""

            st.session_state.documents_recorded_selected = []

            st.session_state.inquiry_committee_data = {}

            st.rerun()


# ============================================================
# GENERATED DOCUMENT
# ============================================================

if st.session_state.get(
    "generated_draft",
    "",
).strip():

    st.divider()

    st.markdown(
        "## 📄 Generated Document"
    )

    # --------------------------------------------------------
    # LIVE PREVIEW
    # --------------------------------------------------------

    st.markdown(
        "### 👁️ Document Preview"
    )

    st.caption(
        "Questions and Answers are displayed as a "
        "two-column table."
    )

    render_document_markdown(
        st.session_state.editable_draft
    )

    st.divider()

    # --------------------------------------------------------
    # EDIT DOCUMENT
    # --------------------------------------------------------

    st.markdown(
        "### ✏️ Review & Modify Document"
    )

    st.caption(
        "You can directly edit the generated document "
        "before exporting it."
    )

    edited_document = st.text_area(
        "Generated Document",
        key="editable_draft",
        height=650,
        label_visibility="collapsed",
    )

    col1, col2 = st.columns(2)

    with col1:

        if st.button(
            "💾 Save Changes",
            type="primary",
            use_container_width=True,
        ):

            st.session_state.generated_draft = (
                edited_document
            )

            save_history(
                st.session_state.document_type,
                edited_document,
            )

            st.success(
                "Your changes have been saved."
            )

    with col2:

        if st.button(
            "↩️ Restore Generated Version",
            use_container_width=True,
        ):

            st.session_state.editable_draft = (
                st.session_state.generated_draft
            )

            st.rerun()

    # --------------------------------------------------------
    # AI EDITING
    # --------------------------------------------------------

    st.markdown(
        "### 🤖 Ask AI to Modify the Document"
    )

    st.caption(
        "Describe exactly what you want changed."
    )

    st.text_area(
        "Modification instructions",
        key="edit_instruction",
        height=100,
        placeholder=(
            "Example: Make the recommendations more "
            "formal without changing their meaning."
        ),
    )

    if st.button(
        "🤖 Apply AI Changes",
        use_container_width=True,
    ):

        instruction = (
            st.session_state.edit_instruction.strip()
        )

        if not instruction:

            st.warning(
                "Please enter the changes you want AI to make."
            )

        elif not edited_document.strip():

            st.warning(
                "There is no document to modify."
            )

        else:

            with st.spinner(
                "Applying requested changes..."
            ):

                try:

                    modified = (
                        modify_generated_document(
                            edited_document,
                            instruction,
                        )
                    )

                    st.session_state.editable_draft = (
                        modified
                    )

                    st.session_state.generated_draft = (
                        modified
                    )

                    st.session_state.edit_instruction = ""

                    save_history(
                        st.session_state.document_type,
                        modified,
                    )

                    st.success(
                        "AI changes applied."
                    )

                    st.rerun()

                except Exception as e:

                    st.error(
                        f"Could not modify document: {e}"
                    )

    # --------------------------------------------------------
    # EXPORT
    # --------------------------------------------------------

    st.markdown(
        "### 📤 Export Document"
    )

    final_document = (
        st.session_state.get(
            "editable_draft",
            st.session_state.generated_draft,
        )
    )

    col1, col2, col3, col4 = st.columns(
        4
    )

    with col1:

        st.download_button(
            "📄 PDF",
            data=export_pdf(
                final_document
            ),
            file_name="DraftForge_Document.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    with col2:

        st.download_button(
            "📝 DOCX",
            data=export_docx(
                final_document
            ),
            file_name="DraftForge_Document.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            use_container_width=True,
        )

    with col3:

        st.download_button(
            "📃 TXT",
            data=export_txt(
                final_document
            ),
            file_name="DraftForge_Document.txt",
            mime="text/plain",
            use_container_width=True,
        )

    with col4:

        st.download_button(
            "🖼️ PNG",
            data=export_png(
                final_document
            ),
            file_name="DraftForge_Document.png",
            mime="image/png",
            use_container_width=True,
    )
