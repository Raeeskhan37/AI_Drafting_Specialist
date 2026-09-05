import io
import os
import re
import json
import textwrap
import requests
import streamlit as st

from datetime import datetime
from groq import Groq

from docx import Document
from docx.shared import Pt

from fpdf import FPDF

from PIL import Image, ImageDraw, ImageFont


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="DraftForge — AI Document Composer",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# CONSTANTS
# ============================================================

GROQ_MODEL = "openai/gpt-oss-120b"
WHISPER_MODEL = "whisper-large-v3"
GEMINI_MODEL = "gemini-2.0-flash"

PROFILE_FILE = "user_profile.json"


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


SPECIAL_INDEXES = {
    "Documents Recorded",
    "Inquiry Committee",
}


QA_INDEX_PREFIX = "Questions / Answers with the Accused"


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
    "Statement of Accused",
    "Statement of Witness",
    "Other",
]


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
        margin-bottom: 0.2rem;
    }

    .subtitle {
        color: #666;
        margin-bottom: 1.5rem;
    }

    .developer-box {
        padding: 12px;
        border-radius: 10px;
        border: 1px solid rgba(128,128,128,0.25);
        margin-top: 10px;
    }

    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# PROFILE
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


def get_signature(profile):

    lines = []

    if profile.get("name", "").strip():
        lines.append(
            profile["name"].strip()
        )

    if profile.get("designation", "").strip():
        lines.append(
            profile["designation"].strip()
        )

    if profile.get("contact_no", "").strip():
        lines.append(
            f"Contact No.: {profile['contact_no'].strip()}"
        )

    if profile.get("current_station", "").strip():
        lines.append(
            f"Current Station: {profile['current_station'].strip()}"
        )

    if not lines:
        return ""

    return "\n\n" + "\n".join(lines)


# ============================================================
# SESSION STATE
# ============================================================

def initialize_state():

    defaults = {

        "profile": load_user_profile(),

        "profile_name": "",
        "profile_designation": "",
        "profile_contact": "",
        "profile_station": "",

        "document_type": "Email",

        "email_instruction": "",
        "letter_instruction": "",

        "inquiry_indexes": [],
        "inquiry_index_counter": {},

        "generated_draft": "",

        "editable_draft": "",

        "document_editor": "",

        "editor_sync": "",

        "edit_instruction": "",

        # IMPORTANT:
        # Separate state used to clear/synchronize
        # the edit_instruction widget BEFORE creation.
        "edit_instruction_sync": None,

        "history": [],
    }

    for key, value in defaults.items():

        if key not in st.session_state:

            st.session_state[key] = value


    profile = st.session_state["profile"]


    if not st.session_state["profile_name"]:
        st.session_state["profile_name"] = profile.get(
            "name",
            "",
        )


    if not st.session_state["profile_designation"]:
        st.session_state["profile_designation"] = profile.get(
            "designation",
            "",
        )


    if not st.session_state["profile_contact"]:
        st.session_state["profile_contact"] = profile.get(
            "contact_no",
            "",
        )


    if not st.session_state["profile_station"]:
        st.session_state["profile_station"] = profile.get(
            "current_station",
            "",
        )


initialize_state()


# ============================================================
# API
# ============================================================

def get_secret(name):

    try:

        value = st.secrets.get(name)

        if value:
            return value

    except Exception:
        pass

    return os.environ.get(
        name,
        "",
    )


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
# AI CLEANING
# ============================================================

def clean_ai_output(text):

    if not text:
        return ""

    text = text.strip()

    text = re.sub(
        r"^```(?:markdown|text|txt)?\s*",
        "",
        text,
        flags=re.IGNORECASE,
    )

    text = re.sub(
        r"\s*```$",
        "",
        text,
    )

    return text.strip()


# ============================================================
# GROQ
# ============================================================

def call_groq(prompt):

    client = get_groq_client()

    if client is None:
        return None

    try:

        response = client.chat.completions.create(

            model=GROQ_MODEL,

            messages=[
                {
                    "role": "system",
                    "content": (
                        "You are a professional official "
                        "document drafting assistant. "
                        "Use only information supplied by "
                        "the user. Never invent facts."
                    ),
                },
                {
                    "role": "user",
                    "content": prompt,
                },
            ],

            temperature=0.2,
        )

        return clean_ai_output(
            response.choices[0].message.content
        )

    except Exception:

        return None


# ============================================================
# GEMINI
# ============================================================

def call_gemini(prompt):

    api_key = get_secret(
        "GEMINI_API_KEY"
    )

    if not api_key:
        return None

    url = (
        "https://generativelanguage.googleapis.com/"
        "v1beta/models/"
        f"{GEMINI_MODEL}:generateContent"
        f"?key={api_key}"
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
        ],

        "generationConfig": {
            "temperature": 0.2
        },
    }

    try:

        response = requests.post(
            url,
            json=payload,
            timeout=60,
        )

        response.raise_for_status()

        data = response.json()

        text = (
            data
            .get("candidates", [{}])[0]
            .get("content", {})
            .get("parts", [{}])[0]
            .get("text", "")
        )

        return clean_ai_output(
            text
        )

    except Exception:

        return None


def ask_ai(prompt):

    result = call_groq(prompt)

    if result:
        return result

    result = call_gemini(prompt)

    if result:
        return result

    return None


# ============================================================
# VOICE TRANSCRIPTION
# ============================================================

def transcribe_audio(audio_file):

    client = get_groq_client()

    if client is None:

        st.error(
            "GROQ_API_KEY is not configured."
        )

        return ""

    try:

        audio_bytes = (
            audio_file.getvalue()
        )

        transcription = (
            client.audio.transcriptions.create(
                file=(
                    "recording.wav",
                    audio_bytes,
                    "audio/wav",
                ),
                model=WHISPER_MODEL,
            )
        )

        return transcription.text.strip()

    except Exception as e:

        st.error(
            f"Voice transcription failed: {e}"
        )

        return ""


# ============================================================
# EMAIL
# ============================================================

def generate_email(instruction):

    prompt = f"""
Draft a professional official email based ONLY on
the user's instructions.

USER INSTRUCTIONS:

{instruction}

Rules:

1. Use professional official English.
2. Correct spelling, grammar and punctuation.
3. Correct obvious voice-transcription mistakes.
4. Preserve the user's intended meaning.
5. Do not invent facts.
6. Do not invent names, dates, reference numbers,
   allegations, events or decisions.
7. Do not create a sender signature.
8. The application will append the official profile.
9. Return only the email body.
"""

    return ask_ai(prompt)


# ============================================================
# LETTER
# ============================================================

def generate_letter(instruction):

    prompt = f"""
Draft a professional official letter based ONLY on
the user's instructions.

USER INSTRUCTIONS:

{instruction}

Rules:

1. Use professional official English.
2. Correct spelling, grammar and punctuation.
3. Correct obvious voice-transcription mistakes.
4. Preserve the user's intended meaning.
5. Do not invent facts.
6. Do not invent names, dates, reference numbers,
   allegations, events or decisions.
7. Do not create a sender signature.
8. The application will append the official profile.
9. Return only the letter content.
"""

    return ask_ai(prompt)


# ============================================================
# INDEX MANAGEMENT
# ============================================================

def next_index_number(index_name):

    current = (
        st.session_state
        .inquiry_index_counter
        .get(
            index_name,
            0,
        )
    )

    current += 1

    st.session_state.inquiry_index_counter[
        index_name
    ] = current

    return current


def add_inquiry_index(index_name):

    number = next_index_number(
        index_name
    )

    unique_id = (
        f"{index_name}-{number}-"
        f"{datetime.now().timestamp()}"
    )

    st.session_state.inquiry_indexes.append(
        {
            "name": index_name,
            "number": number,
            "id": unique_id,
        }
    )


def remove_inquiry_index(index_id):

    st.session_state.inquiry_indexes = [

        item

        for item in st.session_state.inquiry_indexes

        if item["id"] != index_id
    ]


# ============================================================
# VOICE + TEXT
# ============================================================

def render_voice_text_input(
    label,
    text_key,
    audio_key,
    height=180,
):

    st.markdown(
        f"**{label}**"
    )

    # MICROPHONE FIRST

    audio = st.audio_input(
        "🎤 Record Voice",
        key=audio_key,
    )

    if audio is not None:

        transcript = transcribe_audio(
            audio
        )

        if transcript:

            current = st.session_state.get(
                text_key,
                "",
            )

            if current.strip():

                st.session_state[text_key] = (
                    current.rstrip()
                    + "\n"
                    + transcript
                )

            else:

                st.session_state[text_key] = (
                    transcript
                )


    text = st.text_area(
        "Type or review your instructions",
        key=text_key,
        height=height,
        label_visibility="collapsed",
    )

    return text


# ============================================================
# INDEX INPUT
# ============================================================

def render_index_input(index_item):

    name = index_item["name"]
    number = index_item["number"]
    unique_id = index_item["id"]

    heading = name

    if number > 1:
        heading += f" — No. {number}"

    st.markdown(
        f"### {heading}"
    )


    # ========================================================
    # DOCUMENTS RECORDED
    # ========================================================

    if name == "Documents Recorded":

        st.info(
            "Select the documents that are part of the "
            "inquiry record. The application will automatically "
            "assign annexures as Annex-A, Annex-B, Annex-C "
            "and so on."
        )

        selected = st.multiselect(
            "Select Documents Recorded",
            DOCUMENTS_RECORDED,
            key=f"documents_{unique_id}",
        )

        other = st.text_input(
            "Other document, if any",
            key=f"documents_other_{unique_id}",
        )

        annexure_items = list(selected)

        if other.strip():
            annexure_items.append(
                other.strip()
            )

        if annexure_items:

            st.markdown(
                "**Automatic Annexure Numbering**"
            )

            for idx, document in enumerate(
                annexure_items
            ):

                annexure_letter = chr(
                    65 + idx
                )

                st.write(
                    f"{idx + 1}. "
                    f"{document} — "
                    f"**Annex-{annexure_letter}**"
                )

        return {
            "type": "documents",
            "selected": selected,
            "other": other,
        }


    # ========================================================
    # INQUIRY COMMITTEE
    # ========================================================

    if name == "Inquiry Committee":

        committee_data = {}

        for role in COMMITTEE_ROLES:

            st.markdown(
                f"**{role}**"
            )

            col1, col2, col3 = st.columns(3)

            with col1:

                erp = st.text_input(
                    "ERP#",
                    key=f"erp_{unique_id}_{role}",
                )

            with col2:

                member_name = st.text_input(
                    "Name",
                    key=f"name_{unique_id}_{role}",
                )

            with col3:

                designation = st.text_input(
                    "Designation",
                    key=f"designation_{unique_id}_{role}",
                )

            committee_data[role] = {
                "erp": erp,
                "name": member_name,
                "designation": designation,
            }

        return {
            "type": "committee",
            "data": committee_data,
        }


    # ========================================================
    # NORMAL INDEX
    # ========================================================

    text_key = (
        f"inquiry_text_{unique_id}"
    )

    audio_key = (
        f"inquiry_audio_{unique_id}"
    )

    text = render_voice_text_input(
        f"Instructions / Information for {name}",
        text_key,
        audio_key,
    )

    return {
        "type": "text",
        "text": text,
    }


# ============================================================
# QA
# ============================================================

def is_qa_heading(line):

    normalized = line.strip().lower()

    return normalized.startswith(
        QA_INDEX_PREFIX.lower()
    )


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
            for cell in line.strip("|").split("|")
        ]

        if all(
            re.fullmatch(
                r":?-+:?",
                cell.replace(" ", ""),
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
# BUILD SELECTED SECTION LIST
# ============================================================

def get_selected_section_names(index_data):

    names = []

    for item in index_data:

        name = item["name"]

        if item["number"] > 1:

            heading = (
                f"{name} No. {item['number']}"
            )

        else:

            heading = name

        names.append(heading)

    return names


# ============================================================
# E&D REPORT
# ============================================================

def generate_ed_report(index_data):

    if not index_data:
        return None


    # ========================================================
    # IMPORTANT:
    # ONLY USER-SELECTED INDEXES ARE USED.
    # ========================================================

    selected_sections = (
        get_selected_section_names(
            index_data
        )
    )


    allowed_sections_text = "\n".join(
        f"{i + 1}. {name}"
        for i, name in enumerate(
            selected_sections
        )
    )


    # ========================================================
    # NORMAL INDEX DATA
    # ========================================================

    normal_parts = []

    for item in index_data:

        name = item["name"]

        if name in SPECIAL_INDEXES:
            continue

        text = item.get(
            "text",
            "",
        ).strip()

        heading = name

        if item["number"] > 1:

            heading += (
                f" No. {item['number']}"
            )

        # IMPORTANT:
        # Even empty selected indexes are represented.
        if text:

            normal_parts.append(
                f"""
SECTION:
{heading}

USER INFORMATION:
{text}
"""
            )

        else:

            normal_parts.append(
                f"""
SECTION:
{heading}

USER INFORMATION:
No information was provided by the user for this selected index.
"""
            )


    normal_information = "\n".join(
        normal_parts
    )


    # ========================================================
    # DOCUMENTS RECORDED
    # ========================================================

    documents = []

    documents_selected = False

    for item in index_data:

        if item["name"] == "Documents Recorded":

            documents_selected = True

            documents.extend(
                item.get(
                    "selected",
                    [],
                )
            )

            other = item.get(
                "other",
                "",
            ).strip()

            if other:
                documents.append(
                    other
                )


    annexure_lines = []

    for idx, document in enumerate(
        documents
    ):

        letter = chr(
            65 + idx
        )

        annexure_lines.append(
            f"{idx + 1}. "
            f"{document} — "
            f"Annex-{letter}"
        )


    if annexure_lines:

        documents_text = "\n".join(
            annexure_lines
        )

    else:

        documents_text = (
            "No documents were selected "
            "by the user."
        )


    # ========================================================
    # COMMITTEE
    # ========================================================

    committee_selected = False
    committee_lines = []

    for item in index_data:

        if item["name"] == "Inquiry Committee":

            committee_selected = True

            committee = item.get(
                "data",
                {},
            )

            for role in COMMITTEE_ROLES:

                member = committee.get(
                    role,
                    {},
                )

                erp = member.get(
                    "erp",
                    "",
                ).strip()

                name = member.get(
                    "name",
                    "",
                ).strip()

                designation = member.get(
                    "designation",
                    "",
                ).strip()

                if (
                    erp
                    or name
                    or designation
                ):

                    committee_lines.append(
                        f"{role}: "
                        f"ERP# {erp}; "
                        f"Name: {name}; "
                        f"Designation: "
                        f"{designation}"
                    )


    if committee_lines:

        committee_text = "\n".join(
            committee_lines
        )

    else:

        committee_text = (
            "No committee member details "
            "were provided."
        )


    # ========================================================
    # STRUCTURED SECTIONS
    # ========================================================

    structured_information = ""

    if documents_selected:

        structured_information += f"""

STRUCTURED SECTION:
Documents Recorded

The application has already assigned the following
exact annexure labels:

{documents_text}

The AI MUST reproduce these documents and annexure
labels exactly.
"""


    if committee_selected:

        structured_information += f"""

STRUCTURED SECTION:
Inquiry Committee

{committee_text}

Include Inquiry Committee ONLY because the user
selected this index.
"""


    # ========================================================
    # CRITICAL AI PROMPT
    # ========================================================

    prompt = f"""
Prepare a professional E&D Inquiry Report using ONLY
the information supplied by the user.

============================================================
CRITICAL SECTION CONTROL
============================================================

The following are the ONLY sections that the user
actually selected:

{allowed_sections_text}

These are the ONLY sections permitted in the final report.

DO NOT create any other section.

In particular:

- DO NOT create Introduction unless it is listed above.
- DO NOT create Summary of Evidence unless it is listed above.
- DO NOT create Findings unless it is listed above.
- DO NOT create Conclusions unless it is listed above.
- DO NOT create Recommendations unless it is listed above.
- DO NOT create Discussion unless it is listed above.
- DO NOT create Inquiry Committee unless it is listed above.
- DO NOT create Documents Recorded unless it is listed above.
- DO NOT create any generic E&D report sections.
- DO NOT add an opening overview.
- DO NOT add a closing summary.
- DO NOT add sections because they are normally present
  in an E&D inquiry report.

The final report must contain ONLY the selected sections
listed above and must follow their order.

============================================================
FACTUAL RULES
============================================================

1. Use professional official English.
2. Correct grammar, spelling and punctuation.
3. Correct obvious voice-transcription mistakes.
4. Preserve the user's intended meaning.
5. NEVER invent facts.
6. NEVER invent names.
7. NEVER invent dates.
8. NEVER invent allegations.
9. NEVER invent witnesses.
10. NEVER invent evidence.
11. NEVER invent findings.
12. NEVER invent recommendations.
13. NEVER assume missing information.
14. Do not add facts merely because they are common in
    E&D inquiry reports.
15. Do not add information from your own knowledge.

============================================================
SELECTED INDEX INFORMATION
============================================================

{normal_information}

{structured_information}

============================================================
DOCUMENTS RECORDED RULES
============================================================

If and ONLY IF "Documents Recorded" appears in the
selected section list:

- Include Documents Recorded exactly once.
- Include ONLY documents actually selected.
- Preserve the exact document names.
- Preserve the exact automatically assigned annexures.
- Do NOT renumber annexures.
- Do NOT change Annex-A into Annex-1.
- Do NOT skip letters.
- Do NOT create additional documents.

The annexures are controlled by the application.

============================================================
INQUIRY COMMITTEE RULES
============================================================

If and ONLY IF "Inquiry Committee" appears in the
selected section list:

- Include Inquiry Committee exactly once.
- Use only the supplied committee information.
- Do not invent members.
- Do not invent ERP numbers.
- Do not invent designations.
- Do not create an empty committee section if the index
  was not selected.

============================================================
QUESTIONS / ANSWERS RULE
============================================================

Whenever a selected section is:

Questions / Answers with the Accused

or:

Questions / Answers with the Accused No. 2
Questions / Answers with the Accused No. 3

the section MUST use exactly this two-column Markdown
table structure:

| Questions | Answers |
|---|---|
| 1. Question | Answer |
| 2. Question | Answer |

Do not merge questions and answers.

Do not invent questions.

Do not invent answers.

============================================================
FINAL OUTPUT RULE
============================================================

Return ONLY the final report.

Do not provide explanations about your drafting.

Do not mention which sections were omitted.

Do not mention these instructions.

Do not create headings that are not present in the
selected section list.
"""


    result = ask_ai(
        prompt
    )


    if not result:
        return None


    return result.strip()


# ============================================================
# PROFILE UPDATE
# ============================================================

def update_profile_from_widgets():

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

    st.session_state.profile = profile

    save_user_profile(profile)

    return profile


def profile_complete():

    profile = (
        st.session_state.profile
    )

    return all(
        [
            profile.get("name", "").strip(),
            profile.get("designation", "").strip(),
            profile.get("contact_no", "").strip(),
            profile.get("current_station", "").strip(),
        ]
    )


# ============================================================
# HISTORY
# ============================================================

def save_to_history(
    document_type,
    content,
):

    if not content:
        return

    item = {

        "type":
        document_type,

        "date":
        datetime.now().strftime(
            "%Y-%m-%d %H:%M"
        ),

        "content":
        content,

    }

    history = (
        st.session_state.history
    )

    history.insert(
        0,
        item,
    )

    st.session_state.history = (
        history[:20]
    )


# ============================================================
# MODIFY DOCUMENT
# ============================================================

def modify_generated_document(
    current_document,
    instruction,
):

    prompt = f"""
Modify the following official document according to
the user's requested changes.

USER REQUEST:

{instruction}

CURRENT DOCUMENT:

{current_document}

Rules:

1. Preserve all existing facts.
2. Do not invent facts.
3. Do not invent names, dates, allegations, evidence,
   findings or recommendations.
4. Make only the requested changes.
5. Maintain professional official English.
6. Correct grammar and punctuation where appropriate.
7. Preserve factual meaning.
8. DO NOT introduce new sections that were not present
   in the current document.
9. DO NOT create Introduction, Summary, Findings,
   Recommendations, Conclusion or any other section
   unless that section already exists in the document.
10. If the document contains Documents Recorded,
    preserve the exact document names and Annex-A,
    Annex-B, Annex-C etc.
11. Never change annexure labels.
12. If the document contains Questions / Answers,
    retain the exact two-column Markdown table:

| Questions | Answers |
|---|---|
| 1. Question | Answer |
| 2. Question | Answer |

13. Return the complete modified document.
"""

    return ask_ai(
        prompt
    )


# ============================================================
# DOCUMENT PREVIEW
# ============================================================

def display_generated_document(text):

    if not text:
        return

    lines = text.splitlines()

    i = 0

    while i < len(lines):

        line = lines[i]

        if is_qa_heading(line):

            st.markdown(
                f"### {line.strip()}"
            )

            rows, end = (
                parse_markdown_table(
                    lines,
                    i + 1,
                )
            )

            if rows:

                cleaned = []

                for q, a in rows:

                    if (
                        q.lower() == "questions"
                        and
                        a.lower() == "answers"
                    ):
                        continue

                    cleaned.append(
                        {
                            "Questions": q,
                            "Answers": a,
                        }
                    )

                if cleaned:
                    st.table(cleaned)

                i = end

                continue

        st.markdown(line)

        i += 1


# ============================================================
# TXT
# ============================================================

def create_txt(text):

    return text.encode(
        "utf-8"
    )


# ============================================================
# DOCX
# ============================================================

def add_markdown_table_to_docx(
    document,
    rows,
):

    if not rows:
        return

    table = document.add_table(
        rows=1,
        cols=2,
    )

    table.style = "Table Grid"

    header = table.rows[0].cells

    header[0].text = "Questions"
    header[1].text = "Answers"

    for question, answer in rows:

        if (
            question.lower() == "questions"
            and
            answer.lower() == "answers"
        ):
            continue

        cells = (
            table.add_row().cells
        )

        cells[0].text = question
        cells[1].text = answer


def create_docx(text):

    document = Document()

    lines = text.splitlines()

    i = 0

    while i < len(lines):

        line = lines[i].strip()

        if not line:

            document.add_paragraph("")

            i += 1

            continue

        if is_qa_heading(line):

            paragraph = (
                document.add_paragraph()
            )

            run = paragraph.add_run(
                line
            )

            run.bold = True
            run.font.size = Pt(13)

            rows, end = (
                parse_markdown_table(
                    lines,
                    i + 1,
                )
            )

            if rows:

                add_markdown_table_to_docx(
                    document,
                    rows,
                )

                i = end

                continue

        paragraph = (
            document.add_paragraph(line)
        )

        for run in paragraph.runs:
            run.font.size = Pt(11)

        i += 1

    output = io.BytesIO()

    document.save(output)

    output.seek(0)

    return output.getvalue()


# ============================================================
# PDF
# ============================================================

class PDFDocument(FPDF):

    def footer(self):

        self.set_y(-15)

        self.set_font(
            "Arial",
            size=8,
        )

        self.cell(
            0,
            10,
            f"Page {self.page_no()}",
            align="C",
        )


def pdf_safe(text):

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
        "←": "<-",

    }

    for old, new in replacements.items():

        text = text.replace(
            old,
            new,
        )

    return (
        text
        .encode(
            "latin-1",
            "replace",
        )
        .decode(
            "latin-1"
        )
    )


# ============================================================
# SAFE PDF TEXT WRITER
# ============================================================

def pdf_write_safe(
    pdf,
    text,
    height=6,
):

    text = pdf_safe(
        str(text)
    )

    if not text:
        pdf.ln(3)
        return

    # --------------------------------------------------------
    # Break very long unbroken strings.
    # This prevents:
    #
    # Not enough horizontal space to render a single character
    # --------------------------------------------------------

    chunks = []

    for paragraph in text.split("\n"):

        if not paragraph:

            chunks.append("")

            continue

        wrapped = textwrap.wrap(
            paragraph,
            width=110,
            break_long_words=True,
            break_on_hyphens=False,
        )

        if wrapped:
            chunks.extend(wrapped)
        else:
            chunks.append("")


    for line in chunks:

        if not line:

            pdf.ln(3)

            continue

        pdf.multi_cell(
            0,
            height,
            line,
        )


# ============================================================
# PDF QA TABLE
# ============================================================

def add_pdf_qa_table(
    pdf,
    rows,
):

    if not rows:
        return

    col1 = 80
    col2 = 105

    # Header

    pdf.set_font(
        "Arial",
        "B",
        9,
    )

    pdf.cell(
        col1,
        8,
        "Questions",
        border=1,
    )

    pdf.cell(
        col2,
        8,
        "Answers",
        border=1,
    )

    pdf.ln()

    pdf.set_font(
        "Arial",
        "",
        8,
    )

    for question, answer in rows:

        if (
            question.lower() == "questions"
            and
            answer.lower() == "answers"
        ):
            continue

        question = pdf_safe(
            question
        )

        answer = pdf_safe(
            answer
        )

        # Wrap question and answer safely.

        q_lines = textwrap.wrap(
            question,
            width=48,
            break_long_words=True,
            break_on_hyphens=False,
        )

        a_lines = textwrap.wrap(
            answer,
            width=62,
            break_long_words=True,
            break_on_hyphens=False,
        )

        if not q_lines:
            q_lines = [""]

        if not a_lines:
            a_lines = [""]

        row_count = max(
            len(q_lines),
            len(a_lines),
        )

        row_height = 5

        x = pdf.get_x()
        y = pdf.get_y()

        for row_index in range(
            row_count
        ):

            q_text = (
                q_lines[row_index]
                if row_index < len(q_lines)
                else ""
            )

            a_text = (
                a_lines[row_index]
                if row_index < len(a_lines)
                else ""
            )

            pdf.set_xy(
                x,
                y + row_index * row_height,
            )

            pdf.cell(
                col1,
                row_height,
                q_text,
                border=1,
            )

            pdf.set_xy(
                x + col1,
                y + row_index * row_height,
            )

            pdf.cell(
                col2,
                row_height,
                a_text,
                border=1,
            )

        pdf.set_xy(
            x,
            y + row_count * row_height,
        )


# ============================================================
# CREATE PDF
# ============================================================

def create_pdf(text):

    pdf = PDFDocument()

    pdf.set_auto_page_break(
        auto=True,
        margin=20,
    )

    pdf.add_page()

    lines = text.splitlines()

    i = 0

    while i < len(lines):

        line = lines[i].strip()

        if not line:

            pdf.ln(3)

            i += 1

            continue


        # ----------------------------------------------------
        # Q&A SECTION
        # ----------------------------------------------------

        if is_qa_heading(line):

            pdf.set_font(
                "Arial",
                "B",
                12,
            )

            pdf_write_safe(
                pdf,
                line,
                height=7,
            )

            pdf.ln(2)

            rows, end = (
                parse_markdown_table(
                    lines,
                    i + 1,
                )
            )

            if rows:

                add_pdf_qa_table(
                    pdf,
                    rows,
                )

                pdf.ln(4)

                i = end

                continue


        # ----------------------------------------------------
        # NORMAL TEXT
        # ----------------------------------------------------

        # Detect likely headings.

        if (
            len(line) < 100
            and (
                line.endswith(":")
                or line in ED_INDEXES
                or re.match(
                    r"^.+ No\. \d+$",
                    line,
                )
            )
        ):

            pdf.set_font(
                "Arial",
                "B",
                12,
            )

            pdf_write_safe(
                pdf,
                line,
                height=7,
            )

        else:

            pdf.set_font(
                "Arial",
                "",
                10,
            )

            pdf_write_safe(
                pdf,
                line,
                height=6,
            )

        i += 1


    return bytes(
        pdf.output()
    )


# ============================================================
# PNG
# ============================================================

def create_png(text):

    width = 1600
    margin = 70
    line_height = 32

    try:

        font = ImageFont.truetype(
            "DejaVuSans.ttf",
            22,
        )

    except Exception:

        font = ImageFont.load_default()

    lines = []

    for raw_line in text.splitlines():

        if not raw_line:

            lines.append("")

            continue

        words = raw_line.split()

        current = ""

        for word in words:

            test = (
                current
                + " "
                + word
            ).strip()

            try:

                bbox = font.getbbox(
                    test
                )

                width_test = (
                    bbox[2]
                    - bbox[0]
                )

            except Exception:

                width_test = (
                    len(test) * 12
                )

            if width_test > (
                width - margin * 2
            ):

                if current:
                    lines.append(
                        current
                    )

                current = word

            else:

                current = test

        if current:
            lines.append(current)


    height = max(
        500,
        margin * 2
        + len(lines) * line_height,
    )


    image = Image.new(
        "RGB",
        (
            width,
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


    output = io.BytesIO()

    image.save(
        output,
        format="PNG",
    )

    output.seek(0)

    return output.getvalue()


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

    st.divider()


    # ========================================================
    # PROFILE
    # ========================================================

    st.markdown(
        "### 👤 User Profile"
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
        "💾 Save Profile",
        use_container_width=True,
    ):

        update_profile_from_widgets()

        st.success(
            "Profile saved."
        )


    st.divider()


    # ========================================================
    # GUIDE
    # ========================================================

    with st.expander(
        "📖 Quick Guide",
        expanded=False,
    ):

        st.markdown(
            """
            **1. Select a document type**

            Choose Email, Letter or Inquiry.

            **2. Give instructions**

            Type naturally or use the microphone.

            **3. Generate**

            AI converts your instructions into
            professional official English.

            **4. Review**

            Edit the generated document manually.

            **5. Ask AI to modify**

            Give AI instructions for further changes.

            **6. Export**

            Download PDF, DOCX, TXT or PNG.
            """
        )


    # ========================================================
    # SUGGESTIONS
    # ========================================================

    with st.expander(
        "💡 Suggestions for Improvement",
        expanded=False,
    ):

        st.markdown(
            """
            - Document templates
            - Edit generated document
            - Regenerate with changes
            - Urdu / English support
            - Attachment support
            - Saved document library
            - User login and secure profiles
            - Searchable document history
            - Inquiry progress tracking
            - Improved official printing format
            """
        )


    # ========================================================
    # HISTORY
    # ========================================================

    with st.expander(
        "🕘 Document History",
        expanded=False,
    ):

        if st.session_state.history:

            for index, item in enumerate(
                st.session_state.history
            ):

                st.markdown(
                    f"**{index + 1}. "
                    f"{item['type']}**"
                )

                st.caption(
                    item["date"]
                )

                if st.button(
                    "Restore",
                    key=f"restore_history_{index}",
                ):

                    content = item["content"]

                    st.session_state.generated_draft = (
                        content
                    )

                    st.session_state.editable_draft = (
                        content
                    )

                    st.session_state.editor_sync = (
                        content
                    )

                    st.rerun()

        else:

            st.caption(
                "No documents generated yet."
            )


    st.divider()


    # ========================================================
    # DEVELOPER
    # ========================================================

    st.markdown(
        """
        <div class="developer-box">
        <b>About the Developer</b><br><br>
        Developed by: Raees Khan<br>
        Assistant Director, NADRA
        </div>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# MAIN HEADER
# ============================================================

st.markdown(
    '<div class="main-title">'
    '📝 DraftForge — AI Document Composer'
    '</div>',
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="subtitle">'
    'Create professional official documents using '
    'natural language or voice input.'
    '</div>',
    unsafe_allow_html=True,
)


# ============================================================
# DOCUMENT TYPE
# ============================================================

document_type = st.selectbox(
    "Select Document Type",
    [
        "Email",
        "Letter",
        "Inquiry",
    ],
    key="document_type",
)


# ============================================================
# EMAIL
# ============================================================

if document_type == "Email":

    st.markdown(
        "### 📧 Email"
    )

    st.info(
        "You can type your instructions or record them "
        "using the microphone."
    )

    audio = st.audio_input(
        "🎤 Record Email Instructions",
        key="email_audio",
    )

    if audio is not None:

        transcript = transcribe_audio(
            audio
        )

        if transcript:

            current = (
                st.session_state.email_instruction
            )

            if current.strip():

                st.session_state.email_instruction = (
                    current.rstrip()
                    + "\n"
                    + transcript
                )

            else:

                st.session_state.email_instruction = (
                    transcript
                )


    st.text_area(
        "Email Instructions",
        key="email_instruction",
        height=220,
        placeholder=(
            "Example: Draft an email to the regional "
            "office requesting urgent action regarding..."
        ),
    )


    # GENERATE LAST

    if st.button(
        "✨ Generate Email",
        type="primary",
        use_container_width=True,
    ):

        if not profile_complete():

            st.error(
                "Please complete and save your User Profile "
                "before generating an Email."
            )

        elif not (
            st.session_state.email_instruction.strip()
        ):

            st.warning(
                "Please enter or record your instructions."
            )

        else:

            with st.spinner(
                "Generating professional email..."
            ):

                result = generate_email(
                    st.session_state.email_instruction
                )

            if result:

                final_document = (
                    result
                    + get_signature(
                        st.session_state.profile
                    )
                )

                st.session_state.generated_draft = (
                    final_document
                )

                st.session_state.editable_draft = (
                    final_document
                )

                st.session_state.editor_sync = (
                    final_document
                )

                save_to_history(
                    "Email",
                    final_document,
                )

                st.rerun()

            else:

                st.error(
                    "AI generation failed. "
                    "Please check your API key."
                )


# ============================================================
# LETTER
# ============================================================

elif document_type == "Letter":

    st.markdown(
        "### 📄 Letter"
    )

    st.info(
        "You can type your instructions or record them "
        "using the microphone."
    )

    audio = st.audio_input(
        "🎤 Record Letter Instructions",
        key="letter_audio",
    )

    if audio is not None:

        transcript = transcribe_audio(
            audio
        )

        if transcript:

            current = (
                st.session_state.letter_instruction
            )

            if current.strip():

                st.session_state.letter_instruction = (
                    current.rstrip()
                    + "\n"
                    + transcript
                )

            else:

                st.session_state.letter_instruction = (
                    transcript
                )


    st.text_area(
        "Letter Instructions",
        key="letter_instruction",
        height=220,
        placeholder=(
            "Example: Draft a letter to the Director "
            "requesting..."
        ),
    )


    # GENERATE LAST

    if st.button(
        "✨ Generate Letter",
        type="primary",
        use_container_width=True,
    ):

        if not profile_complete():

            st.error(
                "Please complete and save your User Profile "
                "before generating a Letter."
            )

        elif not (
            st.session_state.letter_instruction.strip()
        ):

            st.warning(
                "Please enter or record your instructions."
            )

        else:

            with st.spinner(
                "Generating professional letter..."
            ):

                result = generate_letter(
                    st.session_state.letter_instruction
                )

            if result:

                final_document = (
                    result
                    + get_signature(
                        st.session_state.profile
                    )
                )

                st.session_state.generated_draft = (
                    final_document
                )

                st.session_state.editable_draft = (
                    final_document
                )

                st.session_state.editor_sync = (
                    final_document
                )

                save_to_history(
                    "Letter",
                    final_document,
                )

                st.rerun()

            else:

                st.error(
                    "AI generation failed. "
                    "Please check your API key."
                )


# ============================================================
# INQUIRY
# ============================================================

elif document_type == "Inquiry":

    st.markdown(
        "### 🔎 Inquiry"
    )

    inquiry_type = st.radio(
        "Select Inquiry Type",
        [
            "E&D Inquiry",
            "FFI Inquiry",
        ],
        horizontal=True,
    )


    # ========================================================
    # FFI
    # ========================================================

    if inquiry_type == "FFI Inquiry":

        st.info(
            "FFI Inquiry is currently Under Construction / "
            "Under Process."
        )


    # ========================================================
    # E&D
    # ========================================================

    else:

        st.markdown(
            "### E&D Inquiry"
        )

        st.caption(
            "Add the required indexes. "
            "The same index can be added multiple times."
        )

        available_indexes = st.selectbox(
            "Select Index to Add",
            ED_INDEXES,
            key="selected_ed_index",
        )

        if st.button(
            "➕ Add Index",
            use_container_width=True,
        ):

            add_inquiry_index(
                available_indexes
            )

            st.rerun()


        st.divider()


        if not st.session_state.inquiry_indexes:

            st.info(
                "No indexes added yet. "
                "Select an index above and click Add Index."
            )

        else:

            collected_data = []

            for item in (
                st.session_state.inquiry_indexes
            ):

                col1, col2 = st.columns(
                    [8, 1]
                )

                with col1:

                    result = render_index_input(
                        item
                    )

                with col2:

                    st.write("")

                    if st.button(
                        "🗑️",
                        key=f"delete_{item['id']}",
                        help="Remove this index",
                    ):

                        remove_inquiry_index(
                            item["id"]
                        )

                        st.rerun()


                if result:

                    collected_data.append(
                        {
                            "name": item["name"],
                            "number": item["number"],
                            **result,
                        }
                    )

                st.divider()


            # =================================================
            # GENERATE
            # =================================================

            if st.button(
                "✨ Generate Inquiry Report",
                type="primary",
                use_container_width=True,
            ):

                if not collected_data:

                    st.warning(
                        "Please add at least one inquiry index."
                    )

                else:

                    with st.spinner(
                        "Preparing E&D Inquiry Report..."
                    ):

                        result = generate_ed_report(
                            collected_data
                        )

                    if result:

                        st.session_state.generated_draft = (
                            result
                        )

                        st.session_state.editable_draft = (
                            result
                        )

                        st.session_state.editor_sync = (
                            result
                        )

                        save_to_history(
                            "E&D Inquiry",
                            result,
                        )

                        st.rerun()

                    else:

                        st.error(
                            "AI generation failed. "
                            "Please check your API key."
                        )


# ============================================================
# GENERATED DOCUMENT
# ============================================================

if st.session_state.generated_draft:

    st.divider()

    st.markdown(
        "## 📑 Generated Document"
    )


    # ========================================================
    # EDITOR SYNC
    # ========================================================

    if (
        st.session_state.get(
            "editor_sync",
            "",
        )
        != ""
    ):

        st.session_state.document_editor = (
            st.session_state.editor_sync
        )

        st.session_state.editor_sync = ""


    elif not st.session_state.get(
        "document_editor",
        "",
    ):

        st.session_state.document_editor = (

            st.session_state.editable_draft
            or
            st.session_state.generated_draft

        )


    # ========================================================
    # MANUAL EDITOR
    # ========================================================

    st.markdown(
        "### ✏️ Review & Modify Document"
    )


    edited_document = st.text_area(

        "Generated Document",

        key="document_editor",

        height=650,

        label_visibility="collapsed",

    )


    col1, col2 = st.columns(2)


    with col1:

        if st.button(
            "💾 Save Changes",
            use_container_width=True,
        ):

            st.session_state.editable_draft = (
                edited_document
            )

            st.session_state.generated_draft = (
                edited_document
            )

            save_to_history(
                document_type,
                edited_document,
            )

            st.success(
                "Changes saved."
            )


    with col2:

        if st.button(
            "↩️ Restore Generated Version",
            use_container_width=True,
        ):

            original = (
                st.session_state.generated_draft
            )

            st.session_state.editable_draft = (
                original
            )

            st.session_state.editor_sync = (
                original
            )

            st.rerun()


    # ========================================================
    # AI MODIFICATION
    # ========================================================

    st.markdown(
        "### 🤖 Ask AI to Modify the Document"
    )


    # --------------------------------------------------------
    # CRITICAL FIX:
    #
    # Apply synchronization BEFORE text_area is created.
    # --------------------------------------------------------

    if (
        "edit_instruction_sync"
        in st.session_state
    ):

        sync_value = (
            st.session_state.pop(
                "edit_instruction_sync"
            )
        )

        if sync_value is not None:

            st.session_state.edit_instruction = (
                sync_value
            )


    st.text_area(

        "Modification Instructions",

        key="edit_instruction",

        height=130,

        placeholder=(
            "Example: Make the language more formal, "
            "shorten the introduction, or revise the "
            "recommendation section..."
        ),

    )


    if st.button(
        "🤖 Apply AI Changes",
        type="primary",
        use_container_width=True,
    ):

        instruction = (
            st.session_state.edit_instruction.strip()
        )

        current_document = (

            st.session_state.editable_draft
            or
            st.session_state.document_editor
            or
            st.session_state.generated_draft

        )


        if not instruction:

            st.warning(
                "Please enter instructions for the AI."
            )

        elif not current_document.strip():

            st.warning(
                "There is no document to modify."
            )

        else:

            with st.spinner(
                "AI is modifying the document..."
            ):

                modified = (
                    modify_generated_document(
                        current_document,
                        instruction,
                    )
                )


            if modified:

                st.session_state.editable_draft = (
                    modified
                )

                st.session_state.generated_draft = (
                    modified
                )

                # IMPORTANT:
                # Do NOT modify document_editor here.
                # It is already instantiated.
                st.session_state.editor_sync = (
                    modified
                )

                # IMPORTANT:
                # Do NOT modify edit_instruction directly.
                #
                # Instead clear it before the next widget
                # instantiation.
                st.session_state.edit_instruction_sync = (
                    ""
                )


                save_to_history(
                    document_type,
                    modified,
                )

                st.rerun()


            else:

                st.error(
                    "Could not modify document. "
                    "Please check your API key."
                )


    # ========================================================
    # PREVIEW
    # ========================================================

    st.markdown(
        "### 👁️ Document Preview"
    )

    preview_document = (

        st.session_state.editable_draft
        or
        edited_document
        or
        st.session_state.generated_draft

    )

    display_generated_document(
        preview_document
    )


    # ========================================================
    # EXPORT
    # ========================================================

    st.markdown(
        "### 📤 Export Document"
    )

    export_document = (

        st.session_state.editable_draft
        or
        edited_document
        or
        st.session_state.generated_draft

    )


    col1, col2, col3, col4 = (
        st.columns(4)
    )


    # ========================================================
    # PDF
    # ========================================================

    with col1:

        try:

            pdf_bytes = create_pdf(
                export_document
            )

            st.download_button(
                "📕 PDF",
                data=pdf_bytes,
                file_name="draftforge_document.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

        except Exception as e:

            st.error(
                f"PDF export failed: {e}"
            )


    # ========================================================
    # DOCX
    # ========================================================

    with col2:

        try:

            docx_bytes = create_docx(
                export_document
            )

            st.download_button(
                "📘 DOCX",
                data=docx_bytes,
                file_name="draftforge_document.docx",
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.wordprocessingml.document"
                ),
                use_container_width=True,
            )

        except Exception as e:

            st.error(
                f"DOCX export failed: {e}"
            )


    # ========================================================
    # TXT
    # ========================================================

    with col3:

        txt_bytes = create_txt(
            export_document
        )

        st.download_button(
            "📄 TXT",
            data=txt_bytes,
            file_name="draftforge_document.txt",
            mime="text/plain",
            use_container_width=True,
        )


    # ========================================================
    # PNG
    # ========================================================

    with col4:

        try:

            png_bytes = create_png(
                export_document
            )

            st.download_button(
                "🖼️ PNG",
                data=png_bytes,
                file_name="draftforge_document.png",
                mime="image/png",
                use_container_width=True,
            )

        except Exception as e:

            st.error(
                f"PNG export failed: {e}"
            )


# ============================================================
# FOOTER
# ============================================================

st.markdown("---")

st.caption(
    "DraftForge — AI Document Composer"
)

st.caption(
    "Developed by Raees Khan, Assistant Director, NADRA"
)
