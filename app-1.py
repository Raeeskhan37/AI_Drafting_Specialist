import io
import os
import re
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
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="DraftForge — AI Document Composer",
    page_icon="📝",
    layout="wide",
)


# ============================================================
# CONSTANTS
# ============================================================

GROQ_MODEL = "openai/gpt-oss-120b"
WHISPER_MODEL = "whisper-large-v3"
GEMINI_MODEL = "gemini-2.0-flash"


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


COMMITTEE_ROLES = [
    "Convener of Inquiry",
    "Member 1",
    "Member 2",
    "Departmental Representative",
]


# ============================================================
# CUSTOM CSS
# ============================================================

st.markdown(
    """
<style>

.main-title {
    font-size: 38px;
    font-weight: 800;
    margin-bottom: 0px;
}

.subtitle {
    font-size: 17px;
    color: #777;
    margin-bottom: 22px;
}

.composer-box {
    border: 1px solid rgba(128,128,128,0.35);
    border-radius: 14px;
    padding: 16px;
    margin-top: 8px;
    margin-bottom: 18px;
    background: rgba(128,128,128,0.04);
}

.voice-heading {
    font-size: 15px;
    font-weight: 600;
    margin-bottom: 5px;
}

.voice-help {
    font-size: 13px;
    color: #888;
    margin-bottom: 8px;
}

.index-card {
    border: 1px solid rgba(128,128,128,0.35);
    border-radius: 14px;
    padding: 15px;
    margin-bottom: 15px;
}

.warning-box {
    padding: 18px;
    border-radius: 12px;
    background: rgba(255,165,0,0.10);
    border-left: 5px solid orange;
}

</style>
""",
    unsafe_allow_html=True,
)


# ============================================================
# SESSION STATE INITIALIZATION
# ============================================================

def initialize_state():

    defaults = {
        "document_type": "Email",
        "inquiry_type": "E&D Inquiry",

        "generated_draft": "",

        "email_input": "",
        "letter_input": "",

        "email_audio_seen": "",
        "letter_audio_seen": "",

        "ed_instances": [],

        "history": [],
    }

    for key, value in defaults.items():

        if key not in st.session_state:
            st.session_state[key] = value


initialize_state()


# ============================================================
# SECRET HELPER
# ============================================================

def get_secret(name, default=None):

    try:

        value = st.secrets.get(name)

        if value:
            return value

    except Exception:
        pass

    return os.getenv(name, default)


# ============================================================
# GROQ CLIENT
# ============================================================

@st.cache_resource
def get_groq_client():

    api_key = get_secret("GROQ_API_KEY")

    if not api_key:
        return None

    return Groq(api_key=api_key)


# ============================================================
# TEXT APPEND HELPER
# ============================================================

def append_text(old_text, new_text):

    old_text = (old_text or "").strip()
    new_text = (new_text or "").strip()

    if not new_text:
        return old_text

    if not old_text:
        return new_text

    return old_text + "\n" + new_text


# ============================================================
# VOICE TRANSCRIPTION
# ============================================================

def transcribe_audio(audio_file):

    client = get_groq_client()

    if client is None:

        raise RuntimeError(
            "GROQ_API_KEY is missing from Streamlit Secrets."
        )

    audio_bytes = audio_file.getvalue()

    audio_buffer = io.BytesIO(audio_bytes)

    audio_buffer.name = "recording.wav"

    result = client.audio.transcriptions.create(
        file=audio_buffer,
        model=WHISPER_MODEL,
        response_format="text",
    )

    if hasattr(result, "text"):

        return result.text.strip()

    return str(result).strip()


# ============================================================
# AUDIO SIGNATURE
# ============================================================

def audio_signature(audio):

    if audio is None:
        return ""

    data = audio.getvalue()

    return hashlib.md5(data).hexdigest()


# ============================================================
# MARKUP CLEANING
# ============================================================

def clean_markup(text):

    if not text:
        return ""

    text = re.sub(
        r"\*\*<u>(.*?)</u>\*\*",
        r"\1",
        text,
    )

    text = re.sub(
        r"<u>(.*?)</u>",
        r"\1",
        text,
    )

    text = text.replace("**", "")
    text = text.replace("__", "")

    return text


# ============================================================
# AI GENERATION
# ============================================================

def generate_ai(system_prompt, user_prompt):

    client = get_groq_client()

    if client:

        try:

            response = client.chat.completions.create(
                model=GROQ_MODEL,
                temperature=0.15,
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
            )

            return response.choices[0].message.content.strip()

        except Exception as exc:

            st.warning(
                f"Groq generation failed: {exc}"
            )

    # --------------------------------------------------------
    # Gemini fallback
    # --------------------------------------------------------

    gemini_key = get_secret("GEMINI_API_KEY")

    if gemini_key:

        try:

            url = (
                "https://generativelanguage.googleapis.com/"
                "v1beta/models/"
                f"{GEMINI_MODEL}:generateContent"
                f"?key={gemini_key}"
            )

            payload = {
                "contents": [
                    {
                        "parts": [
                            {
                                "text":
                                    system_prompt
                                    + "\n\n"
                                    + user_prompt
                            }
                        ]
                    }
                ]
            }

            response = requests.post(
                url,
                json=payload,
                timeout=120,
            )

            response.raise_for_status()

            data = response.json()

            return (
                data["candidates"][0]
                ["content"]
                ["parts"][0]
                ["text"]
                .strip()
            )

        except Exception as exc:

            st.error(
                f"AI generation failed: {exc}"
            )

    return None


# ============================================================
# AI RULES
# ============================================================

LANGUAGE_RULES = """
Correct spelling, grammar, punctuation and sentence structure.

Correct obvious typing mistakes.

Correct obvious voice-transcription mistakes where the intended
meaning is clear.

Convert informal wording into professional official language.

DO NOT invent facts.

DO NOT invent:
- names
- dates
- allegations
- evidence
- witnesses
- documents
- findings
- recommendations
- reference numbers
- events

Preserve the user's actual meaning.

For statements of accused persons and witnesses, preserve their
meaning faithfully and do not turn allegations into established facts.
"""


NORMAL_SYSTEM_PROMPT = f"""
You are DraftForge, an AI official-document drafting assistant.

{LANGUAGE_RULES}

Prepare professional government/official correspondence.
"""


ED_SYSTEM_PROMPT = f"""
You are DraftForge, an AI assistant for preparing formal
Departmental Inquiry Reports.

{LANGUAGE_RULES}

The user's selected sections must remain in EXACTLY the same order.

Repeated sections must remain separate.

Do not merge repeated sections.

Do not remove supplied information.

Do not invent additional facts.

Every major heading must be formatted:

**<u>HEADING</u>**
"""


# ============================================================
# E&D INSTANCE FUNCTIONS
# ============================================================

def add_ed_index(index_name):

    instance_id = len(
        st.session_state.ed_instances
    ) + 1

    st.session_state.ed_instances.append(
        {
            "id": instance_id,
            "index": index_name,
            "text": "",
            "audio_seen": "",
        }
    )


def remove_ed_index(instance_id):

    st.session_state.ed_instances = [
        item
        for item in st.session_state.ed_instances
        if item["id"] != instance_id
    ]


def get_occurrence(index_name, instance_id):

    count = 0

    for item in st.session_state.ed_instances:

        if item["index"] == index_name:

            count += 1

        if item["id"] == instance_id:

            return count

    return 1


def get_display_heading(index_name, occurrence):

    if index_name in [
        "Statement of the Accused",
        "Questions / Answers with the Accused",
    ]:

        if occurrence > 1:

            return (
                f"{index_name} No. {occurrence}"
            )

    return index_name


# ============================================================
# E&D HEADING FORMAT
# ============================================================

def format_ed_headings(text):

    if not text:
        return text

    headings = [
        "DEPARTMENTAL INQUIRY REPORT",
        "Inquiry Reference No.",
        "Date",
    ] + ED_INDEXES

    for h in headings:

        pattern = re.compile(
            rf"(?mi)^(\s*){re.escape(h)}\s*$"
        )

        text = pattern.sub(
            rf"\1**<u>{h}</u>**",
            text,
        )

    text = re.sub(
        r"(?mi)^(\s*)Statement of the Accused No\. (\d+)\s*$",
        r"\1**<u>Statement of the Accused No. \2</u>**",
        text,
    )

    text = re.sub(
        r"(?mi)^(\s*)Questions / Answers with the Accused No\. (\d+)\s*$",
        r"\1**<u>Questions / Answers with the Accused No. \2</u>**",
        text,
    )

    return text


# ============================================================
# EMAIL / LETTER COMPOSER
# ============================================================

def render_composer(prefix, title):

    text_key = f"{prefix}_input"
    audio_key = f"{prefix}_audio"
    seen_key = f"{prefix}_audio_seen"

    # Ensure state exists BEFORE audio widget.
    if text_key not in st.session_state:
        st.session_state[text_key] = ""

    if seen_key not in st.session_state:
        st.session_state[seen_key] = ""

    st.markdown(
        '<div class="composer-box">',
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="voice-heading">'
        "🎙️ Voice + Text Input"
        "</div>",
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="voice-help">'
        "Tap the microphone to record. You do NOT need "
        "to type anything first."
        "</div>",
        unsafe_allow_html=True,
    )

    # --------------------------------------------------------
    # AUDIO INPUT
    # --------------------------------------------------------

    audio = st.audio_input(
        "🎙️ Record Voice",
        sample_rate=16000,
        key=audio_key,
        width="stretch",
    )

    # --------------------------------------------------------
    # PROCESS NEW AUDIO BEFORE TEXT AREA
    # --------------------------------------------------------

    if audio is not None:

        signature = audio_signature(audio)

        if signature != st.session_state[seen_key]:

            with st.spinner(
                "🎙️ Transcribing your voice..."
            ):

                try:

                    spoken_text = transcribe_audio(
                        audio
                    )

                    st.session_state[text_key] = (
                        append_text(
                            st.session_state[text_key],
                            spoken_text,
                        )
                    )

                    st.session_state[
                        seen_key
                    ] = signature

                    st.success(
                        "Voice successfully added to the input."
                    )

                except Exception as exc:

                    st.error(str(exc))

    # --------------------------------------------------------
    # MAIN INPUT BOX
    # --------------------------------------------------------

    text = st.text_area(
        title,
        key=text_key,
        height=180,
        placeholder=(
            "Type your instructions here. "
            "You can also use the microphone above. "
            "Typed and spoken information will be combined."
        ),
        label_visibility="collapsed",
    )

    st.markdown(
        "</div>",
        unsafe_allow_html=True,
    )

    return text


# ============================================================
# E&D VOICE + TEXT INPUT
# ============================================================

def render_ed_input(item):

    instance_id = item["id"]

    text_key = f"ed_text_{instance_id}"
    audio_key = f"ed_audio_{instance_id}"

    # --------------------------------------------------------
    # Synchronize initial state
    # --------------------------------------------------------

    if text_key not in st.session_state:

        st.session_state[text_key] = (
            item.get("text", "")
        )

    # --------------------------------------------------------
    # AUDIO
    # --------------------------------------------------------

    audio = st.audio_input(
        "🎙️ Record Voice",
        sample_rate=16000,
        key=audio_key,
        width="stretch",
    )

    if audio is not None:

        signature = audio_signature(audio)

        if item.get("audio_seen") != signature:

            with st.spinner(
                "🎙️ Transcribing voice..."
            ):

                try:

                    spoken_text = (
                        transcribe_audio(audio)
                    )

                    st.session_state[text_key] = (
                        append_text(
                            st.session_state.get(
                                text_key,
                                "",
                            ),
                            spoken_text,
                        )
                    )

                    item["text"] = (
                        st.session_state[text_key]
                    )

                    item["audio_seen"] = signature

                    st.success(
                        "Voice successfully added to this index."
                    )

                except Exception as exc:

                    st.error(str(exc))

    # --------------------------------------------------------
    # TEXT AREA
    # --------------------------------------------------------

    current_text = st.text_area(
        "Information",
        key=text_key,
        height=180,
        placeholder=(
            "Type information naturally here, "
            "or use the microphone above."
        ),
        label_visibility="collapsed",
    )

    item["text"] = current_text

    return current_text


# ============================================================
# EMAIL
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
Create a professional official email.

Recipient:
{recipient}

Subject:
{subject}

User instructions:
{instructions}

Use formal professional language.

Preserve the user's intended meaning.

Do not invent facts.
"""

    return generate_ai(
        NORMAL_SYSTEM_PROMPT,
        prompt,
    )


# ============================================================
# LETTER
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

    today = datetime.now().strftime(
        "%d %B %Y"
    )

    prompt = f"""
Prepare a professional official letter.

Date:
{today}

Recipient / Office:
{recipient}

Subject:
{subject}

User instructions:
{instructions}

Use formal professional official language.

Do not invent facts.
"""

    return generate_ai(
        NORMAL_SYSTEM_PROMPT,
        prompt,
    )


# ============================================================
# DOCUMENTS RECORDED
# ============================================================

def render_documents_recorded():

    st.markdown(
        "### Documents Recorded"
    )

    selected = []

    columns = st.columns(2)

    for i, document in enumerate(
        DOCUMENTS_RECORDED
    ):

        with columns[i % 2]:

            if st.checkbox(
                document,
                key=f"document_recorded_{i}",
            ):

                selected.append(document)

    return selected


# ============================================================
# INQUIRY COMMITTEE
# ============================================================

def render_committee():

    st.markdown(
        "### Inquiry Committee"
    )

    committee = []

    for role in COMMITTEE_ROLES:

        st.markdown(
            f"**{role}**"
        )

        c1, c2, c3 = st.columns(
            [1, 2, 2]
        )

        with c1:

            erp = st.text_input(
                "ERP#",
                key=f"erp_{role}",
            )

        with c2:

            name = st.text_input(
                "Name",
                key=f"name_{role}",
            )

        with c3:

            designation = st.text_input(
                "Designation",
                key=f"designation_{role}",
            )

        committee.append(
            {
                "role": role,
                "erp": erp,
                "name": name,
                "designation": designation,
            }
        )

    return committee


# ============================================================
# E&D REPORT GENERATION
# ============================================================

def generate_ed_report(
    reference_no,
    documents_recorded,
    committee,
):

    today = datetime.now().strftime(
        "%d %B %Y"
    )

    sections = []

    for item in st.session_state.ed_instances:

        occurrence = get_occurrence(
            item["index"],
            item["id"],
        )

        display_heading = (
            get_display_heading(
                item["index"],
                occurrence,
            )
        )

        sections.append(
            f"### {display_heading}\n"
            f"{item.get('text', '').strip()}"
        )

    sections_text = "\n\n".join(
        sections
    )

    documents_text = (
        "\n".join(
            f"- {document}"
            for document in documents_recorded
        )
        if documents_recorded
        else "No documents selected."
    )

    committee_lines = []

    for member in committee:

        committee_lines.append(
            f"- {member['role']}: "
            f"ERP# {member['erp']}; "
            f"Name: {member['name']}; "
            f"Designation: {member['designation']}"
        )

    committee_text = (
        "\n".join(committee_lines)
        if committee_lines
        else "Not provided."
    )

    prompt = f"""
Prepare the Departmental Inquiry Report.

Inquiry Reference No.:
{reference_no}

Date:
{today}

============================================================
SELECTED E&D SECTIONS
============================================================

The following sections were selected by the user.

Preserve their EXACT order.

Preserve repeated occurrences separately.

{sections_text}

============================================================
DOCUMENTS RECORDED
============================================================

{documents_text}

============================================================
INQUIRY COMMITTEE
============================================================

{committee_text}

============================================================

Important instructions:

1. Do not invent facts.
2. Do not omit supplied information.
3. Do not merge repeated sections.
4. Preserve the meaning of accused/witness statements.
5. Correct spelling and grammar.
6. Use formal official language.
7. Do not manufacture findings or conclusions.
8. Use the current date:
   {today}
9. Use:
   DEPARTMENTAL INQUIRY REPORT

10. Format every major heading as:

**<u>HEADING</u>**
"""

    result = generate_ai(
        ED_SYSTEM_PROMPT,
        prompt,
    )

    if result:

        return format_ed_headings(
            result
        )

    return None


# ============================================================
# PDF EXPORT
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
        "…": "...",
        "→": "->",
        "\u00a0": " ",
    }

    for old, new in replacements.items():

        text = text.replace(
            old,
            new,
        )

    text = clean_markup(text)

    return (
        text
        .encode("latin-1", "ignore")
        .decode("latin-1")
    )


def is_heading(line):

    line = clean_markup(
        line
    ).strip()

    if line in [
        "DEPARTMENTAL INQUIRY REPORT",
        "Inquiry Reference No.",
        "Date",
    ]:

        return True

    for heading in ED_INDEXES:

        if line.lower() == heading.lower():

            return True

        if re.match(
            rf"^{re.escape(heading)} No\. \d+$",
            line,
            re.IGNORECASE,
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

        return [""]

    lines = []

    current = ""

    for word in words:

        if pdf.get_string_width(word) > width:

            if current:

                lines.append(current)
                current = ""

            piece = ""

            for char in word:

                test = piece + char

                if (
                    pdf.get_string_width(test)
                    <= width
                ):

                    piece = test

                else:

                    if piece:

                        lines.append(piece)

                    piece = char

            current = piece

            continue

        test = (
            word
            if not current
            else current + " " + word
        )

        if (
            pdf.get_string_width(test)
            <= width
        ):

            current = test

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

    pdf.set_margins(
        15,
        15,
        15,
    )

    pdf.add_page()

    usable_width = (
        pdf.w
        - pdf.l_margin
        - pdf.r_margin
    )

    for raw_line in text.splitlines():

        line = pdf_clean(
            raw_line
        )

        if not line:

            pdf.ln(4)
            continue

        if is_heading(line):

            pdf.set_font(
                "Helvetica",
                "BU",
                11,
            )

        else:

            pdf.set_font(
                "Helvetica",
                "",
                11,
            )

        lines = wrap_pdf_line(
            pdf,
            line,
            usable_width,
        )

        for wrapped in lines:

            pdf.multi_cell(
                usable_width,
                7,
                wrapped,
            )

    return bytes(
        pdf.output()
    )


# ============================================================
# DOCX EXPORT
# ============================================================

def export_docx(text):

    document = Document()

    for raw_line in text.splitlines():

        line = clean_markup(
            raw_line
        )

        paragraph = (
            document.add_paragraph()
        )

        if not line:

            continue

        run = paragraph.add_run(
            line
        )

        run.font.size = Pt(11)

        if is_heading(line):

            run.bold = True
            run.underline = True

    buffer = io.BytesIO()

    document.save(buffer)

    return buffer.getvalue()


# ============================================================
# TXT EXPORT
# ============================================================

def export_txt(text):

    return clean_markup(
        text
    ).encode("utf-8")


# ============================================================
# PNG EXPORT
# ============================================================

def export_png(text):

    clean_text = clean_markup(
        text
    )

    lines = []

    for paragraph in clean_text.splitlines():

        if not paragraph:

            lines.append("")
            continue

        words = paragraph.split()

        current = ""

        for word in words:

            test = (
                word
                if not current
                else current + " " + word
            )

            if len(test) <= 95:

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

    try:

        font = ImageFont.truetype(
            "DejaVuSans.ttf",
            22,
        )

        bold_font = ImageFont.truetype(
            "DejaVuSans-Bold.ttf",
            22,
        )

    except Exception:

        font = ImageFont.load_default()
        bold_font = font

    width = 1600

    height = max(
        500,
        100 + 34 * len(lines),
    )

    image = Image.new(
        "RGB",
        (width, height),
        "white",
    )

    draw = ImageDraw.Draw(
        image
    )

    y = 50

    for line in lines:

        current_font = (
            bold_font
            if is_heading(line)
            else font
        )

        draw.text(
            (50, y),
            line,
            fill="black",
            font=current_font,
        )

        if is_heading(line):

            bbox = draw.textbbox(
                (50, y),
                line,
                font=current_font,
            )

            draw.line(
                (
                    bbox[0],
                    bbox[3] + 2,
                    bbox[2],
                    bbox[3] + 2,
                ),
                fill="black",
                width=1,
            )

        y += 34

    output = io.BytesIO()

    image.save(
        output,
        format="PNG",
    )

    return output.getvalue()


# ============================================================
# HISTORY
# ============================================================

def save_history(
    document_type,
    draft,
):

    if not draft:
        return

    st.session_state.history.insert(
        0,
        {
            "type": document_type,
            "date": datetime.now().strftime(
                "%d %B %Y %H:%M"
            ),
            "draft": draft,
        },
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

    st.divider()

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

**4. Export**

PDF • DOCX • TXT • PNG
"""
    )

    st.divider()

    if st.button(
        "🗑️ Clear Generated Draft",
        use_container_width=True,
    ):

        st.session_state.generated_draft = ""

        st.rerun()

    if st.session_state.history:

        st.divider()

        st.markdown(
            "### Recent Drafts"
        )

        for item in (
            st.session_state.history[:5]
        ):

            with st.expander(
                f"{item['type']} — {item['date']}"
            ):

                st.text(
                    clean_markup(
                        item["draft"]
                    )[:1000]
                )


# ============================================================
# MAIN HEADER
# ============================================================

st.markdown(
    '<div class="main-title">'
    "📝 DraftForge"
    "</div>",
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="subtitle">'
    "AI Document Composer — create professional "
    "official documents using text or voice"
    "</div>",
    unsafe_allow_html=True,
)


# ============================================================
# PROMINENT DOCUMENT SELECTOR
# ============================================================

st.markdown(
    "### 📌 What would you like to create?"
)

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

if document_choice:

    if "Email" in document_choice:

        st.session_state.document_type = "Email"

    elif "Letter" in document_choice:

        st.session_state.document_type = "Letter"

    elif "Inquiry" in document_choice:

        st.session_state.document_type = "Inquiry"


current_type = (
    st.session_state.document_type
)

st.divider()


# ============================================================
# EMAIL UI
# ============================================================

if current_type == "Email":

    st.markdown(
        "## ✉️ Create Email"
    )

    st.text_input(
        "Recipient",
        placeholder="Enter recipient or office",
        key="email_recipient",
    )

    st.text_input(
        "Subject",
        placeholder="Enter email subject",
        key="email_subject",
    )

    st.markdown(
        "### Email Instructions"
    )

    email_text = render_composer(
        "email",
        "Email Instructions",
    )

    # IMPORTANT:
    # Generate button is AFTER the complete input area.

    if st.button(
        "✨ Generate Email",
        type="primary",
        use_container_width=True,
    ):

        if not email_text.strip():

            st.warning(
                "Please type or record your email instructions."
            )

        else:

            with st.spinner(
                "DraftForge is preparing your email..."
            ):

                draft = generate_email()

            if draft:

                st.session_state.generated_draft = (
                    draft
                )

                save_history(
                    "Email",
                    draft,
                )

                st.rerun()


# ============================================================
# LETTER UI
# ============================================================

elif current_type == "Letter":

    st.markdown(
        "## 📄 Create Official Letter"
    )

    st.text_input(
        "Recipient / Office",
        placeholder=(
            "Enter recipient, office or designation"
        ),
        key="letter_recipient",
    )

    st.text_input(
        "Subject",
        placeholder="Enter letter subject",
        key="letter_subject",
    )

    st.markdown(
        "### Letter Instructions"
    )

    letter_text = render_composer(
        "letter",
        "Letter Instructions",
    )

    # IMPORTANT:
    # Generate button is AFTER the complete input area.

    if st.button(
        "✨ Generate Letter",
        type="primary",
        use_container_width=True,
    ):

        if not letter_text.strip():

            st.warning(
                "Please type or record your letter instructions."
            )

        else:

            with st.spinner(
                "DraftForge is preparing your letter..."
            ):

                draft = generate_letter()

            if draft:

                st.session_state.generated_draft = (
                    draft
                )

                save_history(
                    "Letter",
                    draft,
                )

                st.rerun()


# ============================================================
# INQUIRY UI
# ============================================================

elif current_type == "Inquiry":

    st.markdown(
        "## 🔎 Departmental Inquiry"
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

    if inquiry_choice is None:

        inquiry_choice = "⚖️ E&D Inquiry"

    # --------------------------------------------------------
    # FFI
    # --------------------------------------------------------

    if "FFI" in inquiry_choice:

        st.session_state.inquiry_type = (
            "FFI Inquiry"
        )

        st.markdown(
            """
            <div class="warning-box">

            <h3>🔍 FFI Inquiry</h3>

            <b>Under Construction / Under Process</b>

            <br><br>

            The FFI inquiry module is currently being developed.

            </div>
            """,
            unsafe_allow_html=True,
        )

    # --------------------------------------------------------
    # E&D
    # --------------------------------------------------------

    else:

        st.session_state.inquiry_type = (
            "E&D Inquiry"
        )

        st.text_input(
            "Inquiry Reference No.",
            placeholder="e.g. ABC/123",
            key="inquiry_reference_no",
        )

        st.info(
            "Select an index, add it, then provide "
            "information by typing or voice. "
            "The same index can be selected multiple times."
        )

        # ----------------------------------------------------
        # INDEX SELECTOR
        # ----------------------------------------------------

        selected_index = st.selectbox(
            "➕ Select E&D Index",
            ED_INDEXES,
            key="ed_index_selector",
        )

        if st.button(
            "➕ Add Selected Index",
            use_container_width=True,
        ):

            add_ed_index(
                selected_index
            )

            st.rerun()

        # ----------------------------------------------------
        # E&D INSTANCES
        # ----------------------------------------------------

        if st.session_state.ed_instances:

            st.markdown(
                "### 📋 Inquiry Information"
            )

            for item in list(
                st.session_state.ed_instances
            ):

                occurrence = get_occurrence(
                    item["index"],
                    item["id"],
                )

                display_heading = (
                    get_display_heading(
                        item["index"],
                        occurrence,
                    )
                )

                with st.container(
                    border=True
                ):

                    col1, col2 = st.columns(
                        [8, 1]
                    )

                    with col1:

                        st.markdown(
                            f"### {display_heading}"
                        )

                    with col2:

                        if st.button(
                            "🗑️",
                            key=f"delete_{item['id']}",
                            help="Remove this section",
                        ):

                            remove_ed_index(
                                item["id"]
                            )

                            st.rerun()

                    st.caption(
                        "Type naturally or tap 🎙️ to record. "
                        "Voice transcription is added to "
                        "the same input."
                    )

                    render_ed_input(
                        item
                    )

        # ----------------------------------------------------
        # DOCUMENTS RECORDED
        # ----------------------------------------------------

        documents_recorded = []

        if any(
            item["index"]
            == "Documents Recorded"
            for item
            in st.session_state.ed_instances
        ):

            documents_recorded = (
                render_documents_recorded()
            )

        # ----------------------------------------------------
        # INQUIRY COMMITTEE
        # ----------------------------------------------------

        committee = []

        if any(
            item["index"]
            == "Inquiry Committee"
            for item
            in st.session_state.ed_instances
        ):

            committee = render_committee()

        # ----------------------------------------------------
        # FINAL ACTION BUTTONS
        # ----------------------------------------------------

        st.divider()

        if st.button(
            "✨ Generate Inquiry Report",
            type="primary",
            use_container_width=True,
        ):

            reference_no = st.session_state.get(
                "inquiry_reference_no",
                "",
            )

            if not reference_no.strip():

                st.warning(
                    "Please enter the Inquiry Reference No."
                )

            elif not st.session_state.ed_instances:

                st.warning(
                    "Please add at least one inquiry index."
                )

            elif not any(
                item.get("text", "").strip()
                for item
                in st.session_state.ed_instances
            ):

                st.warning(
                    "Please provide information in "
                    "at least one inquiry index."
                )

            else:

                with st.spinner(
                    "DraftForge is preparing the "
                    "Departmental Inquiry Report..."
                ):

                    draft = generate_ed_report(
                        reference_no,
                        documents_recorded,
                        committee,
                    )

                if draft:

                    st.session_state.generated_draft = (
                        draft
                    )

                    save_history(
                        "E&D Inquiry",
                        draft,
                    )

                    st.rerun()

        if st.button(
            "🔄 Start New Inquiry",
            use_container_width=True,
        ):

            st.session_state.ed_instances = []

            st.session_state.generated_draft = ""

            st.rerun()


# ============================================================
# GENERATED DOCUMENT
# ============================================================

if st.session_state.generated_draft:

    st.divider()

    st.markdown(
        "## 📄 Generated Document"
    )

    st.markdown(
        st.session_state.generated_draft,
        unsafe_allow_html=True,
    )

    st.divider()

    st.markdown(
        "### 📥 Download"
    )

    pdf_data = export_pdf(
        st.session_state.generated_draft
    )

    docx_data = export_docx(
        st.session_state.generated_draft
    )

    txt_data = export_txt(
        st.session_state.generated_draft
    )

    png_data = export_png(
        st.session_state.generated_draft
    )

    c1, c2, c3, c4 = st.columns(4)

    with c1:

        st.download_button(
            "📕 PDF",
            data=pdf_data,
            file_name="DraftForge_Document.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    with c2:

        st.download_button(
            "📘 DOCX",
            data=docx_data,
            file_name="DraftForge_Document.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument"
                ".wordprocessingml.document"
            ),
            use_container_width=True,
        )

    with c3:

        st.download_button(
            "📄 TXT",
            data=txt_data,
            file_name="DraftForge_Document.txt",
            mime="text/plain",
            use_container_width=True,
        )

    with c4:

        st.download_button(
            "🖼️ PNG",
            data=png_data,
            file_name="DraftForge_Document.png",
            mime="image/png",
            use_container_width=True,
        )
