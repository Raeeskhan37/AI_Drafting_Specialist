import os
import re
import io
import json
import hashlib
import textwrap
from datetime import datetime

import streamlit as st

# ============================================================
# OPTIONAL DEPENDENCIES
# ============================================================

try:
    from groq import Groq
except Exception:
    Groq = None

try:
    from docx import Document
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
except Exception:
    Document = None

try:
    from fpdf import FPDF
except Exception:
    FPDF = None

try:
    from PIL import Image, ImageDraw, ImageFont
except Exception:
    Image = None


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="DraftForge — AI Document Composer",
    page_icon="✦",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# CONSTANTS
# ============================================================

PROFILE_FILE = "user_profile.json"
HISTORY_FILE = "draftforge_history.json"

GROQ_MODEL = "openai/gpt-oss-120b"
WHISPER_MODEL = "whisper-large-v3"

DOCUMENT_TYPES = [
    "Email",
    "Letter",
    "Inquiry",
]

INQUIRY_TYPES = [
    "E&D Inquiry",
    "FFI Inquiry",
]

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
# DEFAULT STATE
# ============================================================

DEFAULT_STATE = {
    "document_type": "Email",
    "inquiry_type": "E&D Inquiry",
    "index_data": [],
    "generated_draft": "",
    "editable_draft": "",
    "document_editor": "",
    "editor_sync": "",
    "edit_instruction": "",
    "edit_instruction_sync": "",
    "email_instruction": "",
    "letter_instruction": "",
    "email_voice_hash": "",
    "letter_voice_hash": "",
    "profile": {},
    "history": [],
    "show_history": False,
    "show_profile": False,
    "generation_counter": 0,
}


# ============================================================
# INITIALIZE SESSION STATE
# ============================================================

for key, value in DEFAULT_STATE.items():
    if key not in st.session_state:
        if isinstance(value, list):
            st.session_state[key] = []
        elif isinstance(value, dict):
            st.session_state[key] = {}
        else:
            st.session_state[key] = value


# ============================================================
# FILE HELPERS
# ============================================================

def load_json_file(filename, default):
    try:
        if os.path.exists(filename):
            with open(filename, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass

    return default


def save_json_file(filename, data):
    try:
        with open(filename, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return True
    except Exception:
        return False


# ============================================================
# LOAD PROFILE / HISTORY
# ============================================================

if not st.session_state.profile:
    st.session_state.profile = load_json_file(
        PROFILE_FILE,
        {
            "Name": "",
            "Designation": "",
            "Contact No.": "",
            "Current Station": "",
        },
    )

if not st.session_state.history:
    st.session_state.history = load_json_file(
        HISTORY_FILE,
        [],
    )


# ============================================================
# CUSTOM CSS
# ============================================================

st.markdown(
    """
    <style>

    /* ========================================================
       GLOBAL
       ======================================================== */

    .stApp {
        background:
            linear-gradient(
                180deg,
                #f8fafc 0%,
                #ffffff 45%,
                #f8fafc 100%
            );
    }

    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 3rem;
        max-width: 1450px;
    }

    /* ========================================================
       HEADER
       ======================================================== */

    .df-header {
        background:
            linear-gradient(
                135deg,
                #0f172a,
                #1e293b 55%,
                #312e81
            );
        border-radius: 24px;
        padding: 30px 34px;
        margin-bottom: 28px;
        box-shadow:
            0 15px 40px rgba(15, 23, 42, 0.14);
    }

    .df-brand {
        color: white;
        font-size: 2rem;
        font-weight: 850;
        letter-spacing: -0.03em;
    }

    .df-tagline {
        color: #cbd5e1;
        margin-top: 6px;
        font-size: 1rem;
    }

    .df-status {
        display: inline-block;
        margin-top: 18px;
        padding: 7px 13px;
        border-radius: 999px;
        background: rgba(255,255,255,0.10);
        color: #dbeafe;
        font-size: 0.78rem;
        font-weight: 700;
    }

    /* ========================================================
       SECTION HEADINGS
       ======================================================== */

    .section-title {
        font-size: 1.45rem;
        font-weight: 850;
        color: #0f172a;
        margin-top: 24px;
        margin-bottom: 4px;
    }

    .section-subtitle {
        color: #64748b;
        font-size: 0.92rem;
        margin-bottom: 18px;
    }

    /* ========================================================
       DOCUMENT CARDS
       ======================================================== */

    .doc-card {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 18px;
        padding: 22px 20px;
        min-height: 150px;
        box-shadow:
            0 5px 18px rgba(15,23,42,0.05);
        transition: 0.2s ease;
    }

    .doc-icon {
        font-size: 2rem;
        margin-bottom: 10px;
    }

    .doc-title {
        font-size: 1.1rem;
        font-weight: 800;
        color: #0f172a;
    }

    .doc-desc {
        font-size: 0.82rem;
        color: #64748b;
        margin-top: 4px;
    }

    /* ========================================================
       INFORMATION BOXES
       ======================================================== */

    .info-box {
        background:
            linear-gradient(
                135deg,
                #eff6ff,
                #f8fafc
            );
        border: 1px solid #dbeafe;
        border-radius: 16px;
        padding: 16px 18px;
        margin: 12px 0 18px 0;
        color: #334155;
    }

    .voice-help {
        color: #64748b;
        font-size: 0.82rem;
        margin-top: 7px;
        margin-bottom: 10px;
    }

    /* ========================================================
       SELECTED INDEX CARD
       ======================================================== */

    .index-card {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 18px;
        margin: 12px 0;
        box-shadow:
            0 4px 14px rgba(15,23,42,0.04);
    }

    /* ========================================================
       FORM INPUT VISIBILITY
       ======================================================== */

    div[data-testid="stTextInput"] input,
    div[data-testid="stTextArea"] textarea,
    div[data-testid="stSelectbox"] div[data-baseweb="select"] {
        background-color: #ffffff !important;
        color: #111827 !important;
    }

    div[data-testid="stTextInput"] input,
    div[data-testid="stTextArea"] textarea {
        border: 1px solid #cbd5e1 !important;
        border-radius: 10px !important;
    }

    div[data-testid="stTextInput"] label,
    div[data-testid="stTextArea"] label {
        color: #334155 !important;
        font-weight: 650 !important;
    }

    div[data-testid="stTextInput"] input::placeholder,
    div[data-testid="stTextArea"] textarea::placeholder {
        color: #94a3b8 !important;
        opacity: 1 !important;
    }

    div[data-testid="stTextInput"] input:focus,
    div[data-testid="stTextArea"] textarea:focus {
        border-color: #6366f1 !important;
        box-shadow:
            0 0 0 2px rgba(99,102,241,0.12) !important;
    }

    /* ========================================================
       BUTTONS
       ======================================================== */

    .stButton > button {
        border-radius: 10px;
        font-weight: 700;
        min-height: 42px;
    }

    /* ========================================================
       WORKSPACE
       ======================================================== */

    .workspace-card {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 18px;
        padding: 20px;
        box-shadow:
            0 7px 22px rgba(15,23,42,0.05);
        margin-top: 20px;
    }

    /* ========================================================
       FOOTER
       ======================================================== */

    .df-footer {
        margin-top: 45px;
        padding-top: 22px;
        border-top: 1px solid #e2e8f0;
        text-align: center;
        color: #64748b;
        font-size: 0.82rem;
    }

    .df-footer strong {
        color: #334155;
    }

    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# HEADER
# ============================================================

st.markdown(
    """
    <div class="df-header">

        <div class="df-brand">
            ✦ DraftForge
        </div>

        <div class="df-tagline">
            AI Document Composer — create professional official
            documents faster
        </div>

        <div class="df-status">
            ● AI Document Workspace
        </div>

    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.markdown(
        """
        <div style="
            font-size:1.35rem;
            font-weight:850;
            color:#0f172a;
            margin-bottom:4px;
        ">
            ✦ DraftForge
        </div>

        <div style="
            color:#64748b;
            font-size:0.8rem;
            margin-bottom:20px;
        ">
            Official document workspace
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("### 💡 Tips & Templates")

    st.caption(
        "Write naturally. DraftForge will convert your "
        "instructions into professional official English."
    )

    st.caption(
        "🎤 You can speak instead of typing."
    )

    st.caption(
        "📄 Inquiry sections can be added in any order."
    )

    st.caption(
        "🔁 The same inquiry index can be added multiple times."
    )

    st.divider()

    if st.button(
        "📚 My Documents",
        use_container_width=True,
    ):
        st.session_state.show_history = not st.session_state.show_history
        st.session_state.show_profile = False

    if st.button(
        "👤 My Profile",
        use_container_width=True,
    ):
        st.session_state.show_profile = not st.session_state.show_profile
        st.session_state.show_history = False

    st.divider()

    st.markdown("### ✦ About DraftForge")

    st.caption(
        "AI-assisted drafting workspace for professional "
        "official correspondence and inquiry documentation."
    )

    st.caption(
        "Supported: Email • Letter • E&D Inquiry • FFI Inquiry"
    )


# ============================================================
# PROFILE PANEL
# ============================================================

if st.session_state.show_profile:

    st.markdown(
        """
        <div class="workspace-card">

        <div class="section-title">
            👤 My Profile
        </div>

        <div class="section-subtitle">
            Your profile information is automatically appended
            to Email and Letter drafts.
        </div>

        </div>
        """,
        unsafe_allow_html=True,
    )

    profile = st.session_state.profile

    profile_name = st.text_input(
        "Name",
        value=profile.get("Name", ""),
        key="profile_name",
    )

    profile_designation = st.text_input(
        "Designation",
        value=profile.get("Designation", ""),
        key="profile_designation",
    )

    profile_contact = st.text_input(
        "Contact No.",
        value=profile.get("Contact No.", ""),
        key="profile_contact",
    )

    profile_station = st.text_input(
        "Current Station",
        value=profile.get("Current Station", ""),
        key="profile_station",
    )

    if st.button(
        "💾 Save Profile",
        type="primary",
    ):

        st.session_state.profile = {
            "Name": profile_name,
            "Designation": profile_designation,
            "Contact No.": profile_contact,
            "Current Station": profile_station,
        }

        save_json_file(
            PROFILE_FILE,
            st.session_state.profile,
        )

        st.success("Profile saved successfully.")


# ============================================================
# HISTORY PANEL
# ============================================================

if st.session_state.show_history:

    st.markdown(
        """
        <div class="section-title">
            📚 My Documents
        </div>

        <div class="section-subtitle">
            Previously generated documents.
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not st.session_state.history:

        st.info(
            "No saved documents yet. "
            "Generated documents will appear here."
        )

    else:

        for history_index, record in enumerate(
            reversed(st.session_state.history)
        ):

            title = record.get(
                "title",
                "Untitled Document",
            )

            created = record.get(
                "created",
                "",
            )

            with st.expander(
                f"📄 {title} — {created}"
            ):

                st.text(
                    record.get(
                        "document",
                        "",
                    )
                )

                if st.button(
                    "Open Document",
                    key=f"history_open_{history_index}",
                ):

                    original = record.get(
                        "document",
                        "",
                    )

                    st.session_state.generated_draft = original
                    st.session_state.editable_draft = original
                    st.session_state.editor_sync = original

                    st.session_state.show_history = False

                    st.rerun()


# ============================================================
# DOCUMENT TYPE SELECTION
# ============================================================

st.markdown(
    """
    <div class="section-title">
        ① Choose your document
    </div>

    <div class="section-subtitle">
        Select the type of document you want DraftForge to create.
    </div>
    """,
    unsafe_allow_html=True,
)


doc_columns = st.columns(3)

document_cards = [
    (
        "📧",
        "Email",
        "Professional official email",
    ),
    (
        "📄",
        "Letter",
        "Formal official correspondence",
    ),
    (
        "🔎",
        "Inquiry",
        "E&D / FFI inquiry documents",
    ),
]


for column, card in zip(
    doc_columns,
    document_cards,
):

    icon, title, description = card

    with column:

        selected = (
            st.session_state.document_type
            == title
        )

        st.markdown(
            f"""
            <div class="doc-card"
                 style="
                    border:
                    2px solid
                    {'#6366f1' if selected else '#e2e8f0'};
                 ">

                <div class="doc-icon">
                    {icon}
                </div>

                <div class="doc-title">
                    {title}
                </div>

                <div class="doc-desc">
                    {description}
                </div>

            </div>
            """,
            unsafe_allow_html=True,
        )

        if st.button(
            (
                "✓ Selected"
                if selected
                else f"Select {title}"
            ),
            key=f"select_document_{title}",
            use_container_width=True,
        ):

            st.session_state.document_type = title

            if title != "Inquiry":
                st.session_state.index_data = []

            st.rerun()


# ============================================================
# EMAIL
# ============================================================

if st.session_state.document_type == "Email":

    st.markdown(
        """
        <div class="section-title">
            ② Provide your information
        </div>

        <div class="section-subtitle">
            Type naturally, speak naturally, or combine both.
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="info-box">
            💡 <strong>Tell DraftForge what the email should say.</strong>
            You do not need to write perfect English.
        </div>
        """,
        unsafe_allow_html=True,
    )

    # --------------------------------------------------------
    # VOICE FIRST
    # --------------------------------------------------------

    st.markdown("### 🎤 Voice input")

    st.markdown(
        """
        <div class="voice-help">
            You can speak naturally. Your transcription will be
            added to the text below.
        </div>
        """,
        unsafe_allow_html=True,
    )

    audio = st.audio_input(
        "Record voice",
        key="email_audio",
    )

    # Process audio BEFORE text widget is created.
    if audio is not None:

        try:

            audio_bytes = audio.getvalue()

            audio_hash = hashlib.sha256(
                audio_bytes
            ).hexdigest()

            if (
                audio_hash
                != st.session_state.email_voice_hash
            ):

                api_key = None

                try:
                    api_key = st.secrets.get(
                        "GROQ_API_KEY"
                    )
                except Exception:
                    pass

                api_key = (
                    api_key
                    or os.getenv("GROQ_API_KEY")
                )

                if not api_key:

                    st.warning(
                        "GROQ_API_KEY is not configured."
                    )

                elif Groq is None:

                    st.error(
                        "Groq package is not installed."
                    )

                else:

                    client = Groq(
                        api_key=api_key
                    )

                    transcription = (
                        client.audio.transcriptions.create(
                            file=(
                                "voice.wav",
                                audio_bytes,
                            ),
                            model=WHISPER_MODEL,
                        )
                    )

                    transcript = (
                        getattr(
                            transcription,
                            "text",
                            "",
                        )
                        or ""
                    ).strip()

                    if transcript:

                        existing = (
                            st.session_state.email_instruction
                            or ""
                        ).strip()

                        if existing:

                            st.session_state.email_instruction = (
                                existing
                                + "\n"
                                + transcript
                            )

                        else:

                            st.session_state.email_instruction = (
                                transcript
                            )

                    st.session_state.email_voice_hash = (
                        audio_hash
                    )

        except Exception as e:

            st.error(
                f"Voice transcription failed: {e}"
            )

    # --------------------------------------------------------
    # TEXT INPUT
    # --------------------------------------------------------

    email_instruction = st.text_area(
        "Email instructions / information",
        key="email_instruction",
        height=180,
        placeholder=(
            "Example: Inform the regional office that "
            "the backup internet connection is unavailable "
            "and request immediate restoration."
        ),
    )

    st.caption(
        "DraftForge will correct grammar and convert your "
        "natural-language instructions into professional "
        "official English."
    )

    # --------------------------------------------------------
    # GENERATE AT LAST
    # --------------------------------------------------------

    st.markdown("---")

    if st.button(
        "✦ Generate Email",
        type="primary",
        use_container_width=True,
    ):

        if not email_instruction.strip():

            st.warning(
                "Please provide some information first."
            )

        else:

            with st.spinner(
                "DraftForge is preparing your email..."
            ):

                final_document = generate_ai_document(
                    "Email",
                    email_instruction,
                    st.session_state.profile,
                )

            if final_document:

                st.session_state.generated_draft = (
                    final_document
                )

                st.session_state.editable_draft = (
                    final_document
                )

                st.session_state.editor_sync = (
                    final_document
                )

                save_history(
                    "Email",
                    final_document,
                )

                st.rerun()


# ============================================================
# LETTER
# ============================================================

elif st.session_state.document_type == "Letter":

    st.markdown(
        """
        <div class="section-title">
            ② Provide your information
        </div>

        <div class="section-subtitle">
            Type naturally, speak naturally, or combine both.
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="info-box">
            💡 <strong>Tell DraftForge what the letter should say.</strong>
            You do not need to write perfect English.
        </div>
        """,
        unsafe_allow_html=True,
    )

    # --------------------------------------------------------
    # VOICE FIRST
    # --------------------------------------------------------

    st.markdown("### 🎤 Voice input")

    st.markdown(
        """
        <div class="voice-help">
            You can speak naturally. Your transcription will be
            added to the text below.
        </div>
        """,
        unsafe_allow_html=True,
    )

    audio = st.audio_input(
        "Record voice",
        key="letter_audio",
    )

    # Process before text widget.
    if audio is not None:

        try:

            audio_bytes = audio.getvalue()

            audio_hash = hashlib.sha256(
                audio_bytes
            ).hexdigest()

            if (
                audio_hash
                != st.session_state.letter_voice_hash
            ):

                api_key = None

                try:
                    api_key = st.secrets.get(
                        "GROQ_API_KEY"
                    )
                except Exception:
                    pass

                api_key = (
                    api_key
                    or os.getenv("GROQ_API_KEY")
                )

                if not api_key:

                    st.warning(
                        "GROQ_API_KEY is not configured."
                    )

                elif Groq is None:

                    st.error(
                        "Groq package is not installed."
                    )

                else:

                    client = Groq(
                        api_key=api_key
                    )

                    transcription = (
                        client.audio.transcriptions.create(
                            file=(
                                "voice.wav",
                                audio_bytes,
                            ),
                            model=WHISPER_MODEL,
                        )
                    )

                    transcript = (
                        getattr(
                            transcription,
                            "text",
                            "",
                        )
                        or ""
                    ).strip()

                    if transcript:

                        existing = (
                            st.session_state.letter_instruction
                            or ""
                        ).strip()

                        if existing:

                            st.session_state.letter_instruction = (
                                existing
                                + "\n"
                                + transcript
                            )

                        else:

                            st.session_state.letter_instruction = (
                                transcript
                            )

                    st.session_state.letter_voice_hash = (
                        audio_hash
                    )

        except Exception as e:

            st.error(
                f"Voice transcription failed: {e}"
            )

    # --------------------------------------------------------
    # TEXT
    # --------------------------------------------------------

    letter_instruction = st.text_area(
        "Letter instructions / information",
        key="letter_instruction",
        height=180,
        placeholder=(
            "Example: Write a formal letter requesting "
            "immediate repair of the office backup connection."
        ),
    )

    st.caption(
        "DraftForge will convert your natural-language "
        "instructions into professional official English."
    )

    # --------------------------------------------------------
    # GENERATE LAST
    # --------------------------------------------------------

    st.markdown("---")

    if st.button(
        "✦ Generate Letter",
        type="primary",
        use_container_width=True,
    ):

        if not letter_instruction.strip():

            st.warning(
                "Please provide some information first."
            )

        else:

            with st.spinner(
                "DraftForge is preparing your letter..."
            ):

                final_document = generate_ai_document(
                    "Letter",
                    letter_instruction,
                    st.session_state.profile,
                )

            if final_document:

                st.session_state.generated_draft = (
                    final_document
                )

                st.session_state.editable_draft = (
                    final_document
                )

                st.session_state.editor_sync = (
                    final_document
                )

                save_history(
                    "Letter",
                    final_document,
                )

                st.rerun()


# ============================================================
# INQUIRY
# ============================================================

elif st.session_state.document_type == "Inquiry":

    st.markdown(
        """
        <div class="section-title">
            ② Build your inquiry
        </div>

        <div class="section-subtitle">
            Select the inquiry type and add only the sections
            you actually need.
        </div>
        """,
        unsafe_allow_html=True,
    )

    inquiry_type = st.selectbox(
        "Inquiry type",
        INQUIRY_TYPES,
        index=INQUIRY_TYPES.index(
            st.session_state.inquiry_type
        ),
    )

    if inquiry_type != st.session_state.inquiry_type:

        st.session_state.inquiry_type = inquiry_type

        st.session_state.index_data = []

        st.session_state.generated_draft = ""

        st.session_state.editable_draft = ""

        st.session_state.document_editor = ""

        st.rerun()

    if inquiry_type == "FFI Inquiry":

        st.info(
            "FFI Inquiry is currently under construction. "
            "The E&D inquiry workflow is available now."
        )

    # --------------------------------------------------------
    # ADD INDEX
    # --------------------------------------------------------

    st.markdown("### ➕ Add inquiry section")

    add_col1, add_col2 = st.columns(
        [4, 1]
    )

    with add_col1:

        selected_index = st.selectbox(
            "Select section to add",
            ED_INDEXES,
            key="new_inquiry_index",
        )

    with add_col2:

        st.markdown(
            "<div style='height:28px'></div>",
            unsafe_allow_html=True,
        )

        if st.button(
            "＋ Add",
            use_container_width=True,
        ):

            new_item = {
                "name": selected_index,
                "content": "",
            }

            if selected_index == "Documents Recorded":

                new_item["documents"] = []

            elif selected_index == "Inquiry Committee":

                new_item["committee"] = [
                    {
                        "role": role,
                        "erp": "",
                        "name": "",
                        "designation": "",
                    }
                    for role in COMMITTEE_ROLES
                ]

            st.session_state.index_data.append(
                new_item
            )

            st.rerun()

    # --------------------------------------------------------
    # SELECTED SECTIONS
    # --------------------------------------------------------

    if st.session_state.index_data:

        st.markdown(
            """
            <div class="section-title">
                Selected inquiry sections
            </div>

            <div class="section-subtitle">
                Complete the sections below. They will appear
                in the generated report in exactly this order.
            </div>
            """,
            unsafe_allow_html=True,
        )

        occurrence_counter = {}

        for position, item in enumerate(
            st.session_state.index_data
        ):

            base_name = item["name"]

            occurrence_counter[
                base_name
            ] = occurrence_counter.get(
                base_name,
                0,
            ) + 1

            occurrence = occurrence_counter[
                base_name
            ]

            display_name = base_name

            if base_name in [
                "Statement of the Accused",
                "Questions / Answers with the Accused",
            ] and occurrence > 1:

                display_name = (
                    f"{base_name} "
                    f"No. {occurrence}"
                )

            # ------------------------------------------------
            # DOCUMENTS RECORDED
            # ------------------------------------------------

            if base_name == "Documents Recorded":

                st.markdown(
                    f"""
                    <div class="index-card">

                        <div style="
                            font-size:1.05rem;
                            font-weight:800;
                            color:#0f172a;
                        ">
                            {display_name}
                        </div>

                        <div style="
                            font-size:0.8rem;
                            color:#64748b;
                            margin-top:4px;
                        ">
                            Select the documents that were
                            recorded or examined.
                        </div>

                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                documents = item.get(
                    "documents",
                    [],
                )

                selected_documents = st.multiselect(
                    "Documents",
                    DOCUMENTS_RECORDED,
                    default=documents,
                    key=f"documents_recorded_{position}",
                )

                item["documents"] = selected_documents

                continue

            # ------------------------------------------------
            # INQUIRY COMMITTEE
            # ------------------------------------------------

            if base_name == "Inquiry Committee":

                st.markdown(
                    """
                    <div style="
                        background:
                            linear-gradient(
                                135deg,
                                #eef2ff,
                                #f8fafc
                            );
                        border:1px solid #c7d2fe;
                        border-radius:16px;
                        padding:18px;
                        margin-bottom:15px;
                    ">

                        <div style="
                            font-size:1.05rem;
                            font-weight:800;
                            color:#1e293b;
                            margin-bottom:5px;
                        ">
                            👥 Inquiry Committee
                        </div>

                        <div style="
                            font-size:0.82rem;
                            color:#64748b;
                        ">
                            Enter the ERP#, name and designation
                            of each committee member.
                        </div>

                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                committee = item.get(
                    "committee",
                    [],
                )

                while len(committee) < len(
                    COMMITTEE_ROLES
                ):

                    role = COMMITTEE_ROLES[
                        len(committee)
                    ]

                    committee.append(
                        {
                            "role": role,
                            "erp": "",
                            "name": "",
                            "designation": "",
                        }
                    )

                for member_index, role in enumerate(
                    COMMITTEE_ROLES
                ):

                    member = committee[
                        member_index
                    ]

                    st.markdown(
                        f"""
                        <div style="
                            background:#ffffff;
                            border:1px solid #e2e8f0;
                            border-left:5px solid #6366f1;
                            border-radius:12px;
                            padding:12px 15px;
                            margin-top:14px;
                            margin-bottom:8px;
                            box-shadow:
                                0 3px 12px
                                rgba(15,23,42,0.04);
                        ">

                            <div style="
                                font-size:0.98rem;
                                font-weight:800;
                                color:#1e293b;
                            ">
                                {role}
                            </div>

                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

                    c1, c2, c3 = st.columns(3)

                    with c1:

                        member["erp"] = st.text_input(
                            "ERP#",
                            value=member.get(
                                "erp",
                                "",
                            ),
                            key=(
                                f"committee_erp_"
                                f"{position}_"
                                f"{member_index}"
                            ),
                            placeholder="Enter ERP#",
                        )

                    with c2:

                        member["name"] = st.text_input(
                            "Name",
                            value=member.get(
                                "name",
                                "",
                            ),
                            key=(
                                f"committee_name_"
                                f"{position}_"
                                f"{member_index}"
                            ),
                            placeholder="Enter name",
                        )

                    with c3:

                        member["designation"] = st.text_input(
                            "Designation",
                            value=member.get(
                                "designation",
                                "",
                            ),
                            key=(
                                f"committee_designation_"
                                f"{position}_"
                                f"{member_index}"
                            ),
                            placeholder="Enter designation",
                        )

                item["committee"] = committee

                continue

            # ------------------------------------------------
            # NORMAL INDEX
            # ------------------------------------------------

            st.markdown(
                f"""
                <div class="index-card">

                    <div style="
                        font-size:1.02rem;
                        font-weight:800;
                        color:#0f172a;
                    ">
                        {display_name}
                    </div>

                </div>
                """,
                unsafe_allow_html=True,
            )

            # ------------------------------------------------
            # Q&A
            # ------------------------------------------------

            if base_name == (
                "Questions / Answers with the Accused"
            ):

                rows = item.get(
                    "qa_rows",
                    [
                        {
                            "question": "",
                            "answer": "",
                        }
                    ],
                )

                if not rows:

                    rows = [
                        {
                            "question": "",
                            "answer": "",
                        }
                    ]

                for row_index, row in enumerate(
                    rows
                ):

                    q_col, a_col = st.columns(2)

                    with q_col:

                        row["question"] = st.text_area(
                            f"Question {row_index + 1}",
                            value=row.get(
                                "question",
                                "",
                            ),
                            key=(
                                f"qa_question_"
                                f"{position}_"
                                f"{row_index}"
                            ),
                            height=100,
                            placeholder=(
                                "Enter question..."
                            ),
                        )

                    with a_col:

                        row["answer"] = st.text_area(
                            f"Answer {row_index + 1}",
                            value=row.get(
                                "answer",
                                "",
                            ),
                            key=(
                                f"qa_answer_"
                                f"{position}_"
                                f"{row_index}"
                            ),
                            height=100,
                            placeholder=(
                                "Enter answer..."
                            ),
                        )

                item["qa_rows"] = rows

                if st.button(
                    "＋ Add Question / Answer",
                    key=f"add_qa_{position}",
                ):

                    rows.append(
                        {
                            "question": "",
                            "answer": "",
                        }
                    )

                    item["qa_rows"] = rows

                    st.rerun()

            else:

                # ------------------------------------------------
                # VOICE INPUT FOR INQUIRY INDEX
                # ------------------------------------------------

                voice_key = (
                    f"inquiry_voice_{position}"
                )

                hash_key = (
                    f"inquiry_voice_hash_{position}"
                )

                text_key = (
                    f"inquiry_text_{position}"
                )

                if text_key not in st.session_state:

                    st.session_state[
                        text_key
                    ] = item.get(
                        "content",
                        "",
                    )

                st.markdown(
                    "🎤 Voice input"
                )

                audio = st.audio_input(
                    "Record voice",
                    key=voice_key,
                )

                # Process audio before text widget.
                if audio is not None:

                    try:

                        audio_bytes = (
                            audio.getvalue()
                        )

                        audio_hash = hashlib.sha256(
                            audio_bytes
                        ).hexdigest()

                        previous_hash = (
                            st.session_state.get(
                                hash_key,
                                "",
                            )
                        )

                        if (
                            audio_hash
                            != previous_hash
                        ):

                            api_key = None

                            try:
                                api_key = (
                                    st.secrets.get(
                                        "GROQ_API_KEY"
                                    )
                                )
                            except Exception:
                                pass

                            api_key = (
                                api_key
                                or os.getenv(
                                    "GROQ_API_KEY"
                                )
                            )

                            if not api_key:

                                st.warning(
                                    "GROQ_API_KEY is not "
                                    "configured."
                                )

                            elif Groq is None:

                                st.error(
                                    "Groq package is "
                                    "not installed."
                                )

                            else:

                                client = Groq(
                                    api_key=api_key
                                )

                                transcription = (
                                    client.audio.transcriptions.create(
                                        file=(
                                            "voice.wav",
                                            audio_bytes,
                                        ),
                                        model=WHISPER_MODEL,
                                    )
                                )

                                transcript = (
                                    getattr(
                                        transcription,
                                        "text",
                                        "",
                                    )
                                    or ""
                                ).strip()

                                if transcript:

                                    existing = (
                                        st.session_state.get(
                                            text_key,
                                            "",
                                        )
                                        or ""
                                    ).strip()

                                    if existing:

                                        st.session_state[
                                            text_key
                                        ] = (
                                            existing
                                            + "\n"
                                            + transcript
                                        )

                                    else:

                                        st.session_state[
                                            text_key
                                        ] = transcript

                                st.session_state[
                                    hash_key
                                ] = audio_hash

                    except Exception as e:

                        st.error(
                            f"Voice transcription failed: {e}"
                        )

                current_content = st.text_area(
                    display_name,
                    key=text_key,
                    height=180,
                    placeholder=(
                        "Enter information for this section "
                        "in your own words..."
                    ),
                )

                item["content"] = current_content

            # ------------------------------------------------
            # REMOVE INDEX
            # ------------------------------------------------

            if st.button(
                "🗑 Remove this section",
                key=f"remove_index_{position}",
            ):

                del st.session_state.index_data[
                    position
                ]

                st.rerun()

    else:

        st.info(
            "No inquiry sections added yet. "
            "Use the selector above to add the sections "
            "you need."
        )

    # --------------------------------------------------------
    # GENERATE INQUIRY
    # --------------------------------------------------------

    st.markdown("---")

    if st.button(
        "✦ Generate Inquiry Report",
        type="primary",
        use_container_width=True,
    ):

        if not st.session_state.index_data:

            st.warning(
                "Please add at least one inquiry section."
            )

        else:

            with st.spinner(
                "DraftForge is preparing your inquiry report..."
            ):

                final_document = generate_inquiry_report(
                    st.session_state.index_data,
                    st.session_state.inquiry_type,
                )

            if final_document:

                st.session_state.generated_draft = (
                    final_document
                )

                st.session_state.editable_draft = (
                    final_document
                )

                st.session_state.editor_sync = (
                    final_document
                )

                save_history(
                    "Inquiry Report",
                    final_document,
                )

                st.rerun()


# ============================================================
# DOCUMENT EDITOR / EXPORT
# ============================================================

if st.session_state.generated_draft:

    st.markdown(
        """
        <div class="section-title">
            ③ Generate & Export
        </div>

        <div class="section-subtitle">
            Review, edit and export your generated document.
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="workspace-card">
        """,
        unsafe_allow_html=True,
    )

    # --------------------------------------------------------
    # EDITOR SYNC BEFORE WIDGET
    # --------------------------------------------------------

    if st.session_state.get(
        "editor_sync",
        "",
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
            or st.session_state.generated_draft
        )

    edited_document = st.text_area(
        "Document Editor",
        key="document_editor",
        height=550,
    )

    # --------------------------------------------------------
    # SAVE CHANGES
    # --------------------------------------------------------

    edit_col1, edit_col2 = st.columns(2)

    with edit_col1:

        if st.button(
            "💾 Save Changes",
            type="primary",
            use_container_width=True,
        ):

            st.session_state.editable_draft = (
                edited_document
            )

            st.session_state.generated_draft = (
                edited_document
            )

            st.success(
                "Changes saved."
            )

    with edit_col2:

        if st.button(
            "↩ Restore Original",
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

    # --------------------------------------------------------
    # AI EDITOR
    # --------------------------------------------------------

    st.markdown("### ✦ AI Editing Assistant")

    if st.session_state.get(
        "edit_instruction_sync",
        "",
    ):

        st.session_state.edit_instruction = (
            st.session_state.pop(
                "edit_instruction_sync"
            )
        )

    edit_instruction = st.text_area(
        "Tell DraftForge what you want to change",
        key="edit_instruction",
        height=120,
        placeholder=(
            "Example: Make this more concise and formal."
        ),
    )

    if st.button(
        "✦ Apply AI Edit",
        use_container_width=True,
    ):

        if not edit_instruction.strip():

            st.warning(
                "Please describe the change you want."
            )

        else:

            with st.spinner(
                "Applying requested changes..."
            ):

                modified = ai_edit_document(
                    edited_document,
                    edit_instruction,
                )

            if modified:

                st.session_state.editable_draft = (
                    modified
                )

                st.session_state.generated_draft = (
                    modified
                )

                st.session_state.editor_sync = (
                    modified
                )

                st.session_state.edit_instruction_sync = (
                    ""
                )

                st.rerun()

    # --------------------------------------------------------
    # EXPORT
    # --------------------------------------------------------

    st.markdown("### 📤 Export")

    export_col1, export_col2, export_col3, export_col4 = (
        st.columns(4)
    )

    current_document = (
        st.session_state.editable_draft
        or edited_document
    )

    with export_col1:

        if st.button(
            "📄 PDF",
            use_container_width=True,
        ):

            pdf_bytes = create_pdf(
                current_document
            )

            if pdf_bytes:

                st.download_button(
                    "⬇ Download PDF",
                    data=pdf_bytes,
                    file_name="DraftForge_Document.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )

    with export_col2:

        if st.button(
            "📝 DOCX",
            use_container_width=True,
        ):

            docx_bytes = create_docx(
                current_document
            )

            if docx_bytes:

                st.download_button(
                    "⬇ Download DOCX",
                    data=docx_bytes,
                    file_name="DraftForge_Document.docx",
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.wordprocessingml.document"
                    ),
                    use_container_width=True,
                )

    with export_col3:

        st.download_button(
            "⬇ TXT",
            data=current_document,
            file_name="DraftForge_Document.txt",
            mime="text/plain",
            use_container_width=True,
        )

    with export_col4:

        if Image is not None:

            if st.button(
                "🖼 PNG",
                use_container_width=True,
            ):

                png_bytes = create_png(
                    current_document
                )

                if png_bytes:

                    st.download_button(
                        "⬇ Download PNG",
                        data=png_bytes,
                        file_name="DraftForge_Document.png",
                        mime="image/png",
                        use_container_width=True,
                    )

    st.markdown(
        """
        </div>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# FOOTER
# ============================================================

st.markdown(
    """
    <div class="df-footer">

        <div style="
            font-size:1rem;
            font-weight:800;
            color:#334155;
            margin-bottom:8px;
        ">
            ✦ DraftForge — AI Document Composer
        </div>

        <div>
            Developed by:
            <strong>Raees Khan — Assistant Director, NADRA</strong>
        </div>

    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# AI HELPERS
# ============================================================

def get_groq_client():

    api_key = None

    try:
        api_key = st.secrets.get(
            "GROQ_API_KEY"
        )
    except Exception:
        pass

    api_key = (
        api_key
        or os.getenv("GROQ_API_KEY")
    )

    if not api_key:
        st.error(
            "GROQ_API_KEY is not configured. "
            "Please add it to Streamlit Secrets."
        )
        return None

    if Groq is None:
        st.error(
            "The Groq Python package is not installed."
        )
        return None

    return Groq(
        api_key=api_key
    )


def clean_ai_text(text):

    if not text:
        return ""

    text = text.strip()

    # Remove accidental markdown code fences.
    text = re.sub(
        r"^```(?:text|markdown)?\s*",
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


def profile_signature(profile):

    name = profile.get(
        "Name",
        "",
    ).strip()

    designation = profile.get(
        "Designation",
        "",
    ).strip()

    contact = profile.get(
        "Contact No.",
        "",
    ).strip()

    station = profile.get(
        "Current Station",
        "",
    ).strip()

    lines = []

    if name:
        lines.append(name)

    if designation:
        lines.append(designation)

    if contact:
        lines.append(
            f"Contact No.: {contact}"
        )

    if station:
        lines.append(
            f"Current Station: {station}"
        )

    if not lines:
        return ""

    return "\n\n" + "\n".join(lines)


def generate_ai_document(
    document_type,
    user_instruction,
    profile,
):

    client = get_groq_client()

    if client is None:
        return ""

    system_prompt = f"""
You are DraftForge, an AI assistant for preparing
professional official documents.

DOCUMENT TYPE:
{document_type}

TASK:
Convert the user's natural-language instructions into
professional official English.

STRICT RULES:

1. Correct spelling, grammar, punctuation and obvious
   voice-transcription errors.

2. Preserve the user's intended meaning.

3. Do not invent names, dates, facts, allegations,
   evidence, events, reference numbers or other information.

4. Do not add unsupported facts.

5. Do not add a sender signature.

6. Do not explain what you changed.

7. Return ONLY the finished document.

8. Use professional official English.

9. Keep the document concise where appropriate.

10. The user's instructions may be informal or incomplete.
    Improve the language without inventing missing facts.
"""

    user_prompt = f"""
Prepare the {document_type} based only on the following
information:

{user_instruction}
"""

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

        result = clean_ai_text(
            response.choices[0].message.content
        )

        if document_type in [
            "Email",
            "Letter",
        ]:

            signature = profile_signature(
                profile
            )

            if signature:
                result += signature

        return result

    except Exception as e:

        st.error(
            f"AI generation failed: {e}"
        )

        return ""


# ============================================================
# INQUIRY GENERATION
# ============================================================

def build_documents_recorded_text(documents):

    if not documents:
        return "No information was provided for this index."

    lines = []

    for index, document in enumerate(
        documents
    ):

        annex_letter = chr(
            ord("A") + index
        )

        lines.append(
            f"Annex-{annex_letter}: {document}"
        )

    return "\n".join(lines)


def build_committee_text(committee):

    if not committee:
        return "No information was provided for this index."

    lines = []

    for member in committee:

        role = member.get(
            "role",
            "",
        ).strip()

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

        if not (
            erp
            or name
            or designation
        ):
            continue

        line = role

        details = []

        if erp:
            details.append(
                f"ERP#: {erp}"
            )

        if name:
            details.append(
                f"Name: {name}"
            )

        if designation:
            details.append(
                f"Designation: {designation}"
            )

        if details:

            line += " — " + ", ".join(
                details
            )

        lines.append(line)

    if not lines:
        return "No information was provided for this index."

    return "\n".join(lines)


def format_qa_rows(rows):

    if not rows:
        return (
            "| Questions | Answers |\n"
            "|---|---|\n"
            "| No information was provided for this index. | |"
        )

    output = [
        "| Questions | Answers |",
        "|---|---|",
    ]

    question_number = 1

    for row in rows:

        question = (
            row.get(
                "question",
                "",
            )
            or ""
        ).strip()

        answer = (
            row.get(
                "answer",
                "",
            )
            or ""
        ).strip()

        if not question and not answer:
            continue

        question = question.replace(
            "|",
            "\\|",
        )

        answer = answer.replace(
            "|",
            "\\|",
        )

        output.append(
            f"| {question_number}. "
            f"{question} | {answer} |"
        )

        question_number += 1

    if len(output) == 2:

        output.append(
            "| No information was provided for this index. | |"
        )

    return "\n".join(output)


def generate_inquiry_report(
    index_data,
    inquiry_type,
):

    client = get_groq_client()

    if client is None:
        return ""

    # --------------------------------------------------------
    # IMPORTANT:
    # Build ONLY the indexes selected by the user.
    # --------------------------------------------------------

    selected_sections = []

    for item in index_data:

        name = item.get(
            "name",
            "",
        )

        section = {
            "name": name,
        }

        if name == "Documents Recorded":

            section["content"] = (
                build_documents_recorded_text(
                    item.get(
                        "documents",
                        [],
                    )
                )
            )

        elif name == "Inquiry Committee":

            section["content"] = (
                build_committee_text(
                    item.get(
                        "committee",
                        [],
                    )
                )
            )

        elif name == (
            "Questions / Answers with the Accused"
        ):

            section["content"] = (
                format_qa_rows(
                    item.get(
                        "qa_rows",
                        [],
                    )
                )
            )

        else:

            content = (
                item.get(
                    "content",
                    "",
                )
                or ""
            ).strip()

            if not content:

                content = (
                    "No information was provided "
                    "for this index."
                )

            section["content"] = content

        selected_sections.append(
            section
        )

    if not selected_sections:

        st.warning(
            "No inquiry sections were selected."
        )

        return ""

    # --------------------------------------------------------
    # Give the AI ONLY the selected sections.
    # --------------------------------------------------------

    raw_sections = []

    for section in selected_sections:

        raw_sections.append(
            f"""
INDEX:
{section["name"]}

USER INFORMATION:
{section["content"]}
"""
        )

    source_text = "\n".join(
        raw_sections
    )

    system_prompt = f"""
You are DraftForge, an AI assistant preparing an
official {inquiry_type} inquiry document.

CRITICAL REQUIREMENT:

The user selected specific inquiry indexes.

You MUST produce ONLY those selected indexes.

Do NOT create any index that the user did not select.

Do NOT add:
- Introduction
- Background
- Summary of Evidence
- Findings
- Findings on Each Charge
- Discussion
- Conclusion
- Recommendations
- Inquiry Committee
- Documents Recorded
- any other section

unless that exact index was supplied by the user.

Do not invent facts.

Do not invent names.

Do not invent dates.

Do not invent allegations.

Do not invent evidence.

Do not invent findings.

Do not invent recommendations.

Correct spelling, grammar, punctuation and obvious
voice transcription errors while preserving meaning.

Keep the selected indexes in exactly the same order
provided by the user.

For a selected index with no information, write:

No information was provided for this index.

The final output must contain ONLY the selected indexes.

Use professional official English.

For Questions / Answers with the Accused, preserve the
two-column Markdown table format:

| Questions | Answers |
|---|---|
| 1. Question | Answer |

For Documents Recorded, preserve contiguous annexure
labels such as Annex-A, Annex-B, Annex-C.

For Inquiry Committee, preserve the supplied committee
roles and details.

Do not add a signature.
"""

    user_prompt = f"""
The following are the ONLY inquiry indexes selected by
the user:

{source_text}

Prepare the final inquiry document now.

Remember:
- Include ONLY these indexes.
- Preserve their order.
- Do not invent absent sections.
"""

    try:

        response = client.chat.completions.create(
            model=GROQ_MODEL,
            temperature=0.10,
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

        result = clean_ai_text(
            response.choices[0].message.content
        )

        return result

    except Exception as e:

        st.error(
            f"Inquiry generation failed: {e}"
        )

        return ""


# ============================================================
# AI DOCUMENT EDIT
# ============================================================

def ai_edit_document(
    document,
    instruction,
):

    client = get_groq_client()

    if client is None:
        return ""

    system_prompt = """
You are an official-document editing assistant.

Edit the supplied document according to the user's
instruction.

Rules:

1. Preserve the original meaning unless the user explicitly
   requests a change in meaning.

2. Do not invent facts.

3. Do not invent names, dates, evidence or events.

4. Correct grammar, spelling and punctuation when appropriate.

5. Preserve existing section structure unless the user asks
   for structural changes.

6. Return ONLY the revised document.

7. Do not explain your changes.
"""

    user_prompt = f"""
CURRENT DOCUMENT:

{document}

USER'S EDITING INSTRUCTION:

{instruction}

Return the complete revised document.
"""

    try:

        response = client.chat.completions.create(
            model=GROQ_MODEL,
            temperature=0.10,
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

        return clean_ai_text(
            response.choices[0].message.content
        )

    except Exception as e:

        st.error(
            f"AI editing failed: {e}"
        )

        return ""


# ============================================================
# HISTORY
# ============================================================

def save_history(
    title,
    document,
):

    if not document:
        return

    record = {
        "title": title,
        "document": document,
        "created": datetime.now().strftime(
            "%Y-%m-%d %H:%M"
        ),
    }

    history = st.session_state.history

    history.append(
        record
    )

    # Keep recent documents manageable.
    st.session_state.history = history[-50:]

    save_json_file(
        HISTORY_FILE,
        st.session_state.history,
    )


# ============================================================
# PDF EXPORT
# ============================================================

def create_pdf(
    document_text,
):

    if FPDF is None:

        st.error(
            "FPDF is not installed."
        )

        return None

    try:

        pdf = FPDF()
        pdf.set_auto_page_break(
            auto=True,
            margin=18,
        )

        pdf.add_page()

        pdf.set_font(
            "Arial",
            "B",
            15,
        )

        pdf.cell(
            0,
            10,
            "DraftForge Document",
            ln=True,
            align="C",
        )

        pdf.ln(5)

        pdf.set_font(
            "Arial",
            "",
            10,
        )

        left_margin = 15
        right_margin = 15

        pdf.set_left_margin(
            left_margin
        )

        pdf.set_right_margin(
            right_margin
        )

        usable_width = (
            pdf.w
            - left_margin
            - right_margin
        )

        lines = document_text.splitlines()

        for line in lines:

            # ------------------------------------------------
            # Markdown table
            # ------------------------------------------------

            if (
                "|" in line
                and line.strip()
            ):

                parts = [
                    p.strip()
                    for p in line.strip().strip("|").split("|")
                ]

                if (
                    len(parts) == 2
                    and not all(
                        set(p) <= {"-", ":"}
                        for p in parts
                    )
                ):

                    col1 = usable_width * 0.42
                    col2 = usable_width * 0.58

                    x_before = pdf.get_x()

                    pdf.set_x(
                        left_margin
                    )

                    pdf.set_font(
                        "Arial",
                        "",
                        9,
                    )

                    pdf.multi_cell(
                        col1,
                        7,
                        safe_pdf_text(
                            parts[0]
                        ),
                        border=1,
                    )

                    y_after = pdf.get_y()

                    pdf.set_xy(
                        left_margin + col1,
                        y_after - 7,
                    )

                    pdf.multi_cell(
                        col2,
                        7,
                        safe_pdf_text(
                            parts[1]
                        ),
                        border=1,
                    )

                    continue

            # ------------------------------------------------
            # Normal text
            # ------------------------------------------------

            if not line.strip():

                pdf.ln(4)
                continue

            wrapped_lines = textwrap.wrap(
                line,
                width=105,
                break_long_words=True,
                break_on_hyphens=True,
            )

            if not wrapped_lines:
                wrapped_lines = [""]

            for wrapped in wrapped_lines:

                pdf.set_x(
                    left_margin
                )

                pdf.multi_cell(
                    usable_width,
                    6,
                    safe_pdf_text(
                        wrapped
                    ),
                )

        return bytes(
            pdf.output(
                dest="S"
            )
        )

    except Exception as e:

        st.error(
            f"PDF export failed: {e}"
        )

        return None


def safe_pdf_text(text):

    if text is None:
        return ""

    # FPDF core fonts are not Unicode.
    return (
        str(text)
        .replace("–", "-")
        .replace("—", "-")
        .replace("’", "'")
        .replace("“", '"')
        .replace("”", '"')
        .replace("•", "-")
        .encode(
            "latin-1",
            "replace",
        )
        .decode(
            "latin-1"
        )
    )


# ============================================================
# DOCX EXPORT
# ============================================================

def create_docx(
    document_text,
):

    if Document is None:

        st.error(
            "python-docx is not installed."
        )

        return None

    try:

        doc = Document()

        section = doc.sections[0]

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

        lines = document_text.splitlines()

        index = 0

        while index < len(lines):

            line = lines[index]

            # ------------------------------------------------
            # Detect Markdown table
            # ------------------------------------------------

            if (
                "|" in line
                and index + 1 < len(lines)
                and "|" in lines[index + 1]
                and re.match(
                    r"^\s*\|?\s*:?-+:?\s*\|",
                    lines[index + 1],
                )
            ):

                table_rows = []

                while (
                    index < len(lines)
                    and "|" in lines[index]
                ):

                    current = lines[index].strip()

                    if re.match(
                        r"^\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)+\|?$",
                        current,
                    ):

                        index += 1
                        continue

                    parts = [
                        p.strip()
                        for p in current.strip(
                            "|"
                        ).split("|")
                    ]

                    if len(parts) >= 2:

                        table_rows.append(
                            parts[:2]
                        )

                    index += 1

                if table_rows:

                    table = doc.add_table(
                        rows=len(table_rows),
                        cols=2,
                    )

                    table.style = (
                        "Table Grid"
                    )

                    table.alignment = (
                        WD_TABLE_ALIGNMENT.CENTER
                    )

                    for r, row in enumerate(
                        table_rows
                    ):

                        for c in range(2):

                            cell = table.cell(
                                r,
                                c,
                            )

                            cell.vertical_alignment = (
                                WD_CELL_VERTICAL_ALIGNMENT.CENTER
                            )

                            cell.text = (
                                row[c]
                                if c < len(row)
                                else ""
                            )

                            for paragraph in (
                                cell.paragraphs
                            ):

                                for run in (
                                    paragraph.runs
                                ):

                                    run.font.size = (
                                        Pt(10)
                                    )

                    doc.add_paragraph()

                    continue

            # ------------------------------------------------
            # Normal paragraph
            # ------------------------------------------------

            if line.strip():

                paragraph = doc.add_paragraph()

                run = paragraph.add_run(
                    line
                )

                run.font.name = (
                    "Arial"
                )

                run.font.size = (
                    Pt(11)
                )

            else:

                doc.add_paragraph()

            index += 1

        output = io.BytesIO()

        doc.save(output)

        output.seek(0)

        return output.getvalue()

    except Exception as e:

        st.error(
            f"DOCX export failed: {e}"
        )

        return None


# ============================================================
# PNG EXPORT
# ============================================================

def create_png(
    document_text,
):

    if Image is None:

        st.error(
            "Pillow is not installed."
        )

        return None

    try:

        font = None

        font_candidates = [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
        ]

        bold_candidates = [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
        ]

        for path in font_candidates:

            if os.path.exists(path):

                font = ImageFont.truetype(
                    path,
                    28,
                )

                break

        bold_font = font

        for path in bold_candidates:

            if os.path.exists(path):

                bold_font = ImageFont.truetype(
                    path,
                    34,
                )

                break

        if font is None:

            font = ImageFont.load_default()
            bold_font = font

        lines = []

        for raw_line in document_text.splitlines():

            wrapped = textwrap.wrap(
                raw_line,
                width=75,
                break_long_words=True,
            )

            if wrapped:
                lines.extend(wrapped)
            else:
                lines.append("")

        line_height = 42

        width = 1400

        height = (
            max(
                len(lines),
                1,
            )
            * line_height
            + 180
        )

        image = Image.new(
            "RGB",
            (width, height),
            "white",
        )

        draw = ImageDraw.Draw(
            image
        )

        draw.text(
            (60, 40),
            "DraftForge",
            font=bold_font,
            fill="black",
        )

        y = 110

        for line in lines:

            draw.text(
                (60, y),
                line,
                font=font,
                fill="black",
            )

            y += line_height

        output = io.BytesIO()

        image.save(
            output,
            format="PNG",
        )

        output.seek(0)

        return output.getvalue()

    except Exception as e:

        st.error(
            f"PNG export failed: {e}"
        )

        return None
