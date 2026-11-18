"""Kidus BID bot — Exact Claude.ai dark-mode conversational proposal interface."""
from datetime import datetime
import copy
import io
import json
import os
from pathlib import Path
import re
import shutil
from tempfile import TemporaryDirectory
import uuid

import openpyxl
from docx import Document
import streamlit as st

from bidagent.config import load_config
from bidagent.ingest import ingest_docx
from bidagent.library import Library
from bidagent.llm import LLM, load_env
from bidagent.pipeline import Pipeline
from bidagent.util import slug


PROJECT_ROOT = Path(__file__).resolve().parent
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
TENDER_TYPES = ["pdf", "docx", "txt", "md", "csv", "xlsx", "xlsm"]
CONFIG = load_config(PROJECT_ROOT / "config.yaml")
TEMPLATE_PATH = CONFIG["_paths"]["library"] / "templates" / CONFIG["template"]["name"]
LIBRARY_PATH = CONFIG["_paths"]["library"]
load_env(PROJECT_ROOT / ".env")
load_env(PROJECT_ROOT.parent / ".env")

MODEL_AUTO = "Auto (recommended)"
MODEL_OPTIONS = (
    MODEL_AUTO,
    "Groq · GPT-OSS 20B (fast)",
    "Gemini · 3.8 Flash",
    "OpenRouter · Free",
)
MODEL_CHAINS = {
    "Groq · GPT-OSS 20B (fast)": [
        {"provider": "groq", "model": "openai/gpt-oss-20b"},
        {"provider": "gemini", "model": "gemini-3.8-flash"},
        {"provider": "gemini", "model": "gemini-3.5-flash"},
        {"provider": "openrouter", "model": "openrouter/free"},
    ],
    "Gemini · 3.8 Flash": [
        {"provider": "gemini", "model": "gemini-3.8-flash"},
        {"provider": "gemini", "model": "gemini-3.5-flash"},
        {"provider": "groq", "model": "openai/gpt-oss-20b"},
        {"provider": "openrouter", "model": "openrouter/free"},
    ],
    "OpenRouter · Free": [
        {"provider": "openrouter", "model": "openrouter/free"},
        {"provider": "gemini", "model": "gemini-3.8-flash"},
        {"provider": "groq", "model": "openai/gpt-oss-20b"},
    ],
}


def config_for_model_choice(choice: str = MODEL_AUTO) -> dict:
    """Return the normal task-aware config or apply a user-selected provider first."""
    cfg = copy.deepcopy(CONFIG)
    selected_chain = MODEL_CHAINS.get(choice)
    if selected_chain:
        for agent in ("chat", "default", "tender", "datasheet", "writer", "compliance", "reviewer"):
            cfg["agents"][agent] = copy.deepcopy(selected_chain)
    return cfg


# ==============================================================================
# Helper Utilities
# ==============================================================================

def format_bytes(size: int) -> str:
    """Format byte size into human readable string."""
    if not size or size < 0:
        return "0 B"
    for unit in ["B", "KB", "MB", "GB"]:
        if size < 1024.0:
            return f"{size:.1f} {unit}" if unit != "B" else f"{size} B"
        size /= 1024.0
    return f"{size:.1f} TB"


def ensure_starter_template(template_path: Path, lib: Library) -> Path:
    """Create a standard starter Word template and ingest it into the library if not present."""
    if template_path.exists():
        return template_path

    template_path.parent.mkdir(parents=True, exist_ok=True)
    doc = Document()
    doc.add_paragraph("{{PROJECT_NAME}}", style="Title")
    doc.add_paragraph("Prepared for {{CLIENT}}  |  Tender No: {{TENDER_NO}}  |  Date: {{DATE}}")

    doc.add_heading("Company Background", 1)
    doc.add_paragraph(
        "We are an established engineering, technology, and systems integration organization "
        "delivering high-reliability turnkey infrastructure and lifecycle support."
    )

    doc.add_heading("Project Introduction", 1)
    doc.add_heading("Project Overview", 2)
    doc.add_paragraph(
        "This technical proposal is submitted in full response to the client requirements, "
        "providing a fully integrated, scalable, and standards-compliant technical solution."
    )

    doc.add_heading("Purpose of the Project", 2)
    doc.add_paragraph(
        "The primary purpose is to modernize operational capability, enhance security, and deliver "
        "uncompromising system availability and operational efficiency."
    )

    doc.add_heading("Technical Proposal", 1)
    doc.add_heading("Bill of Materials", 1)

    table = doc.add_table(rows=1, cols=5)
    table.style = "Table Grid"
    headers = ["No", "Item", "Model", "Unit", "Quantity"]
    for idx, text in enumerate(headers):
        table.rows[0].cells[idx].text = text

    doc.save(str(template_path))
    ingest_docx(template_path, lib, None, template_path.name)
    return template_path


def create_demo_bom_bytes() -> bytes:
    """Generate a sample Bill of Materials workbook for demonstration."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BOM"
    ws.append(["Campus Security & Surveillance Upgrade - Bill of Materials"])
    ws.append([])
    ws.append(["No", "Item", "Model", "Unit", "Quantity"])
    sample_rows = [
        ("1", "Surveillance System", "", "", ""),
        ("1.1", "4K Network Camera IP66 Outdoor", "Acme NC4K200", "Nos", 24),
        ("1.2", "Thermal Detection Camera", "Zeta TH900X", "Nos", 2),
        ("2", "Network Infrastructure", "", "", ""),
        ("2.1", "24-Port Gigabit PoE+ Switch", "Acme SW24P100", "Nos", 3),
        ("3", "Server & Enclosure", "", "", ""),
        ("3.1", "42U Server Rack Enclosure", "Generic 42U Rack RK42U", "Nos", 1),
    ]
    for row in sample_rows:
        ws.append(list(row))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


DEMO_TENDER_TEXT = """TENDER REQUIREMENTS SPECIFICATION
Project: Campus Security & Surveillance Upgrade
Client: Sample Organization
Tender No: DEMO-SEC-001
Location: Sample Operations Campus

1. Scope of Work
Supply, configure, deploy, and commission an IP-based video surveillance and networking system.

2. Technical Specifications
- Camera Sensor: Minimum 8 MP CMOS / 4K optical resolution.
- Camera Weatherproofing: IP66 or IP67 ingress protection for outdoor units.
- Camera Power: PoE IEEE 802.3af/at compliant.
- Network Switching: At least 24 Gigabit Ethernet PoE+ ports with 370W minimum total PoE budget.
- Uplinks: Minimum 2 SFP/SFP+ optical uplink ports.
- Warranty: All equipment shall be covered by 3-year warranty and standard vendor support.
"""


def parse_chat_val(val):
    """Safely unpack text and uploaded files from Streamlit's chat_input."""
    if val is None:
        return "", []
    if isinstance(val, str):
        return val, []
    text = getattr(val, "text", "") or ""
    files = []
    if hasattr(val, "files"):
        try:
            files = val.files or []
        except AttributeError:
            files = []
    elif hasattr(val, "get"):
        files = val.get("files", []) or []
    return text, files


def run_pipeline(
    bom_name: str,
    bom_bytes: bytes,
    tender_files_data: list,
    template_bytes: bytes | None,
    meta: dict,
    instructions: str,
    status_container,
    model_choice: str = MODEL_AUTO,
) -> dict:
    """Execute the end-to-end BidAgent pipeline and return generated artifacts."""
    upload_root = PROJECT_ROOT / "work" / "uploads"
    upload_root.mkdir(parents=True, exist_ok=True)
    logs = []

    with TemporaryDirectory(prefix="bid-", dir=upload_root) as temp_dir:
        temp_root = Path(temp_dir)
        bom_ext = Path(bom_name).suffix.lower() or ".xlsx"
        bom_path = temp_root / f"bom{bom_ext}"
        bom_path.write_bytes(bom_bytes)

        tender_paths = []
        for index, item in enumerate(tender_files_data, start=1):
            t_name = item["name"]
            t_ext = Path(t_name).suffix.lower()
            if t_ext not in {"." + ext for ext in TENDER_TYPES}:
                continue
            path = temp_root / f"tender_{index}{t_ext}"
            path.write_bytes(item["bytes"])
            tender_paths.append(path)

        cfg = config_for_model_choice(model_choice)
        lib = Library(cfg["_paths"]["library"])

        if template_bytes is not None:
            status_container.write("Ingesting Word template & styles...")
            template_upload = temp_root / "uploaded_template.docx"
            template_upload.write_bytes(template_bytes)
            ingest_docx(template_upload, lib, None, cfg["template"]["name"])
        elif not TEMPLATE_PATH.exists():
            status_container.write("Initializing starter Word template...")
            ensure_starter_template(TEMPLATE_PATH, lib)

        status_container.write("Analyzing Bill of Materials...")
        status_container.write("Extracting tender clauses & specifications...")
        status_container.write("Researching product specs & library matching...")
        status_container.write("Drafting technical proposal sections...")
        status_container.write("Auditing compliance matrix against requirements...")
        status_container.write("Assembling Word document (.docx) & audit report (.md)...")

        pipeline = Pipeline(cfg, force=True, log=logs.append)
        built = pipeline.run(
            bom_path,
            tender_paths=tender_paths,
            name=meta.get("name", "").strip() or Path(bom_name).stem,
            client=meta.get("client", "").strip(),
            tender_no=meta.get("tender_no", "").strip(),
            location=meta.get("location", "").strip(),
            instructions=instructions.strip(),
        )

    docx_path = Path(built["docx"])
    report_path = Path(built["report"])

    return {
        "docx_name": docx_path.name,
        "docx_bytes": docx_path.read_bytes(),
        "docx_size": format_bytes(docx_path.stat().st_size),
        "report_name": report_path.name,
        "report_text": report_path.read_text(encoding="utf-8"),
        "report_bytes": report_path.read_bytes(),
        "report_size": format_bytes(report_path.stat().st_size),
        "stats": built["stats"],
        "findings": built.get("findings", []),
        "logs": logs,
    }


def generate_conversational_reply(
    query: str,
    history: list,
    bid_meta: dict,
    active_bom: dict = None,
    active_tenders: list = None,
    last_result: dict = None,
    model_choice: str = MODEL_AUTO,
) -> str:
    """Generate an adaptive, context-aware LLM response that learns from conversation history."""
    # Keep routine social messages instant instead of waiting on a model round-trip.
    quick_message = re.sub(r"[^a-z0-9']+", " ", query.lower()).strip()
    quick_replies = {
        "hi": "Hey! What are we working on today?",
        "hy": "Hey! What are we working on today?",
        "hi there": "Hey! What are we working on today?",
        "hello": "Hey! What are we working on today?",
        "hello there": "Hey! What are we working on today?",
        "hey": "Hey! What are we working on today?",
        "hey there": "Hey! What are we working on today?",
        "yo": "Yo! What are we working on today?",
        "sup": "Yo! What are we working on today?",
        "thanks": "Anytime!",
        "thank you": "Anytime!",
        "ok": "Got you.",
        "okay": "Got you.",
        "you good": "All good! What are we tackling?",
        "how's it going": "All good! What are we tackling?",
        "good morning": "Morning! What are we working on?",
        "good afternoon": "Afternoon! What are we working on?",
        "good evening": "Evening! What are we working on?",
    }
    if quick_message in quick_replies:
        return quick_replies[quick_message]

    active_bom_name = active_bom.get("name") if active_bom else "None attached yet"
    tenders_summary = ", ".join(t.get("name", "Document") for t in (active_tenders or [])) if active_tenders else "None attached yet"

    system_prompt = f"""You are Kidus BID bot, an advanced AI technical proposal and compliance engineer.
You speak with Claude's distinct conversational voice: thoughtful, direct, articulate, concise, and helpful.

Key Persona & Operating Rules:
1. Memory & Adaptability: Pay careful attention to whatever the user tells you (client specifics, company name, equipment preferences, warranties, tone, project criteria). Acknowledge and remember them, and incorporate them in this and subsequent proposal drafting.
2. Natural Conversation: For greetings ("hi", "you good", "how's it going"), reply warmly, naturally, and conversationally without repeating boilerplate lists. Ask what bid or project they're tackling.
3. Bid & Engineering Expertise: Answer technical questions with solid engineering domain knowledge across IT, networking, CCTV, telecom, electrical, and commercial bidding.
4. Active Project Awareness: Use the project details below to answer questions about the active tender, BOM, or compliance checks accurately.
5. Proposal Generation: When the user wants the formal Word proposal (.docx), let them know they can click 'Load Sample Bid' in the sidebar or upload their BOM (.xlsx) and say "Generate proposal".

Active Project Details:
- Project Name: {bid_meta.get('name') or 'No project selected'}
- Client: {bid_meta.get('client') or 'Not provided'}
- Tender Number: {bid_meta.get('tender_no') or 'Not provided'}
- Location: {bid_meta.get('location') or 'Not provided'}
- Staged BOM: {active_bom_name}
- Staged Tender Documents: {tenders_summary}

Do not assume a sample or previous project is active. If no project details or bid files are provided, ask what project the user is working on without naming any specific tender.
"""

    if last_result:
        stats = last_result.get("stats", {})
        findings = last_result.get("findings", [])
        high_findings = [f for f in findings if f.get("sev") == "high"]
        system_prompt += f"""
Generated Proposal Details:
- Proposal File: {last_result.get('docx_name', 'None')}
- BOM Items Processed: {stats.get('BOM items', 0)}
- Product Sections Written: {stats.get('Products written', 0)}
- Compliance Rows Audited: {stats.get('Compliance rows', 0)}
- Total Discrepancies: {len(findings)} ({len(high_findings)} High Priority)
"""
        if high_findings:
            issues_sample = "; ".join(f"{f.get('where', 'Item')}: {f.get('issue', '')}" for f in high_findings[:5])
            system_prompt += f"- Key Discrepancies: {issues_sample}\n"

    # Build conversation messages for multi-turn context
    conversation_messages = [{"role": "system", "content": system_prompt}]
    for m in (history or [])[-14:]:
        role = "assistant" if m.get("role") == "assistant" else "user"
        content = (m.get("content") or "").strip()
        if content:
            conversation_messages.append({"role": role, "content": content})

    if not conversation_messages or conversation_messages[-1].get("content") != query.strip():
        conversation_messages.append({"role": "user", "content": query.strip()})

    # Fast chat route: Groq first, then Gemini and an OpenRouter free model.
    # Each provider receives the same compact conversation transcript and uses the
    # shared retry/cache/logging behavior from the proposal pipeline.
    conversation_text = "\n\n".join(
        f"{message['role'].title()}: {message['content']}"
        for message in conversation_messages[1:]
    )
    try:
        return LLM(config_for_model_choice(model_choice), PROJECT_ROOT / "work").complete(
            agent="chat",
            system=system_prompt,
            user=conversation_text,
            temperature=0.7,
            max_tokens=1500,
        )
    except Exception:
        pass

    # Fallback contextual replies if API call fails
    q_lower = query.lower()
    if any(k in q_lower for k in ["compliance", "comply", "deviation", "fail", "gap", "issue", "finding"]):
        findings = last_result.get("findings", []) if last_result else []
        if not findings:
            return "No compliance deviations were flagged in the last proposal build. All items either comply or have passed initial static checks."
        high = [f for f in findings if f.get("sev") == "high"]
        lines = [f"### Compliance Findings ({len(findings)} total)\n"]
        if high:
            lines.append(f"**High Priority Discrepancies ({len(high)}):**")
            for item in high:
                lines.append(f"- **{item.get('where', 'General')}**: {item.get('issue', '')}\n  *Action:* {item.get('fix', 'Review specification')}")
        return "\n".join(lines)

    return (
        f"I hear you! I am ready to work on the **{bid_meta.get('name', 'project')}**. "
        "Attach your BOM or tender documents whenever you're ready, or tell me about your bidding specifications!"
    )


# ==============================================================================
# Persistent Session & Memory Management
# ==============================================================================

SESSIONS_DIR = PROJECT_ROOT / "work" / "sessions"
SESSIONS_DIR.mkdir(parents=True, exist_ok=True)


def list_sessions() -> list:
    """List all saved chat sessions sorted by updated_at descending."""
    sessions = []
    for meta_file in SESSIONS_DIR.glob("*/session.json"):
        try:
            data = json.loads(meta_file.read_text(encoding="utf-8"))
            if data.get("id"):
                sessions.append(data)
        except Exception:
            continue
    sessions.sort(key=lambda s: s.get("updated_at", ""), reverse=True)
    return sessions


def load_session(session_id: str) -> bool:
    """Load session messages, metadata, and files into st.session_state."""
    session_dir = SESSIONS_DIR / session_id
    meta_file = session_dir / "session.json"
    if not meta_file.exists():
        return False
    try:
        data = json.loads(meta_file.read_text(encoding="utf-8"))
        st.session_state["current_session_id"] = session_id
        st.session_state["bid_meta"] = data.get("bid_meta", {
            "name": "",
            "client": "",
            "tender_no": "",
            "location": "",
        })

        # Reload BOM
        bom_info = data.get("active_bom")
        if bom_info and (session_dir / bom_info["filename"]).exists():
            st.session_state["active_bom"] = {
                "name": bom_info["name"],
                "bytes": (session_dir / bom_info["filename"]).read_bytes(),
            }
        else:
            st.session_state["active_bom"] = None

        # Reload Tenders
        st.session_state["active_tenders"] = []
        for t_info in data.get("active_tenders", []):
            t_path = session_dir / t_info["filename"]
            if t_path.exists():
                st.session_state["active_tenders"].append({
                    "name": t_info["name"],
                    "bytes": t_path.read_bytes(),
                })

        # Reload messages & restore artifact bytes
        messages = []
        for msg in data.get("messages", []):
            msg_copy = dict(msg)
            artifacts = []
            for art in msg.get("artifacts", []):
                art_copy = dict(art)
                art_filename = art.get("filename", "")
                if art_filename and (session_dir / art_filename).exists():
                    art_copy["bytes"] = (session_dir / art_filename).read_bytes()
                artifacts.append(art_copy)
            msg_copy["artifacts"] = artifacts
            messages.append(msg_copy)

        st.session_state["messages"] = messages
        st.session_state["last_result"] = data.get("last_result")
        return True
    except Exception as exc:
        st.error(f"Error loading session: {exc}")
        return False


def save_current_session(title: str = None):
    """Save current conversation, memories, and documents to disk."""
    if not st.session_state.get("messages"):
        return

    # Don't clutter Recents with chats that contain only greetings or acknowledgements.
    user_prompts = [
        (m.get("content") or "").strip()
        for m in st.session_state.get("messages", [])
        if m.get("role") == "user" and (m.get("content") or "").strip()
    ]
    smalltalk = {"hi", "hy", "hi there", "hello", "hello there", "hey", "hey there", "yo", "sup", "thanks", "thank you", "ok", "okay", "you good", "how's it going", "good morning", "good afternoon", "good evening"}
    if user_prompts and all(re.sub(r"[^a-z0-9']+", " ", p.lower()).strip() in smalltalk for p in user_prompts):
        return

    session_id = st.session_state.get("current_session_id")
    if not session_id:
        session_id = f"chat_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6]}"
        st.session_state["current_session_id"] = session_id

    session_dir = SESSIONS_DIR / session_id
    session_dir.mkdir(parents=True, exist_ok=True)

    if not title:
        existing = session_dir / "session.json"
        if existing.exists():
            try:
                title = json.loads(existing.read_text("utf-8")).get("title")
            except Exception:
                pass

    # If title is empty or generic greeting, refine with meaningful message
    is_generic = (not title) or title.strip().lower() in ["hi", "hello", "hey", "proposal chat", "new chat", "you good"]
    if is_generic:
        substantive = [p.strip() for p in user_prompts if len(p.strip()) > 3 and p.strip().lower() not in ["hi", "hello", "hey", "you good"]]
        if substantive:
            p = substantive[0].strip()
            title = p[:28] + ("..." if len(p) > 28 else "")
        elif st.session_state.get("bid_meta", {}).get("name"):
            title = st.session_state["bid_meta"]["name"]
        elif user_prompts:
            title = user_prompts[0].strip()[:28]
        else:
            title = "Proposal Chat"

    # Save active BOM
    bom_info = None
    if st.session_state.get("active_bom"):
        bom_ext = Path(st.session_state["active_bom"]["name"]).suffix or ".xlsx"
        bom_filename = "active_bom" + bom_ext
        (session_dir / bom_filename).write_bytes(st.session_state["active_bom"]["bytes"])
        bom_info = {"name": st.session_state["active_bom"]["name"], "filename": bom_filename}

    # Save active tenders
    tenders_info = []
    for idx, t in enumerate(st.session_state.get("active_tenders", [])):
        t_ext = Path(t["name"]).suffix or ".txt"
        t_filename = f"tender_{idx}" + t_ext
        (session_dir / t_filename).write_bytes(t["bytes"])
        tenders_info.append({"name": t["name"], "filename": t_filename})

    # Serialize messages & save artifact files
    serialized_messages = []
    for m_idx, msg in enumerate(st.session_state.get("messages", [])):
        m_copy = {k: v for k, v in msg.items() if k not in ("artifacts",)}
        saved_artifacts = []
        for a_idx, art in enumerate(msg.get("artifacts", [])):
            a_copy = {k: v for k, v in art.items() if k != "bytes"}
            art_ext = ".docx" if art.get("type") == "docx" else ".md"
            art_filename = f"art_{m_idx}_{a_idx}" + art_ext
            if "bytes" in art and art["bytes"]:
                (session_dir / art_filename).write_bytes(art["bytes"])
            a_copy["filename"] = art_filename
            saved_artifacts.append(a_copy)
        m_copy["artifacts"] = saved_artifacts
        serialized_messages.append(m_copy)

    meta_data = {
        "id": session_id,
        "title": title,
        "updated_at": datetime.now().isoformat(),
        "bid_meta": st.session_state.get("bid_meta", {}),
        "active_bom": bom_info,
        "active_tenders": tenders_info,
        "messages": serialized_messages,
        "last_result": st.session_state.get("last_result"),
    }
    (session_dir / "session.json").write_text(json.dumps(meta_data, indent=2, default=str), encoding="utf-8")


# ==============================================================================
# Interface appearance and theme palette
# ==============================================================================

st.set_page_config(
    page_title="Kidus BID bot",
    layout="wide",
    initial_sidebar_state="auto",
)

THEMES = {
    "Dark": {
        "bg": "#071B2B", "sidebar": "#051420", "input": "#0B2236", "card": "#0A1E30",
        "text": "#F4F6FB", "secondary": "#C4D1E2", "muted": "rgba(196,209,226,.68)",
        "border": "rgba(196,209,226,.15)", "border_hover": "rgba(196,209,226,.30)",
        "accent": "#6875F5", "accent_hover": "#7D87FF", "bubble": "#10253B",
        "nav_hover": "rgba(104,117,245,.16)", "panel": "#0A2033", "panel_hover": "#102B43",
        "chip": "#10263C", "shadow": "rgba(0,0,0,.30)", "soft_accent": "rgba(104,117,245,.18)",
    },
    "Light": {
        "bg": "#F6F7FB", "sidebar": "#EEF0F7", "input": "#FFFFFF", "card": "#FFFFFF",
        "text": "#20263A", "secondary": "#46516A", "muted": "rgba(70,81,106,.68)",
        "border": "rgba(32,38,58,.13)", "border_hover": "rgba(32,38,58,.25)",
        "accent": "#5148D9", "accent_hover": "#4037C5", "bubble": "#ECEEFC",
        "nav_hover": "rgba(81,72,217,.09)", "panel": "#F0F2F8", "panel_hover": "#E5E8F2",
        "chip": "#F0F2F8", "shadow": "rgba(31,42,68,.09)", "soft_accent": "rgba(81,72,217,.10)",
    },
}
appearance = st.session_state.get("appearance_mode", "Dark")
theme = THEMES.get(appearance, THEMES["Dark"])

style_block = """
    <style>
    /* Google Fonts: Newsreader & Plus Jakarta Sans / Inter */
    @import url('https://fonts.googleapis.com/css2?family=Newsreader:ital,opsz,wght@0,6..72,300;0,6..72,400;0,6..72,500;0,6..72,600;1,6..72,400&family=Plus+Jakarta+Sans:wght@300;400;500;600&family=Inter:wght@300;400;500;600&display=swap');

    :root {
        color-scheme: __COLOR_SCHEME__;
        --claude-bg: __BG__;
        --claude-sidebar-bg: __SIDEBAR__;
        --claude-input-bg: __INPUT__;
        --claude-card-bg: __CARD__;
        --claude-text: __TEXT__;
        --claude-secondary: __SECONDARY__;
        --claude-muted: __MUTED__;
        --claude-border: __BORDER__;
        --claude-border-hover: __BORDER_HOVER__;
        --claude-accent: __ACCENT__;
        --claude-accent-hover: __ACCENT_HOVER__;
        --claude-user-bubble: __BUBBLE__;
        --claude-nav-hover: __NAV_HOVER__;
        --claude-panel: __PANEL__;
        --claude-panel-hover: __PANEL_HOVER__;
        --claude-chip: __CHIP__;
        --claude-shadow: __SHADOW__;
        --claude-soft-accent: __SOFT_ACCENT__;
    }

    /* Base Typography without breaking Streamlit icon fonts */
    body, .stMarkdown, p, textarea, input, button, label {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif !important;
        color: var(--claude-text);
    }

    /* Editorial Serif Headings */
    h1, h2, h3, h4, .claude-serif {
        font-family: 'Newsreader', 'Copernicus', 'Charter', 'Iowan Old Style', Georgia, serif !important;
        font-weight: 400 !important;
        letter-spacing: -0.015em !important;
        color: var(--claude-text) !important;
    }

    /* Main App Background */
    .stApp {
        background-color: var(--claude-bg) !important;
        color: var(--claude-text) !important;
        transition: background-color 0.22s ease, color 0.22s ease !important;
    }
    body, [data-testid="stAppViewContainer"], [data-testid="stBottom"],
    [data-testid="stBottom"] > div, [data-testid="stBottomBlockContainer"] {
        background-color: var(--claude-bg) !important;
        transition: background-color 0.22s ease !important;
    }
    [data-testid="stSidebar"], [data-testid="stChatInput"],
    [data-testid="stExpander"], .claude-artifact-card,
    .stChatMessage, .stButton > button, .stDownloadButton > button,
    [data-testid="stSelectbox"] [role="group"] {
        transition: background-color 0.22s ease, color 0.22s ease,
                    border-color 0.22s ease, box-shadow 0.22s ease !important;
    }

    /* Reading Column - Centered */
    [data-testid="stMainBlockContainer"] {
        max-width: 820px !important;
        padding-top: 1.5rem !important;
        padding-bottom: 7rem !important;
    }

    /* Hide unused Streamlit chrome while preserving the collapsed-sidebar control. */
    #MainMenu, footer {
        visibility: hidden;
        display: none !important;
    }
    header[data-testid="stHeader"] {
        display: block !important;
        visibility: visible !important;
        height: 0 !important;
        background: transparent !important;
        pointer-events: none !important;
    }
    header[data-testid="stHeader"] [data-testid="stToolbar"] {
        display: block !important;
        visibility: visible !important;
        height: 0 !important;
        overflow: visible !important;
    }
    header[data-testid="stHeader"] [data-testid="stToolbarActions"],
    header[data-testid="stHeader"] [data-testid="stAppDeployButton"] {
        display: none !important;
    }
    header[data-testid="stHeader"] [data-testid="stExpandSidebarButton"] {
        display: flex !important;
        visibility: visible !important;
        pointer-events: auto !important;
        position: fixed !important;
        top: 0.55rem !important;
        left: 0.55rem !important;
        z-index: 10000 !important;
        width: 44px !important;
        height: 44px !important;
        align-items: center !important;
        justify-content: center !important;
        color: var(--claude-text) !important;
        background: var(--claude-panel) !important;
        border: 1px solid var(--claude-border) !important;
        border-radius: 10px !important;
        box-shadow: 0 3px 12px var(--claude-shadow) !important;
        cursor: pointer !important;
    }
    header[data-testid="stHeader"] [data-testid="stExpandSidebarButton"] * {
        color: var(--claude-text) !important;
    }
    .st-key-theme_toggle_anchor {
        position: fixed !important;
        top: 0.55rem !important;
        right: 0.8rem !important;
        z-index: 10001 !important;
        width: 48px !important;
    }
    .st-key-model_picker_container {
        position: fixed !important;
        top: 0.55rem !important;
        right: 4.2rem !important;
        z-index: 10001 !important;
        width: 250px !important;
        max-width: calc(100vw - 6rem) !important;
    }
    .st-key-model_picker_container [data-testid="stSelectbox"] {
        width: 100% !important;
    }
    .st-key-model_picker_container [data-testid="stSelectbox"] > label {
        display: none !important;
    }
    .st-key-theme_toggle_anchor [data-testid="stBaseButton-secondary"] {
        width: 44px !important;
        min-width: 44px !important;
        height: 44px !important;
        padding: 0 !important;
        border: 1px solid var(--claude-border) !important;
        border-radius: 12px !important;
        background: var(--claude-panel) !important;
        color: var(--claude-text) !important;
        font-size: 1.25rem !important;
        box-shadow: 0 3px 12px var(--claude-shadow) !important;
    }
    .st-key-theme_toggle_anchor [data-testid="stBaseButton-secondary"] * {
        color: var(--claude-text) !important;
    }
    .st-key-theme_toggle_anchor [data-testid="stBaseButton-secondary"]:hover {
        background: var(--claude-panel-hover) !important;
        border-color: var(--claude-border-hover) !important;
    }

    /* Hide expander toggle arrow icons completely */
    [data-testid="stExpander"] summary svg,
    [data-testid="stExpanderToggleIcon"],
    [data-testid="stExpander"] summary [data-testid*="Icon"] {
        display: none !important;
        visibility: hidden !important;
    }
    [data-testid="stExpander"] summary {
        list-style: none !important;
    }
    [data-testid="stExpander"] summary::-webkit-details-marker {
        display: none !important;
    }
    [data-testid="stExpander"] {
        border-color: var(--claude-border) !important;
        background-color: var(--claude-panel) !important;
        border-radius: 8px !important;
    }

    /* Sidebar Styling */
    [data-testid="stSidebar"] {
        background-color: var(--claude-sidebar-bg) !important;
        border-right: 1px solid var(--claude-border) !important;
    }
    [data-testid="stSidebar"] * {
        color: var(--claude-secondary) !important;
    }
    [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {
        color: var(--claude-text) !important;
        font-family: 'Plus Jakarta Sans', sans-serif !important;
        font-weight: 500 !important;
        font-size: 0.95rem !important;
    }

    /* Sidebar Nav Item */
    .claude-nav-item {
        display: flex;
        align-items: center;
        padding: 7px 12px;
        border-radius: 8px;
        font-size: 0.92rem;
        color: var(--claude-secondary);
        cursor: pointer;
        transition: background 0.15s ease;
        margin-bottom: 2px;
    }
    .claude-nav-item:hover {
        background: var(--claude-nav-hover);
        color: var(--claude-text);
    }
    .claude-recent-item {
        padding: 6px 12px;
        font-size: 0.88rem;
        color: var(--claude-muted);
        border-radius: 6px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        cursor: pointer;
        transition: background 0.15s ease;
    }
    .claude-recent-item:hover {
        background: var(--claude-nav-hover);
        color: var(--claude-text);
    }

    /* Sidebar Recent Chat Buttons */
    [data-testid="stSidebar"] div[data-testid="stVerticalBlock"] div:has(button[key*="rec_sess_"]) button {
        background: transparent !important;
        border: 1px solid transparent !important;
        color: var(--claude-muted) !important;
        font-size: 0.88rem !important;
        font-weight: 400 !important;
        text-align: left !important;
        justify-content: flex-start !important;
        padding: 5px 8px !important;
        border-radius: 6px !important;
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
        transition: all 0.15s ease !important;
        box-shadow: none !important;
        margin-bottom: 2px !important;
    }
    [data-testid="stSidebar"] div[data-testid="stVerticalBlock"] div:has(button[key*="rec_sess_"]) button:hover {
        background: var(--claude-nav-hover) !important;
        color: var(--claude-text) !important;
        border-color: var(--claude-border-hover) !important;
    }

    /* '+ New' Button in Sidebar */
    [data-testid="stSidebar"] .claude-new-btn button {
        background: var(--claude-input-bg) !important;
        color: var(--claude-text) !important;
        border: 1px solid var(--claude-border) !important;
        border-radius: 8px !important;
        font-weight: 500 !important;
        font-size: 0.92rem !important;
        width: 100% !important;
        padding: 0.55rem 1rem !important;
        text-align: left !important;
    }
    [data-testid="stSidebar"] .claude-new-btn button * {
        color: var(--claude-text) !important;
    }
    [data-testid="stSidebar"] .claude-new-btn button:hover {
        background: var(--claude-panel-hover) !important;
        border-color: var(--claude-border-hover) !important;
    }

    /* Sidebar Profile Pill at Bottom */
    .claude-profile-pill {
        display: flex;
        align-items: center;
        gap: 10px;
        padding: 10px 12px;
        border-radius: 10px;
        background: transparent;
        margin-top: 1rem;
        border-top: 1px solid var(--claude-border);
    }
    .claude-avatar-circle {
        width: 26px;
        height: 26px;
        border-radius: 50%;
        background: var(--claude-accent);
        color: var(--claude-text);
        font-weight: 600;
        font-size: 0.82rem;
        display: flex;
        align-items: center;
        justify-content: center;
    }

    /* Claude Hero Greeting */
    .claude-hero-container {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 14px;
        margin-top: 5rem;
        margin-bottom: 2rem;
    }
    .claude-star-icon {
        font-size: 2.1rem;
        color: var(--claude-accent);
        display: inline-block;
        transform: translateY(-2px);
    }
    .claude-hero-title {
        font-family: 'Newsreader', 'Copernicus', Georgia, serif;
        font-size: 2.5rem;
        font-weight: 400;
        color: var(--claude-text);
        letter-spacing: -0.02em;
    }

    /* Claude Center Input Box */
    [data-testid="stChatInput"] {
        background-color: var(--claude-input-bg) !important;
        border: 1px solid var(--claude-border) !important;
        border-radius: 18px !important;
        box-shadow: 0 4px 20px var(--claude-shadow) !important;
        padding: 8px 14px !important;
    }
    [data-testid="stChatInput"] > div:first-child {
        background-color: var(--claude-input-bg) !important;
        border-color: var(--claude-border) !important;
    }
    [data-testid="stChatInput"] [data-testid="stChatInputSubmitButton"] {
        background-color: var(--claude-accent) !important;
        color: #FFFFFF !important;
    }
    [data-testid="stChatInput"]:focus-within {
        border-color: var(--claude-accent) !important;
    }
    [data-testid="stChatInput"] textarea {
        background-color: var(--claude-input-bg) !important;
        color: var(--claude-text) !important;
        font-size: 1rem !important;
        line-height: 1.5 !important;
    }
    [data-testid="stChatInput"] textarea::placeholder {
        color: var(--claude-muted) !important;
        font-size: 1rem !important;
    }
    [data-testid="stSelectbox"] [role="group"] {
        background-color: var(--claude-input-bg) !important;
        border-color: var(--claude-border) !important;
    }
    [data-testid="stSelectbox"] input {
        color: var(--claude-text) !important;
    }

    /* Message Feed */
    .stChatMessage {
        background: transparent !important;
        border: none !important;
        padding: 0.9rem 0 !important;
    }

    /* User Message Bubble */
    [data-testid="stChatMessage"]:has([data-testid="chatAvatarIcon-user"]) {
        background: var(--claude-user-bubble) !important;
        border-radius: 18px !important;
        padding: 12px 18px !important;
        border: 1px solid var(--claude-border) !important;
        margin-left: 2rem !important;
    }

    /* Action buttons use a cool indigo accent */
    .stButton > button, .stDownloadButton > button {
        background-color: var(--claude-accent) !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: 500 !important;
        font-size: 0.9rem !important;
        padding: 0.45rem 1rem !important;
        transition: background-color 0.15s ease, transform 0.1s ease !important;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        background-color: var(--claude-accent-hover) !important;
        transform: translateY(-1px);
    }
    [data-testid="stSidebar"] button[data-testid="stBaseButton-secondary"] {
        background-color: var(--claude-input-bg) !important;
        color: var(--claude-text) !important;
        border: 1px solid var(--claude-border) !important;
    }
    [data-testid="stSidebar"] button[data-testid="stBaseButton-secondary"] * {
        color: var(--claude-text) !important;
    }

    /* Claude Artifact Delivery Card in Dark Mode */
    .claude-artifact-card {
        background: var(--claude-card-bg) !important;
        border: 1px solid var(--claude-border) !important;
        border-radius: 14px !important;
        padding: 14px 18px !important;
        margin: 12px 0 8px 0 !important;
        box-shadow: 0 4px 16px var(--claude-shadow) !important;
    }
    .claude-artifact-badge {
        font-size: 0.72rem;
        background: var(--claude-soft-accent);
        color: var(--claude-secondary);
        border: 1px solid var(--claude-border-hover);
        padding: 2px 8px;
        border-radius: 5px;
        font-weight: 600;
        letter-spacing: 0.03em;
    }

    /* File Attachment & Stat Pills */
    .user-file-chip {
        display: inline-flex;
        align-items: center;
        background: var(--claude-chip);
        color: var(--claude-secondary);
        border: 1px solid var(--claude-border);
        border-radius: 6px;
        padding: 3px 9px;
        margin: 2px 5px 2px 0px;
        font-size: 0.82rem;
        font-weight: 500;
    }
    .stat-pill {
        display: inline-block;
        background: var(--claude-card-bg);
        color: var(--claude-secondary);
        border: 1px solid var(--claude-border);
        border-radius: 6px;
        padding: 3px 8px;
        font-size: 0.78rem;
        font-weight: 500;
        margin-right: 6px;
        margin-top: 4px;
    }

    /* Responsive layout for phones and smaller tablets */
    @media (max-width: 768px) {
        [data-testid="stMainBlockContainer"] {
            max-width: 100% !important;
            padding: 0.9rem 1rem calc(6.5rem + env(safe-area-inset-bottom)) !important;
        }
        .claude-hero-container {
            gap: 10px;
            margin-top: 2.2rem;
            margin-bottom: 1.4rem;
            padding: 0 0.25rem;
        }
        .claude-hero-title {
            font-size: clamp(1.7rem, 7vw, 2.25rem);
            line-height: 1.15;
        }
        .claude-star-icon { font-size: 1.8rem; }
        [data-testid="stChatMessage"] {
            padding: 0.55rem 0 !important;
            max-width: 100% !important;
        }
        [data-testid="stChatMessage"]:has([data-testid="chatAvatarIcon-user"]) {
            margin-left: 0.4rem !important;
            padding: 10px 12px !important;
            border-radius: 15px !important;
        }
        [data-testid="stChatInput"] {
            width: 100% !important;
            border-radius: 15px !important;
            padding: 6px 9px !important;
        }
        [data-testid="stChatInput"] textarea {
            font-size: 16px !important;
        }
        .stButton > button, .stDownloadButton > button {
            min-height: 44px !important;
            white-space: normal !important;
        }
        .claude-artifact-card { padding: 12px !important; }
        [data-testid="stSidebar"] {
            max-width: min(86vw, 340px) !important;
        }
        [data-testid="stSidebar"] [data-testid="stVerticalBlock"] {
            min-width: 0 !important;
        }
        [data-testid="stSidebar"] button {
            min-height: 42px !important;
        }
    }

    @media (max-width: 420px) {
        [data-testid="stMainBlockContainer"] {
            padding-left: 0.75rem !important;
            padding-right: 0.75rem !important;
        }
        [data-testid="stChatMessage"] [data-testid="stMarkdownContainer"] {
            overflow-wrap: anywhere;
        }
        .claude-hero-container { margin-top: 1.6rem; }
    }
    </style>
    """
style_block = style_block.replace("__COLOR_SCHEME__", "dark" if appearance == "Dark" else "light")
for key, value in theme.items():
    style_block = style_block.replace(f"__{key.upper()}__", value)
st.markdown(style_block, unsafe_allow_html=True)


# ==============================================================================
# Session State Initialization
# ==============================================================================

if "current_session_id" not in st.session_state:
    st.session_state["current_session_id"] = None

if "messages" not in st.session_state:
    st.session_state["messages"] = []

if "active_bom" not in st.session_state:
    st.session_state["active_bom"] = None

if "active_tenders" not in st.session_state:
    st.session_state["active_tenders"] = []

if "active_template" not in st.session_state:
    st.session_state["active_template"] = None

if "last_result" not in st.session_state:
    st.session_state["last_result"] = None

if "bid_meta" not in st.session_state:
    st.session_state["bid_meta"] = {
        "name": "",
        "client": "",
        "tender_no": "",
        "location": "",
    }


# ==============================================================================
# Claude Dark Mode Sidebar
# ==============================================================================

with st.sidebar:
    st.markdown(
        """
        <div style="padding: 0.6rem 0 0.8rem 0;">
            <span style="font-size: 1.15rem; font-weight: 600; color: var(--claude-text); letter-spacing: -0.01em;">Kidus BID bot</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # 'New chat' button
    st.markdown('<div class="claude-new-btn">', unsafe_allow_html=True)
    if st.button("New chat", key="new_chat_btn", use_container_width=True):
        if st.session_state.get("messages"):
            save_current_session()
        st.session_state["current_session_id"] = None
        st.session_state["messages"] = []
        st.session_state["active_bom"] = None
        st.session_state["active_tenders"] = []
        st.session_state["last_result"] = None
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    # Sidebar Nav Links
    st.markdown(
        """
        <div style="margin: 0.8rem 0;">
            <div class="claude-nav-item">Search</div>
            <div class="claude-nav-item">Projects</div>
            <div class="claude-nav-item">Artifacts</div>
            <div class="claude-nav-item">Customize</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("<hr style='border: none; border-top: 1px solid var(--claude-border); margin: 0.6rem 0;' />", unsafe_allow_html=True)

    # Dynamic Recents List (Real Stored Memories)
    st.markdown(
        """
        <div style="font-size: 0.78rem; font-weight: 500; color: var(--claude-secondary); padding: 0 4px 6px 4px; display: flex; justify-content: space-between;">
            <span>Recents</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    saved_sessions = list_sessions()
    if saved_sessions:
        for s in saved_sessions[:12]:
            s_id = s["id"]
            s_title = s.get("title") or "Proposal Chat"
            is_active = (s_id == st.session_state.get("current_session_id"))
            label = f"• {s_title}" if is_active else s_title
            if st.button(label, key=f"rec_sess_{s_id}", use_container_width=True):
                load_session(s_id)
                st.rerun()
    else:
        st.markdown(
            """
            <div style="font-size: 0.82rem; color: var(--claude-muted); padding: 4px 6px;">
                No recent chats yet
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<hr style='border: none; border-top: 1px solid var(--claude-border); margin: 0.8rem 0;' />", unsafe_allow_html=True)

    # Project controls inside neat expanders
    with st.expander("Project Metadata", expanded=False):
        st.session_state["bid_meta"]["name"] = st.text_input(
            "Project Name",
            value=st.session_state["bid_meta"]["name"],
        )
        st.session_state["bid_meta"]["client"] = st.text_input(
            "Client",
            value=st.session_state["bid_meta"]["client"],
        )
        st.session_state["bid_meta"]["tender_no"] = st.text_input(
            "Tender Number",
            value=st.session_state["bid_meta"]["tender_no"],
        )
        st.session_state["bid_meta"]["location"] = st.text_input(
            "Location",
            value=st.session_state["bid_meta"]["location"],
        )

    with st.expander("Staged Documents", expanded=False):
        if st.session_state["active_bom"]:
            st.markdown(f"**BOM:** `{st.session_state['active_bom']['name']}` ({format_bytes(len(st.session_state['active_bom']['bytes']))})")
            if st.button("Remove BOM", key="clear_bom_btn", use_container_width=True):
                st.session_state["active_bom"] = None
                st.rerun()
        else:
            st.caption("No BOM attached.")

        if st.session_state["active_tenders"]:
            st.markdown(f"**Tenders ({len(st.session_state['active_tenders'])}):**")
            for t_file in st.session_state["active_tenders"]:
                st.caption(f"• `{t_file['name']}` ({format_bytes(len(t_file['bytes']))})")
            if st.button("Remove Tenders", key="clear_tenders_btn", use_container_width=True):
                st.session_state["active_tenders"] = []
                st.rerun()

    lib = Library(LIBRARY_PATH)
    if not TEMPLATE_PATH.exists():
        if st.button("Init Word Template", key="init_starter_tpl_btn", use_container_width=True):
            ensure_starter_template(TEMPLATE_PATH, lib)
            st.rerun()

    if st.button("Load Sample Bid", key="load_sample_btn", use_container_width=True):
        if st.session_state.get("messages"):
            save_current_session()
        st.session_state["current_session_id"] = None
        st.session_state["bid_meta"] = {
            "name": "Campus Security & Surveillance Upgrade",
            "client": "Sample Organization",
            "tender_no": "DEMO-SEC-001",
            "location": "Sample Operations Campus",
        }
        st.session_state["active_bom"] = {
            "name": "sample_security_bom.xlsx",
            "bytes": create_demo_bom_bytes(),
        }
        st.session_state["active_tenders"] = [
            {"name": "tender_specifications.txt", "bytes": DEMO_TENDER_TEXT.encode("utf-8")}
        ]
        ensure_starter_template(TEMPLATE_PATH, lib)
        st.session_state.messages.append({
            "role": "user",
            "content": "Loaded the demo project files! Generate the technical proposal for the Sample Operations Campus.",
            "files": [
                {"name": "sample_security_bom.xlsx", "size": format_bytes(len(st.session_state["active_bom"]["bytes"])), "ext": ".xlsx"},
                {"name": "tender_specifications.txt", "size": format_bytes(len(DEMO_TENDER_TEXT.encode("utf-8"))), "ext": ".txt"},
            ],
            "artifacts": [],
            "timestamp": datetime.now().strftime("%H:%M"),
        })
        st.session_state["trigger_generation"] = "Loaded demo project. Generate complete proposal with executive summary."
        st.rerun()

    # User profile pill at the bottom of the sidebar
    st.markdown(
        """
        <div class="claude-profile-pill">
            <div class="claude-avatar-circle">D</div>
            <div style="flex: 1; font-size: 0.88rem; color: var(--claude-text); font-weight: 500;">
                Dybala
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ==============================================================================
# Main Page View: Exact Claude.ai Center Screen
# ==============================================================================

def toggle_appearance_mode():
    st.session_state["appearance_mode"] = (
        "Light" if st.session_state.get("appearance_mode", "Dark") == "Dark" else "Dark"
    )

with st.container(key="model_picker_container"):
    st.selectbox(
        "AI model",
        options=MODEL_OPTIONS,
        key="model_choice",
        label_visibility="collapsed",
        help="Choose a preferred model for chat and proposal work. The system falls back automatically if it is unavailable or rate limited. Web research continues through Gemini.",
    )

with st.container(key="theme_toggle_anchor"):
    current_mode = st.session_state.get("appearance_mode", "Dark")
    next_mode = "Light" if current_mode == "Dark" else "Dark"
    toggle_icon = "☀" if next_mode == "Light" else "☾"
    st.button(
        toggle_icon,
        key="theme_toggle_btn",
        help=f"Switch to {next_mode} mode",
        on_click=toggle_appearance_mode,
    )

# Calculate time of day greeting
hour = datetime.now().hour
greeting_time = "Morning" if hour < 12 else ("Afternoon" if hour < 18 else "Evening")

# If no messages yet, display the exact Claude welcome hero
if not st.session_state.messages:
    st.markdown(
        f"""
        <div style="height: 3rem;"></div>
        <div class="claude-hero-container">
            <span class="claude-star-icon">✦</span>
            <span class="claude-hero-title">{greeting_time}, Dybala</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

else:
    # Top bar when chat is active
    st.markdown(
        """
        <div style="display: flex; align-items: center; gap: 8px; padding-top: 0.2rem; margin-bottom: 1.5rem; border-bottom: 1px solid var(--claude-border); padding-bottom: 0.6rem;">
            <span style="color: var(--claude-accent); font-size: 1.3rem;">✦</span>
            <span style="font-family: 'Newsreader', Georgia, serif; font-size: 1.3rem; font-weight: 400; color: var(--claude-text);">Kidus BID bot</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Render message history
    for msg_idx, msg in enumerate(st.session_state.messages):
        avatar = "👤" if msg["role"] == "user" else "🔷"
        with st.chat_message(msg["role"], avatar=avatar):
            if msg.get("files"):
                st.markdown(
                    "".join(
                        f'<span class="user-file-chip">📎 {f["name"]} ({f["size"]})</span>'
                        for f in msg["files"]
                    ),
                    unsafe_allow_html=True,
                )

            st.markdown(msg["content"])

            if msg.get("stats"):
                stats = msg["stats"]
                st.markdown(
                    f'<span class="stat-pill">BOM Items: {stats.get("BOM items", 0)}</span>'
                    f'<span class="stat-pill">Sections Written: {stats.get("Products written", 0)}</span>'
                    f'<span class="stat-pill">Compliance Checks: {stats.get("Compliance rows", 0)}</span>'
                    f'<span class="stat-pill">Draft Specs: {stats.get("Draft products used", 0)}</span>',
                    unsafe_allow_html=True,
                )

            if msg.get("artifacts"):
                st.write("")
                for art_idx, art in enumerate(msg["artifacts"]):
                    if art["type"] == "docx":
                        st.markdown(
                            f"""
                            <div class="claude-artifact-card">
                                <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 4px;">
                                    <div style="display: flex; align-items: center; gap: 8px;">
                                        <span style="font-size: 1.25rem;">📄</span>
                                        <span style="font-weight: 600; color: var(--claude-text); font-size: 0.98rem;">{art['name']}</span>
                                    </div>
                                    <span class="claude-artifact-badge">PROPOSAL DOCUMENT</span>
                                </div>
                                <div style="font-size: 0.85rem; color: var(--claude-secondary); margin-bottom: 6px;">
                                    {art['description']} &bull; {art['size']}
                                </div>
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )
                        st.download_button(
                            label=f"Download {art['name']}",
                            data=art["bytes"],
                            file_name=art["name"],
                            mime=DOCX_MIME,
                            type="primary",
                            key=f"dl_docx_{msg_idx}_{art_idx}",
                            use_container_width=True,
                        )

                    elif art["type"] == "report":
                        st.markdown(
                            f"""
                            <div class="claude-artifact-card">
                                <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 4px;">
                                    <div style="display: flex; align-items: center; gap: 8px;">
                                        <span style="font-size: 1.25rem;">📋</span>
                                        <span style="font-weight: 600; color: var(--claude-text); font-size: 0.98rem;">{art['name']}</span>
                                    </div>
                                    <span class="claude-artifact-badge">AUDIT REPORT</span>
                                </div>
                                <div style="font-size: 0.85rem; color: var(--claude-secondary); margin-bottom: 6px;">
                                    {art['description']} &bull; {art['size']}
                                </div>
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )
                        st.download_button(
                            label=f"Download {art['name']}",
                            data=art["bytes"],
                            file_name=art["name"],
                            mime="text/markdown",
                            key=f"dl_report_{msg_idx}_{art_idx}",
                            use_container_width=True,
                        )
                        with st.expander(f"Review Report Summary: {art['name']}"):
                            st.markdown(art["text"])

            if msg.get("findings"):
                high_findings = [f for f in msg["findings"] if f.get("sev") == "high"]
                if high_findings:
                    with st.expander(f"⚠️ {len(high_findings)} High-Priority Compliance Warnings"):
                        for hf in high_findings:
                            st.markdown(f"**{hf.get('where', 'Item')}**: {hf.get('issue', '')}\n\n*Action:* {hf.get('fix', 'Review specification')}")

            if msg.get("logs"):
                with st.expander("Engine Execution Logs"):
                    st.code("\n".join(msg["logs"]), language=None)


# ==============================================================================
# Chat Input & Bottom Controls (Exact Match to Screenshot)
# ==============================================================================

triggered_instruction = st.session_state.pop("trigger_generation", None)

chat_val = st.chat_input(
    placeholder="How can I help you today?",
    accept_file="multiple",
    file_type=["xlsx", "xlsm", "docx", "pdf", "txt", "md", "csv"],
)

if chat_val or triggered_instruction:
    raw_text, uploaded_files = parse_chat_val(chat_val)
    if triggered_instruction:
        raw_text = triggered_instruction

    attached_chips = []
    new_bom_uploaded = False
    new_template_uploaded = False

    for up_file in uploaded_files:
        name = up_file.name
        ext = Path(name).suffix.lower()
        content = up_file.getvalue()
        attached_chips.append({
            "name": name,
            "size": format_bytes(len(content)),
            "ext": ext,
        })

        if ext in [".xlsx", ".xlsm"]:
            st.session_state["active_bom"] = {"name": name, "bytes": content}
            new_bom_uploaded = True
        elif ext == ".docx":
            if "template" in name.lower() or not TEMPLATE_PATH.exists():
                st.session_state["active_template"] = {"name": name, "bytes": content}
                new_template_uploaded = True
            else:
                st.session_state["active_tenders"].append({"name": name, "bytes": content})
        elif ext in ["." + t for t in TENDER_TYPES]:
            st.session_state["active_tenders"].append({"name": name, "bytes": content})

    user_prompt = raw_text.strip() or ("Generate technical proposal from the attached documents." if attached_chips else "")

    st.session_state.messages.append({
        "role": "user",
        "content": user_prompt,
        "files": attached_chips,
        "artifacts": [],
        "timestamp": datetime.now().strftime("%H:%M"),
    })

    active_bom = st.session_state.get("active_bom")
    intent_generate = (
        new_bom_uploaded
        or triggered_instruction
        or any(k in user_prompt.lower() for k in ["generate", "build", "create proposal", "proposal", "draft proposal", "run", "produce"])
    )

    if intent_generate and active_bom:
        with st.chat_message("assistant", avatar="🔷"):
            with st.status("Processing proposal documents...", expanded=True) as status:
                try:
                    result = run_pipeline(
                        bom_name=active_bom["name"],
                        bom_bytes=active_bom["bytes"],
                        tender_files_data=st.session_state["active_tenders"],
                        template_bytes=(
                            st.session_state["active_template"]["bytes"]
                            if st.session_state["active_template"]
                            else None
                        ),
                        meta=st.session_state["bid_meta"],
                        instructions=user_prompt,
                        status_container=status,
                        model_choice=st.session_state.get("model_choice", MODEL_AUTO),
                    )
                    status.update(label="Proposal package generated successfully", state="complete", expanded=False)
                    st.session_state["last_result"] = result

                    assistant_msg = {
                        "role": "assistant",
                        "content": (
                            f"### Proposal Package: **{st.session_state['bid_meta']['name']}**\n\n"
                            f"I have reviewed the Bill of Materials (`{active_bom['name']}`), evaluated tender requirements, "
                            f"researched technical specifications, and assembled your formal proposal documents."
                        ),
                        "stats": result["stats"],
                        "artifacts": [
                            {
                                "title": "Word Technical Proposal",
                                "name": result["docx_name"],
                                "bytes": result["docx_bytes"],
                                "size": result["docx_size"],
                                "type": "docx",
                                "description": f"Word proposal with {result['stats']['Products written']} product sections and {result['stats']['Compliance rows']} compliance rows.",
                            },
                            {
                                "title": "Compliance & Review Report",
                                "name": result["report_name"],
                                "bytes": result["report_bytes"],
                                "text": result["report_text"],
                                "size": result["report_size"],
                                "type": "report",
                                "description": "Full audit checklist with severity rankings, missing information alerts, and product research citations.",
                            },
                        ],
                        "findings": result["findings"],
                        "logs": result["logs"],
                        "timestamp": datetime.now().strftime("%H:%M"),
                    }
                    st.session_state.messages.append(assistant_msg)
                except Exception as exc:
                    status.update(label=f"Generation stopped: {exc}", state="error")
                    st.session_state.messages.append({
                        "role": "assistant",
                        "content": f"Could not generate proposal: {exc}\n\nPlease check your BOM and tender documents.",
                        "files": [],
                        "artifacts": [],
                        "timestamp": datetime.now().strftime("%H:%M"),
                    })

    elif intent_generate and not active_bom:
        st.session_state.messages.append({
            "role": "assistant",
            "content": (
                "**No Bill of Materials (BOM) found.**\n\n"
                "Please attach your BOM spreadsheet (`.xlsx` or `.xlsm`) using the `+` button in the input bar below, "
                "or click **Load Sample Bid (Demo)** in the sidebar to test with a pre-configured bid."
            ),
            "files": [],
            "artifacts": [],
            "timestamp": datetime.now().strftime("%H:%M"),
        })

    elif new_template_uploaded:
        lib = Library(LIBRARY_PATH)
        tpl_bytes = st.session_state["active_template"]["bytes"]
        with TemporaryDirectory() as tmp_d:
            tmp_p = Path(tmp_d) / "tpl.docx"
            tmp_p.write_bytes(tpl_bytes)
            ingest_docx(tmp_p, lib, None, CONFIG["template"]["name"])
        st.session_state.messages.append({
            "role": "assistant",
            "content": f"**Word Template `{st.session_state['active_template']['name']}` Ingested.**\n\nStyles and starter product entries have been updated.",
            "files": [],
            "artifacts": [],
            "timestamp": datetime.now().strftime("%H:%M"),
        })

    else:
        with st.chat_message("assistant", avatar="🔷"):
            with st.spinner("On it, cooking…"):
                reply = generate_conversational_reply(
                    query=user_prompt,
                    history=st.session_state.messages,
                    bid_meta=st.session_state.get("bid_meta", {}),
                    active_bom=st.session_state.get("active_bom"),
                    active_tenders=st.session_state.get("active_tenders", []),
                    last_result=st.session_state.get("last_result"),
                    model_choice=st.session_state.get("model_choice", MODEL_AUTO),
                )
        st.session_state.messages.append({
            "role": "assistant",
            "content": reply,
            "files": [],
            "artifacts": [],
            "timestamp": datetime.now().strftime("%H:%M"),
        })

    save_current_session()
    st.rerun()
