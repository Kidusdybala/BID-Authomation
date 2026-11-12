"""Local browser interface for uploading bid files and downloading proposals."""
from pathlib import Path
from tempfile import TemporaryDirectory
import os

import streamlit as st

from bidagent.config import load_config
from bidagent.ingest import ingest_docx
from bidagent.library import Library
from bidagent.llm import load_env
from bidagent.pipeline import Pipeline
from bidagent.util import slug


PROJECT_ROOT = Path(__file__).resolve().parent
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
TENDER_TYPES = ["pdf", "docx", "txt", "md", "csv", "xlsx", "xlsm"]
CONFIG = load_config(PROJECT_ROOT / "config.yaml")
TEMPLATE_PATH = CONFIG["_paths"]["library"] / "templates" / CONFIG["template"]["name"]
load_env(PROJECT_ROOT / ".env")


def show_result():
    result = st.session_state.get("bidagent_result")
    if not result:
        return
    st.divider()
    with st.chat_message("assistant", avatar="📄"):
        st.markdown("### Your proposal is ready")
        stats = result["stats"]
        st.write(
            f"Created the Word proposal with **{stats['BOM items']} BOM items**, "
            f"**{stats['Products written']} product sections**, and "
            f"**{stats['Compliance rows']} compliance rows**."
        )
        left, right = st.columns(2)
        left.download_button(
            "Download Word proposal",
            data=result["docx_bytes"],
            file_name=result["docx_name"],
            mime=DOCX_MIME,
            type="primary",
            width="stretch",
            key="download_proposal",
        )
        right.download_button(
            "Download review report",
            data=result["report_text"],
            file_name=result["report_name"],
            mime="text/markdown",
            width="stretch",
            key="download_review",
        )
        with st.expander("Read the review report"):
            st.markdown(result["report_text"])
        if result["logs"]:
            with st.expander("Build activity"):
                st.code("\n".join(result["logs"]), language=None)


st.set_page_config(page_title="BidAgent", page_icon="📄", layout="centered")
st.title("BidAgent")
st.caption("Turn your BOM and tender documents into a reviewed technical proposal.")

with st.chat_message("assistant", avatar="📄"):
    st.markdown(
        "Upload the bill of materials and any tender or bid documents. Add instructions if you want "
        "to guide the proposal, then generate and download the Word document."
    )

with st.form("generate_bid", clear_on_submit=False):
    st.subheader("1 · Bid details")
    name = st.text_input("Project or bid name", placeholder="e.g. MOD Antivirus")
    col_a, col_b = st.columns(2)
    client = col_a.text_input("Client", placeholder="Client or organization")
    tender_no = col_b.text_input("Tender number", placeholder="Optional")
    location = st.text_input("Project location", placeholder="Optional")

    st.subheader("2 · Upload files")
    bom_file = st.file_uploader(
        "Bill of materials (required)",
        type=["xlsx", "xlsm"],
        help="Upload the Excel workbook containing the items and quantities to offer.",
        max_upload_size=200,
    )
    tender_files = st.file_uploader(
        "Tender or bid documents",
        type=TENDER_TYPES,
        accept_multiple_files=True,
        help="Optional. Upload one or more tender files. Product datasheet PDFs are not required; the system can research models online.",
        max_upload_size=200,
    )
    template_file = st.file_uploader(
        "Finished bid to use as the Word template",
        type=["docx"],
        help="Required on first use. Optional if you have already ingested a template; upload another finished bid to replace it and refresh the starter library.",
        max_upload_size=200,
    )

    st.subheader("3 · Tell the system what you need")
    instructions = st.text_area(
        "Proposal instructions",
        placeholder="Add context or preferences, for example: use a formal tone and include an implementation approach.",
        height=110,
        help="Instructions guide the introduction. They are not treated as evidence for product specifications.",
    )
    submitted = st.form_submit_button("Generate technical proposal", type="primary", width="stretch")

if submitted:
    st.session_state.pop("bidagent_result", None)
    if bom_file is None:
        st.error("Upload the BOM Excel workbook to continue.")
    elif template_file is None and not TEMPLATE_PATH.exists():
        st.error("Upload a finished bid .docx to set up the Word template before the first proposal.")
    else:
        suffix = Path(bom_file.name).suffix.lower()
        if suffix not in (".xlsx", ".xlsm"):
            st.error("The BOM must be an .xlsx or .xlsm file.")
        else:
            bid_name = name.strip() or Path(bom_file.name).stem
            bid_slug = slug(bid_name)
            upload_root = PROJECT_ROOT / "work" / "uploads"
            upload_root.mkdir(parents=True, exist_ok=True)
            logs = []
            try:
                with st.spinner("Reading your files, researching products, checking compliance, and writing the proposal…"):
                    with TemporaryDirectory(prefix="bid-", dir=upload_root) as temp_dir:
                        temp_root = Path(temp_dir)
                        bom_path = temp_root / ("bom" + suffix)
                        bom_path.write_bytes(bom_file.getvalue())
                        tender_paths = []
                        for index, uploaded in enumerate(tender_files or [], start=1):
                            file_suffix = Path(uploaded.name).suffix.lower()
                            if file_suffix not in {"." + ext for ext in TENDER_TYPES}:
                                raise ValueError(f"Unsupported tender file type: {uploaded.name}")
                            path = temp_root / f"tender_{index}{file_suffix}"
                            path.write_bytes(uploaded.getvalue())
                            tender_paths.append(path)

                        cfg = CONFIG
                        if template_file is not None:
                            template_upload = temp_root / "finished_bid_template.docx"
                            template_upload.write_bytes(template_file.getvalue())
                            ingest_docx(template_upload, Library(cfg["_paths"]["library"]),
                                        None, cfg["template"]["name"])
                        pipeline = Pipeline(cfg, force=True, log=logs.append)
                        built = pipeline.run(
                            bom_path,
                            tender_paths=tender_paths,
                            name=bid_name,
                            client=client.strip(),
                            tender_no=tender_no.strip(),
                            location=location.strip(),
                            instructions=instructions.strip(),
                        )

                report_path = Path(built["report"])
                docx_path = Path(built["docx"])
                st.session_state["bidagent_result"] = {
                    "stats": built["stats"],
                    "docx_bytes": docx_path.read_bytes(),
                    "docx_name": docx_path.name,
                    "report_text": report_path.read_text(encoding="utf-8"),
                    "report_name": report_path.name,
                    "logs": logs,
                }
                st.session_state["bidagent_prompt"] = instructions.strip()
            except Exception as exc:
                st.error(f"Could not generate the proposal: {exc}")
                with st.expander("Technical details"):
                    st.exception(exc)

if st.session_state.get("bidagent_prompt"):
    with st.chat_message("user", avatar="👤"):
        st.write(st.session_state["bidagent_prompt"])

show_result()

with st.sidebar:
    st.subheader("Before you begin")
    st.write("Upload a finished bid in the form once to create the Word template and starter product library.")
    if os.environ.get("GEMINI_API_KEY"):
        st.success("Gemini key found — online product research is enabled.")
    else:
        st.info("Add GEMINI_API_KEY to .env to enable Google Search product research.")
    st.caption("Web-researched products are drafts. Check their linked sources and specifications before submitting.")
