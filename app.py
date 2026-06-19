import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional


APP_PATH = Path(__file__).resolve()
VENV_PYTHON = APP_PATH.parent / ".venv" / "Scripts" / "python.exe"


def launch_with_streamlit(python_executable: Path) -> None:
    """Run this file through Streamlit and keep the console attached."""
    streamlit_environment = os.environ.copy()
    streamlit_environment["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
    completed = subprocess.run(
        [
            str(python_executable),
            "-m",
            "streamlit",
            "run",
            "--server.headless",
            "true",
            "--browser.gatherUsageStats",
            "false",
            str(APP_PATH),
        ],
        check=False,
        env=streamlit_environment,
    )
    raise SystemExit(completed.returncode)


try:
    import streamlit as st
except ModuleNotFoundError as exc:
    if exc.name == "streamlit" and VENV_PYTHON.exists():
        launch_with_streamlit(VENV_PYTHON)
    raise SystemExit(
        "Streamlit is not installed. Create the project virtual environment and "
        "run: python -m pip install -r requirements.txt"
    ) from exc


if __name__ == "__main__":
    from streamlit.runtime.scriptrunner import get_script_run_ctx

    if get_script_run_ctx(suppress_warning=True) is None:
        runtime_python = VENV_PYTHON if VENV_PYTHON.exists() else Path(sys.executable)
        launch_with_streamlit(runtime_python)

from dotenv import dotenv_values, load_dotenv


ENV_PATH = APP_PATH.parent / ".env"
load_dotenv(dotenv_path=ENV_PATH, override=False)
ENV_FILE_HAS_API_KEY = bool(dotenv_values(ENV_PATH).get("OPENAI_API_KEY"))

from updated_rag_pipeline import (
    CATEGORY_TITLES,
    DOCUMENT_CATALOG,
    export_outputs,
    run_nonprofit_assistant,
)


st.set_page_config(
    page_title="Nonprofit Document Assistant",
    layout="wide",
)

st.title("Nonprofit Document Assistant")
st.caption(
    "Ask grounded questions or create fundraising, strategy, policy, program, "
    "compliance, communications, and community engagement documents."
)


def get_configured_api_key() -> Optional[str]:
    key = os.getenv("OPENAI_API_KEY")
    if key:
        return key.strip()
    try:
        return st.secrets.get("OPENAI_API_KEY")
    except Exception:
        return None


def save_uploaded_files(uploaded_files: Any) -> List[str]:
    upload_dir = Path(tempfile.mkdtemp(prefix="nonprofit-rag-uploads-"))
    saved_paths: List[str] = []
    for uploaded_file in uploaded_files:
        safe_name = Path(uploaded_file.name).name
        file_path = upload_dir / safe_name
        file_path.write_bytes(uploaded_file.getbuffer())
        saved_paths.append(str(file_path))
    return saved_paths


def safe_filename(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_-]+", "_", value).strip("_")
    return cleaned[:80] or "nonprofit_output"


def build_downloads(result: Dict[str, Any], filename_stem: str, title: str) -> Dict[str, Dict[str, Any]]:
    export_dir = Path(tempfile.mkdtemp(prefix="nonprofit-rag-exports-"))
    mime_types = {
        "pdf": "application/pdf",
        "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }
    downloads: Dict[str, Dict[str, Any]] = {}
    errors: Dict[str, str] = {}
    for output_format in mime_types:
        try:
            paths = export_outputs(
                result=result,
                output_dir=str(export_dir),
                formats=[output_format],
                filename_stem=filename_stem,
                title=title,
            )
            exported_path = Path(paths[output_format])
            downloads[output_format] = {
                "data": exported_path.read_bytes(),
                "file_name": exported_path.name,
                "mime": mime_types[output_format],
            }
        except Exception as exc:
            errors[output_format] = str(exc)
    if errors:
        downloads["_errors"] = errors
    return downloads


def render_citations(citations: List[Dict[str, Any]]) -> None:
    if not citations:
        st.info("No grounded citations were available for this response.")
        return

    for citation in citations:
        chunk_id = citation.get("chunk_id", citation["id"])
        st.markdown(
            f'<span id="citation-{chunk_id}"></span>',
            unsafe_allow_html=True,
        )
        label = (
            f"{citation['id']} | {citation['document_title']} | "
            f"{citation['location']}"
        )
        with st.expander(label):
            st.caption(
                f"Relevance: {citation.get('relevance', 0):.3f} | "
                f"Chunk: {chunk_id}"
            )
            st.markdown("**Exact excerpt**")
            st.info(citation["excerpt"])


def render_downloads(downloads: Dict[str, Dict[str, Any]]) -> None:
    st.subheader("Download Output")
    columns = st.columns(4)
    labels = {"pdf": "PDF", "docx": "Word", "pptx": "PowerPoint", "xlsx": "Excel"}
    for column, output_format in zip(columns, labels):
        with column:
            item = downloads.get(output_format)
            if item:
                st.download_button(
                    label=f"Download {labels[output_format]}",
                    data=item["data"],
                    file_name=item["file_name"],
                    mime=item["mime"],
                    key=f"download-{output_format}",
                    use_container_width=True,
                )
            else:
                st.button(
                    f"{labels[output_format]} unavailable",
                    disabled=True,
                    key=f"unavailable-{output_format}",
                    use_container_width=True,
                )
    for output_format, message in downloads.get("_errors", {}).items():
        st.warning(f"{labels.get(output_format, output_format)} export failed: {message}")


if "chat_history" not in st.session_state:
    st.session_state.chat_history = []
if "last_result" not in st.session_state:
    st.session_state.last_result = None
if "downloads" not in st.session_state:
    st.session_state.downloads = {}


st.sidebar.header("Settings")
configured_api_key = get_configured_api_key()
manual_api_key = st.sidebar.text_input(
    "OpenAI API key",
    type="password",
    help="Used only for this session. Environment variables and Streamlit secrets take priority.",
)
api_key = configured_api_key or manual_api_key or None
if ENV_FILE_HAS_API_KEY and configured_api_key:
    st.sidebar.success("OpenAI API key loaded from .env")
elif configured_api_key:
    st.sidebar.success("OpenAI API key configured in the environment")
elif manual_api_key:
    st.sidebar.success("Session API key entered")
else:
    st.sidebar.warning("Add OPENAI_API_KEY=your_key to the project .env file")

category_options = ["Question Answering"] + list(CATEGORY_TITLES.values())
selected_category_title = st.sidebar.selectbox("Task", category_options)
selected_document_type: Optional[str] = None

if selected_category_title != "Question Answering":
    selected_category = next(
        key for key, title in CATEGORY_TITLES.items() if title == selected_category_title
    )
    type_options = {
        spec["title"]: key
        for key, spec in DOCUMENT_CATALOG.items()
        if spec["category"] == selected_category
    }
    selected_type_title = st.sidebar.selectbox("Document type", list(type_options))
    selected_document_type = type_options[selected_type_title]
else:
    selected_type_title = "Grounded Answer"

st.sidebar.markdown("### Supported uploads")
st.sidebar.caption("PDF, Word, PowerPoint, Excel, CSV, TSV, text, and Markdown")

if st.sidebar.button("Clear history", use_container_width=True):
    st.session_state.chat_history = []
    st.session_state.last_result = None
    st.session_state.downloads = {}
    st.rerun()


uploaded_files = st.file_uploader(
    "Upload source documents",
    type=["pdf", "docx", "pptx", "xlsx", "csv", "tsv", "txt", "md"],
    accept_multiple_files=True,
)
if uploaded_files:
    st.success(f"{len(uploaded_files)} source file(s) ready.")

placeholder = (
    "Ask a grounded question about the uploaded documents."
    if selected_category_title == "Question Answering"
    else f"Describe the audience, purpose, and requirements for the {selected_type_title}."
)
user_query = st.text_area("Prompt", placeholder=placeholder, height=150)

if st.button("Generate", type="primary", use_container_width=True):
    if not uploaded_files:
        st.error("Upload at least one source document first.")
    elif not user_query.strip():
        st.error("Enter a question or document request.")
    elif not api_key:
        st.error("Configure an OpenAI API key before generating.")
    else:
        try:
            with st.spinner("Reading sources and generating a grounded response..."):
                saved_file_paths = save_uploaded_files(uploaded_files)
                result = run_nonprofit_assistant(
                    user_query=user_query,
                    uploaded_file_paths=saved_file_paths,
                    mode="qa" if selected_document_type is None else "generate",
                    document_type=selected_document_type,
                    api_key=api_key,
                    citation_base_url="#citation-{chunk_id}",
                )
                title = selected_type_title
                filename_stem = safe_filename(selected_type_title.lower())
                downloads = build_downloads(result, filename_stem, title)
                st.session_state.last_result = result
                st.session_state.downloads = downloads
                st.session_state.chat_history.append(
                    {
                        "task": selected_type_title,
                        "query": user_query,
                        "answer": result["answer"],
                        "citations": result.get("citations", []),
                    }
                )
        except Exception as exc:
            st.exception(exc)


if st.session_state.last_result:
    result = st.session_state.last_result
    st.subheader("Response")
    st.markdown(result["answer"])
    confidence_column, fallback_column, citation_column = st.columns(3)
    confidence_column.metric("Confidence", result.get("confidence", "unknown").title())
    fallback_column.metric("Fallback Used", "Yes" if result.get("fallback_used") else "No")
    citation_column.metric("Citations", len(result.get("citations", [])))
    render_downloads(st.session_state.downloads)
    st.subheader("Grounded Citations")
    render_citations(result.get("citations", []))


if st.session_state.chat_history:
    st.subheader("History")
    for index, item in enumerate(reversed(st.session_state.chat_history), 1):
        with st.expander(f"{index}. {item['task']} - {item['query'][:80]}"):
            st.markdown(f"**Prompt:** {item['query']}")
            st.markdown(item["answer"])
            render_citations(item["citations"])
