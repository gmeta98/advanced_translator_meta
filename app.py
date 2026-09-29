from __future__ import annotations

import base64
import hashlib
import hmac
import os
import json
import zipfile
from collections import Counter
from dataclasses import dataclass
from io import BytesIO
from typing import Iterable
from uuid import uuid4

import streamlit as st
import streamlit.components.v1 as components
from docx import Document
from docx.document import Document as DocumentObject
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
from docx.oxml.text.run import CT_Text
from dotenv import load_dotenv

try:
    from openai import OpenAI
except ImportError:  # pragma: no cover - handled in the UI
    OpenAI = None


load_dotenv(override=False)

LANGUAGES = [
    "Albanian",
    "Italian",
    "English",
    "French",
    "German",
    "Spanish",
    "Portuguese",
    "Greek",
    "Custom",
]

DOCUMENT_TYPE_CONTEXTS = {
    "Legal / judicial file": (
        "Formal legal or judicial file. Preserve legal force, procedural meaning, "
        "institutional names, dates, protocol numbers, article references, party roles, "
        "official titles, and document hierarchy."
    ),
    "Police / investigative file": (
        "Police or investigative file. Preserve investigative terminology, speaker roles, "
        "surveillance or transcript context, dates, times, locations, identifying details, "
        "and concise official register."
    ),
    "Financial police / Guardia di Finanza file": (
        "Italian Guardia di Finanza or financial-police material. Preserve authority names, "
        "tax/financial/legal terminology, protocol references, offices, ranks, investigative "
        "acts, and formal administrative style."
    ),
    "Court / prosecution document": (
        "Court, prosecution, or criminal-procedure document. Preserve procedural terms, "
        "charges, article references, names, roles, chronology, exhibits, and formal legal register."
    ),
    "Business document": (
        "Business or administrative document. Preserve commercial terms, company names, "
        "dates, amounts, obligations, references, and professional tone."
    ),
    "Medical document": (
        "Medical document. Preserve clinical terminology, measurements, dates, diagnoses, "
        "medications, and professional tone."
    ),
    "Academic document": (
        "Academic document. Preserve terminology, citations, titles, names, references, and formal tone."
    ),
    "Immigration document": (
        "Immigration or civil-status document. Preserve authority names, personal data, dates, "
        "case references, legal status terms, and official register."
    ),
    "Technical document": (
        "Technical document. Preserve technical terminology, measurements, labels, references, "
        "and concise professional tone."
    ),
    "General professional document": (
        "General professional document. Preserve meaning, names, dates, references, and professional tone."
    ),
}

DOCUMENT_TYPE_PROMPT_LABELS = {
    "Legal / judicial file": "legal/judicial",
    "Police / investigative file": "police/investigative",
    "Financial police / Guardia di Finanza file": "financial-police/Guardia di Finanza",
    "Court / prosecution document": "court/prosecution",
    "Business document": "business/administrative",
    "Medical document": "medical",
    "Academic document": "academic",
    "Immigration document": "immigration/civil-status",
    "Technical document": "technical",
    "General professional document": "general professional",
}

DEFAULT_SYSTEM_PROMPT = """You are a professional legal and business document translator.
Translate faithfully from {source_language} to {target_language}.
Preserve meaning, names, numbers, dates, addresses, references, section numbering, and punctuation intent.
Use natural, professional {target_language} as used by native speakers in official translation practice.
Prefer the meaning required by the local context over literal word-for-word cognates.
Avoid unnatural borrowed terms when a standard, idiomatic {target_language} term fits the sentence better.
For legal, judicial, police, financial-police, prosecution, or administrative documents, preserve legal effect and procedural nuance. Use standard official/legal terminology in {target_language}; do not simplify, soften, intensify, or paraphrase legal meaning.
For transcripts, police/legal notes, speaker descriptions, and conversational summaries, translate the practical meaning in context.
If a text item is a table-cell fragment rather than a full sentence, infer its role from nearby items and translate it as a natural fragment.
Do not summarize, explain, add commentary, or omit content.
Return only the requested translation output."""

BATCH_MAX_ITEMS = 40
BATCH_MAX_CHARS = 6000
ROLLING_CONTEXT_PAIRS = 12
DOCUMENT_CONTEXT_ITEMS = 80
DOCUMENT_CONTEXT_CHARS = 5000
MAX_TRANSLATION_ATTEMPTS = 2
XML_SPACE = qn("xml:space")
TEXT_BOUNDARY_TAGS = {qn(tag) for tag in ("w:tab", "w:ptab", "w:br", "w:cr", "w:noBreakHyphen")}


@dataclass(frozen=True)
class TextUnit:
    location: str
    text: str
    target_index: int


@dataclass(frozen=True)
class DocumentAnalysis:
    text_items: int
    unique_text_items: int
    repeated_text_items: int
    estimated_batches: int
    repeated_header_rows: list[tuple[tuple[str, ...], int]]


def iter_text_targets(document: DocumentObject) -> Iterable[tuple[str, Paragraph]]:
    seen_paragraphs: set[str] = set()

    for location, paragraph in iter_raw_text_targets(document):
        paragraph_id = f"{paragraph.part.partname}:{paragraph._p.getroottree().getpath(paragraph._p)}"
        if paragraph_id in seen_paragraphs:
            continue
        seen_paragraphs.add(paragraph_id)
        yield location, paragraph


def iter_raw_text_targets(document: DocumentObject) -> Iterable[tuple[str, Paragraph]]:
    yield from iter_block_paragraphs(document.element.body, document, "Body")

    for part_index, (part_type, part) in enumerate(iter_existing_header_footer_parts(document), start=1):
        yield from iter_block_paragraphs(part.element, part, f"{part_type} {part_index}")


def iter_existing_header_footer_parts(document: DocumentObject) -> Iterable[tuple[str, object]]:
    seen_parts: set[str] = set()
    for relationship in document.part.rels.values():
        if relationship.reltype not in (RT.HEADER, RT.FOOTER):
            continue

        part = relationship.target_part
        part_name = str(part.partname)
        if part_name in seen_parts:
            continue

        seen_parts.add(part_name)
        part_type = "Header" if relationship.reltype == RT.HEADER else "Footer"
        yield part_type, part


def iter_block_paragraphs(container, parent, prefix: str) -> Iterable[tuple[str, Paragraph]]:
    paragraph_index = 0
    table_index = 0

    for child in container.iterchildren():
        if child.tag == qn("w:p"):
            paragraph_index += 1
            yield f"{prefix} paragraph {paragraph_index}", Paragraph(child, parent)
        elif child.tag == qn("w:tbl"):
            table_index += 1
            yield from iter_table_paragraphs(Table(child, parent), f"{prefix} table {table_index}")


def iter_table_paragraphs(table: Table, prefix: str) -> Iterable[tuple[str, Paragraph]]:
    for row_index, row in enumerate(table.rows, start=1):
        for cell_index, cell in enumerate(row.cells, start=1):
            yield from iter_cell_paragraphs(cell, f"{prefix}, row {row_index}, cell {cell_index}")


def iter_cell_paragraphs(cell: _Cell, prefix: str) -> Iterable[tuple[str, Paragraph]]:
    yield from iter_block_paragraphs(cell._tc, cell, prefix)


def iter_paragraph_segments(paragraph: Paragraph) -> Iterable[list[CT_Text]]:
    """Group text nodes without crossing an existing tab or break."""
    nodes: list[CT_Text] = []
    # Match the runs and hyperlinks read by python-docx. Descendant text in
    # drawings/text boxes belongs to other paragraphs and must not be overwritten.
    for node in paragraph._p.xpath("./w:r/* | ./w:hyperlink/w:r/*"):
        if node.tag == qn("w:t"):
            nodes.append(node)
        elif node.tag in TEXT_BOUNDARY_TAGS and nodes:
            yield nodes
            nodes = []
    if nodes:
        yield nodes


def iter_translation_targets(document: DocumentObject) -> Iterable[tuple[str, list[CT_Text]]]:
    for location, paragraph in iter_text_targets(document):
        segments = list(iter_paragraph_segments(paragraph))
        for segment_index, nodes in enumerate(segments, start=1):
            segment_location = f"{location}, segment {segment_index}" if len(segments) > 1 else location
            yield segment_location, nodes


def collect_text_units(document: DocumentObject) -> list[TextUnit]:
    units: list[TextUnit] = []

    for text_target_index, (location, nodes) in enumerate(iter_translation_targets(document)):
        text = "".join(node.text or "" for node in nodes).strip()
        if text:
            units.append(TextUnit(location, text, text_target_index))

    return units


def collect_table_header_rows(document: DocumentObject) -> list[tuple[str, ...]]:
    rows: list[tuple[str, ...]] = []

    for table in document.tables:
        if not table.rows:
            continue

        seen_paragraphs: set[str] = set()
        row_text: list[str] = []
        for cell in table.rows[0].cells:
            cell_parts: list[str] = []
            for paragraph in cell.paragraphs:
                paragraph_id = paragraph._p.getroottree().getpath(paragraph._p)
                if paragraph_id in seen_paragraphs:
                    continue
                seen_paragraphs.add(paragraph_id)
                text = paragraph.text.strip()
                if text:
                    cell_parts.append(text)

            cell_text = " / ".join(cell_parts).strip()
            if cell_text:
                row_text.append(cell_text)

        if row_text:
            rows.append(tuple(row_text))

    return rows


def analyze_document(document: DocumentObject) -> DocumentAnalysis:
    units = collect_text_units(document)
    unique_units = list({unit.text: unit for unit in units}.values())
    header_counts = Counter(collect_table_header_rows(document))
    repeated_headers = [
        (row, count)
        for row, count in header_counts.most_common()
        if count > 1
    ]

    return DocumentAnalysis(
        text_items=len(units),
        unique_text_items=len(unique_units),
        repeated_text_items=len(units) - len(unique_units),
        estimated_batches=len(build_batches(unique_units)),
        repeated_header_rows=repeated_headers,
    )


def set_text_segment(text_nodes: list[CT_Text], text: str) -> None:
    if not text_nodes:
        return

    replacement = " ".join(text.replace("\t", " ").splitlines()).strip()
    if not replacement:
        raise ValueError("A translation was empty; the original text was not replaced.")
    original_text = "".join(node.text or "" for node in text_nodes)
    leading_space = original_text[:len(original_text) - len(original_text.lstrip())]
    trailing_space = original_text[len(original_text.rstrip()):]
    replacement = leading_space + replacement + trailing_space
    original_lengths = [len(node.text or "") for node in text_nodes]
    chunks = split_text_for_existing_nodes(replacement, original_lengths)

    for node, chunk in zip(text_nodes, chunks):
        set_text_node_text(node, chunk)


def set_text_node_text(node, text: str) -> None:
    node.text = text
    if text:
        node.set(XML_SPACE, "preserve")
    elif XML_SPACE in node.attrib:
        del node.attrib[XML_SPACE]


def split_text_for_existing_nodes(text: str, original_lengths: list[int]) -> list[str]:
    if len(original_lengths) == 1:
        return [text]

    total_original = sum(original_lengths)
    if total_original <= 0:
        return [text] + [""] * (len(original_lengths) - 1)

    chunks: list[str] = []
    cursor = 0
    text_length = len(text)

    for index, original_length in enumerate(original_lengths):
        if index == len(original_lengths) - 1:
            chunks.append(text[cursor:])
            break

        target_length = round(text_length * (original_length / total_original))
        remaining_nodes = len(original_lengths) - index - 1
        max_cursor = max(cursor, text_length - remaining_nodes)
        next_cursor = min(cursor + target_length, max_cursor)

        if cursor < next_cursor < text_length and not text[next_cursor].isspace():
            left_space = text.rfind(" ", cursor, next_cursor)
            right_space = text.find(" ", next_cursor)
            if left_space > cursor and next_cursor - left_space <= 12:
                next_cursor = left_space
            elif right_space != -1 and right_space - next_cursor <= 12:
                next_cursor = right_space

        chunks.append(text[cursor:next_cursor])
        cursor = next_cursor

    return chunks


def replace_text_unit(document: DocumentObject, unit: TextUnit, translated_text: str) -> None:
    text_targets = [nodes for _, nodes in iter_translation_targets(document)]
    set_text_segment(text_targets[unit.target_index], translated_text)


def build_preview(units: list[TextUnit], max_items: int = 12) -> str:
    preview_lines = []
    for unit in units[:max_items]:
        preview_lines.append(f"{unit.location}\n{unit.text}")
    return "\n\n".join(preview_lines)


def format_header_row(row: tuple[str, ...]) -> str:
    return " | ".join(row)


def get_config_value(name: str) -> str:
    environment_value = os.getenv(name, "").strip()
    if environment_value:
        return environment_value

    try:
        secret_value = st.secrets[name]
    except (FileNotFoundError, KeyError):
        return ""

    return str(secret_value).strip()


def get_configured_client() -> OpenAI | None:
    api_key = get_config_value("OPENAI_API_KEY")
    if not api_key or OpenAI is None:
        return None
    return OpenAI(api_key=api_key)


def build_batches(
    items: list[TextUnit],
    *,
    max_items: int = BATCH_MAX_ITEMS,
    max_chars: int = BATCH_MAX_CHARS,
) -> list[list[TextUnit]]:
    batches: list[list[TextUnit]] = []
    current: list[TextUnit] = []
    current_chars = 0

    for item in items:
        item_chars = len(item.text)
        would_exceed_items = len(current) >= max_items
        would_exceed_chars = current and current_chars + item_chars > max_chars
        if would_exceed_items or would_exceed_chars:
            batches.append(current)
            current = []
            current_chars = 0

        current.append(item)
        current_chars += item_chars

    if current:
        batches.append(current)

    return batches


def build_document_context(units: list[TextUnit], user_context: str) -> str:
    sampled_text = "\n".join(
        f"- {unit.text}"
        for unit in units[:DOCUMENT_CONTEXT_ITEMS]
        if len(unit.text) > 2
    )
    if len(sampled_text) > DOCUMENT_CONTEXT_CHARS:
        sampled_text = sampled_text[:DOCUMENT_CONTEXT_CHARS].rsplit("\n", 1)[0]

    return f"""User document context:
{user_context or "No extra context provided."}

Early document text for context:
{sampled_text or "No preview text available."}"""


def build_profile_context(subject_matter: str) -> str:
    return DOCUMENT_TYPE_CONTEXTS.get(
        subject_matter,
        DOCUMENT_TYPE_CONTEXTS["General professional document"],
    )


def build_recent_context(pairs: list[tuple[TextUnit, str]]) -> str:
    if not pairs:
        return "No previous translations in this job yet."

    recent_pairs = pairs[-ROLLING_CONTEXT_PAIRS:]
    return "\n".join(
        f"- Source: {unit.text}\n  Translation: {translated_text}"
        for unit, translated_text in recent_pairs
    )


def parse_translation_response(response_text: str, units: list[TextUnit]) -> list[str]:
    start = response_text.find("{")
    end = response_text.rfind("}")
    if start == -1 or end == -1 or end <= start:
        raise ValueError("The translation response was not valid JSON.")

    payload = json.loads(response_text[start : end + 1])
    if not isinstance(payload, dict):
        raise ValueError("The translation response was not a JSON object.")
    translations = payload.get("translations")
    if not isinstance(translations, list):
        raise ValueError("The translation response did not include a translations list.")

    by_id: dict[int, str] = {}
    expected_ids = {unit.target_index for unit in units}
    for item in translations:
        if not isinstance(item, dict):
            raise ValueError("The translation response contained an invalid text item.")
        item_id = item.get("id")
        translated_text = item.get("text")
        if type(item_id) is not int or item_id not in expected_ids:
            raise ValueError("The translation response contained an unexpected text item id.")
        if item_id in by_id:
            raise ValueError("The translation response contained a duplicate text item id.")
        if not isinstance(translated_text, str) or not translated_text.strip():
            raise ValueError("The translation response contained an empty or invalid translation.")
        by_id[item_id] = translated_text.strip()

    missing = [unit.target_index for unit in units if unit.target_index not in by_id]
    if missing:
        raise ValueError(f"The translation response missed {len(missing)} text items.")

    return [by_id[unit.target_index].strip() for unit in units]


def translate_units(
    units: list[TextUnit],
    *,
    source_language: str,
    target_language: str,
    subject_matter: str,
    style_instruction: str,
    glossary: str,
    document_context: str,
    recent_context: str,
    model: str,
) -> list[str]:
    client = get_configured_client()
    if client is None:
        return [f"[Preview translation to {target_language}] {unit.text}" for unit in units]

    system_prompt = DEFAULT_SYSTEM_PROMPT.format(
        source_language=source_language,
        target_language=target_language,
    )

    items_json = json.dumps(
        [{"id": unit.target_index, "text": unit.text} for unit in units],
        ensure_ascii=False,
        indent=2,
    )
    user_prompt = f"""Subject matter: {subject_matter or "General professional document"}
Document profile:
{build_profile_context(subject_matter)}

Style requirements: {style_instruction or "Professional, accurate, and natural"}
Glossary and required terminology:
{glossary or "None provided"}

Document-wide context:
{document_context}

Recent translation choices from this same job:
{recent_context}

Translate each JSON item from {source_language} to {target_language}.
Use nearby items in this batch as context for terminology and tone, but translate each item separately.
An item may be part of a paragraph separated by a tab or line break. Those separators are preserved by the app; do not add tabs or line breaks inside a translation.
Choose context-appropriate local wording, not literal cognates, when the source term has several possible meanings.
Do not preserve source-language syntax when it would sound unnatural in {target_language}.
Keep terminology, register, speaker labels, and recurring phrases consistent with the document-wide context and recent translation choices.
For legal prose, translate sentence-by-sentence with exact legal meaning; preserve obligations, negations, conditions, procedural terms, and references.
For names of public bodies, offices, laws, articles, forms, exhibits, protocol numbers, addresses, emails, phone numbers, codes, and file references, preserve the identifying content exactly unless a standard translated institution name is clearly appropriate.
Keep the same id values. Do not merge, split, remove, renumber, or reorder items.
Return only valid JSON in this exact shape:
{{"translations":[{{"id":0,"text":"translated text"}}]}}

Text items:
{items_json}"""

    last_error: Exception | None = None
    for attempt in range(1, MAX_TRANSLATION_ATTEMPTS + 1):
        repair_instruction = ""
        if attempt > 1:
            repair_instruction = (
                "\nThe previous response could not be parsed. Return only valid JSON, "
                "with every original id present exactly once and a non-empty translation for each item."
            )

        response = client.responses.create(
            model=model,
            input=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt + repair_instruction},
            ],
        )
        try:
            return parse_translation_response(response.output_text.strip(), units)
        except ValueError as exc:
            last_error = exc

    raise ValueError(f"Could not parse translation response after retry: {last_error}")


def translate_document(
    source_bytes: bytes,
    *,
    source_language: str,
    target_language: str,
    subject_matter: str,
    style_instruction: str,
    glossary: str,
    user_context: str,
    model: str,
) -> tuple[bytes, list[tuple[TextUnit, str]]]:
    document = Document(BytesIO(source_bytes))
    units = collect_text_units(document)
    translated_pairs: list[tuple[TextUnit, str]] = []
    text_targets = [nodes for _, nodes in iter_translation_targets(document)]
    document_context = build_document_context(units, user_context)
    translation_cache: dict[str, str] = {}
    recent_pairs: list[tuple[TextUnit, str]] = []

    progress = st.progress(0, text="Preparing document")
    translated_count = 0
    unique_units: list[TextUnit] = []
    seen_texts: set[str] = set()
    for unit in units:
        if unit.text in seen_texts:
            continue
        seen_texts.add(unit.text)
        unique_units.append(unit)

    try:
        batches = build_batches(unique_units)
        for batch_index, batch in enumerate(batches, start=1):
            progress.progress(
                translated_count / max(len(unique_units), 1),
                text=f"Translating batch {batch_index} of {len(batches)}",
            )
            translated_texts = translate_units(
                batch,
                source_language=source_language,
                target_language=target_language,
                subject_matter=subject_matter,
                style_instruction=style_instruction,
                glossary=glossary,
                document_context=document_context,
                recent_context=build_recent_context(recent_pairs),
                model=model,
            )
            for unit, translated_text in zip(batch, translated_texts):
                translation_cache[unit.text] = translated_text
                recent_pairs.append((unit, translated_text))
                translated_count += 1

            progress.progress(
                translated_count / max(len(unique_units), 1),
                text=f"Translated {translated_count} of {len(unique_units)} unique text items",
            )

        progress.progress(1.0, text="Applying translations to the Word file")
        for unit in units:
            translated_text = translation_cache[unit.text]
            set_text_segment(text_targets[unit.target_index], translated_text)
            translated_pairs.append((unit, translated_text))

        output = BytesIO()
        document.save(output)
        return output.getvalue(), translated_pairs
    finally:
        progress.empty()


def resolve_language(choice: str, custom: str) -> str:
    return custom.strip() if choice == "Custom" and custom.strip() else choice


def start_automatic_download(data: bytes, file_name: str, mime: str) -> None:
    """Request a browser download once, when a translation job finishes."""
    encoded_data = base64.b64encode(data).decode("ascii")
    # Escape '<' so a file name cannot close the script element.
    download_name = json.dumps(file_name).replace("<", "\\u003c")
    content_type = json.dumps(mime).replace("<", "\\u003c")
    # A fresh marker reloads the iframe even when another job has identical output.
    components.html(
        f"""<!doctype html>
<html><body>
<script data-download-id="{uuid4().hex}">
    const binary = atob("{encoded_data}");
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {{
        bytes[i] = binary.charCodeAt(i);
    }}
    const url = URL.createObjectURL(new Blob([bytes], {{type: {content_type}}}));
    const link = document.createElement("a");
    link.href = url;
    link.download = {download_name};
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
</script>
</body></html>""",
        height=0,
        scrolling=False,
    )


def document_type_prompt_label(subject_matter: str) -> str:
    return DOCUMENT_TYPE_PROMPT_LABELS.get(subject_matter, "general professional")


def default_style_requirements(subject_matter: str) -> str:
    document_label = document_type_prompt_label(subject_matter)
    return (
        "Formal, accurate, natural, and suitable for certified translation review. "
        f"Preserve {document_label} meaning, official tone, names, dates, article references, "
        "protocol numbers, and institutional terminology."
    )


def default_document_context(source_language: str, subject_matter: str) -> str:
    document_label = document_type_prompt_label(subject_matter)
    language_label = source_language if source_language != "Custom" else "Source-language"
    return (
        f"{language_label} {document_label} file. Some documents are prose-heavy "
        f"{document_label} records; others contain tables, transcript summaries, speaker "
        "descriptions, dates, names, offices, protocol numbers, and conversational notes."
    )


def refresh_defaults_for_document_type() -> None:
    subject_matter = st.session_state.get("subject_matter", "Legal / judicial file")
    source_language = resolve_language(
        st.session_state.get("source_choice", "Italian"),
        st.session_state.get("source_custom", ""),
    )
    st.session_state["style_instruction"] = default_style_requirements(subject_matter)
    st.session_state["user_context"] = default_document_context(source_language, subject_matter)


def refresh_document_context_for_source() -> None:
    subject_matter = st.session_state.get("subject_matter", "Legal / judicial file")
    source_language = resolve_language(
        st.session_state.get("source_choice", "Italian"),
        st.session_state.get("source_custom", ""),
    )
    st.session_state["user_context"] = default_document_context(source_language, subject_matter)


def translation_job_signature(uploaded_files: Iterable, **settings: str | bool) -> str:
    files = [
        (file.name, hashlib.sha256(file.getvalue()).hexdigest())
        for file in uploaded_files
    ]
    payload = json.dumps({"files": files, "settings": settings}, sort_keys=True)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def clear_translation_results() -> None:
    for key in (
        "translated_docx",
        "translated_zip",
        "translated_pairs",
        "download_name",
        "translation_job_signature",
    ):
        st.session_state.pop(key, None)


def log_out() -> None:
    clear_translation_results()
    for key in (
        "authenticated",
        "login_error",
        "login_password",
        "uploaded_files",
    ):
        st.session_state.pop(key, None)


def check_password() -> None:
    configured_password = get_config_value("APP_PASSWORD")
    submitted_password = st.session_state.get("login_password", "")
    password_matches = bool(configured_password) and hmac.compare_digest(
        submitted_password.encode("utf-8"),
        configured_password.encode("utf-8"),
    )

    st.session_state["authenticated"] = password_matches
    st.session_state["login_error"] = not password_matches
    st.session_state.pop("login_password", None)


def require_password() -> None:
    configured_password = get_config_value("APP_PASSWORD")
    if not configured_password:
        st.error(
            "App password is not configured. Add APP_PASSWORD to your local .env file "
            "or to Streamlit Cloud Secrets."
        )
        st.stop()

    if st.session_state.get("authenticated"):
        return

    st.title("Document Translation Office")
    st.subheader("Sign in")

    with st.form("login_form"):
        st.text_input("Password", type="password", key="login_password")
        st.form_submit_button(
            "Sign in",
            type="primary",
            on_click=check_password,
            use_container_width=True,
        )

    if st.session_state.get("login_error"):
        st.error("Incorrect password.")

    st.stop()


def main() -> None:
    st.set_page_config(page_title="Document Translation Office", page_icon=":material/description:", layout="wide")
    require_password()
    st.title("Document Translation Office")

    if "source_choice" not in st.session_state:
        st.session_state["source_choice"] = "Italian"
    if "source_custom" not in st.session_state:
        st.session_state["source_custom"] = ""
    if "subject_matter" not in st.session_state:
        st.session_state["subject_matter"] = "Legal / judicial file"
    if "style_instruction" not in st.session_state:
        st.session_state["style_instruction"] = default_style_requirements(
            st.session_state["subject_matter"]
        )
    if "user_context" not in st.session_state:
        st.session_state["user_context"] = default_document_context(
            resolve_language(st.session_state["source_choice"], st.session_state["source_custom"]),
            st.session_state["subject_matter"],
        )

    with st.sidebar:
        st.button(
            "Log out",
            icon=":material/logout:",
            on_click=log_out,
            use_container_width=True,
        )

        st.header("Translation")
        source_choice = st.selectbox(
            "Source language",
            LANGUAGES,
            key="source_choice",
            on_change=refresh_document_context_for_source,
        )
        source_custom = st.text_input(
            "Custom source language",
            key="source_custom",
            disabled=source_choice != "Custom",
            on_change=refresh_document_context_for_source,
        )
        target_choice = st.selectbox("Target language", LANGUAGES, index=0)
        target_custom = st.text_input("Custom target language", disabled=target_choice != "Custom")

        st.header("Professional Settings")
        subject_matter = st.selectbox(
            "Document type",
            [
                "Legal / judicial file",
                "Police / investigative file",
                "Financial police / Guardia di Finanza file",
                "Court / prosecution document",
                "Business document",
                "Medical document",
                "Academic document",
                "Immigration document",
                "Technical document",
                "General professional document",
            ],
            key="subject_matter",
            on_change=refresh_defaults_for_document_type,
        )
        style_instruction = st.text_area(
            "Style requirements",
            key="style_instruction",
            height=110,
        )
        user_context = st.text_area(
            "Document context",
            key="user_context",
            height=110,
        )
        glossary = st.text_area(
            "Required terminology",
            placeholder="Example: Residence Permit = Permesso di soggiorno",
            height=130,
        )

        st.header("API")
        model = st.selectbox("OpenAI model", ["gpt-6-sol", "gpt-6-luna"], index=0)
        api_ready = bool(get_config_value("OPENAI_API_KEY"))
        st.caption("API connected" if api_ready else "Preview mode")

    uploaded_files = st.file_uploader(
        "Word documents",
        type=["docx"],
        accept_multiple_files=True,
        key="uploaded_files",
    )

    if not uploaded_files:
        clear_translation_results()
        st.info("Upload one or more .docx files to begin.")
        return

    source_language = resolve_language(source_choice, source_custom)
    target_language = resolve_language(target_choice, target_custom)
    job_signature = translation_job_signature(
        uploaded_files,
        source_language=source_language,
        target_language=target_language,
        subject_matter=subject_matter,
        style_instruction=style_instruction,
        glossary=glossary,
        user_context=user_context,
        model=model,
        api_ready=api_ready,
    )
    if st.session_state.get("translation_job_signature") != job_signature:
        clear_translation_results()
    uploaded_file = uploaded_files[0]
    source_bytes = uploaded_file.getvalue()
    document = Document(BytesIO(source_bytes))
    units = collect_text_units(document)
    analysis = analyze_document(document)

    left, right = st.columns([1, 1])
    with left:
        st.subheader("Document")
        metric_cols = st.columns(3)
        metric_cols[0].metric("Text items", analysis.text_items)
        metric_cols[1].metric("Unique text", analysis.unique_text_items)
        metric_cols[2].metric("Batches", analysis.estimated_batches)
        if analysis.repeated_text_items:
            st.caption(f"{analysis.repeated_text_items} repeated text items will reuse an exact-match translation.")

        with st.expander("Repeated table headers"):
            if analysis.repeated_header_rows:
                for row, count in analysis.repeated_header_rows[:10]:
                    st.write(f"{count} times: {format_header_row(row)}")
            else:
                st.write("No repeated first-row table header patterns found. Changed headers will be translated separately.")

        if len(uploaded_files) > 1:
            st.caption(f"Showing analysis for {uploaded_file.name}. {len(uploaded_files)} files are queued for translation.")
        st.text_area("Extracted text preview", build_preview(units), height=440)

    with right:
        st.subheader("Translation Job")
        st.write(f"{source_language} to {target_language}")
        st.write(subject_matter)
        st.caption("Layout, tables, paragraphs, runs, and existing text formatting are preserved; only text node contents are replaced.")
        if not model:
            st.warning("Set OPENAI_MODEL in .env or enter a model before live API translation.")

        can_translate = bool(units) and bool(model or not api_ready)
        if st.button("Translate document", type="primary", disabled=not can_translate, use_container_width=True):
            clear_translation_results()
            translated_outputs: list[tuple[str, bytes]] = []
            all_pairs: list[tuple[TextUnit, str]] = []

            for file_index, current_file in enumerate(uploaded_files, start=1):
                st.write(f"Translating file {file_index} of {len(uploaded_files)}: {current_file.name}")
                try:
                    translated_docx, translated_pairs = translate_document(
                        current_file.getvalue(),
                        source_language=source_language,
                        target_language=target_language,
                        subject_matter=subject_matter,
                        style_instruction=style_instruction,
                        glossary=glossary,
                        user_context=user_context,
                        model=model,
                    )
                except ValueError as exc:
                    st.error(f"Could not translate {current_file.name}: {exc}")
                    return
                translated_name = current_file.name.replace(".docx", f"_{target_language}.docx")
                translated_outputs.append((translated_name, translated_docx))
                if file_index == 1:
                    all_pairs = translated_pairs

            if len(translated_outputs) == 1:
                st.session_state["translated_docx"] = translated_outputs[0][1]
                st.session_state["download_name"] = translated_outputs[0][0]
                st.session_state.pop("translated_zip", None)
            else:
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
                    for translated_name, translated_docx in translated_outputs:
                        archive.writestr(translated_name, translated_docx)

                st.session_state["translated_zip"] = zip_buffer.getvalue()
                st.session_state["download_name"] = f"translated_{target_language}.zip"
                st.session_state.pop("translated_docx", None)

            st.session_state["translated_pairs"] = all_pairs
            st.session_state["translation_job_signature"] = job_signature
            if "translated_zip" in st.session_state:
                start_automatic_download(
                    st.session_state["translated_zip"],
                    st.session_state["download_name"],
                    "application/zip",
                )
            else:
                start_automatic_download(
                    st.session_state["translated_docx"],
                    st.session_state["download_name"],
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

        if "translated_docx" in st.session_state or "translated_zip" in st.session_state:
            st.success("Translated document ready.")
            st.caption("The download starts automatically. If it doesn’t start, use the button below.")
            if "translated_zip" in st.session_state:
                st.download_button(
                    "Download translated ZIP",
                    data=st.session_state["translated_zip"],
                    file_name=st.session_state["download_name"],
                    mime="application/zip",
                    use_container_width=True,
                )
            else:
                st.download_button(
                    "Download translated Word document",
                    data=st.session_state["translated_docx"],
                    file_name=st.session_state["download_name"],
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )

            with st.expander("Review translated text"):
                for unit, translated_text in st.session_state["translated_pairs"][:25]:
                    st.markdown(f"**{unit.location}**")
                    st.text_area(
                        f"Translated {unit.location}",
                        translated_text,
                        height=120,
                        label_visibility="collapsed",
                    )


if __name__ == "__main__":
    main()
