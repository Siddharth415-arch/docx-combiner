# -*- coding: utf-8 -*-
"""
Internal Tools — Streamlit Web App
Three tools in one:
  1. DocX Combiner & Formatter  — sorts, merges, and formats chapter headings
  2. Benchmark Converter        — converts docx files to benchmark format
  3. Document Translator        — translates .docx via Gemini AI (upload or Drive link)
"""

import io
import os
import re
import shutil
import tempfile
import threading
import time
import traceback
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import streamlit as st
from docx import Document
from docx.shared import Inches
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

try:
    import google.generativeai as genai
    GEMINI_AVAILABLE = True
except ImportError:
    GEMINI_AVAILABLE = False

try:
    import gdown
    GDOWN_AVAILABLE = True
except ImportError:
    GDOWN_AVAILABLE = False


# ══════════════════════════════════════════════════════════════════════
#  SHARED PAGE CONFIG  (must be first Streamlit call)
# ══════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="Tools",
    page_icon="🛠️",
    layout="centered",
)


# ══════════════════════════════════════════════════════════════════════
#  SIDEBAR — TOOL SELECTOR
# ══════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("### Select Tool")
    selected_tool = st.radio(
        label="Tool",
        options=["📚 DocX Combiner", "⚙️ Benchmark Converter", "🌐 Document Translator"],
        label_visibility="collapsed",
    )


# ══════════════════════════════════════════════════════════════════════
#  TOOL 1 — DOCX COMBINER & FORMATTER
# ══════════════════════════════════════════════════════════════════════

CHAPTER_PATTERNS = [
    r'^§§§第\d+章.*$',
    r'^第\d+章.*$',
    r'^第[一二三四五六七八九十百千]+章.*$',
    r'^Chapter\s+\d+.*$',
    r'^CHAPTER\s+\d+.*$',
    r'^Ch\.?\s+\d+.*$',
    r'^卷\d+.*$',
]


def extract_start_number(filepath: Path) -> int:
    name = filepath.stem
    numbers = re.findall(r'\d+', name)
    if not numbers:
        return 0
    nums = [int(n) for n in numbers]
    return min(nums) if len(nums) > 1 else nums[0]


def sort_batch_files(files: list) -> list:
    return sorted(files, key=lambda f: extract_start_number(f))


def unpack_docx(docx_path: Path, unpack_dir: Path):
    unpack_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(docx_path, 'r') as z:
        z.extractall(unpack_dir)
    for xml_file in unpack_dir.rglob("*.xml"):
        try:
            tree = etree.parse(str(xml_file))
            with open(xml_file, "wb") as f:
                f.write(etree.tostring(tree, pretty_print=True,
                        encoding="UTF-8", xml_declaration=False))
        except Exception:
            pass


def pack_docx(unpack_dir: Path, output_docx: Path):
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_docx, "w", zipfile.ZIP_DEFLATED) as z:
        for f in unpack_dir.rglob("*"):
            if f.is_file():
                z.write(f, f.relative_to(unpack_dir))


def get_body_paragraphs_raw(doc_xml: Path) -> list:
    try:
        tree = etree.parse(str(doc_xml))
    except Exception:
        return []
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    body = tree.find('.//w:body', ns)
    if body is None:
        return []
    paragraphs = []
    for child in body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'sectPr':
            continue
        paragraphs.append(etree.tostring(child, encoding='unicode'))
    return paragraphs


def get_sectPr_raw(doc_xml: Path) -> str:
    try:
        tree = etree.parse(str(doc_xml))
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        sectPr = tree.find('.//w:body/w:sectPr', ns)
        if sectPr is not None:
            return etree.tostring(sectPr, encoding='unicode')
    except Exception:
        pass
    return ''


def merge_docx_files(file_list: list, output_path: Path, progress_callback=None):
    if len(file_list) == 1:
        shutil.copy(file_list[0], output_path)
        if progress_callback:
            progress_callback(1, 1, file_list[0].name)
        return

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        base_dir = tmp / "base"
        unpack_docx(file_list[0], base_dir)
        base_xml = base_dir / "word" / "document.xml"

        all_paragraphs = []
        for i, f in enumerate(file_list):
            if i == 0:
                paras = get_body_paragraphs_raw(base_xml)
            else:
                fdir = tmp / f"doc_{i}"
                unpack_docx(f, fdir)
                fxml = fdir / "word" / "document.xml"
                paras = get_body_paragraphs_raw(fxml)
            all_paragraphs.extend(paras)
            if progress_callback:
                progress_callback(i + 1, len(file_list), f.name)

        last_dir = tmp / "last"
        unpack_docx(file_list[-1], last_dir)
        sect_pr = get_sectPr_raw(last_dir / "word" / "document.xml")

        body_content = "\n".join(all_paragraphs)
        if sect_pr:
            body_content += f"\n{sect_pr}"

        new_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
mc:Ignorable="w14">
<w:body>
{body_content}
</w:body>
</w:document>'''

        with open(base_xml, 'w', encoding='utf-8') as f:
            f.write(new_xml)

        pack_docx(base_dir, output_path)


def format_chapter_headings(xml_path: Path) -> int:
    with open(xml_path, 'r', encoding='utf-8') as f:
        content = f.read()

    chapter_count = 0

    for pattern in CHAPTER_PATTERNS:
        core = pattern.strip('^$')
        regex = r'(<w:br/>)\s*(<w:t[^>]*>)(' + core + r')(</w:t>)\s*(<w:br/>)'

        def replace_heading(match):
            nonlocal chapter_count
            chapter_count += 1
            return f'''{match.group(1)}
</w:r>
</w:p>
<w:p>
<w:pPr>
<w:pStyle w:val="Heading1"/>
</w:pPr>
<w:r>
<w:t>{match.group(3)}</w:t>
</w:r>
</w:p>
<w:p>
<w:r>
{match.group(5)}'''

        content = re.sub(regex, replace_heading, content, flags=re.IGNORECASE)

    for pattern in CHAPTER_PATTERNS:
        core = pattern.strip('^$')
        regex2 = (
            r'(<w:t[^>]*>第[一二三四五六七八九十百千\d]+卷[^<]*</w:t>)\s*'
            r'(<w:br/>)\s*(<w:t[^>]*>)(' + core + r')(</w:t>)\s*(<w:br/>)'
        )

        def replace_volume(match):
            nonlocal chapter_count
            chapter_count += 1
            return f'''{match.group(1)}
</w:r>
</w:p>
<w:p>
<w:pPr>
<w:pStyle w:val="Heading1"/>
</w:pPr>
<w:r>
<w:t>{match.group(4)}</w:t>
</w:r>
</w:p>
<w:p>
<w:r>
{match.group(6)}'''

        content = re.sub(regex2, replace_volume, content, flags=re.IGNORECASE)

    with open(xml_path, 'w', encoding='utf-8') as f:
        f.write(content)

    return chapter_count


def process_combiner_files(sorted_files: list, output_path: Path, progress_callback=None) -> int:
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        merged_path = tmp / "merged.docx"
        merge_docx_files(sorted_files, merged_path, progress_callback)

        unpack_dir = tmp / "unpacked"
        unpack_docx(merged_path, unpack_dir)
        doc_xml = unpack_dir / "word" / "document.xml"

        if not doc_xml.exists():
            raise ValueError("No document.xml found — the .docx files may be corrupted.")

        count = format_chapter_headings(doc_xml)
        pack_docx(unpack_dir, output_path)
        return count


def render_docx_combiner():
    st.title("📚 DocX Combiner & Formatter")
    st.markdown(
        "Upload multiple `.docx` files and this tool will automatically **sort them "
        "by episode/chapter number**, **merge them in the correct order**, and "
        "**format chapter headings** as Heading 1 — ready to download in seconds."
    )
    st.divider()

    uploaded_files = st.file_uploader(
        "Upload your .docx files",
        type=["docx"],
        accept_multiple_files=True,
        help="Select all the chapter/episode files at once. They will be sorted automatically.",
    )

    output_name = st.text_input(
        "Output file name",
        value="Combined",
        help="Your file will be saved as <name>_Combined_Formatted.docx",
    )
    st.divider()

    if uploaded_files:
        st.subheader(f"📂 {len(uploaded_files)} file(s) detected")
        file_preview = sorted(
            [(uf.name, extract_start_number(Path(uf.name))) for uf in uploaded_files],
            key=lambda x: x[1],
        )
        with st.expander("🔍 Preview detected merge order", expanded=True):
            for fname, num in file_preview:
                label = str(num) if num > 0 else "?"
                st.markdown(f"&nbsp;&nbsp;`[{label:>6}]`&nbsp;&nbsp; {fname}")
            st.markdown("")

        if st.button("🔗 Merge & Format", type="primary", use_container_width=True):
            progress_bar = st.progress(0, text="Preparing files...")
            result_area = st.empty()

            with tempfile.TemporaryDirectory() as tmp:
                tmp_path = Path(tmp)
                input_dir = tmp_path / "input"
                input_dir.mkdir()

                for uf in uploaded_files:
                    (input_dir / uf.name).write_bytes(uf.getvalue())

                docx_files = list(input_dir.glob("*.docx"))
                sorted_files = sort_batch_files(docx_files)

                safe_name = re.sub(r'[^\w\-_. ]', '_', output_name.strip() or "Combined")
                output_filename = f"{safe_name}_Combined_Formatted.docx"
                output_path = tmp_path / output_filename

                def update_progress(current, total, filename):
                    pct = int((current / total) * 80)
                    progress_bar.progress(pct, text=f"Merging ({current}/{total}): {filename}")

                try:
                    count = process_combiner_files(sorted_files, output_path, update_progress)
                    progress_bar.progress(100, text="Complete!")

                    with open(output_path, "rb") as f:
                        file_bytes = f.read()

                    result_area.success(
                        f"✅ Done! Merged **{len(sorted_files)} file(s)** · "
                        f"**{count}** chapter heading(s) formatted as Heading 1."
                    )
                    st.download_button(
                        label=f"⬇️ Download {output_filename}",
                        data=file_bytes,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                        type="primary",
                    )

                except Exception as e:
                    progress_bar.empty()
                    result_area.error(f"❌ Something went wrong: {e}")
                    with st.expander("Show error details"):
                        st.code(traceback.format_exc())
    else:
        st.info("👆 Upload your `.docx` files above to get started.")

    st.divider()
    st.caption(
        "Supported chapter heading formats: &nbsp;"
        "`第X章` &nbsp;·&nbsp; `Chapter X` &nbsp;·&nbsp; `CHAPTER X` &nbsp;·&nbsp; "
        "`Ch. X` &nbsp;·&nbsp; `卷X` &nbsp;·&nbsp; Chinese numeral chapters"
    )


# ══════════════════════════════════════════════════════════════════════
#  TOOL 2 — BENCHMARK CONVERTER
# ══════════════════════════════════════════════════════════════════════

def extract_show_info_from_filename(filename: str):
    name = re.sub(r'\.docx$', '', filename, flags=re.IGNORECASE)
    name = re.sub(r'\s*\(\d+\)$', '', name)
    match = re.search(r'^(.+?)_(\d+)-(\d+)$', name)
    if match:
        return match.group(1), int(match.group(2)), int(match.group(3))
    alt = re.search(r'^(.+?)_(\d+)-(\d+)', name)
    if alt:
        return alt.group(1), int(alt.group(2)), int(alt.group(3))
    return name, 1, 500


def set_document_background_pagination(doc):
    try:
        settings_element = doc.settings.element
        bg_repag = settings_element.find(qn('w:displayBackgroundShape'))
        if bg_repag is None:
            bg_repag = OxmlElement('w:displayBackgroundShape')
            bg_repag.set(qn('w:val'), '1')
            settings_element.append(bg_repag)
        view_element = settings_element.find(qn('w:view'))
        if view_element is None:
            view_element = OxmlElement('w:view')
            view_element.set(qn('w:val'), 'print')
            settings_element.append(view_element)
    except Exception:
        pass


def convert_single_file_to_benchmark(file_bytes: bytes, filename: str) -> tuple[bytes, str]:
    show_name, first_ep, last_ep = extract_show_info_from_filename(filename)
    doc = Document(io.BytesIO(file_bytes))
    new_doc = Document()
    set_document_background_pagination(new_doc)
    for section in new_doc.sections:
        section.top_margin    = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin   = Inches(1)
        section.right_margin  = Inches(1)
    styles = new_doc.styles
    try:
        styles['Heading 1']
    except KeyError:
        styles.add_style('Heading 1', WD_STYLE_TYPE.PARAGRAPH)

    episode_patterns = [
        re.compile(r'^Episode\s+(\d+)', re.IGNORECASE),
        re.compile(r'^Chapter\s+(\d+)',  re.IGNORECASE),
        re.compile(r'^Ep\.?\s+(\d+)',    re.IGNORECASE),
        re.compile(r'^Ch\.?\s+(\d+)',    re.IGNORECASE),
        re.compile(r'^(\d+)\.'),
        re.compile(r'^Episode\s+(\d+):', re.IGNORECASE),
        re.compile(r'^Chapter\s+(\d+):', re.IGNORECASE),
    ]
    title_pattern = re.compile(r':\s*(.+)$')
    current_episode  = first_ep
    paragraphs_to_add = []

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        is_heading  = False
        episode_num = None
        for pattern in episode_patterns:
            m = pattern.match(text)
            if m:
                episode_num = int(m.group(1))
                is_heading  = True
                break
        if not is_heading and para.style.name.startswith('Heading'):
            is_heading = True
            num_m = re.search(r'\d+', text)
            if num_m:
                episode_num = int(num_m.group())
        if is_heading:
            heading_text = f"Episode {episode_num}" if episode_num else f"Episode {current_episode}"
            current_episode = (episode_num or current_episode) + 1
            paragraphs_to_add.append(('heading1', heading_text))
            title_m = title_pattern.search(text)
            if title_m:
                paragraphs_to_add.append(('heading2', title_m.group(1).strip()))
        else:
            paragraphs_to_add.append(('normal', text))

    for para_type, para_text in paragraphs_to_add:
        if para_type == 'heading1':
            p = new_doc.add_paragraph(para_text)
            p.style = 'Heading 1'
        elif para_type == 'heading2':
            p = new_doc.add_paragraph(para_text)
            p.style = 'Heading 2'
        else:
            new_doc.add_paragraph(para_text)

    output_filename = f"{show_name}_{first_ep}-{last_ep}.docx"
    buf = io.BytesIO()
    new_doc.save(buf)
    return buf.getvalue(), output_filename


def render_benchmark_converter():
    st.title("⚙️ Benchmark Converter")
    st.markdown(
        "Upload one or more `.docx` files — **one per show** — and this tool will "
        "convert each file into the **benchmark format** accepted by PocketFM's "
        "internal review tool."
    )
    st.markdown(
        "**Expected filename format:** `ShowName_FirstEp-LastEp.docx`  \n"
        "Example: `The_Man_They_Forgot_1-500.docx`"
    )
    st.divider()

    uploaded_files = st.file_uploader(
        "Upload .docx files (one per show)",
        type=["docx"],
        accept_multiple_files=True,
        key="bench_uploader",
    )
    st.divider()

    if not uploaded_files:
        st.info("👆 Upload your `.docx` files above to get started.")
        return

    st.subheader(f"📂 {len(uploaded_files)} file(s) ready to convert")
    with st.expander("🔍 Files detected", expanded=True):
        for uf in uploaded_files:
            show_name, first_ep, last_ep = extract_show_info_from_filename(uf.name)
            st.markdown(
                f"&nbsp;&nbsp;📄 **{uf.name}**  →  "
                f"Show: `{show_name}` &nbsp;|&nbsp; Episodes: `{first_ep}–{last_ep}`"
            )
        st.markdown("")

    if not st.button("🚀 Convert All Files", type="primary", use_container_width=True):
        return

    progress_bar = st.progress(0, text="Starting conversion...")
    errors  = []
    results = []

    for i, uf in enumerate(uploaded_files):
        progress_bar.progress(
            int(i / len(uploaded_files) * 100),
            text=f"Converting ({i + 1}/{len(uploaded_files)}): {uf.name}",
        )
        try:
            out_bytes, out_name = convert_single_file_to_benchmark(uf.getvalue(), uf.name)
            results.append((out_bytes, out_name))
        except Exception as e:
            errors.append((uf.name, str(e), traceback.format_exc()))

    progress_bar.progress(100, text="Done!")

    if errors:
        for fname, err_msg, tb in errors:
            st.error(f"❌ Failed to convert **{fname}**: {err_msg}")
            with st.expander(f"Error details — {fname}"):
                st.code(tb)

    if not results:
        return

    st.success(f"✅ Successfully converted **{len(results)}** file(s)!")
    st.markdown("---")

    if len(results) == 1:
        out_bytes, out_name = results[0]
        st.download_button(
            label=f"⬇️ Download {out_name}",
            data=out_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            type="primary",
        )
    else:
        st.markdown("**Download individual files:**")
        for out_bytes, out_name in results:
            st.download_button(
                label=f"⬇️ {out_name}",
                data=out_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        st.markdown("**Or download everything at once:**")
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for out_bytes, out_name in results:
                zf.writestr(out_name, out_bytes)
        st.download_button(
            label=f"📦 Download all {len(results)} files as ZIP",
            data=zip_buf.getvalue(),
            file_name="benchmark_converted.zip",
            mime="application/zip",
            use_container_width=True,
            type="primary",
        )

    st.divider()
    st.caption(
        "Heading patterns detected: `Episode N` · `Chapter N` · `Ep. N` · `Ch. N` · `N.` · "
        "any paragraph already styled as Heading."
    )


# ══════════════════════════════════════════════════════════════════════
#  TOOL 3 — DOCUMENT TRANSLATOR  (Free Google Translate  OR  Gemini AI)
# ══════════════════════════════════════════════════════════════════════

GEMINI_BATCH_SIZE    = 30    # paragraphs per Gemini call
GTRANS_BATCH_SIZE    = 20    # paragraphs per free-API call
GTRANS_BATCH_WORKERS = 10    # concurrent free-API batches per file
TRANS_FILE_WORKERS   = 3     # parallel files (both engines)

GTRANSLATE_URL = "https://translate.googleapis.com/translate_a/single"

LANGUAGES = {
    "Chinese (Simplified)": "Chinese (Simplified)",
    "Chinese (Traditional)": "Chinese (Traditional)",
    "Korean": "Korean",
    "Japanese": "Japanese",
    "Arabic": "Arabic",
    "Spanish": "Spanish",
    "French": "French",
    "German": "German",
    "Hindi": "Hindi",
    "Portuguese": "Portuguese",
    "Russian": "Russian",
    "English": "English",
}

GTRANS_LANG_CODES = {
    "Chinese (Simplified)": "zh-CN",
    "Chinese (Traditional)": "zh-TW",
    "Korean": "ko",
    "Japanese": "ja",
    "Arabic": "ar",
    "Spanish": "es",
    "French": "fr",
    "German": "de",
    "Hindi": "hi",
    "Portuguese": "pt",
    "Russian": "ru",
    "English": "en",
}

GEMINI_MODELS = [
    "gemini-1.5-flash",
    "gemini-2.0-flash",
    "gemini-1.5-pro",
]


# ── Free Google Translate engine ─────────────────────────────────────────────

def _gtrans_call(text: str, src_code: str, tgt_code: str) -> str:
    """Single call to the free Google Translate endpoint."""
    import requests as _req
    r = _req.get(
        GTRANSLATE_URL,
        params={"client": "gtx", "sl": src_code, "tl": tgt_code, "dt": "t", "q": text},
        timeout=20,
    )
    r.raise_for_status()
    data = r.json()
    return "".join(part[0] for part in data[0] if part[0])


def _gtrans_batch(texts: list, src_code: str, tgt_code: str) -> list:
    """Translate a batch via free API using \\n separator; falls back individually."""
    if not texts:
        return texts
    try:
        result = _gtrans_call("\n".join(texts), src_code, tgt_code)
        parts  = result.split("\n")
        if len(parts) == len(texts):
            return [p.strip() for p in parts]
    except Exception:
        pass
    return [_gtrans_call(t, src_code, tgt_code) if t.strip() else t for t in texts]


def _translate_docx_gtrans(file_bytes: bytes, src: str, tgt: str,
                            progress_cb=None) -> bytes:
    """Translate a .docx using the free Google Translate API (parallel batches)."""
    src_code = GTRANS_LANG_CODES.get(src, "auto")
    tgt_code = GTRANS_LANG_CODES.get(tgt, "en")

    src_doc   = Document(io.BytesIO(file_bytes))
    all_texts = [p.text for p in src_doc.paragraphs]
    non_empty = [(i, t) for i, t in enumerate(all_texts) if t.strip()]
    texts     = [t for _, t in non_empty]
    idx_map   = [i for i, _ in non_empty]

    translated_map = {}
    completed      = [0]
    lock           = threading.Lock()

    batches = [(s, texts[s:s + GTRANS_BATCH_SIZE]) for s in range(0, len(texts), GTRANS_BATCH_SIZE)]

    def run_batch(start, batch_texts):
        res = _gtrans_batch(batch_texts, src_code, tgt_code)
        with lock:
            for j, trans in enumerate(res):
                translated_map[idx_map[start + j]] = trans
            completed[0] += 1
            if progress_cb:
                progress_cb(completed[0], len(batches))

    with ThreadPoolExecutor(max_workers=GTRANS_BATCH_WORKERS) as ex:
        futs = [ex.submit(run_batch, s, b) for s, b in batches]
        for f in as_completed(futs):
            f.result()

    dst_doc = Document()
    for i, para in enumerate(src_doc.paragraphs):
        text     = translated_map.get(i, para.text)
        new_para = dst_doc.add_paragraph(text)
        try:
            if para.style and para.style.name:
                new_para.style = dst_doc.styles[para.style.name]
        except Exception:
            pass

    buf = io.BytesIO()
    dst_doc.save(buf)
    return buf.getvalue()


# ── Gemini engine ─────────────────────────────────────────────────────────────

def _gemini_translate_batch(texts: list, src: str, tgt: str, model) -> list:
    """Send a numbered list to Gemini, get back numbered translations."""
    if not texts:
        return texts

    numbered = "\n".join(f"{i+1}. {t}" for i, t in enumerate(texts))
    prompt = (
        f"Translate each numbered item from {src} to {tgt}.\n"
        f"Return ONLY the numbered translations in the exact same format (1. 2. 3. ...).\n"
        f"Do not add explanations, notes, or extra lines.\n\n"
        f"{numbered}"
    )

    for attempt in range(3):
        try:
            response  = model.generate_content(prompt)
            raw       = response.text.strip()
            parsed    = []
            for line in raw.splitlines():
                m = re.match(r'^\d+\.\s*(.*)', line.strip())
                if m:
                    parsed.append(m.group(1).strip())
            if len(parsed) == len(texts):
                return parsed
            # fallback: plain non-empty lines
            lines = [l.strip() for l in raw.splitlines() if l.strip()]
            if len(lines) == len(texts):
                return lines
        except Exception:
            if attempt < 2:
                time.sleep(3)

    return texts   # return originals if all retries fail


def _translate_docx_gemini(file_bytes: bytes, src: str, tgt: str,
                            model, progress_cb=None) -> bytes:
    """Translate all paragraphs in a .docx using Gemini and return new bytes."""
    src_doc   = Document(io.BytesIO(file_bytes))
    all_texts = [p.text for p in src_doc.paragraphs]

    non_empty     = [(i, t) for i, t in enumerate(all_texts) if t.strip()]
    texts         = [t for _, t in non_empty]
    idx_map       = [i for i, _ in non_empty]
    translated_map = {}

    batches     = [texts[s:s + GEMINI_BATCH_SIZE] for s in range(0, len(texts), GEMINI_BATCH_SIZE)]
    total       = len(batches)

    for b_idx, batch in enumerate(batches):
        translated = _gemini_translate_batch(batch, src, tgt, model)
        start      = b_idx * GEMINI_BATCH_SIZE
        for j, trans_text in enumerate(translated):
            translated_map[idx_map[start + j]] = trans_text
        if progress_cb:
            progress_cb(b_idx + 1, total)
        time.sleep(0.25)

    dst_doc = Document()
    for i, para in enumerate(src_doc.paragraphs):
        text     = translated_map.get(i, para.text)
        new_para = dst_doc.add_paragraph(text)
        try:
            if para.style and para.style.name:
                new_para.style = dst_doc.styles[para.style.name]
        except Exception:
            pass

    buf = io.BytesIO()
    dst_doc.save(buf)
    return buf.getvalue()


def _fetch_files_from_drive(folder_url: str) -> dict:
    """
    Download all .docx files from a public Google Drive folder.
    Returns {filename: bytes}.
    Folder must be shared as 'Anyone with the link can view'.
    """
    if not GDOWN_AVAILABLE:
        raise RuntimeError("gdown is not installed. Cannot fetch from Drive.")

    tmp_dir = tempfile.mkdtemp()
    try:
        gdown.download_folder(
            url=folder_url,
            output=tmp_dir,
            quiet=True,
            use_cookies=False,
        )
        result = {}
        for root, _, filenames in os.walk(tmp_dir):
            for fn in filenames:
                if fn.lower().endswith(".docx"):
                    with open(os.path.join(root, fn), "rb") as f:
                        result[fn] = f.read()
        return result
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def render_document_translator():
    st.title("🌐 Document Translator")
    st.markdown(
        "Translate `.docx` files — upload from your computer or paste a Google Drive folder link. "
        "All files process in parallel and download as a single ZIP."
    )
    st.divider()

    # ── Session state init ────────────────────────────────────────────────────
    for key, default in [("tr_results", {}), ("tr_errors", {}), ("tr_api_key", "")]:
        if key not in st.session_state:
            st.session_state[key] = default

    # ── ① Engine selector ────────────────────────────────────────────────────
    engine = st.radio(
        "Translation engine",
        options=["🆓  Free — Google Translate (no key needed)",
                 "✨  Gemini AI — higher quality (API key required)"],
        key="tr_engine",
        horizontal=True,
    )
    use_gemini = engine.startswith("✨")

    # ── ② Gemini API key box (only when Gemini is selected) ──────────────────
    model_name = GEMINI_MODELS[0]

    if use_gemini:
        if not GEMINI_AVAILABLE:
            st.error("⛔ `google-generativeai` package missing. Add it to requirements.txt and redeploy.")
            return

        st.markdown(
            """
            <div style="
                background: linear-gradient(135deg, #667eea22, #764ba222);
                border: 2px solid #667eea;
                border-radius: 12px;
                padding: 18px 20px 10px 20px;
                margin: 12px 0 4px 0;
            ">
            <p style="margin:0 0 6px 0; font-weight:700; font-size:1rem;">
                🔑 &nbsp;Gemini API Key <span style="color:#667eea">(required)</span>
            </p>
            """,
            unsafe_allow_html=True,
        )
        api_key_input = st.text_input(
            label="gemini_key_label",
            label_visibility="collapsed",
            type="password",
            placeholder="AIzaSy…  — paste your key here",
            value=st.session_state.tr_api_key,
            key="tr_api_key_input",
        )
        st.markdown(
            """
            <p style="margin:4px 0 0 0; font-size:0.8rem; color:#888;">
            Get a free key at
            <a href="https://aistudio.google.com/apikey" target="_blank">aistudio.google.com/apikey</a>
            &nbsp;·&nbsp; gemini-1.5-flash is free up to 1,500 requests/day
            </p></div>
            """,
            unsafe_allow_html=True,
        )
        if api_key_input:
            st.session_state.tr_api_key = api_key_input

        st.markdown("<div style='margin-top:10px'></div>", unsafe_allow_html=True)
        model_name = st.selectbox("Gemini model", GEMINI_MODELS, index=0, key="tr_model")
    else:
        st.caption("Uses Google Translate's free public endpoint — no account or key needed.")

    st.divider()

    # ── ③ Language selectors ──────────────────────────────────────────────────
    col1, col2, col3 = st.columns([2, 0.6, 2])
    with col1:
        src_name = st.selectbox("Source language", list(LANGUAGES.keys()),
                                index=0, key="tr_src")
    with col2:
        st.markdown("<br><div style='text-align:center;font-size:22px'>→</div>",
                    unsafe_allow_html=True)
    with col3:
        tgt_name = st.selectbox("Target language", list(LANGUAGES.keys()),
                                index=11, key="tr_tgt")

    src_lang = LANGUAGES[src_name]
    tgt_lang = LANGUAGES[tgt_name]

    st.divider()

    # ── ④ Input — two tabs ────────────────────────────────────────────────────
    tab_upload, tab_drive = st.tabs(["📁  Upload Files from Computer",
                                     "🔗  Google Drive Folder Link"])

    with tab_upload:
        uploaded_files = st.file_uploader(
            "Select one or more .docx files",
            type=["docx"],
            accept_multiple_files=True,
            key="tr_uploader",
            help="Hold Cmd/Ctrl to select multiple files at once",
        )

    with tab_drive:
        st.markdown("")
        drive_link = st.text_input(
            "Paste Google Drive folder link",
            placeholder="https://drive.google.com/drive/folders/…",
            key="tr_drive_link",
        )
        st.caption(
            "⚠️ Folder must be shared as **'Anyone with the link can view'** — "
            "private folders require the Colab notebook instead."
        )

    st.divider()

    # ── ⑤ Readiness check & translate button ─────────────────────────────────
    has_input = bool(uploaded_files) or bool(drive_link)
    has_key   = bool(st.session_state.tr_api_key) if use_gemini else True
    ready     = has_input and has_key

    if use_gemini and not has_key:
        st.warning("⬆️ Enter your Gemini API key above to enable translation.")
    elif not has_input:
        st.info("⬆️ Upload files or paste a Drive folder link to get started.")

    if st.button("🚀  Translate All & Download ZIP", type="primary",
                 use_container_width=True, disabled=not ready, key="tr_go"):

        st.session_state.tr_results = {}
        st.session_state.tr_errors  = {}

        # Set up Gemini model if needed
        gemini_model = None
        if use_gemini:
            genai.configure(api_key=st.session_state.tr_api_key)
            gemini_model = genai.GenerativeModel(model_name)

        # Collect file bytes from upload and/or Drive
        file_bytes = {}
        if uploaded_files:
            for f in uploaded_files:
                file_bytes[f.name] = f.read()

        if drive_link:
            with st.spinner("📥 Fetching files from Google Drive…"):
                try:
                    drive_files = _fetch_files_from_drive(drive_link)
                    if not drive_files:
                        st.error("No .docx files found in that Drive folder.")
                    else:
                        file_bytes.update(drive_files)
                        st.success(f"✅ Downloaded {len(drive_files)} file(s) from Drive.")
                except Exception as e:
                    st.error(f"❌ Could not fetch from Drive: {e}")

        if not file_bytes:
            st.stop()

        engine_label = f"Gemini · {model_name}" if use_gemini else "Google Translate (free)"
        st.info(f"Translating **{len(file_bytes)} file(s)** — "
                f"{src_lang} → {tgt_lang} · {engine_label}")

        # Shared state — workers write only plain dicts, never st.*
        progress_state = {
            name: {"done": 0, "total": 1, "status": "queued"}
            for name in file_bytes
        }
        results    = {}
        errors     = {}
        outer_lock = threading.Lock()

        # Build per-file UI rows in main thread
        ui = {}
        for name in file_bytes:
            st.markdown(f"**{name}**")
            ui[name] = {"bar": st.progress(0), "status": st.empty()}
            ui[name]["status"].text("⌛ Queued…")
        st.divider()

        # Worker — pure computation, zero st.* calls inside
        def file_worker(name, content):
            progress_state[name]["status"] = "running"

            def on_batch(done, total):
                progress_state[name].update(done=done, total=total)

            try:
                if use_gemini:
                    translated = _translate_docx_gemini(
                        content, src_lang, tgt_lang, gemini_model, progress_cb=on_batch
                    )
                else:
                    translated = _translate_docx_gtrans(
                        content, src_lang, tgt_lang, progress_cb=on_batch
                    )
                with outer_lock:
                    results[name] = translated
                progress_state[name]["status"] = "done"
            except Exception as e:
                with outer_lock:
                    errors[name] = str(e)
                progress_state[name]["status"] = "error"

        # Run workers + poll UI from main thread
        with ThreadPoolExecutor(max_workers=TRANS_FILE_WORKERS) as executor:
            futures = {
                executor.submit(file_worker, n, c): n
                for n, c in file_bytes.items()
            }
            while not all(f.done() for f in futures):
                for name, state in progress_state.items():
                    pct = state["done"] / max(state["total"], 1)
                    s   = state["status"]
                    if s == "queued":
                        ui[name]["status"].text("⌛ Queued…")
                    elif s == "running":
                        ui[name]["bar"].progress(pct)
                        ui[name]["status"].text(
                            f"⏳ batch {state['done']}/{state['total']} ({int(pct*100)}%)"
                        )
                    elif s == "done":
                        ui[name]["bar"].progress(1.0)
                        ui[name]["status"].text("✅ Done!")
                    elif s == "error":
                        ui[name]["status"].text(f"❌ {errors.get(name, 'unknown error')}")
                time.sleep(0.3)

        # Final pass to catch any last-second state updates
        for name, state in progress_state.items():
            if state["status"] == "done":
                ui[name]["bar"].progress(1.0)
                ui[name]["status"].text("✅ Done!")
            elif state["status"] == "error":
                ui[name]["status"].text(f"❌ {errors.get(name, 'unknown error')}")

        st.session_state.tr_results = results
        st.session_state.tr_errors  = errors

    # ── ⑥ Download — persists across re-runs via session_state ───────────────
    if st.session_state.get("tr_results"):
        results = st.session_state.tr_results
        errors  = st.session_state.tr_errors

        st.subheader("⬇️ Download Translated Files")

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for name, data in results.items():
                zf.writestr(name.replace(".docx", "_EN.docx"), data)
        zip_buf.seek(0)

        st.download_button(
            label=f"📦  Download ALL {len(results)} file(s) as ZIP",
            data=zip_buf.getvalue(),
            file_name="translated_docs.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True,
            key="tr_zip",
        )

        if len(results) > 1:
            st.caption("Or download individual files:")
            cols = st.columns(min(len(results), 3))
            for i, (name, data) in enumerate(results.items()):
                out_name = name.replace(".docx", "_EN.docx")
                with cols[i % len(cols)]:
                    st.download_button(
                        label=f"⬇️ {out_name}",
                        data=data,
                        file_name=out_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"tr_dl_{name}",
                    )

        if errors:
            st.warning(f"⚠️ {len(errors)} file(s) failed: {', '.join(errors.keys())}")

    st.divider()
    st.caption(
        "🆓 Free mode: Google Translate public endpoint, 20 paragraphs/batch, 10 parallel batches · "
        "✨ Gemini mode: numbered-list batches of 30 paragraphs, higher accuracy"
    )


# ══════════════════════════════════════════════════════════════════════
#  ROUTER
# ══════════════════════════════════════════════════════════════════════

if selected_tool == "📚 DocX Combiner":
    render_docx_combiner()
elif selected_tool == "⚙️ Benchmark Converter":
    render_benchmark_converter()
else:
    render_document_translator()
