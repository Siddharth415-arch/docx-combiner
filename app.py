#!/usr/bin/env python3
"""
DocX Combiner & Formatter — Streamlit Web App
Sorts .docx files by episode/chapter number, merges them in order,
and formats chapter headings as Heading 1.
"""

import re
import shutil
import tempfile
import traceback
import zipfile
from pathlib import Path

import streamlit as st
from lxml import etree

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────

CHAPTER_PATTERNS = [
    r'^§§§第\d+章.*$',
    r'^第\d+章.*$',
    r'^第[一二三四五六七八九十百千]+章.*$',
    r'^Chapter\s+\d+.*$',
    r'^CHAPTER\s+\d+.*$',
    r'^Ch\.?\s+\d+.*$',
    r'^卷\d+.*$',
]

# ─────────────────────────────────────────────
# CORE LOGIC (adapted from combine_and_format.py)
# ─────────────────────────────────────────────

def extract_start_number(filepath: Path) -> int:
    """Extract the first (smallest) number from a filename for sorting."""
    name = filepath.stem
    numbers = re.findall(r'\d+', name)
    if not numbers:
        return 0
    nums = [int(n) for n in numbers]
    if len(nums) == 1:
        return nums[0]
    return min(nums)


def sort_batch_files(files: list) -> list:
    """Sort .docx files by their chapter/episode start number."""
    return sorted(files, key=lambda f: extract_start_number(f))


def unpack_docx(docx_path: Path, unpack_dir: Path):
    """Unpack a .docx and pretty-print XML for reliable processing."""
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
    """Pack a directory back into a .docx file."""
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_docx, "w", zipfile.ZIP_DEFLATED) as z:
        for f in unpack_dir.rglob("*"):
            if f.is_file():
                z.write(f, f.relative_to(unpack_dir))


def get_body_paragraphs_raw(doc_xml: Path) -> list:
    """Extract body content from document.xml as raw XML strings (skips sectPr)."""
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
    """Extract <w:sectPr> (page/section properties) from document.xml."""
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
    """
    Merge multiple .docx files in order.
    - Styles/settings from the FIRST file
    - Body content from ALL files concatenated
    - Page layout (sectPr) from the LAST file
    """
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
    """Format chapter headings as Heading 1 style in document.xml."""
    with open(xml_path, 'r', encoding='utf-8') as f:
        content = f.read()

    chapter_count = 0

    # Pattern 1: Chapter heading between <w:br/> tags
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

    # Pattern 2: Chapter heading after a volume marker
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


def process_files(sorted_files: list, output_path: Path, progress_callback=None) -> int:
    """Merge a sorted list of .docx files, format headings, save to output_path."""
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        merged_path = tmp / "merged.docx"

        merge_docx_files(sorted_files, merged_path, progress_callback)

        unpack_dir = tmp / "unpacked"
        unpack_docx(merged_path, unpack_dir)

        doc_xml = unpack_dir / "word" / "document.xml"
        if not doc_xml.exists():
            raise ValueError("No document.xml found after merge — the .docx files may be corrupted.")

        count = format_chapter_headings(doc_xml)
        pack_docx(unpack_dir, output_path)

    return count


# ─────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────

def main():
    st.set_page_config(
        page_title="DocX Combiner",
        page_icon="📚",
        layout="centered"
    )

    # Header
    st.title("📚 DocX Combiner & Formatter")
    st.markdown(
        "Upload multiple `.docx` files and this tool will automatically **sort them "
        "by episode/chapter number**, **merge them in the correct order**, and "
        "**format chapter headings** as Heading 1 — ready to download in seconds."
    )
    st.divider()

    # ── File Upload ──
    uploaded_files = st.file_uploader(
        "Upload your .docx files",
        type=["docx"],
        accept_multiple_files=True,
        help="Select all the chapter/episode files at once. They will be sorted automatically by the numbers in their filenames."
    )

    output_name = st.text_input(
        "Output file name",
        value="Combined",
        help='Your file will be saved as  <name>_Combined_Formatted.docx'
    )

    st.divider()

    if uploaded_files:
        st.subheader(f"📂 {len(uploaded_files)} file(s) detected")

        # Preview detected sort order (client-side, before processing)
        file_preview = sorted(
            [(uf.name, extract_start_number(Path(uf.name))) for uf in uploaded_files],
            key=lambda x: x[1]
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

                # Write uploaded files to temp directory
                for uf in uploaded_files:
                    (input_dir / uf.name).write_bytes(uf.getvalue())

                docx_files = [f for f in input_dir.glob("*.docx")]
                sorted_files = sort_batch_files(docx_files)

                safe_name = re.sub(r'[^\w\-_. ]', '_', output_name.strip() or "Combined")
                output_filename = f"{safe_name}_Combined_Formatted.docx"
                output_path = tmp_path / output_filename

                def update_progress(current, total, filename):
                    pct = int((current / total) * 80)
                    progress_bar.progress(pct, text=f"Merging ({current}/{total}): {filename}")

                try:
                    count = process_files(sorted_files, output_path, update_progress)
                    progress_bar.progress(100, text="Complete!")

                    with open(output_path, "rb") as f:
                        file_bytes = f.read()

                    result_area.success(
                        f"✅ Done! Merged **{len(sorted_files)} file(s)** · "
                        f"**{count}** chapter heading(s) formatted as Heading 1."
                    )

                    st.download_button(
                        label=f"⬇️  Download  {output_filename}",
                        data=file_bytes,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                        type="primary"
                    )

                except Exception as e:
                    progress_bar.empty()
                    result_area.error(f"❌ Something went wrong: {e}")
                    with st.expander("Show error details"):
                        st.code(traceback.format_exc())

    else:
        st.info("👆 Upload your `.docx` files above to get started.")

    # Footer
    st.divider()
    st.caption(
        "Supported chapter heading formats: &nbsp;"
        "`第X章` &nbsp;·&nbsp; `Chapter X` &nbsp;·&nbsp; `CHAPTER X` &nbsp;·&nbsp; "
        "`Ch. X` &nbsp;·&nbsp; `卷X` &nbsp;·&nbsp; Chinese numeral chapters"
    )


if __name__ == "__main__":
    main()
