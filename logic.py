import os
import glob


VALID_FORMATS = {'bold', 'italic', 'underline'}


def process_files(file_folder, mapping_folder, remove_header=False):
    try:
        if not os.path.exists(file_folder):
            return "error", f"Files folder does not exist: {file_folder}"
        if not os.path.exists(mapping_folder):
            return "error", f"Mapping ref folder does not exist: {mapping_folder}"

        docx_files = glob.glob(os.path.join(file_folder, "*.docx"))
        doc_files = glob.glob(os.path.join(file_folder, "*.doc"))
        word_files = docx_files + doc_files

        if not word_files:
            return "no_file", "No file is found."

        missing_mapping_files = []
        processed_count = 0

        for word_file in word_files:
            basename = os.path.splitext(os.path.basename(word_file))[0]
            mapping_file = os.path.join(mapping_folder, f"ref_mapping_{basename}.txt")

            if not os.path.exists(mapping_file):
                missing_mapping_files.append(basename)
                continue

            mappings = read_mapping_file(mapping_file)
            if mappings:
                process_word_file(word_file, mappings, remove_header)
                processed_count += 1

        if missing_mapping_files:
            msg_lines = [f"No ref_mapping_{name}.txt is found." for name in missing_mapping_files]
            return "missing_mapping", "\n".join(msg_lines)

        return "success", "Completed."

    except Exception as e:
        return "error", str(e)


def read_mapping_file(mapping_file):
    mappings = []
    with open(mapping_file, 'r', encoding='utf-8') as f:
        for line_num, line in enumerate(f, 1):
            line = line.strip()
            if not line or line.startswith('#'):
                continue
            if '<<<' not in line:
                continue

            parts = line.split('<<<', 1)
            if len(parts) != 2:
                raise ValueError(f"Invalid mapping line {line_num}: {line}")

            sTo = parts[0].strip()
            sFrom_part = parts[1].strip()

            if not sTo or not sFrom_part:
                raise ValueError(f"Invalid mapping line {line_num}: {line}")

            formats_parts = sFrom_part.split('|')
            sFrom = formats_parts[0].strip()
            formats = frozenset(f.strip().lower() for f in formats_parts[1:])

            for fmt in formats:
                if fmt and fmt not in VALID_FORMATS:
                    raise ValueError(f"Unknown format: {fmt}")

            mappings.append((sTo, sFrom, formats))

    return mappings


def process_word_file(word_file, mappings, remove_header=False):
    ext = os.path.splitext(word_file)[1].lower()
    basename = os.path.splitext(os.path.basename(word_file))[0]
    output_file = os.path.join(os.path.dirname(word_file), f"{basename}-masked{ext}")

    if ext == '.docx':
        process_docx(word_file, output_file, mappings, remove_header)
    elif ext == '.doc':
        process_doc(word_file, output_file, mappings, remove_header)


def process_docx(input_path, output_path, mappings, remove_header=False):
    from docx import Document

    doc = Document(input_path)

    if remove_header:
        _remove_headers_docx(doc)

    for paragraph in doc.paragraphs:
        _replace_in_paragraph(paragraph, mappings)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    _replace_in_paragraph(paragraph, mappings)

    doc.save(output_path)


def _remove_headers_docx(doc):
    for section in doc.sections:
        for header in [section.header, section.first_page_header, section.even_page_header]:
            for paragraph in header.paragraphs:
                paragraph.text = ""


def _replace_in_paragraph(paragraph, mappings):
    for run in paragraph.runs:
        for sTo, sFrom, formats in mappings:
            if sFrom in run.text:
                if _format_matches(run, formats):
                    run.text = run.text.replace(sFrom, sTo)


def _format_matches(run, formats):
    if not formats:
        return True
    if 'bold' in formats:
        font = run.font
        if font.bold is None or not font.bold:
            return False
    if 'italic' in formats:
        font = run.font
        if font.italic is None or not font.italic:
            return False
    if 'underline' in formats:
        font = run.font
        if font.underline is None or not font.underline:
            return False
    return True


def process_doc(input_path, output_path, mappings, remove_header=False):
    import win32com.client
    import pythoncom

    pythoncom.CoInitialize()

    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

        doc = word.Documents.Open(os.path.abspath(input_path))

        if remove_header:
            _remove_headers_doc(doc)

        for sTo, sFrom, formats in mappings:
            find_and_replace(word, sFrom, sTo, formats)

        doc.SaveAs(os.path.abspath(output_path), FileFormat=16)
        doc.Close()
        word.Quit()

    finally:
        pythoncom.CoUninitialize()


def _remove_headers_doc(doc):
    from win32com.client import constants
    try:
        wdHeaderFooter = constants.wdHeaderFooter
    except:
        import win32com.client.constants
        wdHeaderFooter = win32com.client.constants.wdHeaderFooter

    for section in doc.Sections:
        for header_type in [wdHeaderFooter.wdHeaderFooterAllPages,
                           wdHeaderFooter.wdHeaderFooterFirst,
                           wdHeaderFooter.wdHeaderFooterPrimary]:
            try:
                header = section.Headers.Item(header_type)
                header.Range.Text = ""
            except:
                pass


def find_and_replace(word_app, sFrom, sTo, formats=None):
    word_app.Selection.Find.ClearFormatting()
    word_app.Selection.Find.Text = sFrom
    word_app.Selection.Find.Replacement.Text = sTo
    word_app.Selection.Find.Forward = True
    word_app.Selection.Find.Wrap = 1
    word_app.Selection.Find.Format = False
    word_app.Selection.Find.MatchCase = False
    word_app.Selection.Find.MatchWholeWord = False

    if formats and 'bold' in formats:
        word_app.Selection.Find.Font.Bold = True
    if formats and 'italic' in formats:
        word_app.Selection.Find.Font.Italic = True
    if formats and 'underline' in formats:
        word_app.Selection.Find.Font.Underline = True

    word_app.Selection.Find.Execute(Replace=2)