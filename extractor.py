# extractor.py
import pymupdf
import csv
import os
import re
from io import StringIO
import textwrap

EXCLUSION_STRING = 'N/A (No Comment or Highlighted Text Found)'


# -------------------- TEXT WRAPPER --------------------
def wrap_text(text, width=70):
    """
    Wrap comment text without breaking words.
    Makes CSV cleaner with multi-line cells.
    """
    wrapped = textwrap.fill(text, width=width)
    return wrapped


# -------------------- ANNOTATION EXTRACTION --------------------
def get_annotation_content(page, annot, info):
    comment = info.get('content') or info.get('text') or info.get('subject')
    if comment and comment.strip():
        return comment

    annot_type = annot.type[1]
    markup = ['Highlight', 'Underline', 'StrikeOut']

    if annot_type in markup:
        try:
            highlighted = page.get_text(clip=annot.rect)
            if highlighted:
                return f"[Highlighted Text]: {highlighted.strip()}"
        except:
            pass

    return EXCLUSION_STRING


# -------------------- DOC INFO --------------------
def extract_doc_info(doc, file_path):

    """

    Attempts to extract the document number and revision using metadata and regex on the first page.

    The revision extraction is targeted towards finding the last entry in a revision history table.

    """

    doc_num = 'N/A'

    revision = 'N/A'

    latest_revision_value = 'N/A' # Initializing to satisfy subsequent return

    

    # --- Strategy 1: Check Metadata (Title) ---

    metadata = doc.metadata

    if metadata and metadata.get('title'):

        # Example: look for a pattern that matches the file name's structured part

        match = re.search(r'(\d+[_/-]\d+[_/-]\d+[_/-]\d+)', metadata.get('title', ''))

        if match:

            doc_num = match.group(1)



    # --- Strategy 2: Check File Path (Simple Extraction - Preferred if structured) ---

    # Extract the pattern 5936_22-4600015019-00224 from the filepath

    path_match = re.search(r'(\d+_\d+-\d+-\d+)', os.path.basename(file_path))

    if path_match:

        doc_num = path_match.group(1)

    

    # --- Strategy 3: Find last revision in the Revision History table ---

    try:

        if doc.page_count > 0:

            tables = doc[0].find_tables()



        for t_idx, table in enumerate(tables.tables):

            print(f"\n--- TABLE {t_idx+1} ---\n")

            

            # Extract table rows

            rows = table.extract()

            

            # Print rows (optional)

            for row in rows:

                print(row)

            

            # ----- FIND POSITION OF "REV." -----

            rev_pos = None

            

            for r_idx, row in enumerate(rows):

                for c_idx, cell in enumerate(row):

                    if cell and "REV." in cell:

                        rev_pos = (r_idx, c_idx)

                        break

                if rev_pos:

                    break

            

            if not rev_pos:

                print("\n[!] 'REV.' not found in this table.\n")

                continue

            

            print(f"\n[+] 'REV.' found at -> Row {rev_pos[0]}, Column {rev_pos[1]}\n")

            

            # ----- EXTRACT REVISION ROWS ABOVE "REV." -----

            revision_rows = rows[:rev_pos[0]]

            

            print("[+] Revision Rows Found:")

            for rr in revision_rows:

                print(rr)

            

            # ----- LATEST REVISION (Topmost Entry) -----

            if revision_rows:
                for row in reversed(revision_rows):
                    raw_val = str(row[0]).strip()   # make everything string

                    if raw_val != '':
                        # If it's a single digit like "0"–"9", pad with leading zero
                        if raw_val.isdigit() and len(raw_val) == 1:
                            latest_revision_value = "0" + raw_val
                        else:
                            latest_revision_value = raw_val  # keep as is
                    else:
                        break
            else:
                print("\n[!] No revision rows above 'REV.' header.\n")


            # The original logic included commented out regex code here. We keep it commented as requested.

            

    except Exception as e:

        print(f"Warning: Could not extract text from first page for revision: {e}")



    return doc_num, latest_revision_value # Note: latest_revision_value is used here, matching the logic in the provided snippet.


def pad_width(text, min_width=100):
    """
    Ensures the CSV cell has at least `min_width` characters,
    so Excel expands the column automatically.
    """
    length = len(text.split("\n")[0])  # check first line only
    if length < min_width:
        return text + " " * (min_width - length)
    return text


# -------------------- MAIN EXTRACTION --------------------
def extract_pdf_comments(files: list):
    """
    INPUT: List of (company_code, file_bytes, filename)
    OUTPUT: CSV string
    """
    output = StringIO()
    writer = csv.writer(output, delimiter=',')

    writer.writerow(["", "", "", "", "", "Supplier's Answers"])
    writer.writerow([
       "Document Number", "Revision", "Page", "LNT Comment", "Reply By", "Supplier Response"
    ])

    sr = 1

    for company_code, pdf_bytes, filename in files:
        try:
            doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
            doc_number, rev = extract_doc_info(doc, filename)

            for p in range(doc.page_count):
                page = doc[p]
                annots = page.annots()

                for annot in annots:
                    if annot.type[1] not in ["Text", "FreeText"]:
                        continue

                    info = annot.info
                    content = get_annotation_content(page, annot, info)

                    if content == EXCLUSION_STRING:
                        continue

                    author = info.get('title', 'Unknown Author')

                    # Replace commas for safer CSV but keep newlines
                    cleaned = content.replace(",", " -- ").strip()

                    # 🔥 APPLY WRAPPING HERE
                    wrapped = wrap_text(cleaned, width=70)
                    wrapped_comment = pad_width(wrapped, min_width=120)                    

                    if p <9: 
                        page_no = f'\t0{p + 1} of {doc.page_count}'
                    else:
                        page_no = f'\t{p + 1} of {doc.page_count}'

                    writer.writerow([
                        doc_number,
                        f'\t{rev}',
                        page_no,
                        wrapped_comment,
                        "",
                        ""
                    ])

                    sr += 1

        except Exception as e:
            print(f"Error processing {filename}: {e}")

    return output.getvalue()
