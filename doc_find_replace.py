import os
import re

import pandas as pd
from docx import Document


def docx_replace_regex(doc_obj: Document, regex_to_replace: re.compile, replacement: str) -> None:
    """
    Replace the regex in a docx.Document object.

    Args:
        doc_obj: Template `docx.Document` object.
        regex_to_replace: Regex compile to replace in the document.
        replacement : Replacement string to substitute for the regex.
    """
    for p in doc_obj.paragraphs:
        if regex_to_replace.search(p.text):
            inline = p.runs
            # Loop added to work with runs (strings with same style)
            for i in range(len(inline)):
                if regex_to_replace.search(inline[i].text):
                    text = regex_to_replace.sub(replacement, inline[i].text)
                    inline[i].text = text

    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                docx_replace_regex(cell, regex_to_replace, replacement)


def filename_ext(filename: str) -> str:
    return os.path.splitext(filename)[1]


def batch_replace(
    *,
    template_docx: str,
    replacements_csv: str,
    output_dir: str,
    output_base_fn: str,
    output_filetype: str,
):
    """
    Generate multiple output files by executing multiple sets of find-replace operations.

    Args:
        template_docx: Path to the template document.
        replacements_csv: Path to the replacements spec document. Each row
            specifies a document version to generate. Each column header
            specifies the text to replace in the template. The column values
            specify the text to substitute during the replacement.
        output_dir: Path to the output directory.
        output_base_fn: Base filename for the generated output files.
        output_filetype: Output filetype (.docx, .pdf)
    """

    replacements_df = pd.read_csv(replacements_csv)
    template_fn = template_docx

    # TODO add extension checking for template_docx and replacements_csv

    # Loop through each replacement row
    for doc_number, doc_replacements in replacements_df.iterrows():
        print(f"Document {doc_number} - creating replacements for:")
        doc = Document(template_fn)
        output_fn_additions = []

        # Loop through each replacement in the row for the document
        for str_to_replace, replacement_str in doc_replacements.items():
            print(f"\t{str_to_replace} -> {replacement_str}")

            # Add column values that are not excluded to list to be added to
            # output filename
            column_substrings_excl_from_fn = [
                "date",
                "industry",
            ]
            if not any(substr in str_to_replace.lower() for substr in column_substrings_excl_from_fn):
                output_fn_additions.append(replacement_str)

            # Execute document replacement
            regex = re.compile(re.escape(str_to_replace))
            docx_replace_regex(doc, regex, replacement_str)

        # Save file
        output_fn_addition_str = " - ".join(output_fn_additions)
        if output_fn_addition_str:
            output_fn_addition_str = f" - {output_fn_addition_str}"

        # TODO handle .pdf output filetype

        output_fn = os.path.join(output_dir, f"{output_base_fn}{output_fn_addition_str}.docx")
        doc.save(output_fn)
        print(f"Document {doc_number} - file saved to '{output_fn}'.")
        print()


if __name__ == "__main__":
    # TODO add req args
    batch_replace(replacements_csv="replacements.csv", template_docx="cover_letter_test.docx")
