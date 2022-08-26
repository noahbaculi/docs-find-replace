import os
import re

import pandas as pd
from docx import Document


def docx_replace_regex(doc_obj, regex, replace):
    for p in doc_obj.paragraphs:
        if regex.search(p.text):
            inline = p.runs
            # Loop added to work with runs (strings with same style)
            for i in range(len(inline)):
                if regex.search(inline[i].text):
                    text = regex.sub(replace, inline[i].text)
                    inline[i].text = text

    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                docx_replace_regex(cell, regex, replace)


def batch_replace(replacements_csv: str, replacement_docx: str):
    replacements_df = pd.read_csv(replacements_csv)
    template_fn = replacement_docx

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

        template_fn_no_ext = os.path.splitext(template_fn)[0]
        output_fn = f"{template_fn_no_ext}{output_fn_addition_str}.docx"
        doc.save(output_fn)
        print(f"Document {doc_number} - file saved to '{output_fn}'.")
        print()


if __name__ == "__main__":
    batch_replace(replacements_csv="replacements.csv", replacement_docx="cover_letter_test.docx")
