import os
from os.path import basename
from time import sleep
from zipfile import ZipFile

import gunicorn
from flask import Flask, redirect, render_template, request, url_for, send_from_directory, abort, after_this_request
from werkzeug.utils import secure_filename

import doc_find_replace

app = Flask(__name__)


def clear_folder(dir: str):
    for file in os.listdir(dir):
        try:
            os.remove(os.path.join(dir, file))
        except PermissionError as error:
            print(error)


@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        upload_dir = "uploads"
        output_dir = "created"

        # Delete all working files
        clear_folder(upload_dir)
        clear_folder(output_dir)

        template_file = request.files.get("template_file")
        replacements_file = request.files.get("replacements_file")

        if template_file and replacements_file:
            template_fn = secure_filename(template_file.filename)
            template_file.save(os.path.join(upload_dir, template_fn))

            # if output_base_fn == "":  # if no base filename is provided
            output_base_fn = secure_filename(request.form.get("output_base_fn")) or template_fn
            output_filetype = request.form.get("output_filetype") or ".docx"

            replacements_fn = secure_filename(replacements_file.filename)
            replacements_file.save(os.path.join(upload_dir, replacements_fn))

            eprint((template_fn, replacements_fn, output_base_fn, output_filetype))

            output_file_paths = doc_find_replace.batch_replace(
                replacements_csv=replacements_fn,
                template_docx=template_fn,
                output_dir=output_dir,
                output_base_fn=output_base_fn,
                output_filetype=output_filetype,
            )

            output_zip_path = os.path.join(output_dir, "generated_documents.zip")
            with ZipFile(output_zip_path, "w") as zip_obj:
                for generated_doc_path in output_file_paths:
                    zip_obj.write(generated_doc_path, basename(generated_doc_path))

            # Delete all working files after request
            @after_this_request
            def clear__files(response):
                clear_folder(upload_dir)
                clear_folder(output_dir)
                return response

            """Download a file."""
            try:
                return send_from_directory(output_dir, "generated_documents.zip", as_attachment=True)
            except FileNotFoundError:
                abort(404)

            ## SUCCESS

            return "", 201

    return render_template("main.html")


if __name__ == "__main__":
    # Quick test configuration. Please use proper Flask configuration options
    # in production settings, and use a separate file or environment variables
    # to manage the secret key!
    app.secret_key = os.urandom(12).hex()
    app.config["SESSION_TYPE"] = "filesystem"

    import sys
    from pprint import pprint as pp

    def eprint(*args, **kwargs):
        """
        Printing to stderr so visible when using flask dev server
        """
        pp(*args, sys.stderr, **kwargs)

    app.debug = True
    app.run()
