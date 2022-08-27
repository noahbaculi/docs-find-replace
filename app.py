import os
from os.path import basename
from time import sleep
from zipfile import ZipFile

import gunicorn
from flask import Flask, abort, after_this_request, redirect, render_template, request, send_from_directory, url_for
from flask_cors import CORS
from werkzeug.utils import secure_filename

import doc_find_replace

app = Flask(__name__)
CORS(app)


def eprint(*args, **kwargs):
    if __name__ == "__main__":
        """
        Printing to stderr so visible when using flask dev server
        """
        import sys
        from pprint import pprint as pp

        pp(*args, sys.stderr, **kwargs)
    else:
        print(*args, **kwargs)


def clear_folder(dir: str):
    if not os.path.exists(dir):
        os.makedirs(dir)

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

        eprint(f"template_file : {template_file}")
        eprint(f"replacements_file : {replacements_file}")

        if template_file and replacements_file:
            template_fn = secure_filename(template_file.filename)
            template_file.save(os.path.join(upload_dir, template_fn))

            output_base_fn = secure_filename(request.form.get("output_base_fn", "")) or template_fn

            replacements_fn = secure_filename(replacements_file.filename)
            replacements_file.save(os.path.join(upload_dir, replacements_fn))

            eprint((template_fn, replacements_fn, output_base_fn))

            output_file_paths = doc_find_replace.batch_replace(
                replacements_csv=replacements_fn,
                template_docx=template_fn,
                output_dir=output_dir,
                output_base_fn=output_base_fn,
                output_filetype=".docx",
            )

            print(f"output_file_paths : {output_file_paths}")
            print(os.listdir(output_dir))

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

            # Download the file
            try:
                return send_from_directory(output_dir, "generated_documents.zip", as_attachment=True)
            except FileNotFoundError:
                abort(404)

    return render_template("main.html")


if __name__ == "__main__":
    # Quick test configuration. Please use proper Flask configuration options
    # in production settings, and use a separate file or environment variables
    # to manage the secret key!
    app.secret_key = os.urandom(12).hex()
    app.config["SESSION_TYPE"] = "filesystem"

    app.debug = True
    app.run()
