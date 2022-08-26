import os
import re

from flask import Flask, flash, redirect, render_template, request, url_for
from werkzeug.utils import secure_filename

import doc_find_replace

UPLOAD_FOLDER = "uploads"
ALLOWED_EXTENSIONS = {"txt", "pdf", "png", "jpg", "jpeg", "gif"}

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        # check if the post request has the file part

        eprint(request.files)

        if "file" not in request.files:
            flash("No file part")
            return redirect(request.url)
        file = request.files["file"]
        # if user does not select file, browser also
        # submit an empty part without filename
        if file.filename == "":
            flash("No selected file")
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)

            file.save(os.path.join(app.config["UPLOAD_FOLDER"], filename))
            # doc_find_replace.batch_replace(
            #     replacements_csv="replacements.csv", replacement_docx="cover_letter_test.docx"
            # )
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
