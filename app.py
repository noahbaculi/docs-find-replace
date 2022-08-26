from flask import Flask
import gunicorn  # must be in requirements.txt for Heroku deployment

app = Flask(__name__)


@app.route("/")
def index():
    return "Hello World"


if __name__ == "__main__":
    app.run()
