# docs-find-replace

## Web API

* [Flask API](https://doc-find-replace.herokuapp.com/)
* Hosted on [Heroku](https://www.heroku.com/)
* Does not support `.pdf` output since it is not running on a Windows server

```shell
# setup
heroku login
git push heroku main

# open app
heroku open

# check logs
heroku logs --tail
```

Heroku uses `requirements.txt` and `Procfile` files to build the Python app.

Can use the `pipreqs` package to write only the packages used in project to
`requirements.txt`.

## Local tool

* Local Python script run
* Can output `.pdf` files if run on Windows machine

```shell
# clone repo
git clone https://github.com/noahbaculi/docs-find-replace.git

# install dependencies
python -m venv ./venv
source venv/bin/activate
pip install -r requirements.txt

# edit args as desired and run script
python doc_find_replace.py
```

The `ThreadPoolExecutor` implementation has the most impact when outputting
`.pdf` files.
