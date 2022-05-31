from flask import Flask

app = Flask(__name__)
app.config["SECRET_KEY"] = "22650567d247eb344bf23637"
app.config["UPLOAD_FOLDER"] = "converter/static/files"

from converter import routes
