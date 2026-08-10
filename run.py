import os
import secrets

from flask import Flask

from main.routes import bp_main

app = Flask(__name__)
app.config["SECRET_KEY"] = os.getenv("SECRET_KEY", secrets.token_hex(32))
app.register_blueprint(bp_main, url_prefix="")

if __name__ == "__main__":
    app.run(debug=False, port=5000)
