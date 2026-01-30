from flask import Flask
from flask import render_template

app = Flask(__name__)

@app.route("/")
def init_page():
    return "<p>initpage</p>"

@app.route("/test")
def hello_world(text="TEST text"):
    return render_template('index.html', innertext=text)

if __name__ == "__main__":
    app.run(debug=True)