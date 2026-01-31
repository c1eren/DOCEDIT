import os
from io import BytesIO
from flask import Flask, request, flash, redirect, render_template, send_file
from werkzeug.utils import secure_filename
from werkzeug.exceptions import RequestEntityTooLarge

from docx import Document

app = Flask(__name__)

ALLOWED_EXTENSIONS = {'txt', 'pdf', 'docx'}
UPLOAD_FOLDER = os.path.join(os.getcwd(), "uploads")

app.secret_key = "dev-secret-key"
app.config['MAX_CONTENT_LENGTH'] = 10 * 1000 * 1000 #10 MB
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/")
def init_page():
    return '<a href="/upload">Go to upload page</a>'

@app.route("/test")
def hello_world(text="Upload document"):
    return render_template('index.html', innertext=text)

@app.route('/upload', methods=['GET', 'POST'])
def upload_document():
    if request.method == 'POST':

        # Check if the post request has the file part
        try:
            if 'file' not in request.files:
                flash('No file part')
                return redirect(request.url)
        except RequestEntityTooLarge:
            flash('Maximum filesize: ' + str(app.config['MAX_CONTENT_LENGTH'] / (1000 * 1000)) + ' MB' )
            return redirect(request.url)
        
        # Get file
        file = request.files['file']

        # Check if file actually submitted
        if file.filename == '':
            flash('No file selected')
            return redirect(request.url)
        
        # Check file exists and is allowed
        if file and allowed_file(file.filename):
            source_stream = BytesIO(file.read())
            document = Document(source_stream)
            source_stream.close()

            # Replace chosen text in paragraph (with styling these are called "Run"s)
            for paragraph in document.paragraphs:
                for run in paragraph.runs:
                    if "Cieren" in run.text:
                        run.text = run.text.replace("Cieren", "ME")

            target_stream = BytesIO()
            document.save(target_stream)
            target_stream.seek(0)

            return send_file(
            target_stream,
            as_attachment=True,
            download_name="modified.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
             
                
            filename = secure_filename(file.filename)
            # file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            # return '<p>good job</p>'
            # return redirect(url_for('download_file', name=filename))
        else:
            flash('Accepted filetypes: ' + str(ALLOWED_EXTENSIONS).strip('}{'))
            return redirect(request.url)            
    else:
        return render_template('index.html')

if __name__ == "__main__":
    app.run(debug=True)