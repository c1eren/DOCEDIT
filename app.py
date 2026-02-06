import os
from io import BytesIO
from flask import Flask, request, flash, redirect, render_template, send_file, session, abort
from werkzeug.utils import secure_filename
from werkzeug.exceptions import RequestEntityTooLarge
from werkzeug.datastructures import FileStorage

import json
from docx import Document

app = Flask(__name__)

ALLOWED_EXTENSIONS = {'docx'}
UPLOAD_FOLDER = os.path.join(os.getcwd(), "uploads")

app.secret_key = "dev-secret-key"
app.config['MAX_CONTENT_LENGTH'] = 10 * 1000 * 1000 #10 MB
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Might need to save the document temporarily to keep it in memory or something so it can be altered later

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/")
def init_page():
    return '<a href="/upload">Go to upload page</a>'

@app.route('/upload', methods=['GET', 'POST'])
def upload_document():
    if request.method == 'POST':

        # Check if the post request has the file part
        try:
            if 'file' in request.files:
                # Check if file actually submitted
                # Get file
                file = request.files['file']

                if file.filename == '':
                    flash('No file selected')
                    return redirect(request.url)
                
                # Check file exists and is allowed
                if file and allowed_file(file.filename):
                    source_stream = BytesIO(file.read())
                    document = Document(source_stream)
                    source_stream.close()

                    # Save uploaded file temporarily and the path in session
                    fName = secure_filename(file.filename)
                    path = os.path.join(app.config['UPLOAD_FOLDER'], fName)
                    file.stream.seek(0)
                    file.save(path)
                    session['doc_path'] = path

                    # Extract text from all paragraphs
                    full_text = []
                    for paragraph in document.paragraphs:
                        full_text.append(paragraph.text)

                    return render_template('display.html', text=full_text)

        except RequestEntityTooLarge:
            flash('Maximum filesize: ' + str(app.config['MAX_CONTENT_LENGTH'] / (1000 * 1000)) + ' MB' )
            return redirect(request.url)
        
        if request.form.get('selections'):
            selections = json.loads(request.form['selections'])
            return f"<pre>{selections}</pre>"



            doc_path = session.get('doc_path')
            if not doc_path:
                abort(400)

            if not os.path.isfile(doc_path):
                app.logger.error(f"Missing docx: {doc_path}")
                flash('Document file does not exist, try a different file')
                return redirect(request.url)

            if os.path.getsize(doc_path) == 0:
                flash('Document file is empty, try a different file')
                return redirect(request.url)    
            
            # try:
            document = Document(doc_path)
            F_text = []
            for item in selections:
                # Each item in the list is a dictionary
                F_text.append(item['text'])
            for paragraph in document.paragraphs:
                for run in paragraph.runs:
                    for text in F_text:
                        if text in run.text:
                            run.text = run.text.replace(text, "$$HERE$$")
            full_text = []
            for paragraph in document.paragraphs:
                full_text.append(paragraph.text)
            # F_text = [p.text for p in document.paragraphs]
            return render_template('show_fields.html', text=full_text)

            # except FileNotFoundError:
                # abort(404, description="Document not found") 

            # t = []
            # for item in selections:
            #     # Each item in the list is a dictionary
            #     t.append(item['text'])
            # for paragraph in document.paragraphs:
            #     for run in paragraph.runs:
            #         for text in t:
            #             if text in run.text:
            #                 run.text = run.text.replace(text, "HERE")

            # target_stream = BytesIO()
            # document.save(target_stream)
            # target_stream.seek(0)
            # Extract text from all paragraphs

            return render_template('show_fields.html', text=F_text)

            return render_template('show_fields.html', fields=t)
    
        

        
        # Need to POST an array with the fields to edit

            # Replace chosen text in paragraph (with styling these are called "Run"s)
            # TODO Work out how to pull word list from fields.
            for paragraph in document.paragraphs:
                for run in paragraph.runs:
                    if "Cieren" in run.text:
                        run.text = run.text.replace("Cieren", "ME")

            target_stream = BytesIO()
            document.save(target_stream)
            target_stream.seek(0)

            # This will be returned from download function or page or whatever
            # return send_file(
            # target_stream,
            # as_attachment=True,
            # download_name="modified.docx",
            # mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            # )
             
                
            filename = secure_filename(file.filename)
            # file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            # return '<p>good job</p>'
            # return redirect(url_for('download_file', name=filename))
        else:
            flash('Accepted filetypes: ' + str(ALLOWED_EXTENSIONS).strip('}{'))
            return redirect(request.url)            
    else:
        return render_template('display.html')

if __name__ == "__main__":
    app.run(debug=True)