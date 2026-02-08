import os
from io import BytesIO
import copy
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

                    return render_template('create_template.html', text=full_text, fileName=file.filename, allowed_extensions=ALLOWED_EXTENSIONS)

        except RequestEntityTooLarge:
            flash('Maximum filesize: ' + str(app.config['MAX_CONTENT_LENGTH'] / (1000 * 1000)) + ' MB' )
            return redirect(request.url)
        
        if request.form.get('selections'):
            selections = json.loads(request.form['selections'])
            # return f"<pre>{selections[0]}</pre>"

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

            # We're going to march through this char by char and replace exactly our selection, keeping all that formatting crap intact
            # We, in fact, did not do that

            final_fields = []
            for selection in selections:
                selText        = "".join(selection['text'].split())
                pHolderText    = selection['placeholder']
                # paragraphIndex = int(selection['range']['paragraphIndex']['startContainer'])
                # startIdx       = selection['sOffsets']['start']
                # endIdx         = selection['sOffsets']['end']
                paragraphNum = 0

                for paragraph in document.paragraphs:
                    paragraphNum += 1
                    for run in paragraph.runs:
                        stripped_run = []
                        original_index = []
                        for index, char in enumerate(run.text):
                            if not char.isspace():
                                stripped_run.append(char)
                                original_index.append(index)

                        stripped_run = "".join(stripped_run)
                        # Finding the starting char in matching subby
                        start_stripped = stripped_run.find(selText)
                        # In a function we'll just return -1
                        if start_stripped == -1:
                            continue
                        start_original = original_index[start_stripped]
                        end_original = original_index[start_stripped + len(selText) - 1]

                        # Create a new string by combining slices and a new character
                        run.text = run.text[:start_original] + pHolderText + run.text[end_original + 1:]

                        temp_selection = copy.deepcopy(selection)
                        temp_selection['range']['paragraphIndex']['startContainer'] = paragraphNum
                        final_fields.append(temp_selection)

            full_text = []
            for paragraph in document.paragraphs:
                full_text.append(paragraph.text)


            # This will be returned from download function or page or whatever
            # target_stream = BytesIO()
            # document.save(target_stream)
            # target_stream.seek(0)

            # return send_file(
            # target_stream,
            # as_attachment=True,
            # download_name="modified.pdf"
            # )
            
            # F_text = [p.text for p in document.paragraphs]
            return render_template('show_fields.html', text=full_text, fields=final_fields)

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

        else:
            flash('Accepted filetypes: ' + str(ALLOWED_EXTENSIONS).strip('}{'))
            return redirect(request.url)            
    else:
        return render_template('create_template.html', allowed_extensions=ALLOWED_EXTENSIONS)

if __name__ == "__main__":
    app.run(debug=True)