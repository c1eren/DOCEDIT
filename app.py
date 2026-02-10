import os
from io import BytesIO
import uuid
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
    return render_template("template_hub.html")

@app.route('/create-template', methods=['GET', 'POST'])
def create_template_page():
    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'upload':
            return handle_upload(request)         

        elif action == 'selections':
            return handle_selections(request)
            
        elif action == 'fields':
            return handle_fields(request)

        else:
            abort(400, "Invalid or missing action")
            
    else:
        return render_template('create_template.html', allowed_extensions=ALLOWED_EXTENSIONS)
                
                    # for paragraph in document.paragraphs:
                    #     paragraphNum += 1
                    #     for run in paragraph.runs:
                    #         stripped_run = []
                    #         original_index = []
                    #         for index, char in enumerate(run.text):
                    #             if not char.isspace():
                    #                 stripped_run.append(char)
                    #                 original_index.append(index)

                    #         stripped_run = "".join(stripped_run)
                    #         # Finding the starting char in matching subby
                    #         start_stripped = stripped_run.find(selText)
                    #         # In a function we'll just return -1
                    #         if start_stripped == -1:
                    #             continue
                    #         start_original = original_index[start_stripped]
                    #         end_original = original_index[start_stripped + len(selText) - 1]

                    #         # Create a new string by combining slices and a new character
                    #         run.text = run.text[:start_original] + pHolderText + run.text[end_original + 1:]

                    #         temp_selection = {
                    #             "text": selection['text'],
                    #             "paragraphIndex": paragraphNum,
                    #             "index": {"start": start_original, "end": end_original},
                    #             "field_text": pHolderText
                    #             }
                    #         final_fields.append(temp_selection)

                    # for paragraph in document.paragraphs:
                    #     paragraphNum += 1
                    #     for run in paragraph.runs:
                    #         stripped_run = []
                    #         original_index = []
                    #         for index, char in enumerate(run.text):
                    #             if not char.isspace():
                    #                 stripped_run.append(char)
                    #                 original_index.append(index)

                    #         stripped_run = "".join(stripped_run)
                    #         # Finding the starting char in matching subby
                    #         start_stripped = stripped_run.find(selText)
                    #         # In a function we'll just return -1
                    #         if start_stripped == -1:
                    #             continue
                    #         start_original = original_index[start_stripped]
                    #         end_original = original_index[start_stripped + len(selText) - 1]

                    #         # Create a new string by combining slices and a new character
                    #         run.text = run.text[:start_original] + pHolderText + run.text[end_original + 1:]

                    #         temp_selection = copy.deepcopy(selection)
                    #         temp_selection['range']['paragraphIndex']['startContainer'] = paragraphNum
                    #         final_fields.append(temp_selection)

                # for paragraph in document.paragraphs:
                #     full_text.append(paragraph.text)


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

def get_current_doc_path():
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
    
    return doc_path


def handle_upload(request):
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
                    else:
                        flash('Accepted filetypes: ' + str(ALLOWED_EXTENSIONS).strip('}{'))
                        return redirect(request.url)
                else:
                    flash('File not found in request')
                    return redirect(request.url)         

            except RequestEntityTooLarge:
                flash('Maximum filesize: ' + str(app.config['MAX_CONTENT_LENGTH'] / (1000 * 1000)) + ' MB' )
                return redirect(request.url)
            
def handle_selections(request):
    try:
        selections = json.loads(request.form['selections'])
        # return f"<div>{selections}</div>"

        doc_path = get_current_doc_path()    
        document = Document(doc_path)

        # We're going to march through this char by char and replace exactly our selection, keeping all that formatting crap intact
        # We, in fact, did not do that
        full_text = []
        final_fields = []


        for pIndex, paragraph in enumerate(document.paragraphs, start=1):

            preview_text = ""
            reference_text = ""
            
            for run in paragraph.runs:
                reference_text += run.text
                preview_text += run.text
                temp_selections = []

            for selection in selections:                
                selText        = "".join(ch for ch in selection['text'] if not ch.isspace()).lower()
                pHolderText    = selection['placeholder']

                stripped_text  = []
                original_index = []

                for index, char in enumerate(reference_text):
                    if not char.isspace():
                        stripped_text.append(char)
                        original_index.append(index)

                stripped_text = "".join(stripped_text).lower()
                # Finding the starting char in matching subby
                start_stripped = stripped_text.find(selText)
                # In a function we'll just return -1
                if start_stripped != -1:
                    start_original = original_index[start_stripped]
                    end_original = original_index[start_stripped + len(selText) - 1]
                    # Create a new string by combining slices and a new character
                    preview_text = preview_text[:start_original] + pHolderText + preview_text[end_original + 1:]

                    temp_selection = {
                        "uuid": str(uuid.uuid4()),
                        "text": selection['text'],
                        "paragraphIndex": pIndex,
                        "index": {"start": start_original, "end": end_original},
                        "field_text": pHolderText
                        }
                    temp_selections.append(temp_selection)
                    final_fields.append(temp_selection)
                    
            full_text.append(preview_text)

        return render_template('save_template.html', text=full_text, fields=final_fields)
    
    except KeyError:
        # Missing 'selections' in POST
        flash("No selection data received from the form.")
        return redirect(request.url)

    except Exception as e:
        # Catch everything else lol
        app.logger.error(f"Unexpected error in handle_selections: {e}")
        flash("An unexpected error occurred. Please try again.")
        return redirect(request.url)
    
def handle_fields(request):
    try:
        fields = json.loads(request.form['fields'])

        doc_path = get_current_doc_path()    
        document = Document(doc_path)

        # TODO: Fix all this logic lol
        
        # Go through each paragraph 
        # for pIndex, paragraph in enumerate(document.paragraphs, start=1):
        #     # Check each field against the run to see if it has a match in there to replace
        #     for field in fields:
        #         if field['paragraphIndex'] == str(pIndex):
        #             for run in paragraph.runs:
        #                 stripped_run = []
        #                 original_index = []
        #                 for index, char in enumerate(run.text):
        #                     if not char.isspace():
        #                         stripped_run.append(char)
        #                         original_index.append(index)

        #                 stripped_run = "".join(stripped_run)
        #                 # Finding the starting char in matching subby
        #                 start_stripped = stripped_run.find(field['text'])
        #                 # In a function we'll just return -1
        #                 if start_stripped == -1:
        #                     continue
        #                 start_original = original_index[start_stripped]
        #                 end_original = original_index[start_stripped + len(field['text']) - 1]

        #                 # Create a new string by combining slices and a new character
        #                 run.text = run.text[:start_original] + field['field_text'] + run.text[end_original + 1:]
        
        
        for pIndex, paragraph in enumerate(document.paragraphs, start=1):
            # Sort fields in reverse based on where they are in paragraph
            para_fields = [f for f in fields if f["paragraphIndex"] == pIndex]
            para_fields.sort(key=lambda f: f["index"]["start"], reverse=True)
            flat = ""
            runs = []
            # Combine runs into flat text, map the start and end indexes per paragraph 
            for run in paragraph.runs:
                start = len(flat)
                flat += run.text
                end = len(flat)
                runs.append({
                    "run": run,
                    "start": start,
                    "end": end
                })
                
            # Using indexes from the flattened list and the field indexes
            for field in para_fields:
                if field['paragraphIndex'] == pIndex:
                    field_start = field["index"]["start"]
                    field_end   = field["index"]["end"] + 1
                    replacement = field["field_text"]

                    inserted = False

                    for r in runs:
                        if r["end"] <= field_start or r["start"] >= field_end:
                            continue

                        run = r["run"]
                        text = run.text

                        local_start = max(field_start, r["start"]) - r["start"]
                        local_end   = min(field_end, r["end"]) - r["start"]

                        if not inserted:
                            run.text = text[:local_start] + replacement + text[local_end:]
                            inserted = True
                        else:
                            run.text = text[:local_start] + text[local_end:]
        full_text = []
        for paragraph in document.paragraphs:
            full_text.append(paragraph.text)
        return f"<div>{full_text}</div>"

    except KeyError:
        # Missing 'selections' in POST
        flash("No selection data received from the form.")
        return redirect(request.url)

    except Exception as e:
        # Catch everything else lol
        app.logger.error(f"Unexpected error in handle_fields: {e}")
        flash("An unexpected error occurred. Please try again.")
        return redirect(request.url)



if __name__ == "__main__":
    app.run(debug=True)