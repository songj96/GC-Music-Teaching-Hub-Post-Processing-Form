from flask import Flask, render_template, request, Response, session, redirect, url_for, flash
from config import SECRET_KEY, USERNAME, PASSWORD, APIKEY
from flask_wtf import FlaskForm
from flask_wtf.file import FileField, FileAllowed, FileRequired
from wtforms import StringField, BooleanField, SubmitField, SelectField
from wtforms.validators import DataRequired
from tempfile import mkdtemp
from flask_session import Session
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import requests
import base64
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
from googleapiclient.discovery import build
from docx2pdf import convert
import ocrmypdf
import pythoncom
import gdown
import zipfile

app = Flask(__name__)
app.config['SESSION_TYPE'] = 'filesystem'
app.config['SESSION_FILE_DIR'] = mkdtemp()
app.config['SESSION_PERMANENT'] = False
app.config['SESSION_USE_SIGNER'] = True
app.config['SECRET_KEY'] = SECRET_KEY
Session(app)

class PostForm(FlaskForm):
    submission_date = StringField('Submission Date', validators=[DataRequired()])
    submitter_name_for_email = StringField('Submitter Name', validators=[DataRequired()])
    title_course = StringField('Course', validators=[DataRequired()])
    title_material = StringField('Syllabus/Name of Material', validators=[DataRequired()])
    title_college = StringField('College', validators=[DataRequired()])
    title_semester = StringField('Semester', validators=[DataRequired()])
    course_description = StringField('Description', validators=[DataRequired()])
    instructor_name = StringField('Instructor Name', validators=[DataRequired()])
    course_name = StringField('Course Name', validators=[DataRequired()])
    vision_requirement = BooleanField('Is there a vision requirement?')
    document_url = StringField('Document URL', validators=[DataRequired()])
    document_url_links = StringField('Document URL', validators=[DataRequired()])
    submit = SubmitField('Submit')
    category = SelectField('Category', choices=[('assignments', 'Assignments'), ('lesson-plans-and-activities', 'Lesson plans and activities'), ('syllabuses', 'Syllabuses')], validators=[DataRequired()])
    category2 = SelectField('Subcategory', choices=[('Ethnomusicology', 'Ethnomusicology'), ('Musicology', 'Musicology'), ('Music Theory', 'Music Theory'), ('Performance and Composition', 'Performance and Composition')], validators=[DataRequired()])
    tags = StringField('Tags (Separate with commas)', validators=[DataRequired()])


class EmailForm(FlaskForm):
    submitter_name_for_email = StringField('Submitter Name', validators=[DataRequired()])
    submitter_email = StringField('Submitter Email', validators=[DataRequired()])
    post_url = StringField('Post URL', validators=[DataRequired()])
    submit = SubmitField('Submit')

class UploadForm(FlaskForm):
    file1 = FileField('File', validators=[FileRequired(), FileAllowed(['pdf'])])
    file2 = FileField('File', validators=[FileAllowed(['pdf'])])
    file3 = FileField('File', validators=[FileAllowed(['pdf'])])
    file4 = FileField('File', validators=[FileAllowed(['pdf'])])
    file5 = FileField('File', validators=[FileAllowed(['pdf'])])

def word_to_pdf(file_name):
    convert(file_name, f"{file_name[:-5]}.pdf")

def ocr_pdf(file_name):
    try:
        ocrmypdf.ocr(file_name, f"{file_name[:-4]}_OCR.pdf")
        return "OCR completed successfully"
    except ocrmypdf.exceptions.PriorOcrFoundError:
        return "The PDF already had OCR information"
    except ocrmypdf.exceptions.InputFileError:
        return "There was something wrong with the input file"
    
def upload_file_to_wordpress(file_path):
    url = 'https://gcmteachinghub.commons.gc.cuny.edu/'
    rest_api_endpoint = url + 'wp-json/wp/v2/media'
    username = USERNAME
    password = PASSWORD
    credentials = username + ':' + password
    token = base64.b64encode(credentials.encode())
    header = {'Authorization': 'Basic ' + token.decode('utf-8')}

    file_name = os.path.basename(file_path)
    with open(file_path, 'rb') as file:
        files = {'file': (file_name, file, 'multipart/form-data')}
        response = requests.post(rest_api_endpoint, headers=header, files=files)

        if response.status_code == 201:
            response_data = response.json()
            return response_data['source_url']

        return None

def create_text_fields(form, values):
    for i, value in enumerate(values):
        field_name = f'field_{i}'
        setattr(form, field_name, StringField(default=value))

def create_message(sender, to, subject, message_text):
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject
    msg = MIMEText(message_text)
    message.attach(msg)

    raw_message = base64.urlsafe_b64encode(message.as_bytes())
    raw_message = raw_message.decode()
    body  = {'raw': raw_message}
    return body

def send_message(service, user_id, body):
    try:
        message = (service.users().messages().send(userId=user_id, body=body).execute())
        print('Message Id: %s' % message['id'])
        return message
    except Exception as e:
        print(f'An error occurred: {str(e)}')
        return None

def send_confirmation_email(name, link, email):
    creds = None
    client_id = ""
    client_secret = ""
    refresh_token = ""

    creds_info = {
        "client_id": client_id,
        "client_secret": client_secret,
        "refresh_token": refresh_token,
    }

    creds = Credentials.from_authorized_user_info(creds_info)

    service = build('gmail', 'v1', credentials=creds)

    sender = ""
    to = email
    subject = "Submissions to the GC Music Teaching Hub"
    message_text = f"""Dear {name},
    
Thank you for submitting your teaching materials to the GC Music Teaching Hub. 
You can now find your submission at this link: {link}.
    
As you continue making changes to your teaching materials, feel free to resubmit.
    
Best,
GC Music Teaching Hub Admins"""

    message = create_message(sender, to, subject, message_text)
    send_message(service, 'me', message)
    
@app.route('/review', methods=['GET', 'POST'])
def review():
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    client = gspread.authorize(creds)
    spreadsheet = client.open("GC Music Teaching Hub Submission (Responses)").sheet1
    all_data = spreadsheet.get_all_values()
    if len(all_data) > 0:
        if 'current_row' not in session:
            session['current_row'] = len(all_data) - 1
        if request.form.get('move') == 'Previous Row' and session['current_row'] > 0:
            session['current_row'] -= 1
        elif request.form.get('move') == 'Next Row' and session['current_row'] < len(all_data) - 1:
            session['current_row'] += 1
        current_row_values = all_data[session['current_row']]
        form_data = {
            'submission_date': current_row_values[0],
            'title_course': current_row_values[12],
            'title_material': '',
            'title_college': current_row_values[10],
            'title_semester': current_row_values[11],
            'course_description': current_row_values[13],
            'instructor_name': current_row_values[8],
            'course_name': current_row_values[12],
            'vision_requirement': '',
            'tags': current_row_values[14],
            'submitter_name_for_email': current_row_values[2],
            'submitter_email': current_row_values[1],
            'document_url_links': current_row_values[7]
        }
        session['submission_date'] = current_row_values[0]
        session['submitter_name_for_email'] = current_row_values[2]
        session['submitter_email'] = current_row_values[1]
        session['document_url_original'] = current_row_values[7].split(", ")
        session['form_data'] = form_data
        form = PostForm(data=form_data)
        create_text_fields(form, session['document_url_original'])
    else:
        form = PostForm()
    return render_template('review.html', form=form, current_row=session['current_row'] + 1)

@app.route('/ocr', methods=['GET', 'POST'])
def ocr():
    api_key = APIKEY
    if 'document_url_original' in session:
        document_urls_ocr = session['document_url_original']
    else:
        document_urls_ocr = ['', '', '', '', '']

    if request.method == 'POST':
        document_urls_ocr = request.form.getlist('document_url')
        session['document_url_original'] = document_urls_ocr
    
        ocr_pdf_files = [] 

        for file_url in document_urls_ocr:
            ocr_pdf_file = None
            pythoncom.CoInitialize()

            file_id = None
            if 'drive.google.com' in file_url:
                if 'id=' in file_url:
                    file_id = file_url.split('id=')[1]
                elif '/d/' in file_url:
                    file_id = file_url.split('/d/')[1].split('/')[0]

            if file_id is not None:

                drive_api_url = f"https://www.googleapis.com/drive/v3/files/{file_id}?fields=name&key={api_key}"
                response = requests.get(drive_api_url)
                response_json = response.json()
                
                if 'name' not in response_json:
                    flash('Failed to get the file name from Google Drive API.')
                    return render_template('ocr.html', document_urls=document_urls_ocr)

                file_name = response_json.get('name')
                file_name = file_name.replace(':', '_')

                gdrive_url = f"https://drive.google.com/uc?id={file_id}"
                gdown.download(gdrive_url, os.path.join(app.static_folder, file_name), quiet=False)

                if not os.path.isfile(os.path.join(app.static_folder, file_name)):
                    flash("The specified file does not exist.")
                elif file_name.endswith('.docx'):
                    word_to_pdf(os.path.join(app.static_folder, file_name))
                    ocr_pdf_file = f"{file_name[:-5]}.pdf"
                    os.remove(os.path.join(app.static_folder, file_name))  
                elif file_name.endswith('.pdf'):
                    ocr_pdf_file = f"{file_name}"
                else:
                    flash("The specified file is not a DOCX or PDF document.")
                    os.remove(os.path.join(app.static_folder, file_name))

                pythoncom.CoUninitialize()
                if ocr_pdf_file:
                    ocr_pdf_files.append(ocr_pdf_file)

        if len(ocr_pdf_files) == 1:

            ocr_pdf_file = ocr_pdf_files[0]
            ocr_pdf_file_path = os.path.join(app.static_folder, ocr_pdf_file)

            def generate():
                with open(ocr_pdf_file_path, "rb") as f:
                    yield from f
                os.remove(ocr_pdf_file_path)

            response = Response(generate(), mimetype='application/octet-stream')
            response.headers.set("Content-Disposition", "attachment", filename=ocr_pdf_file)
            return response
        elif len(ocr_pdf_files) > 1:

            zip_file_path = os.path.join(app.static_folder, "ocr_result.zip")

            with zipfile.ZipFile(zip_file_path, "w") as zip_file:
                for ocr_pdf_file in ocr_pdf_files:
                    zip_file.write(os.path.join(app.static_folder, ocr_pdf_file), ocr_pdf_file)

            def generate():
                with open(zip_file_path, "rb") as f:
                    yield from f
                os.remove(zip_file_path)
                for ocr_pdf_file in ocr_pdf_files:
                    os.remove(os.path.join(app.static_folder, ocr_pdf_file))

            response = Response(generate(), mimetype='application/octet-stream')
            response.headers.set("Content-Disposition", "attachment", filename="ocr_result.zip")
            return response

    return render_template('ocr.html', document_urls_ocr=document_urls_ocr)


@app.route('/upload', methods=['GET', 'POST'])
def upload_file():
    form = UploadForm()
    if form.validate_on_submit():
        document_urls = []
        files = [form.file1.data, form.file2.data, form.file3.data, form.file4.data, form.file5.data]
        for i, file in enumerate(files, start=1):
            if file:
                file_path = os.path.join(app.static_folder, file.filename)
                file.save(file_path)                 

                document_url = upload_file_to_wordpress(file_path)
                if document_url is None:

                    flash("File upload failed.")
                else:
                    document_urls.append(document_url)

                os.remove(file_path)  

        if document_urls:
            formatted_urls = '\n'.join(document_urls)
            flash(f"Files uploaded successfully. URLs:\n{formatted_urls}")
            session['document_urls'] = document_urls
            create_text_fields(form, session['document_urls'])
            return redirect(url_for('create_post'))

    return render_template('upload.html', form=form)

@app.route('/form', methods=['GET', 'POST'])
def create_post():
    form = PostForm(data=session.get('form_data'))
    if 'document_urls' in session:
        document_urls = session['document_urls']
    else:
        document_urls = ['', '', '', '', '']
        
    if form.validate_on_submit():
        document_urls = request.form.getlist('document_url')
        session['document_url_original'] = document_urls   
    
        url = 'https://gcmteachinghub.commons.gc.cuny.edu/'
        rest_api_endpoint = url + 'wp-json/wp/v2/posts'
        username = USERNAME
        password = PASSWORD
        credentials = username + ':' + password
        token = base64.b64encode(credentials.encode())
        header = {'Authorization': 'Basic ' + token.decode('utf-8')}

        vision_statement = "<!-- wp:paragraph -->\n<p>This assignment contains notated music and requires vision.</p>\n<!-- /wp:paragraph -->" if form.vision_requirement.data else ""

        additional_category = None
        if form.category.data == 'assignments':
            if form.category2.data == 'Ethnomusicology':
                additional_category = 'ethnomusicology-assignments'
            elif form.category2.data == 'Musicology':
                additional_category = 'musicology-assignments'
            elif form.category2.data == 'Music Theory':
                additional_category = 'music-theory-assignments'
            elif form.category2.data == 'Performance and Composition':
                additional_category = 'performance-and-composition-assignments'
        if form.category.data == 'lesson-plans-and-activities':
            if form.category2.data == 'Ethnomusicology':
                additional_category = 'ethnomusicology-lpa'
            elif form.category2.data == 'Musicology':
                additional_category = 'musicology-lpa'
            elif form.category2.data == 'Music Theory':
                additional_category = 'music-theory-lpa'
            elif form.category2.data == 'Performance and Composition':
                additional_category = 'performance-and-composition-lpa'
        if form.category.data == 'syllabuses':
            if form.category2.data == 'Ethnomusicology':
                additional_category = 'ethnomusicology-syllabuses'
            elif form.category2.data == 'Musicology':
                additional_category = 'musicology-syllabuses'
            elif form.category2.data == 'Music Theory':
                additional_category = 'music-theory-syllabuses'
            elif form.category2.data == 'Performance and Composition':
                additional_category = 'performance-and-composition-syllabuses'

        tags_string = form.tags.data
        tagslist = [tag.strip() for tag in tags_string.split(',')]

        url_string = ''
        for url in document_urls:
            if url != '':
                embed = f"""
                <p>[embeddoc url="{url}" download="all"]</p>
                """
                url_string += (embed)

        data = {
            'title': f'{form.title_course.data} | {form.title_material.data} | {form.title_college.data} | {form.title_semester.data}',
            'content': f'''
                <!-- wp:paragraph -->
                <p>{form.course_description.data}</p>
                <!-- /wp:paragraph -->

                <!-- wp:paragraph -->
                <p>Instructor: {form.instructor_name.data}</p>
                <!-- /wp:paragraph -->

                <!-- wp:paragraph -->
                <p>Course: {form.course_name.data}</p>
                <!-- /wp:paragraph -->

                {vision_statement}

                {url_string}
            ''',
            'status': 'draft', 
            'categories': [form.category.data, additional_category] if additional_category else [form.category.data],
            'tags': tagslist,
            'author': 34855,
        }

        response = requests.post(rest_api_endpoint, headers=header, json=data)

        if response.status_code == 201:
            response_data = response.json()
            session['post_url'] = response_data['link']
        
        return redirect(url_for('send_email'))
    return render_template('form.html', form=form, document_urls=document_urls)

@app.route('/send-email', methods=['GET', 'POST'])
def send_email():
    form = EmailForm()

    if request.method == 'GET':
        if 'submitter_name_for_email' in session:
            form.submitter_name_for_email.data = session['submitter_name_for_email']
        if 'post_url' in session:
            form.post_url.data = session['post_url']
        if 'submitter_email' in session:
            form.submitter_email.data = session['submitter_email']

    if form.validate_on_submit():
        send_confirmation_email(form.submitter_name_for_email.data, form.post_url.data, form.submitter_email.data)
        session.pop('submitter_name_for_email', None)
        session.pop('post_url', None)
        session.pop('submitter_email', None)
        return redirect(url_for('table'))
    return render_template('email_form.html', form=form)

@app.route('/guide')
def guide():
    return render_template('guide.html')

@app.route('/table')
def table():
    return render_template('table.html')

if __name__ == '__main__':
    app.run(debug=True)