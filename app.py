import os
from flask import Flask, request, render_template, send_from_directory
from main import work_dir, pdf_dir, add_to_excel, add_ack_toexcel, make_log, init


UPLOAD_FOLDER = pdf_dir
ALLOWED_EXTENSIONS = {'pdf'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

redirect = '<a href="/">go back</a>'


@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')


@app.route('/upload_invoice', methods=['POST'])
def upload_file_inv():
    if request.method == 'POST':

        files = request.files.getlist("file_inv[]")

        if len(files[0].filename) == 0:
            return render_template('index.html', data='please upload 1 or more than 1 file')
        response = ''
        count = 0
        make_log('Upload received : ' + str(files))
        for file in files:
            if (file.filename.split('.')[1]).lower() != 'pdf':
                make_log('FOUND other than pdf: ' + file.filename)
                count += 1
                continue

            file.save(os.path.join(app.config['UPLOAD_FOLDER'], file.filename))
            response += add_to_excel(file.filename)

        return str(len(files) - count) + '<hr>' + response + redirect


@app.route('/upload_acknowledgement', methods=['POST'])
def upload_file_ack():
    if request.method == 'POST':

        files = request.files.getlist("file_ack[]")

        if len(files[0].filename) == 0:
            return render_template('index.html', data='please upload 1 or more than 1 file')
        response = ''
        count = 0
        make_log('Upload received : ' + str(files))

        for file in files:
            if file.filename.split('.')[1].lower() != 'pdf':
                make_log('FOUND other than pdf: ' + file.filename)
                count += 1
                continue
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], file.filename))
            response += add_ack_toexcel(file.filename)

        return str(len(files) - count) + '<hr>' + response + redirect


@app.route('/download_inv/<path:filename>')
def send_inv(filename):

    filename = 'invoice.xlsx'
    return send_from_directory(work_dir, filename, as_attachment=True, cache_timeout=0)


@app.route('/download_ack/<path:filename>')
def send_ack(filename):

    filename = 'acknowledgement.xlsx'
    return send_from_directory(work_dir, filename, as_attachment=True, cache_timeout=0)


@app.route('/upload_excel_inv', methods=['POST'])
def upload_excel_inv():
    if request.method == 'POST':
        file = request.files['excel']
        try:
            if file.filename.split('.')[1].lower() != 'xlsx':
                return render_template('index.html', data='please upload file with extension ".xlsx"')
        except IndexError:
            return render_template('index.html', data='please upload 1 or more than 1 file')
        file.filename = 'invoice.xlsx'
        file.save(os.path.join(app.config['UPLOAD_FOLDER'] + '/..', file.filename))
        return "upload successful" + redirect


@app.route('/upload_excel_ack', methods=['POST'])
def upload_excel_ack():
    if request.method == 'POST':

        file = request.files['excel']
        try:
            if file.filename.split('.')[1].lower() != 'xlsx':
                return render_template('index.html', data='please upload file with extension ".xlsx"')
        except IndexError:
            return render_template('index.html', data='please upload 1 or more than 1 file')

        file.filename = 'acknowledgement.xlsx'
        file.save(os.path.join(app.config['UPLOAD_FOLDER'] + '/..', file.filename))
        return "upload successful" + redirect


if __name__ == '__main__':
    init()
    app.secret_key = 'some secret key'
    app.run(host='0.0.0.0', port=5000, debug=True)

