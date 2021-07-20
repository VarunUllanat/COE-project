import os

import flask
from flask import Flask, request, redirect, url_for, send_from_directory
from werkzeug.utils import secure_filename
import pandas as pd
from flask import send_file
import pubchempy as pcp
from zipfile import ZipFile
from collections import deque

UPLOAD_FOLDER = 'static/'
ALLOWED_EXTENSIONS = set(['xlsx'])

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1] in ALLOWED_EXTENSIONS


def pipeline(filename, opt):
    data = pd.read_excel(filename, header=None)
    list_of_compounds = data[0]
    files = []
    no_result = []
    for i in list_of_compounds:
        try:
            pcp.download('SDF', app.config['UPLOAD_FOLDER'] + i + opt + '.sdf', i, 'name', record_type=opt)
            files.append(app.config['UPLOAD_FOLDER'] + i + opt + '.sdf')
        except:
            no_result.append(i)
    os.remove(filename)
    return files, no_result


def return_files_tut(filename):
    try:
        return send_file(filename, attachment_filename='result.zip')
    except Exception as e:
        return str(e)

def zip_files(files):
    with ZipFile(app.config['UPLOAD_FOLDER'] + 'files.zip', mode='w') as zf:
        for f in files:
            zf.write(f)
            os.remove(f)
    return app.config['UPLOAD_FOLDER'] + 'files.zip'


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        upfile = request.files['file']
        option = request.form['options']
        if upfile and allowed_file(upfile.filename):
            dir = app.config['UPLOAD_FOLDER']
            for f in os.listdir(dir):
                os.remove(os.path.join(dir, f))
            filename = secure_filename(upfile.filename)
            upfile.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            files, no_result = pipeline(os.path.join(app.config['UPLOAD_FOLDER'], filename), option)
            zipped = zip_files(files)
            return return_files_tut(zipped)
    return flask.render_template('main.html')


if __name__ == "__main__":
    app.run(port=5000, debug=True)
