from flask import Flask, request, send_file, render_template, flash
from docx import Document
from cStringIO import StringIO
from werkzeug.utils import secure_filename
from models import REPLACEMENT_DICT, replace_paragraph


def create_app():
    app = Flask(__name__)
    app.secret_key = 'SECRET_KEY'  # Insert secret key here
    app.config['DEBUG'] = True
    return app

app = create_app()

ALLOWED_EXTENSIONS = set(['docx'])


# Checks extensions
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def streamline(document, f, mark):
    replacements = 0
    syllables_saved = 0

    for paragraph in document.paragraphs:
        is_card = False
        # Detects if paragraph is a card (has underlined words)
        for run in paragraph.runs:
            if run.underline or run.style.font.underline:
                is_card = True
        r, s = replace_paragraph(paragraph, is_card, mark)
        replacements += r
        syllables_saved += s

    document.save(f)


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['file']
        filename = secure_filename(file.filename)
        if file:
            if allowed_file(file.filename):
                document = Document(file)
                f = StringIO()
                streamline(document, f, request.form.get('mark'))
                f.seek(0)
                return send_file(f, as_attachment=True,
                                 attachment_filename='streamlined_' + filename)
            else:
                flash('Wrong file format. Only ' +
                      ' '.join(ALLOWED_EXTENSIONS) + ' allowed.', 'danger')
                return render_template('index.html')
        else:
            flash('No file selected.', 'danger')

        return render_template('index.html')


if __name__ == "__main__":
    app.run(port=8080, host='0.0.0.0')
