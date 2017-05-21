from flask import Flask, make_response, request, send_file, render_template, flash
from flask_bootstrap import Bootstrap
from docx import Document
from cStringIO import StringIO
 	
def create_app():
	app = Flask(__name__)
	app.secret_key = 'some_secret'
	Bootstrap(app)

	return app

app = create_app()

regular_dict = {
				# Insert dictionary in this format: "orig_str": "new_str"
				}

def transform(document, f):
	for paragraph in document.paragraphs:
		underline_flag = 0

		for run in paragraph.runs:
			if run.underline == True:
				underline_flag = 1

		if underline_flag == 0:
			for run in paragraph.runs:
				wordlist = run.text.split()
				for word in wordlist:
					if word in regular_dict:
						if word.istitle():
							run.text = run.text.replace(word,regular_dict[word].title())
						elif word.isupper():
							run.text = run.text.replace(word,regular_dict[word].upper())
						else:
							run.text = run.text.replace(word,regular_dict[word])

	document.save(f)
	return f

@app.route('/')
def form():
    return render_template('index.html')

@app.route('/transform', methods=['POST'])
def transform_view():
    file = request.files['data_file']
    print file.name
    if not file:
        return "No file"

    document = Document(file)
    f = StringIO()
    result = transform(document, f)

    length = result.tell()
    f.seek (0)
    return send_file(f, as_attachment=True, attachment_filename='report.doc')
