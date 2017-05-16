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

regular_dict = {"because": "since",
				"insofar as": "since",
				"in a world where": "since",
				"therefore": "thus",
				"ergo": "thus",
				"which means that": "so",
				"that means that": "so",
				"which means": "so",
				"resolution": "topic",
				"the affirmative": "the aff",
				"the negative": "the neg",
				"my opponent": "they",
				"my value is": "I value",
				"the criterion": "the standard",
				"the value criterion": "the standard",
				"tells you": "says",
				"today's resolution": "",
				"in today's round": "",
				"in this round": "",
				"in today's debate round": "",
				"in this debate round": "",
				"explains": "writes",
				"firstly": "first",
				"secondly": "second",
				"thirdly": "third",
				"fourthly": "fourth",
				"fifthly": "fifth",
				"sixthly": "sixth",
				"sevently": "seventh",
				"eightly": "eight",
				"ninthly": "ninth",
				"respond to": "answer",
				"responds to": "answers",
				"response": "answer",
				"argument": "reason",
				"implication": "impact",
				"justification": "warrant",
				"reason as to why": "reason",
				"reasons as to why": "reasons",
				"basically": "",
				"inherently": "",
				"fundamentally": "",
				"ultimately": "",
				"essentially": "",
				"obviously": "",
				"intrinsically": "",
				"furthermore": "and",
				"additionally": "and",
				"in addition": "and",
				"moreover": "and",
				"can not": "can't",
				"are not": "aren't",
				"cannot": "can't",
				"do not": "don't",
				"he had": "he'd",
				"he would": "he'd",
				"he shall": "he'll",
				"he will": "he'll",
				"he has": "he's",
				"he is": "he's",
				"how did": "how'd",
				"how would": "how'd",
				"how has": "how's",
				"how is": "how's",
				"I had": "I'd",
				"I would": "I'd",
				"I shall": "I'll",
				"I will": "I'll",
				"I am": "I'm",
				"I have": "I've",
				"it would": "it'd",
				"it shall": "it'll",
				"it will": "it'll",
				"it has": "it's",
				"it is": "it's",
				"of the clock": "o'clock",
				"somebody is": "somebody's",
				"someone had": "someone'd",
				"someone would": "someone'd",
				"someone shall": "someone'll",
				"someone will": "someone'll",
				"someone has": "someone's",
				"someone is": "someone's",
				"something had": "something'd",
				"something would": "something'd",
				"something shall": "something'll",
				"something will": "something'll",
				"something has": "something's",
				"something is": "something's",
				"that is": "that's",
				"that has": "that's",
				"that would": "that'd",
				"that had": "that'd",
				"there had": "there'd",
				"there would": "there'd",
				"there is": "there's",
				"there has": "there's",
				"they had": "they'd",
				"they would": "they'd",
				"they shall / they will": "they'll",
				"they are": "they're",
				"they have": "they've",
				"we had": "we'd",
				"we would": "we'd",
				"we will": "we'll",
				"we are": "we're",
				"we have": "we've",
				"were not": "weren't",
				"what did": "what'd",
				"what shall": "what'll",
				"what will": "what'll",
				"what is": "what's",
				"what has": "what's",
				"when is": "when's",
				"when has": "when's",
				"where did": "where'd",
				"where is": "where's",
				"where has": "where's",
				"where have": "where've",
				"who would": "who'd",
				"who had": "who'd",
				"who did": "who'd",
				"who shall": "who'll",
				"who will": "who'll",
				"who are": "who're",
				"who has": "who's",
				"who is": "who's",
				"who have": "who've",
				"why did": "why'd",
				"why has": "why's",
				"why is": "why's",
				"will not": "won't",
				"will not have": "won't've",
				"you had": "you'd",
				"you would": "you'd",
				"you would have": "you'd've",
				"you shall": "you'll",
				"you will": "you'll",
				"you are": "you're",
				"you are not": "you aren't",
				"you have": "you've",
				"is going to": "will",
				"for example": "for instance",
				"such as": "like",
				"however": "but",
				"whether or not": "if",
				"in order to": "to",
				"government": "state",
				"United States": "US",
				"individuals": "persons",
				"undermines": "harms",
				"action": "act",
				"obligation": "duty",
				"autonomy": "agency",
				"autonomous": "agential",
				"adjudicate": "judge",
				"suggests": "says",
				"incorrect": "wrong",
				"was not able to": "couldn't",
				"was able to": "could",
				"is not able to": "can't",
				"is able to": "can",
				"an individual": "a person",
				"the individual": "the person",
				"means that": "means",
				"we would call": "we call",
				"vote affirmative": "vote aff",
				"vote negative": "vote neg",
				"additional": "further",
				"need to": "must",
				"it is necessary to": "we must",
				"ethical": "moral",
				"allow": "let",
				"a moral code": "an ethic",
				"a moral system": "an ethic",
				"a system of morality": "an ethic",
				"a code of morality": "an ethic",
				"technology": "tech",
				"information": "info",
				"powerful": "potent",
				"particularly": "especially",
				"weapon of mass destruction": "WMD",
				"weapons of mass destruction": "WMDs",
				"eliminates": "takes out",
				"eliminating": "taking out",
				"reducing": "shrinking",
				"reduce": "shrink",
				"decrease": "shrink",
				"decreasing": "shrinking",
				"critical": "key",
				"crucial": "key",
				"combat": "fight",
				"an affirmative": "an aff",
				"a negative": "a neg",
				"preventing": "stopping",
				"prevent": "stop",
				"a moral theory": "the ethic",
				"the moral code": "the ethic",
				"the moral system": "the ethic",
				"the system of morality": "the ethic",
				"the code of morality": "the ethic",
				"the moral theory": "the ethic",
				"demonstrate": "show",
				"show that": "show",
				"prove that": "prove",
				"utilize": "use",
				"proven": "shown",
				"minimize": "shrink",
				"maximize": "increase",
				"lead to": "cause",
				"leading to": "causing",
				"summarize": "explain",
				"results in": "causes",
				"result in": "cause",
				"given that": "since",
				"the proliferation": "the spread",
				"repercussion": "consequence",
				"to do so": "to",
				"protect": "guard",
				"important": "crucial"
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

@app.route('/transform', methods=["POST"])
def transform_view():
    file = request.files['data_file']
    if not file:
        return "No file"

    document = Document(file)
    f = StringIO()
    result = transform(document, f)

    length = result.tell()
    f.seek (0)
    return send_file(f, as_attachment=True, attachment_filename='report.doc')