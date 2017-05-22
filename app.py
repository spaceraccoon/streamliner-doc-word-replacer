from flask import Flask, request, send_file, render_template, flash
from docx import Document
from cStringIO import StringIO
from werkzeug.utils import secure_filename
    
def create_app():
    app = Flask(__name__)
    app.secret_key = 'SECRET_KEY'  # Insert secret key here

    return app

app = create_app()

ALLOWED_EXTENSIONS = set(['docx'])

# Replace this with your own entries in the form old: (new,syllable_diff) 
REPLACEMENT_DICT = {"because": ("since",1),
                    "insofar as": ("since",3),
                    "in a world where": ("since",3),
                    "therefore": ("thus",1),
                    "ergo": ("thus",1),
                    "which means that": ("so",2),
                    "that means that": ("so",2),
                    "which means": ("so",1),
                    "resolution": ("topic",2),
                    "the affirmative": ("the aff",3),
                    "the negative": ("the neg",2),
                    "my opponent": ("they",3),
                    "my value is": ("I value",1),
                    "the criterion": ("the standard",2),
                    "the value criterion": ("the standard",4),
                    "tells you": ("says",1),
                    "today's resolution": ("",6),
                    "in today's round": ("",4),
                    "in this round": ("",3),
                    "in today's debate round": ("",6),
                    "in this debate round": ("",5),
                    "explains": ("writes",1),
                    "firstly": ("first",1),
                    "secondly": ("second",1),
                    "thirdly": ("third",1),
                    "fourthly": ("fourth",1),
                    "fifthly": ("fifth",1),
                    "sixthly": ("sixth",1),
                    "sevently": ("seventh",1),
                    "eightly": ("eight",1),
                    "ninthly ": ("ninth",1),
                    "respond to": ("answer",2),
                    "responds to": ("answers",2),
                    "response": ("answer",1),
                    "argument": ("reason",1),
                    "implication": ("impact",2),
                    "justification": ("warrant",3),
                    "reason as to why": ("reason",3),
                    "reasons as to why": ("reasons",3),
                    "basically": ("",4),
                    "inherently": ("",4),
                    "fundamentally": ("",5),
                    "ultimately": ("",4),
                    "essentially": ("",4),
                    "obviously": ("",4),
                    "intrinsically": ("",5),
                    "furthermore": ("and",2),
                    "additionally": ("and",4),
                    "in addition": ("and",3),
                    "moreover": ("and",2),
                    "can not": ("can't",1),
                    "are not": ("aren't",1),
                    "cannot": ("can't",1),
                    "do not": ("don't",1),
                    "he had": ("he'd",1),
                    "he would": ("he'd",1),
                    "he shall": ("he'll",1),
                    "he will": ("he'll",1),
                    "he has": ("he's",1),
                    "he is": ("he's",1),
                    "how did": ("how'd",1),
                    "how would": ("how'd",1),
                    "how has": ("how's",1),
                    "how is": ("how's",1),
                    "I had": ("I'd",1),
                    "I would": ("I'd",1),
                    "I shall": ("I'll",1),
                    "I will": ("I'll",1),
                    "I am": ("I'm",1),
                    "I have": ("I've",1),
                    "it would": ("it'd",1),
                    "it shall": ("it'll",1),
                    "it will": ("it'll",1),
                    "it has": ("it's",1),
                    "it is": ("it's",1),
                    "of the clock": ("o'clock",1),
                    "somebody is": ("somebody's",1),
                    "someone had": ("someone'd",1),
                    "someone would": ("someone'd",1),
                    "someone shall": ("someone'll",1),
                    "someone will": ("someone'll",1),
                    "someone has": ("someone's",1),
                    "someone is": ("someone's",1),
                    "something had": ("something'd",1),
                    "something would": ("something'd",1),
                    "something shall": ("something'll",1),
                    "something will": ("something'll",1),
                    "something has": ("something's",1),
                    "something is": ("something's",1),
                    "that is": ("that's",1),
                    "that has": ("that's",1),
                    "that would": ("that'd",1),
                    "that had": ("that'd",1),
                    "there had": ("there'd",1),
                    "there would": ("there'd",1),
                    "there is": ("there's",1),
                    "there has": ("there's",1),
                    "they had": ("they'd",1),
                    "they would": ("they'd",1),
                    "they shall / they will": ("they'll",1),
                    "they are": ("they're",1),
                    "they have": ("they've",1),
                    "we had": ("we'd",1),
                    "we would": ("we'd",1),
                    "we will": ("we'll",1),
                    "we are": ("we're",1),
                    "we have": ("we've",1),
                    "were not": ("weren't",1),
                    "what did": ("what'd",1),
                    "what shall": ("what'll",1),
                    "what will": ("what'll",1),
                    "what is": ("what's",1),
                    "what has": ("what's",1),
                    "when is": ("when's",1),
                    "when has": ("when's",1),
                    "where did": ("where'd",1),
                    "where is": ("where's",1),
                    "where has": ("where's",1),
                    "where have": ("where've",1),
                    "who would": ("who'd",1),
                    "who had": ("who'd",1),
                    "who did": ("who'd",1),
                    "who shall": ("who'll",1),
                    "who will": ("who'll",1),
                    "who are": ("who're",1),
                    "who has": ("who's",1),
                    "who is": ("who's",1),
                    "who have": ("who've",1),
                    "why did": ("why'd",1),
                    "why has": ("why's",1),
                    "why is": ("why's",1),
                    "will not": ("won't",1),
                    "will not have": ("won't've",1),
                    "you had": ("you'd",1),
                    "you would": ("you'd",1),
                    "you would have": ("you'd've",1),
                    "you shall": ("you'll",1),
                    "you will": ("you'll",1),
                    "you are": ("you're",1),
                    "you are not": ("you aren't",1),
                    "you have": ("you've",1),
                    "is going to": ("will",2),
                    "for example": ("for instance",1),
                    "such as": ("like",1),
                    "however": ("but",2),
                    "whether or not": ("if",3),
                    "in order to": ("to",3),
                    "government": ("state",1),
                    "United States": ("US",2),
                    "individuals": ("persons",3),
                    "undermines": ("harms",2),
                    "action": ("act",1),
                    "obligation": ("duty",2),
                    "autonomy": ("agency",1),
                    "autonomous": ("agential",1),
                    "adjudicate": ("judge",3),
                    "suggests": ("says",1),
                    "incorrect": ("wrong",2),
                    "was not able to": ("couldn't",3),
                    "was able to": ("could",3),
                    "is not able to": ("can't",5),
                    "is able to": ("can",4),
                    "an individual": ("a person",3),
                    "the individual": ("the person",3),
                    "means that": ("means",1),
                    "we would call": ("we call",1),
                    "vote affirmative": ("vote aff",3),
                    "vote negative": ("vote neg",2),
                    "additional": ("further",2),
                    "need to": ("must",1),
                    "it is necessary to ": ("we must",5),
                    "ethical": ("moral",1),
                    "allow": ("let",1),
                    "a moral code": ("an ethic",1),
                    "a moral system": ("an ethic",2),
                    "a system of morality": ("an ethic",5),
                    "a code of morality": ("an ethic",4),
                    "technology": ("tech",3),
                    "information": ("info",2),
                    "powerful": ("potent",1),
                    "particularly": ("especially",1),
                    "weapon of mass destruction": ("WMD",2),
                    "weapons of mass destruction": ("WMDs",2),
                    "eliminates": ("takes out",2),
                    "eliminating": ("taking out",2),
                    "reducing": ("shrinking",1),
                    "reduce": ("shrink",1),
                    "decrease": ("shrink",1),
                    "decreasing": ("shrinking",1),
                    "critical": ("key",2),
                    "crucial": ("key",1),
                    "combat": ("fight",1),
                    "an affirmative": ("an aff",3),
                    "a negative": ("a neg",2),
                    "preventing": ("stopping",1),
                    "prevent": ("stop",1),
                    "a moral theory": ("the ethic",2),
                    "the moral code": ("the ethic",1),
                    "the moral system": ("the ethic",2),
                    "the system of morality": ("the ethic",5),
                    "the code of morality": ("the ethic",4),
                    "the moral theory": ("the ethic",2),
                    "demonstrate": ("show",2),
                    "show that": ("show",1),
                    "prove that": ("prove",1),
                    "utilize": ("use",2),
                    "proven": ("shown",1),
                    "minimize": ("shrink",2),
                    "maximize": ("increase",1),
                    "lead to": ("cause",1),
                    "leading to": ("causing",1),
                    "summarize": ("explain",2),
                    "results in": ("causes",1),
                    "result in": ("cause",1),
                    "given that": ("since",2),
                    "the proliferation": ("the spread",4),
                    "repercussion": ("consequence",1),
                    "to do so": ("to",2),
                    "protect": ("guard",1),
                    "important": ("crucial",1)
}

# Checks extensions
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def streamline(document, f):
    syllables_saved = 0

    for paragraph in document.paragraphs:
        underline_flag = 0
        # Detects if paragraph is a card (has underlined words)
        for run in paragraph.runs:
            if run.underline == True:
                underline_flag = 1
        # Only replaces words in non-card paragraphs
        if underline_flag == 0:
            for run in paragraph.runs:
                run_lower = run.text.lower()  # Parse to match lowercase dict
                for phrase in REPLACEMENT_DICT:
                    if phrase in run_lower:
                        # Capitalizes replacements if phrase is capitalized 
                        phrase_words = phrase.split()
                        phrase_words[0] = phrase_words[0].title() 
                        title_phrase = ' '.join(phrase_words)
                        value = REPLACEMENT_DICT[phrase]
                        # Calculates syllables saved
                        syllables_saved += (run.text.count(phrase) + 
                            run.text.count(title_phrase)) * value[1]
                        run.text = run.text.replace(phrase,value[0])
                        run.text = run.text.replace(title_phrase,
                            value[0].title())
    # TODO: Find way to re-render template after sending file to show syllables
    # saved message
    # flash('You saved ' + syllables_saved + ' syllables!','success')
    document.save(f)

@app.route('/', methods=['GET','POST'])
def index():
    if request.method == 'POST':
        file = request.files['file']
        filename = secure_filename(file.filename)
        if file:
            if not allowed_file(file.filename):
                flash('Wrong file format. Only ' + ' '.join(ALLOWED_EXTENSIONS)
                    + ' allowed.','danger')
                return render_template('index.html')
            else:
                document = Document(file)
                f = StringIO()
                streamline(document, f)
                f.seek(0)
                return send_file(f, as_attachment=True, 
                    attachment_filename='streamlined_' + filename)
        else:
            flash('No file selected.','danger')
            return render_template('index.html')
    else:
        return render_template('index.html')
