from collections import OrderedDict
from docx.enum.text import WD_COLOR
from docx.shared import Pt


# Replace this with your own entries in the form old: (new,syllable_diff)
replacement_dict = {
                    "because": ("since", 1),
                    "insofar as": ("since", 3),
                    "in a world where": ("since", 3),
                    "therefore": ("thus", 1),
                    "ergo": ("thus", 1),
                    "which means that": ("so", 2),
                    "that means that": ("so", 2),
                    "which means": ("so", 1),
                    "resolution": ("topic", 2),
                    "the affirmative": ("the aff", 3),
                    "the negative": ("the neg", 2),
                    "my opponent": ("they", 3),
                    "my value is": ("I value", 1),
                    "the criterion": ("the standard", 2),
                    "the value criterion": ("the standard", 4),
                    "tells you": ("says", 1),
                    "today's resolution": ("", 6),
                    "in today's round": ("", 4),
                    "in this round": ("", 3),
                    "in today's debate round": ("", 6),
                    "in this debate round": ("", 5),
                    "explains": ("writes", 1),
                    "firstly": ("first", 1),
                    "secondly": ("second", 1),
                    "thirdly": ("third", 1),
                    "fourthly": ("fourth", 1),
                    "fifthly": ("fifth", 1),
                    "sixthly": ("sixth", 1),
                    "sevently": ("seventh", 1),
                    "eightly": ("eight", 1),
                    "ninthly ": ("ninth", 1),
                    "respond to": ("answer", 2),
                    "responds to": ("answers", 2),
                    "response": ("answer", 1),
                    "argument": ("reason", 1),
                    "implication": ("impact", 2),
                    "justification": ("warrant", 3),
                    "reason as to why": ("reason", 3),
                    "reasons as to why": ("reasons", 3),
                    "basically": ("", 4),
                    "inherently": ("", 4),
                    "fundamentally": ("", 5),
                    "ultimately": ("", 4),
                    "essentially": ("", 4),
                    "obviously": ("", 4),
                    "intrinsically": ("", 5),
                    "furthermore": ("and", 2),
                    "additionally": ("and", 4),
                    "in addition": ("and", 3),
                    "moreover": ("and", 2),
                    "can not": ("can't", 1),
                    "are not": ("aren't", 1),
                    "cannot": ("can't", 1),
                    "do not": ("don't", 1),
                    "he had": ("he'd", 1),
                    "he would": ("he'd", 1),
                    "he shall": ("he'll", 1),
                    "he will": ("he'll", 1),
                    "he has": ("he's", 1),
                    "he is": ("he's", 1),
                    "how did": ("how'd", 1),
                    "how would": ("how'd", 1),
                    "how has": ("how's", 1),
                    "how is": ("how's", 1),
                    "i had": ("I'd", 1),
                    "i would": ("I'd", 1),
                    "i shall": ("I'll", 1),
                    "i will": ("I'll", 1),
                    "i am": ("I'm", 1),
                    "i have": ("I've", 1),
                    "it would": ("it'd", 1),
                    "it shall": ("it'll", 1),
                    "it will": ("it'll", 1),
                    "it has": ("it's", 1),
                    "it is": ("it's", 1),
                    "of the clock": ("o'clock", 1),
                    "somebody is": ("somebody's", 1),
                    "someone had": ("someone'd", 1),
                    "someone would": ("someone'd", 1),
                    "someone shall": ("someone'll", 1),
                    "someone will": ("someone'll", 1),
                    "someone has": ("someone's", 1),
                    "someone is": ("someone's", 1),
                    "something had": ("something'd", 1),
                    "something would": ("something'd", 1),
                    "something shall": ("something'll", 1),
                    "something will": ("something'll", 1),
                    "something has": ("something's", 1),
                    "something is": ("something's", 1),
                    "that is": ("that's", 1),
                    "that has": ("that's", 1),
                    "that would": ("that'd", 1),
                    "that had": ("that'd", 1),
                    "there had": ("there'd", 1),
                    "there would": ("there'd", 1),
                    "there is": ("there's", 1),
                    "there has": ("there's", 1),
                    "they had": ("they'd", 1),
                    "they would": ("they'd", 1),
                    "they shall / they will": ("they'll", 1),
                    "they are": ("they're", 1),
                    "they have": ("they've", 1),
                    "we had": ("we'd", 1),
                    "we would": ("we'd", 1),
                    "we will": ("we'll", 1),
                    "we are": ("we're", 1),
                    "we have": ("we've", 1),
                    "were not": ("weren't", 1),
                    "what did": ("what'd", 1),
                    "what shall": ("what'll", 1),
                    "what will": ("what'll", 1),
                    "what is": ("what's", 1),
                    "what has": ("what's", 1),
                    "when is": ("when's", 1),
                    "when has": ("when's", 1),
                    "where did": ("where'd", 1),
                    "where is": ("where's", 1),
                    "where has": ("where's", 1),
                    "where have": ("where've", 1),
                    "who would": ("who'd", 1),
                    "who had": ("who'd", 1),
                    "who did": ("who'd", 1),
                    "who shall": ("who'll", 1),
                    "who will": ("who'll", 1),
                    "who are": ("who're", 1),
                    "who has": ("who's", 1),
                    "who is": ("who's", 1),
                    "who have": ("who've", 1),
                    "why did": ("why'd", 1),
                    "why has": ("why's", 1),
                    "why is": ("why's", 1),
                    "will not": ("won't", 1),
                    "will not have": ("won't've", 1),
                    "you had": ("you'd", 1),
                    "you would": ("you'd", 1),
                    "you would have": ("you'd've", 1),
                    "you shall": ("you'll", 1),
                    "you will": ("you'll", 1),
                    "you are": ("you're", 1),
                    "you are not": ("you aren't", 1),
                    "you have": ("you've", 1),
                    "is going to": ("will", 2),
                    "for example": ("for instance", 1),
                    "such as": ("like", 1),
                    "however": ("but", 2),
                    "whether or not": ("if", 3),
                    "in order to": ("to", 3),
                    "government": ("state", 1),
                    "united states": ("US", 2),
                    "individuals": ("persons", 3),
                    "undermines": ("harms", 2),
                    "action": ("act", 1),
                    "obligation": ("duty", 2),
                    "autonomy": ("agency", 1),
                    "autonomous": ("agential", 1),
                    "adjudicate": ("judge", 3),
                    "suggests": ("says", 1),
                    "incorrect": ("wrong", 2),
                    "was not able to": ("couldn't", 3),
                    "was able to": ("could", 3),
                    "is not able to": ("can't", 5),
                    "is able to": ("can", 4),
                    "an individual": ("a person", 3),
                    "the individual": ("the person", 3),
                    "means that": ("means", 1),
                    "we would call": ("we call", 1),
                    "vote affirmative": ("vote aff", 3),
                    "vote negative": ("vote neg", 2),
                    "additional": ("further", 2),
                    "need to": ("must", 1),
                    "it is necessary to ": ("we must", 5),
                    "ethical": ("moral", 1),
                    "allow": ("let", 1),
                    "a moral code": ("an ethic", 1),
                    "a moral system": ("an ethic", 2),
                    "a system of morality": ("an ethic", 5),
                    "a code of morality": ("an ethic", 4),
                    "technology": ("tech", 3),
                    "information": ("info", 2),
                    "powerful": ("potent", 1),
                    "particularly": ("especially", 1),
                    "weapon of mass destruction": ("WMD", 2),
                    "weapons of mass destruction": ("WMDs", 2),
                    "eliminates": ("takes out", 2),
                    "eliminating": ("taking out", 2),
                    "reducing": ("shrinking", 1),
                    "reduce": ("shrink", 1),
                    "decrease": ("shrink", 1),
                    "decreasing": ("shrinking", 1),
                    "critical": ("key", 2),
                    "crucial": ("key", 1),
                    "combat": ("fight", 1),
                    "an affirmative": ("an aff", 3),
                    "a negative": ("a neg", 2),
                    "preventing": ("stopping", 1),
                    "prevent": ("stop", 1),
                    "a moral theory": ("the ethic", 2),
                    "the moral code": ("the ethic", 1),
                    "the moral system": ("the ethic", 2),
                    "the system of morality": ("the ethic", 5),
                    "the code of morality": ("the ethic", 4),
                    "the moral theory": ("the ethic", 2),
                    "demonstrate": ("show", 2),
                    "show that": ("show", 1),
                    "prove that": ("prove", 1),
                    "utilize": ("use", 2),
                    "proven": ("shown", 1),
                    "minimize": ("shrink", 2),
                    "maximize": ("increase", 1),
                    "lead to": ("cause", 1),
                    "leading to": ("causing", 1),
                    "summarize": ("explain", 2),
                    "results in": ("causes", 1),
                    "result in": ("cause", 1),
                    "given that": ("since", 2),
                    "the proliferation": ("the spread", 4),
                    "repercussion": ("consequence", 1),
                    "to do so": ("to", 2),
                    "protect": ("guard", 1),
                    "important": ("crucial", 1)
}


def linedown(string):
    '''Returns lined down value of string, otherwise returns False

    A "linedown" is when a all of a value's characters appear as consecutive
    but not necessarily contiguous characters in a key.
    This is useful for debate evidence cards, which do not allow modification
    of text.
    '''
    new_replacement = ''
    replacement = REPLACEMENT_DICT[string][0]
    replacement = filter(str.isalpha, replacement)
    length = len(replacement)
    i = 0
    while i < length:
        try:
            index = string.index(replacement[i])
            new_replacement += replacement[i]
            string = string[index+1:]
            i += 1
        except:
            return False
    return new_replacement


def add_set_run(paragraph, run, text, font, style):
    '''Add run to paragraph and set styles according to Word precedence rules.

    Limited to a few properties for optimization.
    '''
    run = paragraph.add_run()
    run.add_text(text)
    if font.bold is False:
        run.font.bold = False
    elif font.bold:
        run.font.bold = True
    if font.underline is False:
        run.font.underline = False
    elif font.underline:
        run.font.underline = True
    if font.italic is False:
        run.font.italic = False
    elif font.italic:
        run.font.italic = True
    if font.size:
        run.font.size = font.size
    if font.name:
        run.font.name = font.name
    run.style = style

    return run


def split_list(string, splitter, is_card, capitalize):
    '''Splits substring into separate elements according to the key.

    These elements will be later packaged into separate runs.
    '''
    output = []
    font, style = string[3], string[4]
    syllables_saved = REPLACEMENT_DICT[splitter][1]

    replacement = splitter if is_card else REPLACEMENT_DICT[splitter][0]
    replacement = replacement.capitalize() if capitalize else replacement

    substrings = string[0].split(splitter)
    for substring in substrings[:-1]:
        output.append((substring, False, 0, font, style))
        output.append((replacement, True, syllables_saved, font, style))
    output.append((substrings[-1], False, 0, font, style))

    return output


def replace_paragraph(paragraph, is_card, mark):
    '''Replaces and formats paragraph, returning replacements and
    syllables_saved.

    Each ele in p_list is formatted as (TEXT, REPLACE_FLAG, SYLLABLES_SAVED,
    FONT, STYLE)
    '''
    p_list = []
    replacements = 0
    syllables_saved = 0

    for run in paragraph.runs:
        p_list.append((run.text, False, 0, run.font, run.style))
    p_lower = paragraph.text.lower()
    for k in REPLACEMENT_DICT:
        # Moves on to next phrase if not in paragraph text or can linedown
        # for a card paragraph
        if (k not in p_lower) or (is_card and k not in LINEDOWN_DICT):
            continue

        for index, substring in enumerate(p_list):
            if not substring[1]:    # Make sure text has not yet been replaced
                # Replaces element with unpacked list of elements, split by
                # the key
                if is_card and (substring[4].font.underline is None or
                                substring[3].underline is False):
                    continue
                if k in substring[0]:
                    p_list[index:index+1] = split_list(substring, k, is_card,
                                                       False)
                if k.capitalize() in substring[0]:
                    p_list[index:index+1] = split_list(substring, k, is_card,
                                                       True)

    # Clear paragraph and insert new formatted runs
    paragraph.clear()
    for ele in p_list:
        substring, was_replaced, s, font, style = [e for e in ele]
        if was_replaced:
            replacements += 1
            syllables_saved += s
            if is_card:
                # Breaks text further into characters to linedown; only
                # characters that are not in the value are downsized
                replacement = LINEDOWN_DICT[substring.lower()]
                for char in replacement:
                    index = substring.lower().index(char)
                    run = add_set_run(paragraph, run, substring[:index], font,
                                      style)
                    run.font.size = Pt(5)
                    run = add_set_run(paragraph, run, substring[index], font,
                                      style)
                    if mark:
                        run.font.highlight_color = WD_COLOR.TURQUOISE
                    substring = substring[index+1:]

                run = add_set_run(paragraph, run, substring, font, style)
                run.font.size = Pt(5)
                continue
            else:
                run = add_set_run(paragraph, run, substring, font, style)
                if mark:
                    run.font.highlight_color = WD_COLOR.TURQUOISE
                continue

        add_set_run(paragraph, run, substring, font, style)

    return replacements, syllables_saved


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


# Initializes dict constants
REPLACEMENT_DICT = OrderedDict(sorted(replacement_dict.items(),
                               key=lambda t: len(t[0]), reverse=True))

linedown_dict = {}

for k in REPLACEMENT_DICT.iterkeys():
    v = linedown(k)
    if v:
        linedown_dict[k] = v

LINEDOWN_DICT = OrderedDict(sorted(linedown_dict.items(),
                            key=lambda t: len(t[0]), reverse=True))