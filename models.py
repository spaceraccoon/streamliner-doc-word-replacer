from collections import OrderedDict
from docx.enum.text import WD_COLOR
from docx.shared import Pt


# Replace this with your own entries in the form KEY: (VALUE,SYLLABLE_DIFF)
replacement_dict = {
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