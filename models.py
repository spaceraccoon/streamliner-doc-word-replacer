from collections import OrderedDict
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import Pt

# Replace this with your own entries in the form old: (new,syllable_diff) 
replacement_dict = {
}


REPLACEMENT_DICT = OrderedDict(sorted(replacement_dict.items(), 
    key=lambda t: len(t[0]), reverse=True))


def can_linedown(string):
    replacement = REPLACEMENT_DICT[string][0]
    replacement = filter(str.isalpha, replacement)
    length = len(replacement)
    i = 0
    output = True
    while i < length:
        try:
            index = string.index(replacement[i])
            string = string[index+1:]
            i += 1
        except:
            output = False
            return output

    return output


def split_list(string, splitter, capitalize, replace):
    output = []
    syllables_saved = REPLACEMENT_DICT[splitter][1]
    if replace:
        replacement = REPLACEMENT_DICT[splitter][0]
    else:
        replacement = splitter
    if capitalize:
        substrings = string.split(splitter.capitalize())
    else:
        substrings = string.split(splitter)

    for substring in substrings[:-1]:
        output.append((substring, False, 0))
        if capitalize:
            output.append((replacement.capitalize(), True, syllables_saved))
        else:
            output.append((replacement, True, syllables_saved))
    output.append((substrings[-1], False, 0))
    return output


def replace_paragraph(paragraph, action):
    p_text = paragraph.text
    p_list = [(p_text, False, 0)]
    replace = True
    for phrase in REPLACEMENT_DICT:
        if action == 'linedown':
            replace = False
            if not can_linedown(phrase):
                continue
        for index, substring in enumerate(p_list):
            if phrase in substring[0] and not substring[1]:
                p_list[index:index+1] = split_list(substring[0],phrase, False, replace)
            if phrase.capitalize() in substring[0] and not substring[1]:
                p_list[index:index+1] = split_list(substring[0], phrase, True, replace)
    
    paragraph.clear()
    replacements = 0
    syllables_saved = 0

    for element in p_list:
        substring = element[0]
        was_replaced = element[1]
        syllables_saved += element[2]
        if was_replaced:
            replacements += 1
            if action == 'highlight':
                paragraph.add_run(substring).font.highlight_color = WD_COLOR_INDEX.YELLOW
            elif action == 'linedown':
                new_substring = ''
                replacement = REPLACEMENT_DICT[substring.lower()][0]
                replacement = filter(str.isalpha, replacement)
                length = len(replacement)
                i = 0
                while i < length:
                    index = substring.lower().index(replacement[i])
                    paragraph.add_run(substring[:index]).font.size = Pt(8)
                    paragraph.add_run(substring[index])
                    substring = substring[index+1:]
                    i += 1
                paragraph.add_run(substring).font.size = Pt(8)
            else:
                paragraph.add_run(substring)
        else:
            paragraph.add_run(substring)

    return replacements, syllables_saved
