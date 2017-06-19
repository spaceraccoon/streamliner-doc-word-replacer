# Streamliner

Flask webapp that replaces words in uploaded docx files and outputs the new version. Useful if you want to quickly replace a dictionary of words. Streamliner was built for a debating website to reduce the number of syllables in a speech, but can be easily repurposed.

![Screenshot](screenshot.png "Screenshot")

## Setup

### Dependencies

Streamliner was built on Python 2.7 and Flask. It uses the python-docx library, which can be easily installed with pip using `pip install python-docx`.

### Quickstart

1. Clone the project and enter the directory
2. `export FLASK_APP=app.py`
3. `flask run`

## Deploy

An integrated, customized version of Streamliner is available that includes logging, Javascript-based feedback (e.g. 'You saved X syllables!'), and more. Contact me at [my website](https://www.pinewebarchitects.com/) to request it.
