# msgtoeml
Front end for converting Outlook formatted emails .msg to .eml

Testing other features:
- Count words and characters by dropping selected text from other apps
- Beautifying .cxml / .xml files when dropping these files

## Setup for Dev
Create virtual env

Activate it
`source msgtoeml_env/bin/activate`

Install requirements.txt
`pip3 install requirements.txt`

Compile with pyinstaller
`pyinstaller convert_msg_to_eml.py`



