# msgtoeml
Front end for converting Outlook formatted emails .msg to .eml

Testing other features:
- Count words and characters by dropping selected text from other apps
- Beautifying .cxml / .xml files when dropping these files

## Setup for Dev
Requires to install Mac xcode, Brew and Python3.7

xcode:

`xcode-select --install`

Brew:

`/usr/bin/ruby -e "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/master/install)"`

Install Python 3.7 with Brew with the Terminal:

`brew install python3`

Install a virtual environment (virtualenv):

`sudo pip3 install virtualenv`

Create virtual environment for this specific app using Python 3.7:

`virtualenv -p python3.7 msgtoeml_env`

Activate the virtual environment:

`source msgtoeml_env/bin/activate`

Install requirements.txt

`pip3 install requirements.txt`

Compile with pyinstaller

`pyinstaller --noconsole Convert\ msg\ to\ eml.py`

Run the created "Convert msg to eml.app", by double click.

