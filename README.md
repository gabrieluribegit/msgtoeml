# msgtoeml
Front end for converting Outlook formatted emails .msg to .eml

Testing other features:
- Count words and characters by dropping selected text from other apps
- Beautifying .cxml / .xml files when dropping these files
- Convert Base64 and urlencode

## Setup for Dev

Install Python 3.7.3 with Homebrew:
brew install https://raw.githubusercontent.com/Homebrew/homebrew-core/c24d6bcd4795e9263617e49f340cc61498eea5d3/Formula/python.rb

Install virtualenv
pip install virtualenv

Create virtual env
virtualenv -p python3 msgtoeml_env
=======
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

pip install -r requirements.txt

TODO: Add step to use modified compoundfiles and convert-outlook-msg-file

Compile with pyinstaller
pyinstaller pyinstaller --noconsole Convert\ msg\ to\ eml.py
=======

`pip3 install requirements.txt`

Compile with pyinstaller:

`pyinstaller --noconsole Convert\ msg\ to\ eml.py`

Run the created "Convert msg to eml.app", by double click.

