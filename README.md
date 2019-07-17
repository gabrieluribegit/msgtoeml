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

Activate it
source msgtoeml_env/bin/activate

Install requirements.txt
pip install -r requirements.txt

TODO: Add step to use modified compoundfiles and convert-outlook-msg-file

Compile with pyinstaller
pyinstaller pyinstaller --noconsole Convert\ msg\ to\ eml.py



