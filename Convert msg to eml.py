import sys, os, xml.dom.minidom

import urllib.parse

import base64

from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QLineEdit, \
QFileDialog, QLabel, QVBoxLayout, QPushButton, QPlainTextEdit, QToolTip, \
QGridLayout, QComboBox, QFrame, QGroupBox
from PyQt5.QtGui import QIcon, QFont
from PyQt5.QtCore import pyqtSlot

# Modified from https://github.com/JoshData/convert-outlook-msg-file
from outlookmsgfile import *

# from compoundfiles import reader_msg_format
#
# print(dir(reader_msg_format))

# Modified compoundfiles https://github.com/waveform-computing/compoundfiles
# Modified compoundfiles on local virtualenv:
# - Added reader_msg_format.py
# - Replaced Raise Error to self.errors and disabled print messages
# - TODO:
# - Changed output of this Class to return self.errors instead of instance:
"""     def __enter__(self):
        # return self
        ### ggg
        return self.errores"""
# - Registered new class CompoundFileReader_msg_format in __init__.py
# from check_msg_format import *

# BUG:
# Drag and Drop within app
# https://forum.qt.io/topic/104598/supporting-dragging-from-a-qwebengineview-to-other-qwidget-in-app


def process_file(fileName):
    if os.path.isfile(fileName):

        ### IF XCML/XML FILE extension
        # Trigger beautify
        if '.cxml' in fileName or '.xml' in fileName:
            dom = xml.dom.minidom.parse(fileName) # or xml.dom.minidom.parseString(xml_string)
            dom_string = dom.toprettyxml(indent="    ", encoding="utf-8") # 4 spaces
            dom_string = b'\n'.join([s for s in dom_string.splitlines() if s.strip()])

            console.clear()
            console.insertPlainText("CXML / XML file\n\n")
            console.insertPlainText("Selected ->" + fileName + "\n\n")

            # Show output on GUI
            console.insertPlainText("...beautifying...\n\n")

            fileName_cxml = fileName + "_beautified.cxml"
            with open(fileName_cxml, "wb") as f:
                f.write(dom_string)
            console.insertPlainText("Created ->" + fileName_cxml + "\n\n")
            console.insertPlainText("Please open it with a text editor")
            return

        ### MSG TO EML CONVERSION
        else:
            msg_errors = ''
            # Test if file is a well formated .msg despite file extension
            print(f'+++ test fileName {fileName}')
            msg_errors = test_msg(fileName)

            # If msg file is OK
            if not msg_errors:
                # Convert file
                print(f'+++ convert fileName {fileName}')
                convert_msg(fileName)

    # IF NOT isFile?
    else:
        self.clear()
        self.insertPlainText("ERROR:\n\nUnrecognized file format\n\n")
        return


def test_msg(fileName):
    """Test file is a well formated .msg despite file extension"""
    console.clear()
    console.insertPlainText("Selected ->" + fileName + "\n\n")
    msg_errors = load_test_format(fileName)
    # Output errors
    if msg_errors:
        console.insertPlainText("Converting file FAILED! Corrupted or not supported format\n\nERROR:\n" + msg_errors)
    print(f'test_msg {msg_errors}')
    return msg_errors


def convert_msg(fileName):
    """Convert file, save eml and open it"""
    # Show output on GUI
    console.insertPlainText("...converting...\n\n")

    # Convert .msg to .eml
    msg = load(fileName)

    # if msg:
    fileName_eml = fileName + ".eml"
    print(f'++++++ fileName_eml {fileName_eml}')
    # print(f'++++++ fileName_eml {msg}')
    with open(fileName_eml, "wb") as f:
        f.write(msg.as_bytes())


    # Check if file created
    if os.path.isfile(fileName_eml):
        console.insertPlainText("Created ->" + fileName_eml + "\n\n")
        console.insertPlainText("Opening with Mail app")

        # Open file with Mac Mail app
        open_mail = "open " + fileName_eml
        os.system(open_mail)
    # It does not return anything at the end, but opens the eml file

def process_string(dropped_string):
    """Count words and characters from string"""
    string_text = dropped_string
    string_len = str(len(string_text))
    words_len = str(len(string_text.split()))
    # Add to UI console
    console.clear()
    console.insertPlainText("String details: \n\n")
    console.insertPlainText("- Characters count: " + string_len + "\n\n")
    console.insertPlainText("- Word count: " + words_len + "\n\n")
    console.insertPlainText("- String checked: \n\n" + string_text)


class TextDrop(QPlainTextEdit):

    def __init__(self, title, parent):
        super().__init__(title, parent)

        self.setAcceptDrops(True)

    def dragEnterEvent(self, t):
        if t.mimeData().hasUrls():
            t.accept()
        elif t.mimeData().hasFormat('text/plain'):
            t.accept()
        else:
            t.ignore()

    def dropEvent(self, t):
        if t.mimeData().urls():
            for url in t.mimeData().urls():
                fileName = url.toLocalFile()

                if os.path.isfile(fileName):
                    # Process file
                    process_file(fileName)

        ### IF text/plain
        else:
            ### STRING COUNTING WORDS AND CHARS
            process_string(t.mimeData().text())

class Encode(QPlainTextEdit):

    def __init__(self, title, parent):
        super().__init__(title, parent)

        self.setAcceptDrops(True)

    def dragEnterEvent(self, e):
        console_encode.clear()
        console_decode.clear()
        if e.mimeData().hasUrls():
            e.accept()
        elif e.mimeData().hasFormat('text/plain'):
            e.accept()
        else:
            e.ignore()

    def dropEvent(self, e):
        if e.mimeData().urls():
            for url in e.mimeData().urls():
                fileName = url.toLocalFile()
                print(f'+++fileName {fileName}')
        text = e.mimeData().text()
        text_encoded = ''
        error = ''

        if combo.currentText() == 'Base64':
            try:
                text_encoded = base64.b64encode(text.encode('utf-8'))
                text_encoded = str(text_encoded)
                text_encoded = text_encoded[2:-1]
                console_encode.clear()
                console_encode.insertPlainText(text_encoded)
                print(f'{combo.currentText()} text_encoded {text_encoded}')
            except Exception as error:
                # console_decode.insertPlainText('Error: ' + str(error))
                print(f'encoding error {str(error)}')
                console_encode.clear()
                console_encode.insertPlainText("Converting string FAILED! Corrupted or not supported\n\nERROR:\n" + str(error))

        elif combo.currentText() == 'urlencode':
            try:
                text_encoded = urllib.parse.quote_plus(text)
                console_encode.clear()
                console_encode.insertPlainText(text_encoded)
                print(f'{combo.currentText()} text_encoded {text_encoded}')
            except Exception as error:
                # console_decode.insertPlainText('Error: ' + str(error))
                print(f'encoding error {str(error)}')
                console_encode.clear()
                console_encode.insertPlainText("Converting string FAILED! Corrupted or not supported\n\nERROR:\n" + str(error))


class Decode(QPlainTextEdit):

    def __init__(self, title, parent):
        super().__init__(title, parent)

        self.setAcceptDrops(True)

    def dragEnterEvent(self, d):
        console_encode.clear()
        console_decode.clear()
        # if d.mimeData().hasUrls():
        #     d.accept()
        if d.mimeData().hasFormat('text/plain'):
            d.accept()
        else:
            d.ignore()

    def dropEvent(self, d):
        text = d.mimeData().text()
        text_decoded = ''
        error = ''
        if combo.currentText() == 'Base64':
            try:
                text_decoded = base64.b64decode(text).decode('utf-8')
                console_decode.clear()
                console_decode.insertPlainText(text_decoded)
                print(f'{combo.currentText()} text_decoded {text_decoded}')
            except Exception as error:
                # console_decode.insertPlainText('Error: ' + str(error))
                print(f'decoding error {str(error)}')
                console_decode.clear()
                console_decode.insertPlainText("Converting string FAILED! Corrupted or not supported\n\nERROR:\n" + str(error))

        elif combo.currentText() == 'urlencode':
            try:
                text_decoded = urllib.parse.unquote_plus(text)
                console_decode.clear()
                console_decode.insertPlainText(text_decoded)
                print(f'{combo.currentText()} text_decoded {text_decoded}')
            except Exception as error:
                print(f'decoding error {str(error)}')
                console_decode.clear()
                console_decode.insertPlainText("Converting string FAILED! Corrupted or not supported\n\nERROR:\n" + str(error))


class App(QWidget):

    def __init__(self):
        super().__init__()
        self.title = 'Convert .msg to .eml'
        self.left = 10
        self.top = 10
        self.width = 500
        self.height = 800

        self.initUI()

    def initUI(self):
        """"UI definition"""
        global console   # Available for different methods
        global console_encode_label
        global console_decode_label
        global console_encode
        global console_decode
        global combo_label
        global combo

        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        self.layout = QVBoxLayout()

        # Label
        # description = QLabel("Convert .msg to .eml\nBeautify CXML/XML\nCount words and characters")
        # description.resize(10,100)
        # instructions = QLabel("Instructions\n\n")
        # instructions.setToolTip('Convert .msg to .eml:\nUse Button or Drop a file\nIt creates an .eml and open with Mail app\n\nBeutify CXML/XML:\nDrop a .cxml or .xml file, \nIt creates a beautified file on same folder\n\nCount words and characters:\nSelect and drop text from any application')
        # instructions.resize(10,100)

        groupbox = QGroupBox("Features")
        # groupbox.setFont(QFont("Sanserif", 15))
        self.layout.addWidget(groupbox)

        vbox = QVBoxLayout()

        msg_to_eml = QLabel("Convert .msg to .eml")
        msg_to_eml.setToolTip('Use Button or Drop a file, it creates an .eml and open with Mail app')
        msg_to_eml.resize(10,100)
        # msg_to_eml.setMouseTracking(True)


        beautify_xml = QLabel("Beautify CXML/XML")
        beautify_xml.setToolTip('Drop a .cxml or .xml file, it creates a beautified file on same folder')
        beautify_xml.resize(10,100)

        count_chars = QLabel("Count words and characters")
        count_chars.setToolTip('Select and drop text from any application')
        count_chars.resize(10,100)

        combo_label_description = QLabel('Convert Base64 and urlencode', self)
        combo_label_description.setToolTip('Select a conversion option and drop a string to encode or decode')
        combo_label_description.resize(10,100)

        # Button
        button = QPushButton('Select .msg', self)
        button.setToolTip('Select an .msg file')
        button.move(10,100)
        button.clicked.connect(self.openFileNameDialog)

        # Console msg, cxml, string count chars
        console = TextDrop("Drop a .msg, .cxml or xml file or selected text from other apps", self)
        console.resize(400,400)

        vbox.addWidget(msg_to_eml)
        vbox.addWidget(beautify_xml)
        vbox.addWidget(count_chars)
        vbox.addWidget(combo_label_description)
        # vbox.addWidget(button)
        # vbox.addWidget(console)
        groupbox.setLayout(vbox)

        # Combo


        combo_label = QLabel('Select Conversion', self)
        combo_label.setToolTip('Select an encoding option')
        combo_label.resize(10,100)

        combo = QComboBox(self)
        combo.addItem("Base64")
        combo.addItem("urlencode")
        combo.activated[str].connect(self.onActivated)


        groupbox_conversion = QGroupBox("Conversion")
        # groupbox.setFont(QFont("Sanserif", 15))


        grid = QGridLayout()

        grid.addWidget(combo_label, 1, 0)
        grid.addWidget(combo, 1, 1, 1, 3)

        groupbox_conversion.setLayout(grid)

        # Console urlEncode
        console_encode_label = QLabel("Encode")
        console_encode = Encode("Drop a string to be encoded", self)
        console_encode.resize(400,300)

        # Console urlDecode
        console_decode_label = QLabel("Decode")
        console_decode = Decode("Drop a encoded string to be decoded", self)
        console_decode.resize(400,300)

        # Change version with every release
        version = "0.7"
        version_text = "Version " + version
        version = QLabel(version_text)


        ### Add UI objects
        ### GRID test
        # grid_layout = QGridLayout()
        # grid_layout.addWidget(msg_to_eml, 1, 0)
        # grid_layout.addWidget(beautify_xml, 2, 0)
        # grid_layout.addWidget(count_chars, 3, 0)
        # grid_layout.addWidget(combo_label_description, 4, 0)
        # grid_layout.addWidget(groupbox)

        ### GROUPBOX test
        # groupbox = QGroupBox("GroupBox Example")
        # groupbox.setFont(QFont("Sanserif", 15))
        # vbox = QVBoxLayout()
        # msg_to_eml = QLabel("Convert .msg to .eml")
        # # msg_to_eml.setToolTip('Use Button or Drop a file, it creates an .eml and open with Mail app')
        # # msg_to_eml.resize(10,100)
        #
        # beautify_xml = QLabel("Beautify CXML/XML")
        # # beautify_xml.setToolTip('Drop a .cxml or .xml file, it creates a beautified file on same folder')
        # # beautify_xml.resize(10,100)
        #
        # count_chars = QLabel("Count words and characters")
        # # count_chars.setToolTip('Select and drop text from any application')
        # # count_chars.resize(10,100)
        #
        # combo_label_description = QLabel('Convert Base64 and urlencode', self)
        # # combo_label_description.setToolTip('Select a conversion option and drop a string to encode or decode')
        # # combo_label_description.resize(10,100)
        # # groupbox.setLayout(grid_layout)
        # self.setLayout(grid_layout)






        # self.layout.addWidget(msg_to_eml)
        # self.layout.addWidget(beautify_xml)
        # self.layout.addWidget(count_chars)
        # self.layout.addWidget(combo_label_description)
        self.layout.addWidget(button)
        self.layout.addWidget(console)

        # self.layout.addWidget(combo_label)
        # self.layout.addWidget(combo)
        self.layout.addWidget(groupbox_conversion)

        self.layout.addWidget(console_encode_label)
        self.layout.addWidget(console_encode)
        self.layout.addWidget(console_decode_label)
        self.layout.addWidget(console_decode)

        self.layout.addWidget(version)

        # self.setFixedSize(self.width, self.height)
        self.setFixedSize(self.size())
        self.setWindowTitle("Email Format Conversion")
        self.setLayout(self.layout)
        # self.setLayout(grid_layout)

        self.show()

    def mouseMoveEvent(self, event):
        print("On Hover")

    def onActivated(self, text):
    #     combo_label
        console_encode_label.setText("Encode")
        console_decode_label.setText("Decode ")
        # combo_label.adjustSize()

    def openFileNameDialog(self):
        console.clear()
        console.insertPlainText("Please select a .msg\n\n")

        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog

        downloads_folder = os.getenv('HOME') + "/Downloads"
        fileName, _ = QFileDialog.getOpenFileName(self, 'Select a file', downloads_folder, "Outlook Files (*.msg)")

        if os.path.isfile(fileName):

            # Process file
            process_file(fileName)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
