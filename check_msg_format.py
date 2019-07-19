# This module converts a Microsoft Outlook .msg file into
# a MIME message that can be loaded by most email programs
# or inspected in a text editor.
#
# This script relies on the Python package compoundfiles
# for reading the .msg container format.
#
# Referencecs:
#
# https://msdn.microsoft.com/en-us/library/cc463912.aspx
# https://msdn.microsoft.com/en-us/library/cc463900(v=exchg.80).aspx
# https://msdn.microsoft.com/en-us/library/ee157583(v=exchg.80).aspx
# https://blogs.msdn.microsoft.com/openspecification/2009/11/06/msg-file-format-part-1/

import re
import sys

from functools import reduce

import urllib.parse
import email.message, email.parser, email.policy
from email.utils import parsedate_to_datetime, formatdate, formataddr

import compoundfiles


def load_test_format(filename_or_stream):
  with compoundfiles.CompoundFileReaderTestMsgFormat(filename_or_stream) as doc_test:
    # doc_test.rtf_attachments = 0   # ggg to check what it does exactly
    return doc_test
