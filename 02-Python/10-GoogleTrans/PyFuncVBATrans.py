# /usr/bin/env python
# coding=utf-8
'''
This Python Code is to define a Function for VBA to call for translate with Google.
ver 1.0, 31-Jul-2021, Author: MaYS

# 2021-08-02： 
# The UDF GoogleTransPyFun() defined in PyFuncVBATrans.py works as a UDF, with xlPython. 
# xlwing not work with Err: "Could not activate Python COM server, hr = -2147221164 1000" )

To be Updated in Next Version.
1. List the full ISO 139 Language Codes, and automatically correct the Language Code if Languge Code is not correct.
2. Use RegEx to find the Short Description for Language Detection
'''

import os
import re
import string
import sys
import time
import math

# import xlwings as xw
# xlwings does not work, use xlpython instead.
from xlpython import *

from google_trans_new import google_translator

translator = google_translator()
detector = google_translator()

# #xlwings does not work, use xlpython instead.
# @xw.func

@xlfunc
def GoogleTransPyFunc(unTranslatedText: str, language_Code_from: str, language_Code_to: str):
# Variables in Function shall not be named as form of _unTranslatedText started with "_".
# Import Function will encounted with Error: VBA Runtime Error '1004': Method 'MacroOptions' of object'_Application' failed

    global textTranslated
    languageCode_to = 'en'
    languageCode_from = ''

    languageCode_to = language_Code_to
    if not language_Code_to or len(language_Code_to) > 2:
        languageCode_to = 'en'

    languageCode_from = language_Code_from
    if not language_Code_from:
        '''
        Reserved for Detection with "Short Description"
        textforLanguageDetection = re.findall('short description.{500}',
                  _subjects_text[i], re.DOTALL|re.MULTILINE)
        '''
        pass
        if len(unTranslatedText) < 1000:
            languageCode_from = detector.detect(unTranslatedText[: -1])

        textforDetection = unTranslatedText[: 200]
        languageCode_from = detector.detect(textforDetection)

    if len(unTranslatedText) < 4950:
        textTranslated = translator.translate(unTranslatedText,
                                              lang_tgt=languageCode_to, lang_src=languageCode_from)

    textTranslated = translator.translate(unTranslatedText[: 4950],
                                          lang_tgt=languageCode_to, lang_src=languageCode_from)

    time.sleep(0.5)
    # wait for at least 0.5 seconds for each request, avoid to be blocked
    # The VBA Script shall wait for even longer time between each Request.

    return textTranslated
