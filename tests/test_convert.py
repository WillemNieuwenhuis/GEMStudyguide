import unittest
import os
from gemguide.gemtodocx import convert2docx, convert2pdf

class TestConvert(unittest.TestCase):

    def test_xlsx2docx(self):
        ''' Not an actual test, but easy way to perform convert
        '''
        fn = 'GEM - LUND - Track 3  - Year 1 - Geographical Information Systems, Basic Course.xlsx'
        doc = 'output.docx'

        convert2docx(fn, doc)
        self.assertTrue(True)

    def test_xlsx2pdf(self):
        ''' Not an actual test, but easy way to perform convert
        '''
        fn = 'GEM - LUND - Track 3  - Year 1 - Geographical Information Systems, Basic Course.xlsx'
        pdf = 'E:/Data/studyguide/output.pdf'

        convert2pdf(fn, pdf)
        self.assertTrue(True)
