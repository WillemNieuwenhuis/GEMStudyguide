import unittest
from gemguide.gemtodocx import convert

class TestConvert(unittest.TestCase):

    def test_xlsx2docx(self):
        fn = 'GEM - LUND - Track 3  - Year 1 - Geographical Information Systems, Basic Course.xlsx'
        doc = 'output.docx'

        convert(fn, doc)
        self.assertTrue(True)

