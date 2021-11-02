import pandas as pd
import numpy as np
from gemguide.constants import GEMADMINKEYS, GEMCONTENTKEYS

class FromExcel:
    XLSADMINKEYS = ['Start date & end date in 2021 or 2022',
            'Credits (ECTS, 28 hours per ECTS)', 'Coordinator(s)', 'Key words']
    XLSCONTENTKEYS = ['Description', 'Learning outcomes', 'Content',
            'Entry requirements (if any)',
            'Teaching and learning approach']

    def __init__(self, fn : str):
        self.filename = fn

        self.workbook = pd.read_excel(fn, sheet_name=None)  # read all sheets
        self.course = self.workbook['1 Course template']
        allocSheet = self.workbook['2 Time allocation']
        testplanSheet = self.workbook['3 Test plan']
        self.extractAllocation(allocSheet)
        self.testplan = self.workbook['3 Test plan']
        self.extractAssessment(testplanSheet)

    def extractAllocation(self, sheet) -> None:
        begin = sheet == 'Type of activity'
        ix_begin = np.where(begin == True)
        row_begin = ix_begin[0][0]
        col_begin = ix_begin[1][0]

        end = sheet == 'Sum of hours for the course'
        ix_end = np.where(end == True)
        row_end = ix_end[0][0]

        alloc = sheet.iloc[row_begin+1:row_end,col_begin:col_begin+2].values # list of lists
        self.allocation = {}
        for r in alloc:
            self.allocation.update({r[0] : str(r[1])})

    def extractAssessment(self, sheet) -> None:
        begin = sheet == 'Test type (descriptive)'
        ix_begin = np.where(begin == True)
        row_begin = ix_begin[0][0]
        col_begin = ix_begin[1][0]

        # extract the test info and transpose it
        assess = sheet.iloc[row_begin:row_begin+3, col_begin:].T
        assess.columns = assess.iloc[0]
        assess = assess[1:]
        wgt_col = list(assess.columns)[1]   # select the weight column
        assess = assess.fillna(0)
        assess[wgt_col] = assess[wgt_col] * 100
        assess[wgt_col] = pd.to_numeric(assess[wgt_col]).astype(int)
        self.assessment = assess
    
    def getCourseItem(self, item : str) ->str:
        key = self.getKeyname(item)
        it = self.course[self.course['Item'] == key]
        val = it[it.keys()[1]]
        return val.values[0]

    def getKeyname(self, key : str) ->str:
        if key in GEMADMINKEYS:
            ix = GEMADMINKEYS.index(key)
            return self.XLSADMINKEYS[ix]
        if key in GEMCONTENTKEYS:
            ix = GEMCONTENTKEYS.index(key)
            return self.XLSCONTENTKEYS[ix]
        
        return key  # no translation


