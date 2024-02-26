
# -*- coding: utf-8 -*-
"""
Creates the sign off and investigation report for Asia T batch
"""


import sys

sys.path.append('Z:\\Fund_Oversight\\OVERSIGHT\\Operations\\Tools\\Processes\\')

import importlib
import pandas as pd

SignOff=importlib.import_module('SignOff')
importlib.reload(SignOff)
from SignOff import SignOff



def dailySignOff_AsiaT(ReportDate):
    
    SignOff(Batch = 'STP Asia (T)',ReportDate=ReportDate,TNABreak=False)

if __name__=='__main__':
    
    t=dailySignOff_AsiaT(pd.to_datetime('2020-04-29'))