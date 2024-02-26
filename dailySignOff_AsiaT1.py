# -*- coding: utf-8 -*-
"""
Creates the SCD vs NAV Agent daily sign off and investigation reports for Asia T+1 funds. 

"""


import sys

sys.path.append('Z:\\Fund_Oversight\\OVERSIGHT\\Operations\\Tools\\Processes\\')

import importlib
import pandas as pd

SignOff=importlib.import_module('SignOff')
importlib.reload(SignOff)
from SignOff import SignOff

def dailySignOff_AsiaT1(ReportDate):

    SignOff(Batch = 'STP Asia (T+1)',ReportDate=ReportDate,TNABreak=False)
  
    
if __name__=='__main__':
    t=dailySignOff_AsiaT1(pd.to_datetime('2020-09-09'))