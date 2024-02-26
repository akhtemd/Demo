# -*- coding: utf-8 -*-
"""
Creates the SCD vs NAV Agent daily sign off and investigation reports for Asia FOF funds. 
"""

import sys

sys.path.append('Z:\\Fund_Oversight\\OVERSIGHT\\Operations\\Tools\\Processes\\')

import importlib
import pandas as pd

SignOff=importlib.import_module('SignOff')
importlib.reload(SignOff)
from SignOff import SignOff


def dailySignOff_FOF(ReportDate):

  
    SignOff(Batch = 'STP Asia (T+2-FOF)',ReportDate=ReportDate,TNABreak=False)
    

if __name__=='__main__':
    dailySignOff_FOF(pd.to_datetime('2020-04-21'))