
"""
Validation of the holding prices for North American funds using independent sources. Outline the main movers and material errors.
"""
# In[13]:
import os
import sys
import importlib
sys.path.append('Z:\\Fund_Oversight\\OVERSIGHT\\Operations\\Processes\\')
from SCDtoNAV_Fund import SCDtoNAV_Fund
import Errors
import pandas as pd
from pandas.tseries.offsets import CustomBusinessDay
from FXMatrix import SCDFX


FXRecon=importlib.import_module('FXRecon')
FilterData=importlib.import_module('FilterData')
ImportHLDG=importlib.import_module('ImportHLDG')
ImportTNA=importlib.import_module('ImportTNA')


importlib.reload(FXRecon)
importlib.reload(ImportHLDG)
importlib.reload(FilterData)
importlib.reload(ImportTNA)

from FilterData import FilterData
from ImportTNA import RBCTNA,SSTNA
from ImportHLDG import RBCHLDG,SSHLDG,CIBCHLDG
from FXRecon import FXImport,FXImpact

def DTD(ReportDate):
    RBC=pd.Series()
    RBCNOFV=pd.Series()
    #Get list of business days of length Windowreg ending last business day
    for i in pd.bdate_range(start ='2019-11-01',end ='2019-11-29'):
        try:
            DateLst = pd.bdate_range(end = i,periods=2)
    
       
              # In[14.5]:
            
            #import and convert prices from bpl files
            bplLst=[]
            for dt in DateLst:
                bpl = pd.read_csv(f'Z:\\Fund_Oversight\\OVERSIGHT\\Operations\\Daily\\{dt.strftime("%Y")}\\{dt.strftime("%Y%m")}\\{dt.strftime("%Y%m%d")}\\Inputs\\bpl_price_master_{dt.strftime("%Y-%m-%d")}.csv',converters={'sedol':str,'cusip':str})
                bpl.dropna(subset=['ml_price'],inplace=True)
                bpl['ml_namr_fair_value_factor']=bpl['ml_namr_fair_value_factor'].fillna(1)
                bpl['ml_price_NOFV']=bpl['ml_price']
                bpl['ml_price']*=bpl['ml_namr_fair_value_factor']
                bplLst.append(bpl) 
                
            # In[15]:
            
            SCDT=[]
            for dt in DateLst:
                try:
                    SCDT.append(pd.read_excel(f'Z:\\Fund_Oversight\\OVERSIGHT\\Operations\\Daily\\{dt.strftime("%Y")}\\{dt.strftime("%Y%m")}\\{dt.strftime("%Y%m%d")}\\Inputs\\SCD\\AsiaT1\\Holdings_NA_{dt.strftime("%Y%m%d")}.xlsx'))
                except:
                    raise Errors.InputMissing(f'Holdings_NA_{dt.strftime("%Y%m%d")}.xlsx')
                try:
                    #SCDT[-1]=SCDtoNAV_Fund(SCDT[-1].dropna(subset=['Portfolio']),'Portfolio')
                    SCDT[-1][['SEDOL (static)','CUSIP','ISIN']]=SCDT[-1][['SEDOL (static)','CUSIP','ISIN']].applymap(lambda x: 'ZZZZ')
                except:
                    raise Errors.InputFormatError(f'Holdings_NA_{dt.strftime("%Y%m%d")}.xlsx')
        
            
            # In[16]:
            
            for df in SCDT:
                df.drop(columns = df.columns[~df.columns.isin(['Security ID','Portfolio','Clean value PC','Balance nominal/number','SEDOL (static)','CUSIP','ISIN'])],inplace = True)
                df=df.drop(df.index[df['Clean value PC']==0],inplace=True)
                
            
            
            # In[17]:
                    
            RBCNAV = RBCTNA(DateLst[0])['TNA']
         
        
    
                
                
            RBCT = RBCHLDG(DateLst[1])
            RBCT_1=RBCHLDG(DateLst[0])
            #Remove funds that do not have a NAV at T-1 (New funds ie)
            RBCT=RBCT.loc[RBCT['FUND'].isin(RBCNAV.index)]
            RBCT_1=RBCT_1.loc[RBCT_1['FUND'].isin(RBCNAV.index)]
            
            RBCT['PCT_MV_FUND']=RBCT['MKT_VAL_FUND']/RBCT['FUND'].map(lambda x:RBCNAV.loc[x])
        
    
        
         
             # In[21]:        
            FX = SCDFX(DateLst[1],'London')
            FXT_1 = SCDFX(DateLst[0],'London')
            
            RBCReport = FilterData(RBCT,RBCT_1,SCDT[1],SCDT[0],bplLst[0].loc[bplLst[0]['id_bpl_pricing_strategy']=='CWFA'],bplLst[1].loc[bplLst[1]['id_bpl_pricing_strategy']=='MAM_GL'],FX,FXT_1,'RBC',DateLst[-1])
          
            RBCReport[0].name = DateLst[1].strftime('%Y-%m-%d')
            RBCReport[1].name = DateLst[1].strftime('%Y-%m-%d')
            RBC = pd.concat([RBC,RBCReport[0]],axis=1)
            RBCNOFV = pd.concat([RBCNOFV,RBCReport[1]],axis=1)
            
            
        except:
            pass
            print('not good')
        
        
    RBC.dropna(how='all',inplace=True)
    RBC.dropna(axis=1,how='all',inplace=True)
    RBCNOFV.dropna(how='all',inplace=True)
    RBCNOFV.dropna(axis=1,how='all',inplace=True)
    
    Diff = RBC-RBCNOFV
    
    writer = pd.ExcelWriter('Z:\\Fund_Oversight\\OVERSIGHT\\Operations\\Tools\\Config\\RBC_Analysis_WoFV.xlsx')
    RBC.to_excel(writer,sheet_name='FV')
    RBCNOFV.to_excel(writer,sheet_name='No FV')  
    Diff.to_excel(writer,sheet_name='Difference')  
    Diff.describe().to_excel(writer,sheet_name='Stats')  
    writer.save()
    
if __name__=='__main__':
    t=DTD(pd.to_datetime('2020-03-06'))