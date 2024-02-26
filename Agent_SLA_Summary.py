# -*- coding: utf-8 -*-
"""
Created on Thu Jul 30 14:31:52 2020

@author: laguema
"""
import sqlite3
import pandas as pd
import getpass
import importlib
import datetime
import os
from pandas.tseries.offsets import CustomBusinessDay
from pandas.tseries.offsets import BDay
ImportTNA=importlib.import_module('ImportTNA')
importlib.reload(ImportTNA)
from ImportTNA import FundList


def Agent_SLA_Summary(ReportDate):
    
    
    FML = FundList(ReportDate)
    sql ="SELECT * FROM NAV"
    holidays = pd.read_csv('Z:\\Fund_Oversight\\OVERSIGHT\\List of Funds\\Holidays.csv',infer_datetime_format =True,squeeze = True,index_col = 'Calendar ID')
    holidays = pd.to_datetime(holidays)
    
    conn=sqlite3.connect(f'C:\\Users\\{getpass.getuser().capitalize()}\\Documents\\Oversight_PROD.db',timeout = 150)
    NAV=pd.read_sql(sql,conn,index_col = ['NAV_AGENT','CLASS_ID'],parse_dates={'LAST_MODIFIED': 's'})
    conn.close()
    NAV=NAV.loc[NAV.index.get_level_values('CLASS_ID').isin(FML['Fund and Series'])]
    
    
    DateLst=pd.date_range(end =ReportDate,periods=31)
    NAV_Summary=pd.DataFrame(columns =[x.strftime("%Y-%m-%d") for x in DateLst],index=set(NAV.index))
    NAV_Summary.index.names=['NAV_AGENT','CLASS_ID']
    
    for date in DateLst:
        
        NAV_=NAV.loc[NAV['DATE']==date.strftime("%Y-%m-%d")]
        NAV_=NAV_.loc[~NAV_.index.duplicated()]
        FML['DAYS']=date+BDay(1) 
        
        #Get different set of business day by country only considering holiday on T-1
        for country in set(holidays[holidays.isin([date+BDay(1)])].index):
            
            Holist = CustomBusinessDay(holidays=holidays[country])
            FML.loc[FML['Bank Holiday']==country,'DAYS']=pd.date_range(start = date+BDay(1),periods=1,freq = Holist)[0]
                    

        FML.loc[FML['NAV Agent']=='CitiLux','DAYS']=date
        #FML.loc[FML['NAV Agent']=='SSB','DAYS']=date
        FML.loc[FML['STP Batch Run']=='STP Asia (T)','DAYS']=date
        
        FML['NAV expected delivery datetime']=FML[['DAYS','NAV expected delivery time']].apply(lambda x:datetime.datetime.combine(x[0],x[1]),axis=1)
      
        NAV_=NAV_.loc[list(map(lambda y:y in [''.join(fund) for fund in zip(FML['NAV Agent'],FML['Fund and Series'])],[''.join(x) for x in NAV_.index]))]
       
        NAV_['Expected']=pd.to_datetime([FML.loc[(FML['Fund and Series']==x) & (FML['NAV Agent']==agent),'NAV expected delivery datetime'].item() for agent,x in NAV_.index])
        NAV_['Segment']=[FML.loc[(FML['Fund and Series']==x) & (FML['NAV Agent']==agent),'SCD Validation segment'].item() for agent,x in NAV_.index]
      
        NAV_.loc[NAV_['LAST_MODIFIED'].isnull(),'LAST_MODIFIED']=NAV_.loc[NAV_['LAST_MODIFIED'].isnull(),'Expected']
        NAV_['Buffer']=NAV_['Expected']-NAV_['LAST_MODIFIED']
        
        NAV_['Delayed']=float('nan')
     
        NAV_.loc[NAV_['Buffer']<pd.Timedelta('-15 minutes'),'Delayed']='DELAYED'
 
        NAV_Summary[date.strftime("%Y-%m-%d")]=NAV_['Delayed']
 
    path = f'Z:\\Fund_Oversight\\OVERSIGHT\\Operations\\Daily\\{ReportDate.strftime("%Y")}\\{ReportDate.strftime("%Y%m")}\\{ReportDate.strftime("%Y%m%d")}\\Simcorp\\STP\\'
    NAV_Summary.to_csv(f'{path}SLA_Month_{ReportDate.strftime("%Y%m%d")}.csv') 
    
    for batch in set(FML['STP Batch Run']):
        
       
        
        if not os.path.isdir(path+'\\batch\\'):
            os.makedirs(path+'\\batch\\')
        
        if not os.path.isfile(f'{path}SLA_Month_{ReportDate.strftime("%Y%m%d")}.csv'):
            NAV_Summary.loc[NAV_Summary.index.get_level_values('CLASS_ID').isin(FML.loc[FML['STP Batch Run']==batch,'Fund and Series'])].to_csv(f'{path}\\{batch}\\SLA_Month_{ReportDate.strftime("%Y%m%d")}.csv') 
        
if __name__=='__main__':
    t=Agent_SLA_Summary(pd.to_datetime('2020-08-04'))