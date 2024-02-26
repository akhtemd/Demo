# -*- coding: utf-8 -*-
"""
Created on Wed Aug 12 09:08:39 2020

@author: laguema
"""


import sqlite3
import pandas as pd
import getpass
import importlib
import os
import numpy as np


ImportTNA=importlib.import_module('ImportTNA')
importlib.reload(ImportTNA)
from ImportTNA import FundList

           
def Distributions_OUT(ReportDate):
    
    FML=FundList(ReportDate,SCD=False)
    FML['Year End']=FML['Year End'].map(lambda x:x.split()[0] if x==x else x)
    
    path = f'Z:\\Fund_Oversight\\OVERSIGHT\\Operations\\Daily\\{ReportDate.strftime("%Y")}\\{ReportDate.strftime("%Y%m")}\\{ReportDate.strftime("%Y%m%d")}\\Performance Control\\Distributions\\'

    if not os.path.isdir(path):
        os.makedirs(path)
        
    sql ="SELECT * FROM NAV WHERE DATE Between ? and ? AND NAV_AGENT = ?"
    
    for agent in ('RBC','SSB'):
        try:
            conn=sqlite3.connect(f'C:\\Users\\{getpass.getuser().capitalize()}\\Documents\\Oversight_PROD.db',timeout = 150)
            
            NAV=pd.read_sql(sql,conn,params = ((ReportDate-pd.Timedelta(180,unit='d')).strftime("%Y-%m-%d"),ReportDate.strftime("%Y-%m-%d"),agent),parse_dates={'LAST_MODIFIED': 's'})
            
            conn.close()
            
            NAV=NAV.merge(FML,how='inner',left_on=['NAV_AGENT','CLASS_ID','CLASS_CURRENCY'],right_on=['NAV Agent','Fund and Series','Class Currency'],suffixes=('','_r'))
            
            nameDay = ReportDate.strftime("%Y-%m-%d")
            
            NAV.replace(0,float('nan'),inplace=True)
            NAV['month'] = [x+pd.Timedelta(1,'d')-pd.offsets.MonthBegin() for x in pd.to_datetime(NAV['DATE'])]
            
            NAV.set_index('CLASS_ID',inplace=True)
            TNA = NAV.loc[NAV['DATE']==sorted(set(NAV['DATE']))[-2],'TNA']
            DISTAMOUNT = NAV.loc[NAV['DATE']==sorted(set(NAV['DATE']))[-1],'SHARES']* \
                NAV.loc[NAV['DATE']==sorted(set(NAV['DATE']))[-1],'DISTRIBUTION_RATE']
            
            DISTAMOUNT.name = 'Amount'
            TNA=DISTAMOUNT/TNA
            TNA.name = 'Distribution %'
            NAV.reset_index(inplace=True)
            count = NAV.loc[NAV['month']==NAV['month'].max()].groupby('CLASS_ID')['DISTRIBUTION_RATE'].agg('count')
            count.name = 'Nb Distributions current month'
            
            NAV=NAV.loc[(NAV['DATE']==nameDay) | (NAV['month']!=NAV['month'].max())]
            
            
            DIST = NAV.groupby(['month','CLASS_ID'],as_index=False)[['DISTRIBUTION_RATE']].agg(np.nanmean)
            DIST.sort_values('month',inplace=True)
          
            
           
            AMOUNT = DIST.pivot(index = 'CLASS_ID',columns = 'month',values='DISTRIBUTION_RATE')
            AMOUNT=AMOUNT.replace(0,float('nan'))
            
        
            AMOUNT.columns = [x.strftime("%B") for x in AMOUNT.columns]
            
            AMOUNT.rename(columns ={AMOUNT.columns[-1]:nameDay},inplace=True)
            
            
            for x,y in zip(FML['Fund and Series'],FML['Year End']):
                if x in AMOUNT.index and y in AMOUNT.columns:
                    AMOUNT.loc[x,y]='Year End'
            
            
            stat = pd.DataFrame()
            stat['mean']=AMOUNT.iloc[:,:-1].replace('Year End',float('nan')).apply(np.nanmean,axis=1)
            stat['std']=AMOUNT.iloc[:,:-1].replace('Year End',float('nan')).apply(np.nanstd,axis=1)
            stat.loc[stat['std']<0.00001,'std']=float('nan')
            stat['MEAN ERR'] =AMOUNT[nameDay].replace('Year End',float('nan'))-stat['mean']
            stat['STD ERR'] =((AMOUNT[nameDay].replace('Year End',float('nan'))-stat['mean'])/stat['std']).fillna(0)
        
            stat['Investigate ?']='NO'
        
          
            AMOUNT = pd.concat([DISTAMOUNT,AMOUNT[nameDay],TNA,stat,AMOUNT.drop(columns =nameDay),count],axis=1)
        
            AMOUNT.dropna(subset = [nameDay],inplace=True)
            AMOUNT.loc[(AMOUNT['STD ERR'].abs()>=2) & (AMOUNT['Distribution %']>0.004),'Investigate ?']='YES'
            
            label={'mean':'Mean of monthly distributions of previous months exluding year end','std':'Standard deviation of monthly distributions of previous months exluding year end',\
                  'MEAN ERR':'Difference between current month distribution and mean of previous months','STD ERR':'Number of STD from the mean of previous months', \
                      'Nb Distributions current month':'Nb of times the fund distributed in the current month'}
            
              
                                
            writer = pd.ExcelWriter(f'{path}Distributions_{agent}_{nameDay}.xlsx', engine='xlsxwriter')
            workbook =writer.book
            
            AMOUNT.to_excel(writer,startrow=1,sheet_name = 'Distibutions Analysis')
            
            head2=workbook.add_format({'bold': True})
            head2.set_align('center')
            head2.set_align('vcenter')
            head2.set_text_wrap()
            
            colormonth =workbook.add_format({'bg_color':'#b7dee8'})
        
            dollarcent = workbook.add_format({'num_format': '$#,##0.00'})
            percent = workbook.add_format({'num_format': '%#,##0.00'})
            dollar = workbook.add_format({'num_format': '$#,##0.000'})
            dollardig = workbook.add_format({'num_format': '$#,##0.000000'})
            number = workbook.add_format({'num_format': '#,##0.000'})
            
            # Write the column headers with the defined format.
            for col_num, value in enumerate(AMOUNT.columns):
                
                if value in label.keys():
                    
                    value=label[value]
                else:
                    value='Mean amount per distribution per share during month'
                    
                writer.sheets['Distibutions Analysis'].write(0, col_num+1, value, head2)
            
            writer.sheets['Distibutions Analysis'].set_column(1,1,20, dollarcent)
            writer.sheets['Distibutions Analysis'].set_column(2,2,20, dollardig)
            writer.sheets['Distibutions Analysis'].set_column(3,3,20, percent)
            writer.sheets['Distibutions Analysis'].set_column(4,4,20, dollar)
            writer.sheets['Distibutions Analysis'].set_column(5,5,20)
            writer.sheets['Distibutions Analysis'].set_column(6,6,20, dollar)
            writer.sheets['Distibutions Analysis'].set_column(7,8,20)
            writer.sheets['Distibutions Analysis'].set_column(9,len(AMOUNT.columns)-2,20, dollar)
            writer.sheets['Distibutions Analysis'].set_column(len(AMOUNT.columns)-1,len(AMOUNT.columns),20)
            
            writer.sheets['Distibutions Analysis'].conditional_format('B2:Z2', {'type': 'cell','criteria': '=',
                                                      'value':  f'"{nameDay}"',
                                                  'format': colormonth})
            
            writer.save()
    
        except:
            pass


if __name__=='__main__':
    test=Distributions_OUT(pd.to_datetime('2020-08-30'))