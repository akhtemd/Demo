# -*- coding: utf-8 -*-
"""
Gets number of breaks over a given month by country. Period is define from start of the month to the date chosen
"""

import pandas as pd
from openpyxl import load_workbook
from ImportTNA import FundList


def Breaks(ReportDate):
    
    ReportDate+=pd.offsets.DateOffset(days=1)
    FML = FundList(ReportDate-pd.offsets.MonthBegin(0))
    daterg =pd.date_range(start=ReportDate-pd.offsets.MonthBegin(1), end = ReportDate-pd.offsets.DateOffset(days=1))
    FML.loc[FML['SCD Validation segment']=='US_LUX','Country of Registration']='MAM US'
    FML.loc[FML['Country of Registration']=='Luxembourg','Country of Registration']='Hong Kong'
    breaks = {}
    
    for ct in set(FML['Country of Registration']):
        
        breaks[ct] = pd.read_excel('Z:\\Fund_Oversight\\OVERSIGHT\\Reporting\\Historical Breaks\\workFile.xlsx',sheet_name=ct,header = 1,index_col = 0)
        breaks[ct][daterg[-1].strftime('%B')] = breaks[ct][daterg[-1].strftime('%B')].map(lambda x:float('nan'))
       
    MYTNA = pd.read_excel('Z:\\Fund_Oversight\\OVERSIGHT\\Reporting\\Historical Breaks\\workFile.xlsx',sheet_name='Malaysia TNA Rec',header = 1,index_col = 0)
    MYTNA[daterg[-1].strftime('%B')]=float('nan')
    Summary = pd.read_excel('Z:\\Fund_Oversight\\OVERSIGHT\\Reporting\\Historical Breaks\\workFile.xlsx',sheet_name='Summary',header = 1,index_col = 0)
    Summary.drop(index = 'Total',inplace=True)

    CAT = pd.read_excel('Z:\\Fund_Oversight\\OVERSIGHT\\Reporting\\Historical Breaks\\workFile.xlsx',sheet_name='Category',header = [0,1])
    CAT.index = CAT.iloc[:,0]
    CAT = CAT.iloc[:,1:]
    CAT.columns.names=['Month','Country']
    CAT.loc[:,(daterg[-1].strftime("%B").upper(),slice(None))]=0
    
    seg =pd.DataFrame(columns = set(FML['SCD Validation segment']),index = CAT.index, data=0)
    
    for x in daterg:
      
        for y in set(FML['STP Batch Run']):
            
            for s in set(FML.loc[FML['STP Batch Run'] == y,'SCD Validation segment']):
                
                try:
                    
                    z = pd.read_excel(f'Z:\\Fund_Oversight\\OVERSIGHT\\Operations\\Daily\\{x.strftime("%Y")}\\{x.strftime("%Y%m")}\\{x.strftime("%Y%m%d")}\\Simcorp\\{y}\\{s}\\{x.strftime("%Y-%m-%d")} Daily sign off - {s}.xlsx',header=1)   
                    z=z.loc[z['SCD vs NAV Agent']=='Break']             
                    
                    for ct in breaks.keys():
                        
                        if breaks[ct].loc[x.day,x.strftime('%B')]!=breaks[ct].loc[x.day,x.strftime('%B')] and ct !='Vietnam':
                            breaks[ct].loc[x.day,x.strftime('%B')] =0
                                      
                        breaks[ct].loc[x.day,x.strftime('%B')] +=sum(z['SCD Class ID'].isin(FML.loc[FML['Country of Registration']==ct,'SCD Liability code']))
                        
                        for cat in CAT.index:

                            CAT.loc[cat,(daterg[-1].strftime("%B").upper(),ct)]+=sum(z.loc[z['SCD Class ID'].isin(FML.loc[FML['Country of Registration']==ct,'SCD Liability code']),'Break Category']==cat)
                    
                    for cat in seg.index:    
                        seg.loc[cat,s]+=sum(z['Break Category']==cat)
                    
                    seg.loc['Miscelaneous',s]+=len(z.index)-z['Break Category'].isin(seg.index).sum()
              
                except:
                                      
                    pass 
                
                #if pd.offsets.BMonthEnd().rollforward(daterg[0])==x and 'MY_' in s:
                    
                if MYTNA[x.strftime('%B')].item()!=MYTNA[x.strftime('%B')].item():
                    
                    MYTNA[x.strftime('%B')]=0  
                                  
                try:
                
                    z = pd.read_excel(f'Z:\\Fund_Oversight\\OVERSIGHT\\Operations\\Daily\\{x.strftime("%Y")}\\{x.strftime("%Y%m")}\\{x.strftime("%Y%m%d")}\\Simcorp\\{y}\\{s}\\{x.strftime("%Y-%m-%d")} Daily sign off TNA rec- {s}.xlsx',header=1)   
                    MYTNA[x.strftime('%B')] +=sum(z['SCD vs NAV Agent']=='Break')
                    
                except:
                    
                    pass 
                
   
    try:
                                  
        z = pd.read_excel(f'Z:\\Fund_Oversight\\OVERSIGHT\\Operations\\Daily\\{x.strftime("%Y")}\\{x.strftime("%Y%m")}\\{x.strftime("%Y%m%d")}\\Simcorp\\STP\\Vietnam\\{x.strftime("%Y-%m-%d")} Daily Sign off - Vietnam MF.xlsx',sheet_name ='Summary')
        z.dropna(how = 'all',inplace=True)
        z.columns = z.iloc[1,:]
          
        breaks['Vietnam'].loc[x.day,x.strftime('%B')] =sum(z['Status']=='Break')        
                          
    except:
                          
        pass
        
              
    for ct in breaks.keys():             
        CAT.loc['Miscelaneous',(daterg[-1].strftime("%B").upper(),ct)]+=breaks[ct][daterg[-1].strftime('%B')].sum()-CAT[(daterg[-1].strftime("%B").upper(),ct)].sum()
    

        
    for ct in breaks.keys():
        
        Summary.loc[ct,daterg[-1].strftime('%B')] = breaks[ct][daterg[-1].strftime('%B')].sum()
    
    Summary.loc['Malaysia TNA',daterg[-1].strftime('%B')]=MYTNA[daterg[-1].strftime('%B')].item()

    book = load_workbook('Z:\\Fund_Oversight\\OVERSIGHT\\Reporting\\Historical Breaks\\workFile.xlsx')
    writer =  pd.ExcelWriter('Z:\\Fund_Oversight\\OVERSIGHT\\Reporting\\Historical Breaks\\workFile.xlsx',engine ='openpyxl')
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    
    for ct in breaks.keys():
       
        breaks[ct].to_excel(writer, sheet_name=ct,index=False,startrow =1,startcol=1)
        
    Summary.to_excel(writer, sheet_name='Summary',index=False,startrow =1,startcol=1)
    MYTNA.to_excel(writer, sheet_name='Malaysia TNA Rec',index=False,startrow =1,startcol=1)
    CAT.columns=CAT.columns.droplevel()
    CAT.to_excel(writer, sheet_name='Category',index=False,startrow =1,startcol=1)
    seg.to_excel(writer, sheet_name='Summary by segment')
    writer.save()
if __name__=='__main__':
    
    t=Breaks(pd.to_datetime('2020-05-21'))