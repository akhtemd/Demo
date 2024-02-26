# -*- coding: utf-8 -*-
"""
Created on Wed Oct  5 22:53:01 2022

@author: chennvi
"""


import pandas as pd
#from openpyxl import load_workbook
from ImportTNA import FundList
#import win32com.client as win32


def BreaksDaily(ReportDate):
    
    #ReportDate+=pd.offsets.DateOffset(days=1)
    FML = FundList(ReportDate-pd.offsets.MonthBegin(0))
    FML.loc[FML['SCD Validation segment']=='US_LUX','Country of Registration']='MAM US'
    FML.loc[FML['Country of Registration']=='Luxembourg','Country of Registration']='Hong Kong'
    
    
    x= ReportDate
    
    path=f'Z:\\Fund_Oversight\\OVERSIGHT\\Operations\\Daily\\{ReportDate.strftime("%Y")}\\{ReportDate.strftime("%Y%m")}\\{ReportDate.strftime("%Y%m%d")}\\Simcorp\\'
    
    
    writer=pd.ExcelWriter(f'{path}DailyBreaks_{ReportDate.strftime("%Y%m%d")}.xlsx', engine='xlsxwriter')
    workbook=writer.book                            
    format1 = workbook.add_format({'num_format': '0.00%'}) #define decimal format 
   
    
    summary_df={}
    summary_df=pd.DataFrame.from_dict(summary_df)
    tot_break=0
    tot_classes=0

    
    #populate summary tab
    
    for y in ('STP Asia (T)', 'STP Asia (T+1)','STP Asia (T+2-FOF)','STP NorthAm'):
        sumry_df={}
        sumry_df=pd.DataFrame.from_dict(sumry_df)
          
        for s in set(FML.loc[FML['STP Batch Run'] == y,'SCD Validation segment']):
            try:
                z = pd.read_excel(f'Z:\\Fund_Oversight\\OVERSIGHT\\Operations\\Daily\\{x.strftime("%Y")}\\{x.strftime("%Y%m")}\\{x.strftime("%Y%m%d")}\\Simcorp\\{y}\\{s}\\{x.strftime("%Y-%m-%d")} Daily sign off - {s}.xlsx',header=1)   
                tot_rows=len(z)
                z=z.loc[z['SCD vs NAV Agent']=='Break']
                count=len(z)
                if count>0:                                       
                    sumry_df=[[y,s,count,tot_rows]]
                    summary_df=summary_df.append(sumry_df,ignore_index=True)
                    tot_break+=count
                    tot_classes+=tot_rows
                    
                
            except:
                 pass
             
    
    #to add dummy dataframe to populate Summary, summary2 as first,second tabs
    df_null={}
    df_null=pd.DataFrame.from_dict(df_null)
    df_null.to_excel(writer, sheet_name='Summary',index=False)
    df_null.to_excel(writer, sheet_name='Summary2',index=False)
    df_null.to_excel(writer, sheet_name='DTD',index=False)
                
    #start - Get counts from performance control -summary tab
    for s in ('RBC','SSB'):
            sumry_df={}
            sumry_df=pd.DataFrame.from_dict(sumry_df)
            try:
                #print(f'Z:\\Fund_Oversight\\OVERSIGHT\\Operations\\Daily\\{x.strftime("%Y")}\\{x.strftime("%Y%m")}\\{x.strftime("%Y%m%d")}\\Performance Control\\{s}\\Performance_Control_{s}_{x.strftime("%Y-%m-%d")}.xlsx')
                z = pd.read_excel(f'Z:\\Fund_Oversight\\OVERSIGHT\\Operations\\Daily\\{x.strftime("%Y")}\\{x.strftime("%Y%m")}\\{x.strftime("%Y%m%d")}\\Performance Control\\{s}\\Performance_Control_{s}_{x.strftime("%Y-%m-%d")}.xlsx',sheet_name='Summary',header=1)   
                tot_rows_perf_cntrl=len(z)
                tot_classes+=tot_rows_perf_cntrl
                nonnull_cnt=pd.notnull(z).sum().sum() #get count of all not null cell values
                sumry_df=[['Performance Control','Perf Control '+s,nonnull_cnt,tot_rows_perf_cntrl]]
                summary_df=summary_df.append(sumry_df,ignore_index=True)
                tot_break+=nonnull_cnt
                
                #create Performance Control (RBC,SSB) summary tabs
                perf_sheet='Perf Control summary '+s
                z.to_excel(writer, sheet_name=perf_sheet,index=False)
                writer.sheets[perf_sheet].set_column(0,25, 20)
                    
                
            except:
                 pass
        
        #end
    
    #start DTD changes
        
    sumry_df_dtd={}
    sumry_df_dtd=pd.DataFrame.from_dict(sumry_df_dtd)
    try:
                #print(f'Z:\\Fund_Oversight\\OVERSIGHT\\Operations\\Daily\\{x.strftime("%Y")}\\{x.strftime("%Y%m")}\\{x.strftime("%Y%m%d")}\\Performance Validation\\DTD\\BBG {x.strftime("%m%d%Y")}.xlsx')
                z = pd.read_excel(f'Z:\\Fund_Oversight\\OVERSIGHT\\Operations\\Daily\\{x.strftime("%Y")}\\{x.strftime("%Y%m")}\\{x.strftime("%Y%m%d")}\\Performance Validation\\DTD\\BBG {x.strftime("%m%d%Y")}.xlsx',sheet_name='Report')   
                tot_rows_dtd=len(z) 
                tot_classes+=tot_rows_dtd
                z=z[['FUND_ID','CUSIP','Error']]
                #-0.001 smaller than or higher than 0.001 - display the records of this condition
                z_error_0_1 = z[(z['Error']< -0.001) | (z['Error'] > 0.001)]
                
                z_error_0_1=z_error_0_1.sort_values(by='Error', ascending=False)
                
                sumry_df_dtd=[['DTD','DTD ',len(z_error_0_1),tot_rows_dtd]]
                summary_df=summary_df.append(sumry_df_dtd,ignore_index=True)
                tot_break+=len(z_error_0_1)                
                dtd_sheet='DTD'
                z_error_0_1.to_excel(writer, sheet_name=dtd_sheet,index=False)
                writer.sheets[dtd_sheet].set_column(0,2, 20)
    except:
            pass
        
    
    #End DTD changes
        
        
    if summary_df.empty:
        pass
    else:
    
        summary_df.columns=['STP Batch Run','SCD Validation segment','Total Breaks','Total No. of classes']
        summary_df['% of Breaks']=summary_df['Total Breaks']/summary_df['Total No. of classes']
        summary_df['% of Breaks']=summary_df['% of Breaks'].round(4)
        
       
        #Add blank row
        s = pd.Series([None,None,None,None,None],index=['STP Batch Run','SCD Validation segment','Total Breaks','Total No. of classes','% of Breaks'])
        summary_df=summary_df.append(s,ignore_index=True)
        
        #populate grand total of Breaks,classes
        s1 = pd.Series([None,'Total',tot_break,tot_classes,None],index=['STP Batch Run','SCD Validation segment','Total Breaks','Total No. of classes','% of Breaks'])
        summary_df=summary_df.append(s1,ignore_index=True)      
               
        summary_df.to_excel(writer, sheet_name='Summary',index=False)
        writer.sheets['Summary'].set_column(0, 6, 25)
        writer.sheets['Summary'].set_column('E:E',20,format1)
        
        
    seg_discrpcy_df={}
    seg_discrpcy_df=pd.DataFrame.from_dict(seg_discrpcy_df)
            
    #populate 'STP Asia (T)', 'STP Asia (T+1)','STP Asia (T+2-FOF)','STP NorthAm' tabs   
    for y in ('STP Asia (T)', 'STP Asia (T+1)','STP Asia (T+2-FOF)','STP NorthAm'):
          
               
        seg={}        
        
        for s in set(FML.loc[FML['STP Batch Run'] == y,'SCD Validation segment']):
         
            
            try:
                    
                    seg=pd.DataFrame.from_dict(seg)
                    seg_df={}
                    seg_df=pd.DataFrame.from_dict(seg_df)
                    z = pd.read_excel(f'Z:\\Fund_Oversight\\OVERSIGHT\\Operations\\Daily\\{x.strftime("%Y")}\\{x.strftime("%Y%m")}\\{x.strftime("%Y%m%d")}\\Simcorp\\{y}\\{s}\\{x.strftime("%Y-%m-%d")} Daily sign off - {s}.xlsx',header=1)   
                    z=z.loc[z['SCD vs NAV Agent']=='Break']
                    seg_df=z[['NAV Agent','NAV Agent Class ID','NAV Agent Fund ID','SCD Class ID','Discrepancy']]
                    seg_df['SCD Validation segment'] = s
                    seg_df=seg_df[['SCD Validation segment','NAV Agent','NAV Agent Class ID','NAV Agent Fund ID','SCD Class ID','Discrepancy']]
                    if seg_df.empty:
                        pass
                    else:
                        s = pd.Series([None,None,None,None,None,None],index=['SCD Validation segment','NAV Agent','NAV Agent Class ID','NAV Agent Fund ID','SCD Class ID','Discrepancy'])
                        seg_df=seg_df.append(s,ignore_index=True)
                        #Eliminate values between -.40 and +.40
                        discrpcy_40 = seg_df[(seg_df['Discrepancy'].round(4)< -0.0040) | (seg_df['Discrepancy'].round(4) > .0040)]
                        
                        if discrpcy_40.empty:
                            pass
                        else:
                            seg_discrpcy_df=seg_discrpcy_df.append(discrpcy_40,ignore_index=True)


                    if seg.empty:
                        seg=seg_df
                    else: 
                        seg=seg.append(seg_df,ignore_index=True)
                    
                    
            except:
                                      
                    pass
                
                
            if seg.empty:
                            pass
            else:
                            seg.to_excel(writer, sheet_name=y,index=False)
                            writer.sheets[y].set_column(0, 6,25)
                            writer.sheets[y].set_column('F:F',20,format1)
                                                       
                            
                                                                                  
    #populate Summary2 tab
    if seg_discrpcy_df.empty:
        pass
    else:
        #display with Discrepancy order from high to low
        seg_discrpcy_df=seg_discrpcy_df.sort_values(by='Discrepancy', ascending=False)
        seg_discrpcy_df.to_excel(writer, sheet_name='Summary2',index=False)
        writer.sheets['Summary2'].set_column(0, 6,25)
        writer.sheets['Summary2'].set_column('F:F',20,format1)
    writer.save()
    
              
                                     

             
if __name__=='__main__':
    
    t=BreaksDaily(pd.to_datetime('2022-10-26'))