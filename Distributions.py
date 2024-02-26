# -*- coding: utf-8 -*-
"""
Recons the distributions in between SCD and NAV Agent for FOF

"""
import pandas as pd
import os
from ImportTNA import SCDTNA,FundList
from ImportHLDG import SCDHLDG,RBCHLDG

def Distributions(ReportDate):
        
        #Function to adjust column width when printing final results
    def get_col_widths(dataframe):
        # First we find the maximum length of the index column   
        idx_max = max([len(str(s)) for s in dataframe.index.values] + [len(str(dataframe.index.name))])
        # Then, we concatenate this to the max of the lengths of column name and its values for each column, left to right
        return [idx_max] + [max([len(str(s)) for s in dataframe[col].values] + [len(col)]) for col in dataframe.columns]

    path =f'Z:\\Fund_Oversight\\OVERSIGHT\\Operations\\Daily\\{ReportDate.strftime("%Y")}\\{ReportDate.strftime("%Y%m")}\\{ReportDate.strftime("%Y%m%d")}\\Simcorp\\Distribution\\'
    if not os.path.isdir(path):
            os.makedirs(path)
    
    SCD = SCDTNA(ReportDate)['TNA']
    
    FML = FundList(ReportDate)
    
    RBCNAVPath='\\\MFCGD.com\\GWAMDFS\\Apps\\AABOR\\FundAdmin\\RBCFA\\Archive\\NAV\\Oversight\\'
    for files in os.listdir(RBCNAVPath):
        if files.find(f'rbc_portnav_MLOVER_{ReportDate.strftime("%Y%m%d")}')!=-1:
            RBC= pd.read_csv(RBCNAVPath + files,sep='|',encoding='latin-1',parse_dates=True)
            break
    RBC=RBC.sort_values(by=['Portfolio','Portfolio_Curr']).drop_duplicates(subset =['Portfolio'])
    RBC.rename(columns = {'Portfolio':'Fund_ID','Tot_Distrib':'Total_Dist_Rate'},inplace=True)
    RBC=pd.Series(index=RBC['Fund_ID'].str.upper().values,data=RBC['Total_Dist_Rate'].values).fillna(0)

   
    SSNAVPath='\\\MFCGD.com\\GWAMDFS\\Apps\\AABOR\\FundAdmin\\SSBFA\\Archive\\NAV\\Oversight\\'  
    for files in os.listdir(SSNAVPath): 
        if files.find(f'SSB_NAV_CANADA_{ReportDate.strftime("%Y%m%d")}')!=-1:
            SS= pd.read_csv(SSNAVPath + files,sep='|',encoding='latin-1',parse_dates=True)
            break  
    SS=SS.sort_values(by=['Fund_ID','Currency']).drop_duplicates(subset =['Fund_ID'])
    SS=pd.Series(index=SS['Fund_ID'].str.upper().values,data=SS['Total_Dist_Rate'].values).fillna(0)
    
    Agent = pd.concat([RBC,SS],sort=True)
    
    
    SCDH=SCDHLDG(ReportDate)
    SCDH['Agent Name']=SCDH['FUND'].map(lambda x:str(FML.loc[FML['SCD Fund ID']==x,'NAV Agent'].iloc[0]) if x in FML['SCD Fund ID'].values else float('nan'))
    SCDH['FUND']=SCDH['FUND'].map(lambda x:str(FML.loc[FML['SCD Fund ID']==x,'NAV Agent portfolio code'].iloc[0]) if x in FML['SCD Fund ID'].values else float('nan'))
    SCDH.dropna(subset=['FUND'],inplace=True)
    
    SCDH['ID2']=SCDH['ID4'].map(lambda x:str(FML.loc[FML['SCD Fund ID']==x,'Fund and Series'].iloc[0]) if x in FML['SCD Fund ID'].values else float('nan'))
    SCDH.dropna(subset=['ID2'],inplace=True)
    
    
    RBCH=RBCHLDG(ReportDate)
    RBCH.dropna(subset=['ID2'],inplace=True)
    RBCH=RBCH.loc[RBCH['SEC_TYPE']=='MUTUAL FUND']
    RBCH['ID2']=RBCH['ID2'].map(lambda x:''.join(x[3:].split()))
    
    HLDG = pd.concat([RBCH,SCDH.loc[SCDH['Agent Name']=='SSB']],sort=True)[['FUND','ID2','SHARES']]
   
    
    HLDG['Agent Rate']=HLDG['ID2'].map(lambda x:Agent[x] if x in Agent.index else float('nan'))
    HLDG.dropna(subset=['Agent Rate'],inplace=True)
    
    

    Dist = pd.read_excel(f'Z:\\Fund_Oversight\\OVERSIGHT\\Operations\\Daily\\{ReportDate.strftime("%Y")}\\{ReportDate.strftime("%Y%m")}\{ReportDate.strftime("%Y%m%d")}\\Inputs\\Dist_IN_SCD_{ReportDate.strftime("%Y%m%d")}.xlsx')
    Dist.dropna(subset=['Portfolio','Security ID'],how='any',inplace=True)
    Dist[['Portfolio','Security ID']]=Dist[['Portfolio','Security ID']].applymap(lambda x:''.join(x.split()).upper())
    Dist['Agent Fund']=Dist['Portfolio'].map(lambda x:str(FML.loc[FML['SCD Fund ID']==x,'NAV Agent portfolio code'].iloc[0]) if x in FML['SCD Fund ID'].values else float('nan'))
    Dist.dropna(subset=['Agent Fund'],inplace=True)
    
    Dist['Agent Class']=Dist['Security ID'].map(lambda x:str(FML.loc[FML['SCD Fund ID']==x,'Fund and Series'].iloc[0]) if x in FML['SCD Fund ID'].values else float('nan'))
    Dist['Agent Name']=Dist['Portfolio'].map(lambda x:str(FML.loc[FML['SCD Fund ID']==x,'NAV Agent'].iloc[0]) if x in FML['SCD Fund ID'].values else float('nan'))
    Dist.dropna(subset=['Agent Class'],inplace=True)


    Dist = Dist.merge(HLDG.drop_duplicates(subset=['FUND','ID2']),how='outer',left_on =['Agent Fund','Agent Class'],right_on = ['FUND','ID2'])
    Dist[['Agent Rate','SHARES','Signed payment PC']]=Dist[['Agent Rate','SHARES','Signed payment PC']].fillna(0)
    Dist['Agent Amount']=Dist['Agent Rate'].values*Dist['SHARES'].values
    Dist['$ Difference']=Dist['Agent Amount'].values-Dist['Signed payment PC'].values
    
    
    Dist = Dist.loc[(Dist['Agent Rate']!=0) | (Dist['Signed payment PC']!=0)]
    
    Dist.loc[Dist['Portfolio'].isnull(),'Portfolio']=Dist.loc[Dist['Portfolio'].isnull(),'FUND'].map(lambda x:FML.loc[FML['NAV Agent ID (Top level)']==x,'Fund ID'].iloc[0] if x in FML['NAV Agent ID (Top level)'].values else float('nan'))
    #Remove fund not in SCD
    Dist.dropna(subset=['Portfolio'],inplace=True)
    
    Dist.loc[Dist['Agent Name'].isnull(),'Agent Name']=Dist.loc[Dist['Agent Name'].isnull(),'FUND'].map(lambda x:FML.loc[FML['NAV Agent ID (Top level)']==x,'NAV Agent'].iloc[0] if x in FML['NAV Agent ID (Top level)'].values else float('nan'))
    
    Dist.loc[Dist['FUND'].isnull(),'FUND']=Dist.loc[Dist['FUND'].isnull(),'Agent Fund']
    Dist.loc[Dist['ID2'].isnull(),'ID2']=Dist.loc[Dist['ID2'].isnull(),'Agent Class']
    Dist.loc[Dist['Security ID'].isnull(),'Security ID']=Dist.loc[Dist['Security ID'].isnull(),'ID2'].map(lambda x:FML.loc[FML['NAV AGENT ID']==x,'SCD Fund ID'].iloc[0] if x in FML['NAV AGENT ID'].values else float('nan'))
    
    
    
    
   
    
    Dist['% Difference']=Dist['$ Difference'].values/Dist['Portfolio'].map(lambda x:SCD[x])
    Dist.rename(columns = {'Signed payment PC':'SCD $ amount','ID2':'CLASS'},inplace=True)
    
    summary = {}
    
    for agent in set(Dist['Agent Name']):
        summary[agent]=Dist.loc[Dist['Agent Name']==agent].groupby(by=['FUND'],as_index=False)['% Difference'].agg(sum)
        summary[agent]=summary[agent].reindex(summary[agent]['% Difference'].abs().sort_values().index)
    
    summary = pd.concat(summary,axis =1,keys =set(Dist['Agent Name']))
    
    writer = pd.ExcelWriter(f'{path}DistRec_{ReportDate.strftime("%Y%m%d")}.xlsx', engine='xlsxwriter')
    
    format2 = writer.book.add_format({'num_format': '0.00%'})
    
    summary.to_excel(writer,sheet_name = 'Summary')
    writer.sheets['Summary'].set_column(0, 3, 20)
    writer.sheets['Summary'].set_column('C:E',None,format2)
    
    for name in set(Dist['Agent Name']):
        to_print = Dist.loc[Dist['Agent Name']==name,['FUND','CLASS','Security ID','Portfolio','SHARES','Agent Rate','Agent Amount','SCD $ amount','$ Difference','% Difference']].round(4)
        to_print['Date']=ReportDate.strftime("%Y-%m-%d")
        to_print['Remarks']=float('nan')
        to_print.to_excel(writer,sheet_name = name,index = False)
        for j, width in enumerate(get_col_widths(to_print)):
            writer.sheets[name].set_column(j, j, width)
        writer.sheets[name].set_column('J:J',None,format2)
    
    writer.save()
