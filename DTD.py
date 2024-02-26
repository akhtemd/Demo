
"""

Validation of the holding prices for North American funds using independent sources. Outline the main movers and material errors.
"""

# In[13]:
import os
import sys
import importlib
sys.path.append('Z:\\Fund_Oversight\\OVERSIGHT\\Operations\\Processes\\')


import Errors
import pandas as pd
from pandas.tseries.offsets import CustomBusinessDay


FXRecon=importlib.import_module('FXRecon')
FilterData=importlib.import_module('FilterData')
ImportTNA=importlib.import_module('ImportTNA')


importlib.reload(FXRecon)
importlib.reload(FilterData)
importlib.reload(ImportTNA)

from FilterData import FilterData
from ImportTNA import FundList,get_TNA,get_HLDG,get_FX
from FXRecon import FXImport,FXImpact

def DTD(ReportDate):
    # In[14]:
    #Define business days based on custom holidays 
    holidays = pd.read_csv('Z:\\Fund_Oversight\\OVERSIGHT\\List of Funds\\Holidays.csv',infer_datetime_format =True,squeeze = True,index_col = 'Calendar ID')
    holidays = holidays['CAD']
    
    holidays = pd.to_datetime(holidays)
    Holist = CustomBusinessDay(holidays=holidays)
    #Get list of business days of length Windowreg ending last business day
    
    DateLst = pd.date_range(end = ReportDate,periods=2,freq=Holist)
    
    path=f'Z:\\Fund_Oversight\\OVERSIGHT\\Operations\\Daily\\{DateLst[1].strftime("%Y")}\\{DateLst[1].strftime("%Y%m")}\\{DateLst[1].strftime("%Y%m%d")}\\Performance Validation\\DTD\\'
    if not os.path.isdir(path):
        os.makedirs(path)
        
    FML=FundList(ReportDate,SCD = False)
        
    # In[14]:
    FinalCheck=pd.read_excel(f'Z:\\Fund_Oversight\\OVERSIGHT\\Operations\\Daily\\{DateLst[1].strftime("%Y")}\\{DateLst[1].strftime("%Y%m")}\\{DateLst[1].strftime("%Y%m%d")}\\Inputs\\ETF_SameDay_{DateLst[1].strftime("%Y%m%d")}.xlsx',sheet_name = 'Investment Detail with FX')
    if pd.Series(FinalCheck['Accounting Period Status\n']=='PRELIM').any():
        raise Errors.InputFormatError(f'ETF_SameDay_{DateLst[1].strftime("%Y%m%d")}.xlsx')
        
     #In[14.5]:
    
    #import and convert prices from bpl files
    bplLst=[]
    SCDT=[]
    for dt in DateLst:
        
        bpl = pd.read_csv(f'Z:\\Fund_Oversight\\OVERSIGHT\\Operations\\Daily\\{dt.strftime("%Y")}\\{dt.strftime("%Y%m")}\\{dt.strftime("%Y%m%d")}\\Inputs\\bpl_price_master_{dt.strftime("%Y-%m-%d")}.csv',converters={'sedol':str,'cusip':str},encoding="ISO-8859-1")
        bpl.dropna(subset=['ml_price'],inplace=True)
        bpl['ml_namr_fair_value_factor']=bpl['ml_namr_fair_value_factor'].fillna(1)
        bpl['ml_price']*=bpl['ml_namr_fair_value_factor']
        bplLst.append(bpl) 
        
        SCDT.append(get_HLDG(dt,'SCD'))
        SCDT[-1].dropna(subset=['CUSIP','ISIN','SEDOL'],how='all',inplace = True)
        SCDT[-1].drop_duplicates(subset=['CUSIP','ISIN','SEDOL','FUND_ID'],inplace = True)
        SCDT[-1][['SEDOL (static)','CUSIP','ISIN']]=SCDT[-1][['SEDOL','CUSIP','ISIN']].astype(str).replace({'nan':float('nan')}).apply(lambda x:x.str.upper())
        SCDT[-1].drop(columns = SCDT[-1].columns[~SCDT[-1].columns.isin(['BBG','FUND_ID','BASE_MV','SHARES','SEDOL','CUSIP','ISIN'])],inplace = True)
        SCDT[-1].drop(SCDT[-1].index[SCDT[-1]['BASE_MV']==0],inplace=True)
    
    # In[17]:
            
    RBCNAV = get_TNA(DateLst[0],'RBC',level='fund')['TNA']
    SSNAV = get_TNA(DateLst[0],'SSB',level='fund')['TNA']
    
    RBCT = get_HLDG(DateLst[1],'RBC')
    RBCT_1=get_HLDG(DateLst[0],'RBC')
    RBCT=RBCT.loc[RBCT['FUND_ID'].isin(FML['NAV Agent portfolio code'])]
    RBCT_1=RBCT_1.loc[RBCT_1['FUND_ID'].isin(FML['NAV Agent portfolio code'])]
    RBCT['FUND_CUR'] = RBCT['FUND_ID'].map(lambda x:FML.loc[FML['NAV Agent portfolio code']==x,'Fund Currency'].iloc[0])
    RBCT_1['FUND_CUR'] = RBCT_1['FUND_ID'].map(lambda x:FML.loc[FML['NAV Agent portfolio code']==x,'Fund Currency'].iloc[0])
    
    
    CIBCNAV_T_1 = pd.read_excel(f'Z:\\Fund_Oversight\\OVERSIGHT\\Operations\\Daily\\{DateLst[0].strftime("%Y")}\\{DateLst[0].strftime("%Y%m")}\\{DateLst[0].strftime("%Y%m%d")}\\Inputs\\ETF_SameDay_{DateLst[0].strftime("%Y%m%d")}.xlsx',sheet_name = 'Daily Net Asset Value - ETFs',index_col='key\n')
    CIBCNAV_T_1.columns = CIBCNAV_T_1.columns.map(lambda x:x.strip())        
    CIBCNAV=CIBCNAV_T_1.loc[[x.find('-0')!=-1 for x in CIBCNAV_T_1.index],'net_assets']


    CIBC=[]
    for num,dt in enumerate(DateLst):
        CIBCPath =f'Z:\\Fund_Oversight\\OVERSIGHT\\Operations\\Daily\\{dt.strftime("%Y")}\\{dt.strftime("%Y%m")}\\{dt.strftime("%Y%m%d")}\\Inputs\\'
    
        if not os.path.isfile(f'{CIBCPath}ETF_SameDay_{dt.strftime("%Y%m%d")}.xlsx'):
            raise Errors.InputMissing(f'{CIBCPath}ETF_SameDay_{dt.strftime("%Y%m%d")}.xlsx')
            
        CIBC_=pd.read_excel(f'{CIBCPath}ETF_SameDay_{dt.strftime("%Y%m%d")}.xlsx',sheet_name = 'Investment Detail')
    
            
        try:
            
            #Remove unnecessary columns and rows and clean headers
            CIBC_.columns = CIBC_.columns.map(lambda x:x.strip())
            CIBC_.dropna(subset=['Reporting Account Short Number'],inplace =True)
            CIBC_['Reporting Account Short Number']=CIBC_['Reporting Account Short Number'].astype(int).astype(str)
            
            #Standardize column names
            colList={'Segment Description':'SEC_TYPE','Market Value Base':'BASE_MV', 'Account Base Currency':'FUND_CUR', \
                     'Security Description':'SEC_NAME','Issue Currency Code':'SEC_CURRENCY', 'Traded Shares/Par':'SHARES', \
                         'Sedol':'SEDOL','Cusip':'CUSIP','ISIN':'ISIN','Reporting Account Short Number':'FUND_ID'}
                
            CIBC_.rename(columns=colList,inplace = True)
        
            CIBC_[['SEDOL','CUSIP','SEDOL']]=CIBC_[['SEDOL','CUSIP','ISIN']].astype(str).apply(lambda x:x.str.upper()).replace({'NAN':float('nan')})
            
            CIBC.append(CIBC_.drop(columns=CIBC_.columns[~CIBC_.columns.isin(colList.values())]))
           
        except:
           
            raise Errors.InputFormatError('ETF_SameDay_{Date.strftime("%Y%m%d")}.xlsx')
            
    #Computes the weight based on base MV and TNA
    CIBC[num]=CIBC[num].loc[CIBC[num]['FUND_ID'].isin(CIBCNAV.index.map(lambda x:x.split('-')[0]))]
    CIBC[num]['PCT_MV_FUND']=CIBC[num]['BASE_MV']/CIBC[num]['FUND_ID'].map(lambda x:CIBCNAV.loc[x+'-0'])
        
        

    #Remove funds that do not have a NAV at T-1 (New funds ie)
    RBCT=RBCT.loc[RBCT['FUND_ID'].isin(RBCNAV.index)]
    RBCT_1=RBCT_1.loc[RBCT_1['FUND_ID'].isin(RBCNAV.index)]
    
    RBCT['PCT_MV_FUND']=RBCT['BASE_MV']/RBCT['FUND_ID'].map(lambda x:RBCNAV.loc[x])

    # In[110]:

    #Get the categories associated with each fund based on the characteristics of the holdings    
    SS =[]

    for dt in DateLst:
        SS.append(get_HLDG(dt,'SSB'))
        #SS[-1].drop(index=SS[-1].index[SS[-1]['INVEST_TYPE_CD']=='45'],inplace=True)
        SS[-1].drop(index=SS[-1].index[[x[0] in ['A','9'] for x in SS[-1]['CUSIP'].values]],inplace=True)
        SS[-1]=SS[-1].loc[SS[-1]['FUND_ID'].isin(SSNAV.index)]
        SS[-1]=SS[-1].loc[SS[-1]['FUND_ID'].isin(FML['NAV Agent portfolio code'])]
        #Computes the weight based on base MV and TNA
        SS[-1]['FUND_CUR'] = SS[-1]['FUND_ID'].map(lambda x:FML.loc[FML['NAV Agent portfolio code']==x,'Fund Currency'].iloc[0])

    SS[-1]['PCT_MV_FUND']=SS[-1]['BASE_MV']/SS[-1]['FUND_ID'].map(lambda x:SSNAV.loc[x])

   
    # In[20]:
    #FX Recon with implicit rate used to value securites
    FXFile = FXImport(DateLst[-1])
    RBCCur =RBCT.groupby(by=['FUND_ID','SEC_CURRENCY','FUND_CUR'],as_index = False)[['BASE_MV','LOCAL_MV','PCT_MV_FUND']].agg(sum)
    RBCCurSummary =RBCT.groupby(by=['SEC_CURRENCY','FUND_CUR'],as_index = False)[['BASE_MV','LOCAL_MV']].agg(sum)
    RBCCurSummary['EXCHANGE_RATE'] = RBCCurSummary['LOCAL_MV']/RBCCurSummary['BASE_MV']
    RBCCur['EXCHANGE_RATE'] = RBCCur['LOCAL_MV']/RBCCur['BASE_MV']
    
    RBCFX = FXImpact(RBCCur,FXFile)
    RBCFX['Agent']='RBC'
    
    RBCCurSummary=FXImpact(RBCCurSummary,FXFile) 
    RBCCurSummary['Agent']='RBC'
    
    
    SSCur = SS[1].groupby(by=['FUND_ID','FUND_CUR','SEC_CURRENCY'],as_index =False)[['BASE_MV','LOCAL_MV','PCT_MV_FUND']].agg(sum)
    
    SSCurSummary =SS[1].groupby(by=['FUND_CUR','SEC_CURRENCY'],as_index =False)[['BASE_MV','LOCAL_MV']].agg(sum)
    
    SSCurSummary['EXCHANGE_RATE'] = SSCurSummary['LOCAL_MV']/SSCurSummary['BASE_MV']
    
    
    
    SSCur['EXCHANGE_RATE'] = SSCur['LOCAL_MV']/SSCur['BASE_MV']
   
    SSFX = FXImpact(SSCur,FXFile)
    SSFX['Agent']='SS'
    
    SSCurSummary=FXImpact(SSCurSummary,FXFile)
    SSCurSummary['Agent']='SS'
    
    FXRec = pd.concat([SSFX,RBCFX]).set_index('FUND_ID')
    
    FXSummary = pd.concat([SSCurSummary,RBCCurSummary])
    


    # In[20]:

    #FilterData contains the rules to apply the treshold and do the mergers
     # In[21]:        
    FX = get_FX(DateLst[1]).set_index('CURRENCY')
    FXT_1 = get_FX(DateLst[0]).set_index('CURRENCY')
    
    
    CIBCReport = FilterData(CIBC[1],CIBC[0],SCDT[1],SCDT[0],bplLst[0].loc[bplLst[0]['id_bpl_pricing_strategy']=='MAM_GL'],bplLst[1].loc[bplLst[1]['id_bpl_pricing_strategy']=='MAM_GL'],FX,FXT_1,'CIBC',DateLst[-1])
    SSReport = FilterData(SS[1],SS[0],SCDT[1],SCDT[0],bplLst[0].loc[bplLst[0]['id_bpl_pricing_strategy']=='MAM_GL'],bplLst[1].loc[bplLst[1]['id_bpl_pricing_strategy']=='MAM_GL'],FX,FXT_1,'SS',DateLst[-1])
    RBCReport = FilterData(RBCT,RBCT_1,SCDT[1],SCDT[0],bplLst[0].loc[bplLst[0]['id_bpl_pricing_strategy']=='CWFA'],bplLst[1].loc[bplLst[1]['id_bpl_pricing_strategy']=='CWFA'],FX,FXT_1,'RBC',DateLst[-1])

   
    # In[21]:    
    #Print result to excel
    ToPrint = pd.concat([SSReport,RBCReport,CIBCReport],sort=True)
    ToPrint.loc[ToPrint['CUSIP'].isnull(),'CUSIP']=ToPrint.loc[ToPrint['CUSIP'].isnull(),'ISIN']
    ToPrint.loc[ToPrint['SEDOL'].isnull(),'SEDOL']=ToPrint.loc[ToPrint['SEDOL'].isnull(),'CUSIP']
    #Print missing SCD securities to retrieve Bloomberg Prices
    
    BBG = ToPrint.loc[ToPrint['SCD VARIATION'].isnull() & ToPrint['Price Variation'].isnull(),['SEC_NAME','SEDOL','CUSIP','ISIN','FUND_CUR_l']]
    
    #Check for valid identifiers to select only one fro Bloomberg extract
    BBG.loc[BBG['CUSIP'].isnull(),'CUSIP']=BBG.loc[BBG['CUSIP'].isnull(),'ISIN']
    BBG.loc[BBG['SEDOL'].isnull(),'SEDOL']=BBG.loc[BBG['SEDOL'].isnull(),'CUSIP']
    BBG.drop(columns=['CUSIP','ISIN'],inplace=True)
    BBG.rename(columns={'SEDOL':'ID','FUND_CUR_l':'PF_ISO_CURRENCY_l'},inplace=True)
    
    BBG.drop_duplicates(inplace=True)
    BBG['VARIABLE'] = 'PX_LAST'
    BBG['DATE']=DateLst[0].strftime("%Y%m%d")               
    BBG1 = BBG.copy()
    BBG1['DATE']=DateLst[1].strftime("%Y%m%d")  
    BBG = pd.concat([BBG,BBG1])
    
    # In[21]:
    
    writer = pd.ExcelWriter(f'{path}BBG {DateLst[1].strftime("%m%d%Y")}.xlsx', engine='xlsxwriter')
    ToPrint.to_excel(writer,sheet_name='Report')
    writer.sheets['Report'].autofilter("A1:HD1")
    #Protect the sheet and allow autofiltering
    writer.sheets['Report'].protect('daytoday',options={'autofilter': True})
    writer.save()
    
    writer = pd.ExcelWriter(f'Z:\\Fund_Oversight\\OVERSIGHT\\Operations\\Tools\\Bloomberg\\Input\\Extract_{DateLst[1].strftime("%m%d%Y")}.xlsx', engine='xlsxwriter')
    BBG.to_excel(writer,sheet_name='BBG',index = False)
    writer.save()
    
    FXSummary.to_csv(f'{path}FXSummary {DateLst[1].strftime("%m%d%Y")}.csv')
    FXT_1.to_csv(f'{path}FXT_1.csv')
    
    if (FXRec['Impact']>0.0005).any():
        FXRec.loc[FXRec['Impact']>0.0005].to_csv(f'{path}FXRec {DateLst[1].strftime("%m%d%Y")}.csv')
        
if __name__=='__main__':
    t=DTD(pd.to_datetime('2020-08-14'))