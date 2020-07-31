import pandas as pd
import os
import matplotlib
matplotlib.use('Agg') #Prevents displaying of graphs until plt.show(). This increases export time
import matplotlib.pyplot as plt
import numpy as np
import datetime as dt

#from PSI_Fiscal_Calendar import Calendar

main_directory = "C:/Users/mfrangos/Desktop/GL Cube"
os.chdir(f"{main_directory}")
import PSI_Fiscal_Calendar
FiscalCalendar = PSI_Fiscal_Calendar.Calendar()

os.chdir(f"{main_directory}")
##Get some data
Mapping_Account_Names_and_EBITDA_Adj1 = pd.read_csv("Mapping - Account Names & EBITDA Adj.csv")
Mapping_Account_Groups2 = pd.read_csv("Mapping - Account Groups.csv")
Account_Mapping_Merged = pd.merge(Mapping_Account_Groups2,Mapping_Account_Names_and_EBITDA_Adj1,on ="ACCOUNT", how = "outer")
Account_Mapping_Merged["EBITDA ADJS"] = Account_Mapping_Merged["EBITDA ADJS"].fillna("No")  #Fill NA's with "No"
#Mapping_Store_List3 = pd.read_csv("Mapping - Store List.csv")
#os.chdir(f"Y:/Store List Updates")
#Store_List_Updates_Accounting_Department = pd.read_excel("Store List Updates (Accounting Department).xlsx")
#
##Get main dataset
#os.chdir(f"//data/accounting/GL Cube")
#GL_Data = pd.read_excel() #PLACE THE GL DOWNLOAD DATA HERE
#GL_Data.columns = ['GL YEAR', 'Fiscal Month', 'Store', 'DEPT NAME', 'ACCOUNT',
#       'ACCOUNT NAME', 'GL AMOUNT', 'EBITDA ADJS', 'Account Group',
#       'Account Group #', 'LOCATION TYPE', 'Year Opened', 'Region', 'District',
#       'Unnamed: 14', 'Unnamed: 15']
##GL_Data = GL_Data.drop(['Unnamed: 14', 'Unnamed: 15'], axis = 1)
#
##Multiple merges
#Merged_Data_and_Mapping1 = GL_Data.merge(Mapping_Account_Names_and_EBITDA_Adj1.drop("ACCOUNT NAME",axis = 1), on = "ACCOUNT")
#Merged_Data_and_Mapping2 = Merged_Data_and_Mapping1.merge(Mapping_Account_Groups2, on = "ACCOUNT")
#Merged_Data = Merged_Data_and_Mapping1.merge(Mapping_Store_List3, on = "Store")



#LOAD DATA - SHORTCUT for now.
os.chdir(f"//data/accounting/GL Cube")
Merged_Data = pd.read_excel("2017-2018-2019 GL Cube.xlsx", "GL Data")
Store_list = pd.read_excel("Y:/Tableau Datasources/Store List.xlsx")


Stores_with__spa = Store_list.loc[Store_list[" Spa"] == "Yes"]
Stores_with__wash = Store_list.loc[Store_list[" Wash"] == "Yes"]
Stores_with_no_wash_no_spa = Store_list.loc[(Store_list[" Wash"] == "No") & (Store_list[" Spa"] == "No")]

Data_with__wash = Merged_Data.loc[Merged_Data['DEPT #'].isin(Stores_with__wash["Store #"])]
Data_with__spa = Merged_Data.loc[Merged_Data['DEPT #'].isin(Stores_with__spa["Store #"])]       
Control_Group = Merged_Data.loc[Merged_Data['DEPT #'].isin(Stores_with_no_wash_no_spa["Store #"])]
len(Control_Group) + len(Data_with__wash) + len(Data_with__spa)

                           

Average__spa_footage = 491
Average_store_footage = 7539

Data_with__wash.loc[Data_with__wash["ACCOUNT NAME"] == "Groomers Expense              "]['GL AMOUNT'].sum()
Data_with__spa.loc[Data_with__spa["ACCOUNT NAME"] == "Groomers Expense              "]['GL AMOUNT'].sum()
Control_Group.loc[Control_Group["ACCOUNT NAME"] == "Groomers Expense              "]['GL AMOUNT'].sum()


SimilarAccounts_spa_and_wash = pd.DataFrame(set(Data_with__wash["ACCOUNT NAME"])).loc[pd.DataFrame(set(Data_with__wash["ACCOUNT NAME"]))[0].isin(pd.DataFrame(set(Data_with__spa["ACCOUNT NAME"]))[0])]                 
#SimilarAccounts_spa_vs_all = pd.DataFrame(set(Data_with__spa["ACCOUNT NAME"])).loc[pd.DataFrame(set(Data_with__spa["ACCOUNT NAME"]))[0].isin(pd.DataFrame(set(Control_Group["ACCOUNT NAME"]))[0])]

def calculate_average_or_sum(Data,function):
  accounts = []
  lines = []
  for account in SimilarAccounts_spa_and_wash[0]:
    accounts.append(account)
    if function == "mean":
      lines.append(Data.loc[Data["ACCOUNT NAME"] == f"{account}"]['GL AMOUNT'].mean())
    elif function == "sum":
      lines.append(Data.loc[Data["ACCOUNT NAME"] == f"{account}"]['GL AMOUNT'].sum())
  return pd.DataFrame({"ACCOUNT NAME":accounts, "GL AMOUNT":lines })

_Spa_Account_Averages = calculate_average_or_sum(Data_with__spa,"mean")
_Wash_Account_Averages = calculate_average_or_sum(Data_with__wash,"mean")
Control_Group_Account_Averages = calculate_average_or_sum(Control_Group,"mean")


#MERGE THE DATA
temp1 = Data_with__spa
temp1.columns = ['GL YEAR', 'Fiscal Month', 'Store #', 'DEPT NAME', 'ACCT #',
       'ACCOUNT NAME', 'GL AMOUNT', 'EBITDA ADJS', 'Account Group',
       'Account Group #', 'LOCATION TYPE', 'Year Opened', 'Region', 'District',
       'Unnamed: 14', 'Unnamed: 15']
Merged__Spa_Data = pd.merge(Data_with__spa,Store_list, on = "Store #")

temp1 = Merged_Data
temp1.columns = ['GL YEAR', 'Fiscal Month', 'Store #', 'DEPT NAME', 'ACCT #',
       'ACCOUNT NAME', 'GL AMOUNT', 'EBITDA ADJS', 'Account Group',
       'Account Group #', 'LOCATION TYPE', 'Year Opened', 'Region', 'District',
       'Unnamed: 14', 'Unnamed: 15']
GL_Data_and_Store_list_merge = pd.merge(Merged_Data,Store_list, on = "Store #", how = 'outer') 

#Remove strings from int data
GL_Data_and_Store_list_merge['Lease SqFt'] = pd.to_numeric(GL_Data_and_Store_list_merge["Lease SqFt"].astype(str).str.replace(',',''), errors='coerce')
#GL_Data_and_Store_list_merge["Store #"]= GL_Data_and_Store_list_merge["Store #"].str.replace("[^0-9]",'')


#Get all account names/numbers
All_Accounts =  pd.DataFrame({"ACCOUNT NAME": list(set(Merged_Data["ACCOUNT NAME"]))})
account_num_list = []
for account_name in All_Accounts.iloc[:,0]:
  account_num_list.append((set(GL_Data_and_Store_list_merge.loc[GL_Data_and_Store_list_merge["ACCOUNT NAME"] == account_name]["ACCT #"] )))
  GL_Data_and_Store_list_merge.loc[GL_Data_and_Store_list_merge["ACCOUNT NAME"] == account_name]
All_Accounts = pd.DataFrame({"ACCOUNT NAME": list(set(Merged_Data["ACCOUNT NAME"])),
                             "ACCT #": [str(item).replace("{","").replace("}","") for item in account_num_list]}).dropna().reset_index(drop = True) #Cleans up some nonsense      
#Sorts the accounts
All_Accounts = All_Accounts.sort_values(by=['ACCT #'], ascending=True).reset_index(drop = True)
                                            
                                                
Services_Department_Salaries = GL_Data_and_Store_list_merge.loc[GL_Data_and_Store_list_merge["Store #"] == 7128]

#Some date variables
CurrentDate = FiscalCalendar.CurrentDate
CurrentFiscalMonth = FiscalCalendar.CurrentFiscalPeriod
CurrentFiscalYear = FiscalCalendar.CurrentFiscalYear
CurrentFiscalDay = FiscalCalendar.CurrentFiscalDay


List_of_Final_DFs = []
#Iterate through all stores.
for store in Store_list["Store #"]:
    Select_Store_Data = Merged_Data.loc[Merged_Data["Store #"] == store]
    print("Store:", store)
                                                    
    #Final Dataframe columns for P&Ls
    AccountNumbers = All_Accounts["ACCT #"].astype(int)
    Account_Names = All_Accounts["ACCOUNT NAME"]
    Column1  = Select_Store_Data.loc[(Select_Store_Data["GL YEAR"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*24)).year) & (Select_Store_Data["Fiscal Month"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*24)).month)].pivot_table(index = ['ACCT #'],   values ='GL AMOUNT', aggfunc='sum')
    Column2  = Select_Store_Data.loc[(Select_Store_Data["GL YEAR"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*23)).year) & (Select_Store_Data["Fiscal Month"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*23)).month)].pivot_table(index = ['ACCT #'],   values ='GL AMOUNT', aggfunc='sum')
    Column3  = Select_Store_Data.loc[(Select_Store_Data["GL YEAR"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*22)).year) & (Select_Store_Data["Fiscal Month"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*22)).month)].pivot_table(index = ['ACCT #'],   values ='GL AMOUNT', aggfunc='sum')
    Column4  = Select_Store_Data.loc[(Select_Store_Data["GL YEAR"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*21)).year) & (Select_Store_Data["Fiscal Month"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*21)).month)].pivot_table(index = ['ACCT #'],   values ='GL AMOUNT', aggfunc='sum')
    Column5  = Select_Store_Data.loc[(Select_Store_Data["GL YEAR"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*20)).year) & (Select_Store_Data["Fiscal Month"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*20)).month)].pivot_table(index = ['ACCT #'],   values ='GL AMOUNT', aggfunc='sum')
    Column6  = Select_Store_Data.loc[(Select_Store_Data["GL YEAR"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*19)).year) & (Select_Store_Data["Fiscal Month"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*19)).month)].pivot_table(index = ['ACCT #'],   values ='GL AMOUNT', aggfunc='sum')
    Column7  = Select_Store_Data.loc[(Select_Store_Data["GL YEAR"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*18)).year) & (Select_Store_Data["Fiscal Month"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*18)).month)].pivot_table(index = ['ACCT #'],   values ='GL AMOUNT', aggfunc='sum')
    Column8  = Select_Store_Data.loc[(Select_Store_Data["GL YEAR"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*17)).year) & (Select_Store_Data["Fiscal Month"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*17)).month)].pivot_table(index = ['ACCT #'],   values ='GL AMOUNT', aggfunc='sum')
    Column9  = Select_Store_Data.loc[(Select_Store_Data["GL YEAR"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*16)).year) & (Select_Store_Data["Fiscal Month"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*16)).month)].pivot_table(index = ['ACCT #'],   values ='GL AMOUNT', aggfunc='sum')
    Column10 = Select_Store_Data.loc[(Select_Store_Data["GL YEAR"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*15)).year) & (Select_Store_Data["Fiscal Month"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*15)).month)].pivot_table(index = ['ACCT #'],   values ='GL AMOUNT', aggfunc='sum')
    Column11 = Select_Store_Data.loc[(Select_Store_Data["GL YEAR"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*14)).year) & (Select_Store_Data["Fiscal Month"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*14)).month)].pivot_table(index = ['ACCT #'],   values ='GL AMOUNT', aggfunc='sum')
    Column12 = Select_Store_Data.loc[(Select_Store_Data["GL YEAR"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*13)).year) & (Select_Store_Data["Fiscal Month"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*13)).month)].pivot_table(index = ['ACCT #'],   values ='GL AMOUNT', aggfunc='sum')
    Column13 = Select_Store_Data.loc[(Select_Store_Data["GL YEAR"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*12)).year) & (Select_Store_Data["Fiscal Month"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*12)).month)].pivot_table(index = ['ACCT #'],   values ='GL AMOUNT', aggfunc='sum')
    Column14 = Select_Store_Data.loc[(Select_Store_Data["GL YEAR"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*11)).year) & (Select_Store_Data["Fiscal Month"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*11)).month)].pivot_table(index = ['ACCT #'],   values ='GL AMOUNT', aggfunc='sum')
    Column15 = Select_Store_Data.loc[(Select_Store_Data["GL YEAR"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*10)).year) & (Select_Store_Data["Fiscal Month"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*10)).month)].pivot_table(index = ['ACCT #'],   values ='GL AMOUNT', aggfunc='sum')
    Column16 = Select_Store_Data.loc[(Select_Store_Data["GL YEAR"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*9 )).year) & (Select_Store_Data["Fiscal Month"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*9 )).month)].pivot_table(index = ['ACCT #'],   values ='GL AMOUNT', aggfunc='sum')
    Column17 = Select_Store_Data.loc[(Select_Store_Data["GL YEAR"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*8 )).year) & (Select_Store_Data["Fiscal Month"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*8 )).month)].pivot_table(index = ['ACCT #'],   values ='GL AMOUNT', aggfunc='sum')
    Column18 = Select_Store_Data.loc[(Select_Store_Data["GL YEAR"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*7 )).year) & (Select_Store_Data["Fiscal Month"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*7 )).month)].pivot_table(index = ['ACCT #'],   values ='GL AMOUNT', aggfunc='sum')
    Column19 = Select_Store_Data.loc[(Select_Store_Data["GL YEAR"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*6 )).year) & (Select_Store_Data["Fiscal Month"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*6 )).month)].pivot_table(index = ['ACCT #'],   values ='GL AMOUNT', aggfunc='sum')
    Column20 = Select_Store_Data.loc[(Select_Store_Data["GL YEAR"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*5 )).year) & (Select_Store_Data["Fiscal Month"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*5 )).month)].pivot_table(index = ['ACCT #'],   values ='GL AMOUNT', aggfunc='sum')
    Column21 = Select_Store_Data.loc[(Select_Store_Data["GL YEAR"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*4 )).year) & (Select_Store_Data["Fiscal Month"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*4 )).month)].pivot_table(index = ['ACCT #'],   values ='GL AMOUNT', aggfunc='sum')
    Column22 = Select_Store_Data.loc[(Select_Store_Data["GL YEAR"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*3 )).year) & (Select_Store_Data["Fiscal Month"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*3 )).month)].pivot_table(index = ['ACCT #'],   values ='GL AMOUNT', aggfunc='sum')
    Column23 = Select_Store_Data.loc[(Select_Store_Data["GL YEAR"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*2 )).year) & (Select_Store_Data["Fiscal Month"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*2 )).month)].pivot_table(index = ['ACCT #'],   values ='GL AMOUNT', aggfunc='sum')
    Column24 = Select_Store_Data.loc[(Select_Store_Data["GL YEAR"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*1 )).year) & (Select_Store_Data["Fiscal Month"] == (FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*1 )).month)].pivot_table(index = ['ACCT #'],   values ='GL AMOUNT', aggfunc='sum')
    
    print("Assigning Columns. . .")
    try:
      Column1.columns =    [f"{FiscalCalendar.int_to_month_string(CurrentFiscalMonth-24) } {str(FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*24))[:7]}" ]
    except:
      pass
    try:
      Column2.columns =    [f"{FiscalCalendar.int_to_month_string(CurrentFiscalMonth-23) } {str(FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*23))[:7]}" ]
    except:
      pass    
    try:
      Column3.columns =    [f"{FiscalCalendar.int_to_month_string(CurrentFiscalMonth-22) } {str(FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*22))[:7]}" ]
    except:
      pass    
    try:
      Column4.columns =    [f"{FiscalCalendar.int_to_month_string(CurrentFiscalMonth-21) } {str(FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*21))[:7]}" ]
    except:
      pass
    try:
      Column5.columns =    [f"{FiscalCalendar.int_to_month_string(CurrentFiscalMonth-20) } {str(FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*20))[:7]}" ]
    except:
      pass
    try:
      Column6.columns =    [f"{FiscalCalendar.int_to_month_string(CurrentFiscalMonth-19) } {str(FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*19))[:7]}" ]                   
    except:
      pass
    try:
      Column7.columns =    [f"{FiscalCalendar.int_to_month_string(CurrentFiscalMonth-18) } {str(FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*18))[:7]}" ]
    except:
      pass
    try:
      Column8.columns =    [f"{FiscalCalendar.int_to_month_string(CurrentFiscalMonth-17) } {str(FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*17))[:7]}" ]
    except:
      pass
    try:
      Column9.columns =    [f"{FiscalCalendar.int_to_month_string(CurrentFiscalMonth-16) } {str(FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*16))[:7]}" ]                
    except:
      pass
    try:
      Column10.columns =   [f"{FiscalCalendar.int_to_month_string(CurrentFiscalMonth-15) } {str(FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*15))[:7]}" ]
    except:
      pass
    try:
      Column11.columns =   [f"{FiscalCalendar.int_to_month_string(CurrentFiscalMonth-14) } {str(FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*14))[:7]}" ]
    except:
      pass
    try:
      Column12.columns =   [f"{FiscalCalendar.int_to_month_string(CurrentFiscalMonth-13) } {str(FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*13))[:7]}" ]
    except:
      pass
    try:
      Column13.columns =   [f"{FiscalCalendar.int_to_month_string(CurrentFiscalMonth-12) } {str(FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*12))[:7]}" ]
    except:
      pass
    try:
      Column14.columns =   [f"{FiscalCalendar.int_to_month_string(CurrentFiscalMonth-11) } {str(FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*11))[:7]}" ]
    except:
      pass
    try:
      Column15.columns =   [f"{FiscalCalendar.int_to_month_string(CurrentFiscalMonth-10) } {str(FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*10))[:7]}" ]
    except:
      pass
    try:
      Column16.columns =   [f"{FiscalCalendar.int_to_month_string(CurrentFiscalMonth-9)  } {str(FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*9 ))[:7]}" ]
    except:
      pass
    try:
      Column17.columns =   [f"{FiscalCalendar.int_to_month_string(CurrentFiscalMonth-8)  } {str(FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*8 ))[:7]}" ]
    except:
      pass
    try:
      Column18.columns =   [f"{FiscalCalendar.int_to_month_string(CurrentFiscalMonth-7)  } {str(FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*7 ))[:7]}" ]
    except:
      pass
    try:
      Column19.columns =   [f"{FiscalCalendar.int_to_month_string(CurrentFiscalMonth-6)  } {str(FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*6 ))[:7]}" ]                    
    except:
      pass
    try:
      Column20.columns =   [f"{FiscalCalendar.int_to_month_string(CurrentFiscalMonth-5)  } {str(FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*5 ))[:7]}" ]
    except:
      pass
    try:
      Column21.columns =   [f"{FiscalCalendar.int_to_month_string(CurrentFiscalMonth-4)  } {str(FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*4 ))[:7]}" ]
    except:
      pass
    try:
      Column22.columns =   [f"{FiscalCalendar.int_to_month_string(CurrentFiscalMonth-3)  } {str(FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*3 ))[:7]}" ]             
    except:
      pass
    try:
      Column23.columns =   [f"{FiscalCalendar.int_to_month_string(CurrentFiscalMonth-2)  } {str(FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*2 ))[:7]}" ]
    except:
      pass
    try:
      Column24.columns =   [f"{FiscalCalendar.int_to_month_string(CurrentFiscalMonth-1)  } {str(FiscalCalendar.CurrentDate + dt.timedelta(weeks = (-52/12)*1 ))[:7]}" ]
    except:
      pass
    
    print(Column1)
    
    #Create the dataframe shell
    print("Created Dataframe shell")
    data_frame = pd.DataFrame({'ACCT #':AccountNumbers,'ACCOUNT NAME':Account_Names})
    Code_to_execute = f"Final_Df{store} = eval('data_frame')"
    #Create the store PL
    exec(Code_to_execute)
    
    
    #Populate the dataframe with the data
    print("Populating Dataframe")
    for i in range(24):
      exec(f'Final_Df{store} = pd.merge(Final_Df{store},eval(f"Column{i+1}"),  how = "outer", on = "ACCT #")')
     
#      #Set the indexes to the account number
#      DF_with_new_index = eval(f'Final_Df{store}.set_index("ACCT #")')
#      exec(f'Final_Df{store} = DF_with_new_index')
      





#Subtotals to make
#Services_Revenue
#Cost of Goods Sold

#Org 142

#To make subtotals
#Cut the big dataframe into groups
#Then make subtotal rows for those groups
#THen append the all in an rbind fashion
      
           
#The ebitda add_back flag is like a depreciation flag.
#EBITDA/SALES FOR BRENDA 5 years



#Strong stores selected by Brenda
#Selected_Stores = [120,132,139,147,159,166,169,233,237,275,103,121,128,153,236,241,245,273,280]

#Select All Stores. Or you can replace this
Selected_Stores = [store for store in Store_list["Store #"]]

os.chdir(f"//data/accounting/Janette/GL Cube")
GL_2015 = pd.read_excel("GL Cube 2015.xlsx", "2015")
GL_2016 = pd.read_excel("GL Cube 2016.xlsx", "2016")

#Rename to match
Merged_Data.columns = ['GL YEAR', 'PERIOD', 'DEPT #', 'DEPT NAME', 'ACCT #',
       'ACCOUNT NAME', 'GL AMOUNT', 'EBITDA Add-back', 'Account Group',
       'Account Group #', 'LOCATION TYPE', 'Year Opened', 'Region', 'District',
       'Unnamed: 14', 'Unnamed: 15']

Merged_DataFY15_FY19 = pd.concat([Merged_Data,GL_2015,GL_2016],join = "outer") 


                               
#Change from EBIT to EBITDA or visa versa
EBITDA_EBIT_Selection = "EBITDA"

if EBITDA_EBIT_Selection == "EBIT":
  Revenue =             Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Revenue")]            [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]                     
  Cost_of_Goods_Sold =  Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Cost of Goods Sold")] [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]   
  Payroll =             Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Payroll")]            [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Rent =                Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Rent")]               [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Other_Occupancy =     Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Other Occupancy")]    [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Advertising =         Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Advertising")]        [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Tender_Expense =      Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Tender Expense")]     [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Store_Expense =       Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Store Expense")]      [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Office_Expense =      Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Office Expense")]     [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Warehouse_Expense =   Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Warehouse Expense")]  [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Maintenance_Expense = Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Maintenace Expense")] [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]             
elif EBITDA_EBIT_Selection == "EBITDA":
  Revenue =             Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Revenue")            & ((Merged_DataFY15_FY19["EBITDA Add-back"] == "No") | (Merged_DataFY15_FY19["EBITDA Add-back"] == "Not in PV Store P&L") )]  [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]                        
  Cost_of_Goods_Sold =  Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Cost of Goods Sold") & ((Merged_DataFY15_FY19["EBITDA Add-back"] == "No") | (Merged_DataFY15_FY19["EBITDA Add-back"] == "Not in PV Store P&L") )]  [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Payroll =             Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Payroll")            & ((Merged_DataFY15_FY19["EBITDA Add-back"] == "No") | (Merged_DataFY15_FY19["EBITDA Add-back"] == "Not in PV Store P&L") )]  [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Rent =                Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Rent")               & ((Merged_DataFY15_FY19["EBITDA Add-back"] == "No") | (Merged_DataFY15_FY19["EBITDA Add-back"] == "Not in PV Store P&L") )]  [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Other_Occupancy =     Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Other Occupancy")    & ((Merged_DataFY15_FY19["EBITDA Add-back"] == "No") | (Merged_DataFY15_FY19["EBITDA Add-back"] == "Not in PV Store P&L") )]  [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Advertising =         Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Advertising")        & ((Merged_DataFY15_FY19["EBITDA Add-back"] == "No") | (Merged_DataFY15_FY19["EBITDA Add-back"] == "Not in PV Store P&L") )]  [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Tender_Expense =      Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Tender Expense")     & ((Merged_DataFY15_FY19["EBITDA Add-back"] == "No") | (Merged_DataFY15_FY19["EBITDA Add-back"] == "Not in PV Store P&L") )]  [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Store_Expense =       Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Store Expense")      & ((Merged_DataFY15_FY19["EBITDA Add-back"] == "No") | (Merged_DataFY15_FY19["EBITDA Add-back"] == "Not in PV Store P&L") )]  [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Office_Expense =      Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Office Expense")     & ((Merged_DataFY15_FY19["EBITDA Add-back"] == "No") | (Merged_DataFY15_FY19["EBITDA Add-back"] == "Not in PV Store P&L") )]  [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Warehouse_Expense =   Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Warehouse Expense")  & ((Merged_DataFY15_FY19["EBITDA Add-back"] == "No") | (Merged_DataFY15_FY19["EBITDA Add-back"] == "Not in PV Store P&L") )]  [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Maintenance_Expense = Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Maintenace Expense") & ((Merged_DataFY15_FY19["EBITDA Add-back"] == "No") | (Merged_DataFY15_FY19["EBITDA Add-back"] == "Not in PV Store P&L") )]  [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]   


Revenue_pivot =             Revenue.pivot_table(            index = "DEPT #", columns = ["GL YEAR","PERIOD"], values = "GL AMOUNT", aggfunc = "sum")
Cost_of_Goods_Sold_Pivot =  Cost_of_Goods_Sold.pivot_table( index = "DEPT #", columns = ["GL YEAR","PERIOD"], values = "GL AMOUNT", aggfunc = "sum")
Payroll_Pivot =             Payroll.pivot_table(            index = "DEPT #", columns = ["GL YEAR","PERIOD"], values = "GL AMOUNT", aggfunc = "sum")
Rent_Pivot =                Rent.pivot_table(               index = "DEPT #", columns = ["GL YEAR","PERIOD"], values = "GL AMOUNT", aggfunc = "sum")
Other_Occupancy_Pivot =     Other_Occupancy.pivot_table(    index = "DEPT #", columns = ["GL YEAR","PERIOD"], values = "GL AMOUNT", aggfunc = "sum")
Advertising_Pivot =         Advertising.pivot_table(        index = "DEPT #", columns = ["GL YEAR","PERIOD"], values = "GL AMOUNT", aggfunc = "sum")
Tender_Expense_Pivot =      Tender_Expense.pivot_table(     index = "DEPT #", columns = ["GL YEAR","PERIOD"], values = "GL AMOUNT", aggfunc = "sum")
Store_Expense_Pivot =       Store_Expense.pivot_table(      index = "DEPT #", columns = ["GL YEAR","PERIOD"], values = "GL AMOUNT", aggfunc = "sum")
Office_Expense_Pivot =      Office_Expense.pivot_table(     index = "DEPT #", columns = ["GL YEAR","PERIOD"], values = "GL AMOUNT", aggfunc = "sum")
Warehouse_Expense_Pivot =   Warehouse_Expense.pivot_table(  index = "DEPT #", columns = ["GL YEAR","PERIOD"], values = "GL AMOUNT", aggfunc = "sum")
Maintenance_Expense_Pivot = Maintenance_Expense.pivot_table(index = "DEPT #", columns = ["GL YEAR","PERIOD"], values = "GL AMOUNT", aggfunc = "sum")                                
                                              
                                                             
EBITDA_Or_EBIT_Account_List = [Payroll_Pivot,Rent_Pivot,Other_Occupancy_Pivot,Advertising_Pivot,Tender_Expense_Pivot,Store_Expense_Pivot,Office_Expense_Pivot,
                               Warehouse_Expense_Pivot,Maintenance_Expense_Pivot]                              

EBITDA_Or_EBIT = Revenue_pivot.fillna(0).add(Cost_of_Goods_Sold_Pivot.fillna(0)) #first instance
for account in EBITDA_Or_EBIT_Account_List:
  Append_Account = account.fillna(0)
  EBITDA_Or_EBIT =   EBITDA_Or_EBIT.add(Append_Account,fill_value = 0)  
  print(Append_Account.iloc[:,-5:])
print(f"Finished calculating {EBITDA_EBIT_Selection}")

os.chdir("C:/Users/mfrangos/Desktop/Michael's GL Cube/Output Numbers")
Revenue_pivot.multiply(-1).to_excel("Revenue.xlsx", merge_cells = False)
EBITDA_Or_EBIT.multiply(-1).to_excel("EBITDA.xlsx", merge_cells = False)









#Monthly pivots of account data
if EBITDA_EBIT_Selection == "EBIT":
  Revenue =             Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Revenue")]            [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]                       
  Cost_of_Goods_Sold =  Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Cost of Goods Sold")] [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]   
  Payroll =             Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Payroll")]            [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Rent =                Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Rent")]               [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Other_Occupancy =     Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Other Occupancy")]    [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Advertising =         Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Advertising")]        [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Tender_Expense =      Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Tender Expense")]     [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Store_Expense =       Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Store Expense")]      [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Office_Expense =      Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Office Expense")]     [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Warehouse_Expense =   Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Warehouse Expense")]  [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Maintenance_Expense = Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Maintenace Expense")] [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]             
elif EBITDA_EBIT_Selection == "EBITDA":
  Revenue =             Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Revenue")            & ((Merged_DataFY15_FY19["EBITDA Add-back"] == "No") | (Merged_DataFY15_FY19["EBITDA Add-back"] == "Not in PV Store P&L") )]  [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]               
  Cost_of_Goods_Sold =  Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Cost of Goods Sold") & ((Merged_DataFY15_FY19["EBITDA Add-back"] == "No") | (Merged_DataFY15_FY19["EBITDA Add-back"] == "Not in PV Store P&L") )]  [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Payroll =             Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Payroll")            & ((Merged_DataFY15_FY19["EBITDA Add-back"] == "No") | (Merged_DataFY15_FY19["EBITDA Add-back"] == "Not in PV Store P&L") )]  [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Rent =                Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Rent")               & ((Merged_DataFY15_FY19["EBITDA Add-back"] == "No") | (Merged_DataFY15_FY19["EBITDA Add-back"] == "Not in PV Store P&L") )]  [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Other_Occupancy =     Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Other Occupancy")    & ((Merged_DataFY15_FY19["EBITDA Add-back"] == "No") | (Merged_DataFY15_FY19["EBITDA Add-back"] == "Not in PV Store P&L") )]  [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Advertising =         Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Advertising")        & ((Merged_DataFY15_FY19["EBITDA Add-back"] == "No") | (Merged_DataFY15_FY19["EBITDA Add-back"] == "Not in PV Store P&L") )]  [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Tender_Expense =      Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Tender Expense")     & ((Merged_DataFY15_FY19["EBITDA Add-back"] == "No") | (Merged_DataFY15_FY19["EBITDA Add-back"] == "Not in PV Store P&L") )]  [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Store_Expense =       Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Store Expense")      & ((Merged_DataFY15_FY19["EBITDA Add-back"] == "No") | (Merged_DataFY15_FY19["EBITDA Add-back"] == "Not in PV Store P&L") )]  [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Office_Expense =      Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Office Expense")     & ((Merged_DataFY15_FY19["EBITDA Add-back"] == "No") | (Merged_DataFY15_FY19["EBITDA Add-back"] == "Not in PV Store P&L") )]  [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Warehouse_Expense =   Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Warehouse Expense")  & ((Merged_DataFY15_FY19["EBITDA Add-back"] == "No") | (Merged_DataFY15_FY19["EBITDA Add-back"] == "Not in PV Store P&L") )]  [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]         
  Maintenance_Expense = Merged_DataFY15_FY19.loc[(Merged_DataFY15_FY19["DEPT #"].isin(Selected_Stores))  & (Merged_DataFY15_FY19["Account Group"] == "Maintenace Expense") & ((Merged_DataFY15_FY19["EBITDA Add-back"] == "No") | (Merged_DataFY15_FY19["EBITDA Add-back"] == "Not in PV Store P&L") )]  [["DEPT #","GL AMOUNT","GL YEAR", "PERIOD"]]   

#Summary Pivots for EBIT/EBITDA
Revenue_pivot =             Revenue.pivot_table(            index = "DEPT #", columns = ["GL YEAR"], values = "GL AMOUNT", aggfunc = "sum")
Cost_of_Goods_Sold_Pivot =  Cost_of_Goods_Sold.pivot_table( index = "DEPT #", columns = ["GL YEAR"], values = "GL AMOUNT", aggfunc = "sum")
Payroll_Pivot =             Payroll.pivot_table(            index = "DEPT #", columns = ["GL YEAR"], values = "GL AMOUNT", aggfunc = "sum")
Rent_Pivot =                Rent.pivot_table(               index = "DEPT #", columns = ["GL YEAR"], values = "GL AMOUNT", aggfunc = "sum")
Other_Occupancy_Pivot =     Other_Occupancy.pivot_table(    index = "DEPT #", columns = ["GL YEAR"], values = "GL AMOUNT", aggfunc = "sum")
Advertising_Pivot =         Advertising.pivot_table(        index = "DEPT #", columns = ["GL YEAR"], values = "GL AMOUNT", aggfunc = "sum")
Tender_Expense_Pivot =      Tender_Expense.pivot_table(     index = "DEPT #", columns = ["GL YEAR"], values = "GL AMOUNT", aggfunc = "sum")
Store_Expense_Pivot =       Store_Expense.pivot_table(      index = "DEPT #", columns = ["GL YEAR"], values = "GL AMOUNT", aggfunc = "sum")
Office_Expense_Pivot =      Office_Expense.pivot_table(     index = "DEPT #", columns = ["GL YEAR"], values = "GL AMOUNT", aggfunc = "sum")
Warehouse_Expense_Pivot =   Warehouse_Expense.pivot_table(  index = "DEPT #", columns = ["GL YEAR"], values = "GL AMOUNT", aggfunc = "sum")
Maintenance_Expense_Pivot = Maintenance_Expense.pivot_table(index = "DEPT #", columns = ["GL YEAR"], values = "GL AMOUNT", aggfunc = "sum")                                
                                              
                                                             
EBITDA_Or_EBIT_Account_List = [Payroll_Pivot,Rent_Pivot,Other_Occupancy_Pivot,Advertising_Pivot,Tender_Expense_Pivot,Store_Expense_Pivot,Office_Expense_Pivot,
                               Warehouse_Expense_Pivot,Maintenance_Expense_Pivot]                              

EBITDA_Or_EBIT = Revenue_pivot.fillna(0).add(Cost_of_Goods_Sold_Pivot.fillna(0)) #first instance
for account in EBITDA_Or_EBIT_Account_List:
  Append_Account = account.fillna(0) #Prepare to append
  EBITDA_Or_EBIT =   EBITDA_Or_EBIT.add(Append_Account, fill_value = 0)  
print(f"Finished calculating {EBITDA_EBIT_Selection}")

os.chdir("C:/Users/mfrangos/Desktop/Michael's GL Cube/Output Numbers")
Revenue_pivot.multiply(-1).to_excel("Revenue Summary.xlsx", merge_cells = False)
EBITDA_Or_EBIT.multiply(-1).to_excel("EBITDA Summary.xlsx", merge_cells = False)

#Debug
for item in [Payroll,Rent,Other_Occupancy,Advertising,Tender_Expense,Store_Expense,Office_Expense,Warehouse_Expense,Maintenance_Expense]:
  print(set(item.loc[item["DEPT #"] == 103]["GL YEAR"]))

Revenue_pivot.add(Cost_of_Goods_Sold_Pivot).add(Payroll_Pivot).add(Rent_Pivot).add(Other_Occupancy_Pivot).add(Advertising_Pivot).add(Tender_Expense_Pivot).add(Store_Expense_Pivot).add(
    Office_Expense_Pivot).add(Warehouse_Expense_Pivot).add(Maintenance_Expense_Pivot)


groomer_sales_by_store = pd.DataFrame()
for store_PL in List_of_Final_DFs:
  Data = store_PL.loc[store_PL["ACCOUNT NAME"] == "Sales - Groomers              "]
  groomer_sales_by_store = pd.concat([groomer_sales_by_store,Data], axis=1)

