# Python script to utilize PLEXOS API to automate FDRE Tariff

import os, sys, re, clr
import subprocess as sp

import pandas as pd
from datetime import datetime
# from openpyxl import *
import openpyxl
import builtins
import xlwings # update formula in excel sheet

from shutil import copyfile

plexospath = r"C:/Program Files/Energy Exemplar/PLEXOS 10.0 API"
sys.path.append(plexospath)
clr.AddReference('PLEXOS_NET.Core')
clr.AddReference('EEUTILITY')
clr.AddReference('EnergyExemplar.PLEXOS.Utility')

# .NET related imports
from PLEXOS_NET.Core import DatabaseCore, Solution, PLEXOSConnect
from EEUTILITY.Enums import *
from EnergyExemplar.PLEXOS.Utility.Enums import *
from System import Enum, DateTime, String

# read values from csv outputs from Scenario A1
def read_values(fn_sales, fn_purchase):
    # empty dataframe for JKM values
    df_JKM = pd.DataFrame()

    # read and store JKM sales at monthly intervals
    df_sales1 = pd.read_csv(fn_sales)
    # print(df_sales1["Value"])
    df_JKM["Year"] = df_sales1["Year"]
    df_JKM["Month"] = df_sales1["Month"]
    df_JKM["Sales"] = df_sales1["Value"]
    df_JKM["Day"] = '1'

    # read and store JKM Purchases at monthly intervals
    df_purchases1 = pd.read_csv(fn_purchase)
    # print(df_purchases1["Value"])
    df_JKM["Purchases"] = df_purchases1["Value"]
    df_sales1["Date"] = pd.to_datetime(dict(year=df_sales1.Year, month=df_sales1.Month, day=df_sales1.Day)) # no warnings, format='%d/%m/%Y'
    df_JKM["Date"] = df_sales1["Date"].dt.strftime('%d/%m/%Y')

    return df_JKM

def update_dataset(original_ds, df_JKM, No_of_Steps, Sc_name, Add_Purch): #
    if os.path.exists(original_ds):

        # # delete the modified file if it already exists
        # if os.path.exists(new_ds):
        #     os.remove(new_ds)

        # # copy the PLEXOS input file
        # copyfile(original_ds, new_ds)
        
        # Create an object to store the input data
        db = DatabaseCore()
        db.Connection(original_ds)

        # Get the System.Generators membership ID for this new generator
        '''
        Int32 GetMembershipID(
            CollectionEnum nCollectionId,
            String strParent,
            String strChild
            )    
        '''


        # Add an object (and the System Membership)
        '''
        Int32 AddObject(
            String strName,
            ClassEnum nClassId,
            Boolean bAddSystemMembership,
            String strCategory[ = None],
            String strDescription[ = None]
            )
        '''
        # create new scenario and model
        classDict = db.FetchAllClassIds()
        db.AddObject(Sc_name, classDict["Scenario"], True)
        #create copy of A3-base and use for this model
        db.CopyObject("Scenario A3 - base", "Model "+Sc_name, classDict["Model"])
        # db.AddObject("Model "+Sc_name, classDict["Model"], True, "Scenario A3")

        # add membership from new model and scenario
        collectionDict = db.FetchAllCollectionIds()
        db.AddMembership(collectionDict["ModelScenarios"], "Model "+Sc_name, Sc_name)

        # mem_id = db.GetMembershipID(CollectionEnum.SystemConstraints, 'System', 'JKM Sales Max')
        # print("Constraint ID is ", mem_id)
        collectionDict = db.FetchAllCollectionIds()
        # mem_id = db.GetMembershipID(collectionDict["SystemConstraints"], 'System', 'JKM Sales Max')
                                    
        # Add properties
        '''
        Int32 AddProperty(
            Int32 MembershipId,
            Int32 EnumId,
            Int32 BandId,
            Double Value,
            Object DateFrom,
            Object DateTo,
            Object Variable,
            Object DataFile,
            Object Pattern,
            Object Scenario,
            Object Action,
            PeriodEnum PeriodTypeId
            )
        
        Also we need to obtain the EnumId for each property
        that we intend to add
        '''
        flag = 0
        for t in range(No_of_Steps):
            # identify constraint - Sales Max
            mem_id = db.GetMembershipID(collectionDict["SystemConstraints"], 'System', 'JKM Sales Max')

            # Add a property
            nPropId = db.PropertyName2EnumId("System", "Constraint", "Constraints", "RHS")
            flag  = db.AddProperty(mem_id, int(SystemConstraintsEnum.RHS), \
                    1, df_JKM["Sales"][t]+1, df_JKM["Date"][t], None, None, None, None, Sc_name, \
                    1, PeriodEnum.Month)

            # identify constraint - Sales Min
            mem_id = db.GetMembershipID(collectionDict["SystemConstraints"], 'System', 'JKM Sales Min')

            # Add a property
            nPropId = db.PropertyName2EnumId("System", "Constraint", "Constraints", "RHS")
            flag = db.AddProperty(mem_id, int(SystemConstraintsEnum.RHS), \
                                  1, df_JKM["Sales"][t]-1, df_JKM["Date"][t], None, None, None, None, Sc_name, \
                                  1, PeriodEnum.Month)

            # identify constraint - Purchase Max
            mem_id = db.GetMembershipID(collectionDict["SystemConstraints"], 'System', 'JKM Purchases Max')

            # Add a property
            nPropId = db.PropertyName2EnumId("System", "Constraint", "Constraints", "RHS")
            flag = db.AddProperty(mem_id, int(SystemConstraintsEnum.RHS), \
                                  1, df_JKM["Purchases"][t]+1+Add_Purch[t], df_JKM["Date"][t], None, None, None, None, Sc_name, \
                                  1, PeriodEnum.Month)

            # identify constraint - Purchases min
            mem_id = db.GetMembershipID(collectionDict["SystemConstraints"], 'System', 'JKM Purchases Min')

            # Add a property
            nPropId = db.PropertyName2EnumId("System", "Constraint", "Constraints", "RHS")
            flag = db.AddProperty(mem_id, int(SystemConstraintsEnum.RHS), \
                                  1, df_JKM["Purchases"][t]-1+Add_Purch[t], df_JKM["Date"][t], None, None, None, None, Sc_name, \
                                  1, PeriodEnum.Month)

        if flag == 4:
            print("A. Property update complete\n")
        
        # save the data set
        db.Close()

#Execute = db.GetModelsToExecute('rts2.xml',\
                            #'Base',\
                           # '')
							
#Execute the model 
def run_model(plexospath, filename, foldername, modelname):
# def run_model(plexospath, filename, foldername, modelname, username, password): #with credentials
    # launch the model on the local desktop
    # The \n argument is very important because it allows the PLEXOS
    # engine to terminate after completing the simulation
    print("B. Run the model")
    sp.call([os.path.join(plexospath, 'PLEXOS64.exe'), filename, r'\n', r'\o', foldername, r'\m', modelname])
    # sp.call([os.path.join(plexospath, 'PLEXOS64.exe'), filename, r'\n', r'\o', foldername, r'\m', modelname, r'\cu', username, r'\cp', password]) #with username
    

def parse_logfile(pattern, foldername, modelname, linecount = 1):
    
    currentlines = 0
    lines = []
    regex = re.compile(pattern)
    
    for line in builtins.open(os.path.join(foldername, 'Model {} Solution'.format(modelname), 'Model ( {} ) Log.txt'.format(modelname))):
        if len(regex.findall(line)) > 0:
            currentlines = linecount
            
        if currentlines > 0:
            lines.append(line)
            currentlines -= 1
            
        if currentlines == 0 and len(lines) > 0:
            retval = '\n'.join(lines)
            lines = []
            yield retval

def query_results(sol_file):
    #Query Results
    # Create a PLEXOS solution file object and load the solution
    if not os.path.exists(sol_file):
        print('No such file')
        return
        
    sol = Solution()
    sol.Connection(sol_file)
    sol.DisplayAlerts = False

    # print("Update EXCEL sheet with new results")

    # # connect to excel
    # book = load_workbook(xlsx_file)
    # writer = pd.ExcelWriter(xlsx_file, mode='a', if_sheet_exists='overlay', engine='openpyxl')

    '''
    Simple query: works similarly to PLEXOS Solution Viewer

    Solution.Query(phase, collection, parent, child, period, series, props)
        phase -> SimulationPhaseEnum
        collection -> CollectionEnum
        parent -> the name of a parent object or ''
        child -> the name of a child object or ''
        period -> PeriodEnum
        series -> SeriesTypeEnum
        props -> a string containing an integer indicating the Property to query or ''
    returns a ADODB recordset... however, you don't *need* to worry about that...
    '''

    '''
    Simple query: works similarly to PLEXOS Solution Viewer
    
    QueryToList(
    	SimulationPhaseEnum SimulationPhaseId,
    	CollectionEnum CollectionId,
    	String ParentName,
    	String ChildName,
    	PeriodEnum PeriodTypeId,
    	SeriesTypeEnum SeriesTypeId,
    	String PropertyList[ = None],
    	Object DateFrom[ = None],
    	Object DateTo[ = None],
    	String TimesliceList[ = None],
    	String SampleList[ = None],
    	String ModelName[ = None],
    	AggregationEnum AggregationType[ = None],
    	String Category[ = None],
    	String Filter[ = None]
    	)
    '''

    # NOTE: Because None is a reserved word in Python we must use the Parse() method to get the value of AggregationEnum.None
    aggregation_none = Enum.Parse(clr.GetClrType(AggregationTypeEnum), "None")
    # aggregation_object = Enum.Parse(clr.GetClrType(AggregationTypeEnum), "Object")

    # read outputs and save to Dataframe
    # # Collect market cost
    propId1 = sol.PropertyName2EnumId("System", "Market", "Markets", "Cost")
    propId2 = sol.PropertyName2EnumId("System", "Market", "Markets", "Revenue")
    # print("Market Sales")
    # print(propId3)
    collectionDict = sol.FetchAllCollectionIds()
    listresultsJKMC = sol.QueryToList(SimulationPhaseEnum.MTSchedule, \
              collectionDict["SystemMarkets"], \
              '', \
              '', \
              PeriodEnum.Month, \
              SeriesTypeEnum.Names, \
              str(propId1), \
              DateTime.Parse('2022-04-01'), \
              DateTime.Parse('2023-03-31'), \
                                     '0', \
                                     '', \
                                     '', \
                                     aggregation_none, \
                                     'Gas Market')
    # print("Printing names of columns")
    # for test in range(35):
    #     print(listresultsMS.Columns[test])
    #     print(test)

    if listresultsJKMC is None:
        print('No results')
    else:
        # Create a DataFrame with a column for each column in the results
        # columns = "Spot market (JKM)"
        # columns = [r'Spot market \B\JKM\D\\']
        columns = [listresultsJKMC.Columns[33]]
        # print(columns)
        df_JKMC = pd.DataFrame([[row.GetProperty.Overloads[String](n) for n in columns] for row in listresultsJKMC], columns=columns)

    # print(df_JKMC)

    # JKM revenue
    listresultsJKMR = sol.QueryToList(SimulationPhaseEnum.MTSchedule, \
              collectionDict["SystemMarkets"], \
              '', \
              '', \
              PeriodEnum.Month, \
              SeriesTypeEnum.Names, \
              str(propId2), \
              DateTime.Parse('2022-04-01'), \
              DateTime.Parse('2023-03-31'), \
                                     '0', \
                                     '', \
                                     '', \
                                     aggregation_none, \
                                     'Gas Market')
    if listresultsJKMR is None:
        print('No results')
    else:
        # Create a DataFrame with a column for each column in the results
        # columns = "Spot market (JKM)"
        # columns = [r'Spot market \B\JKM\D\\']
        columns = [listresultsJKMR.Columns[33]]
        # print(columns)
        df_JKMR = pd.DataFrame([[row.GetProperty.Overloads[String](n) for n in columns] for row in listresultsJKMR], columns=columns)
    # print("revenue")
    # print(df_JKMR)

    # JEPX sales
    listresultsJEPXR = sol.QueryToList(SimulationPhaseEnum.MTSchedule, \
              collectionDict["SystemMarkets"], \
              '', \
              '', \
              PeriodEnum.Month, \
              SeriesTypeEnum.Names, \
              str(propId2), \
              DateTime.Parse('2022-04-01'), \
              DateTime.Parse('2023-03-31'), \
                                     '0', \
                                     '', \
                                     '', \
                                     aggregation_none, \
                                     'Electricity Market')
    # print("Printing names of columns")
    # for test in range(35):
    #     print(listresultsMS.Columns[test])
    #     print(test)

    if listresultsJEPXR is None:
        print('No results')
    else:
        # Create a DataFrame with a column for each column in the results
        # columns = "Spot market (JKM)"
        # columns = [r'Spot market \B\JKM\D\\']
        columns = [listresultsJEPXR.Columns[33]]
        # print(columns)
        df_JEPXR = pd.DataFrame([[row.GetProperty.Overloads[String](n) for n in columns] for row in listresultsJEPXR], columns=columns)

    # print(df_JEPXR)

    #Generator V&OM costs
    # print(str(sol.FetchAllPropertyEnums() ))
    propId3 = sol.PropertyName2EnumId("System", "Generator", "Generators", "VO&M Cost")
    listresultsGenVOM = sol.QueryToList(SimulationPhaseEnum.MTSchedule, \
              collectionDict["SystemGenerators"], \
              '', \
              '', \
              PeriodEnum.Month, \
              SeriesTypeEnum.Names, \
              str(propId3), \
              DateTime.Parse('2022-04-01'), \
              DateTime.Parse('2023-03-31'), \
                                     '0', \
                                     '', \
                                     '', \
                                     aggregation_none)

    if listresultsGenVOM is None:
        print('No results')
    else:
        # Create a DataFrame with a column for each column in the results
        # columns = "Spot market (JKM)"
        # columns = [r'Spot market \B\JKM\D\\']
        columns = ['3M-Power Plant-1', '3M-Power Plant-2', '7M-Power Plant-1']
        # columns = [listresultsGenVOM.Columns[33]]
        # print(columns)
        df_GenVOM = pd.DataFrame([[row.GetProperty.Overloads[String](n) for n in columns] for row in listresultsGenVOM], columns=columns)

    # print(df_GenVOM)

    # add sums of columns to dataframe
    df_outputs = {"JKM Cost": df_JKMC.values.sum(), "JKM Revenue": df_JKMR.values.sum(),
                  "JEPX Revenue": df_JEPXR.values.sum(), "Generation VOM": df_GenVOM.values.sum()}
    return df_outputs

    print("C. Summary for Model is returned")

def main():
    plexospath = r"C:\Program Files\Energy Exemplar\PLEXOS 10.0"
    count = 0
    PLEXOS_db = 'Tokyo_Gas_V1.6 - A1-3.xml'

    No_of_steps = 2

    # read the JKM values from A1 output
    location = './Model Base (Scenario A1) Solution/'
    df_JKM = read_values(location + 'Sales.csv', location + 'Purchases.csv')

    Scenario_names = ["A3-SS", "A3-SB", "A3-BB", "A3-BS"]

    # Purchase_addons = {'Sc_Name': ["A3-SS", "A3-SB", "A3-BB", "A3-BS"],
    #     'Add_on_1': [0, 0, 70000, 70000],
    #                    'Add_on_2': [0, 70000, 70000, 0]}
    Pur_add = {"A3-SS": [0, 0],
               "A3-SB": [0, 10920], #3640*3 in BBTU for 3 ships
               "A3-BB": [ 10920,  10920],
               "A3-BS": [ 10920, 0]}

    # df_pa = pd.DataFrame(Purchase_addons)
    # print(df_pa)

    # empty dataframe to store results
    df_results = pd.DataFrame(columns = ["Sc.Name", "JKM Cost", "JKM Revenue", "JEPX Revenue", "Generation VOM"])


    # set up for loop to cycle through the models for A3
    for sc_name_str in Scenario_names: #Scenario_names
        # add_pur = df_pa.loc[df_pa['Sc_Name'] == sc_str]
        # print(add_pur[["Add_on_1",'Add_on_2']].to_string(index=False, header=False))

        Add_Pur = Pur_add[sc_name_str]

        # # # add the constraints on JKM sales and purchases
        # update_dataset(PLEXOS_db, df_JKM, No_of_steps, sc_name_str, Add_Pur)
        Model_name = "Model "+sc_name_str
        print(Model_name)

        # # Execute model
        # run_model(plexospath, PLEXOS_db, '.', str("Model "+sc_name_str)) # can add the username and password at the end

        # # Check log file
        # for res in parse_logfile('MT Schedule Completed','.','1. LT_Cost_Optimal_90',25):
        #     print(res)

        # # update result to datafrane
        zipfilename = "Model "+Model_name+" Solution/Model "+Model_name+" Solution.zip"
        # print(zipfilename)
        sc_outputs = query_results(zipfilename)
        dict_1 = {"Sc.Name": sc_name_str, "JKM Cost": sc_outputs["JKM Cost"], "JKM Revenue": sc_outputs["JKM Revenue"],
                  "JEPX Revenue": sc_outputs["JEPX Revenue"], "Generation VOM": sc_outputs["Generation VOM"]}
        # print(dict_1)
        df_results = df_results._append(dict_1, ignore_index = True)

    print("Summary of Results across the models:")
    print(df_results)

    # # Check for convergence
    # convg, new_lcoe = check_convg(LCOE_val, Excel_sheet)
    # LCOE_val = new_lcoe
    # tariff_list.append(new_lcoe)
    # if convg == True:
    #     break
    #
    # count += 1
    # print("Iter no: ",count)

    # print("Tariff list: ", tariff_list)


if __name__ == '__main__':
    main()
