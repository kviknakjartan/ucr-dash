from PyQt5.QtWidgets import * 
from PyQt5 import QtCore, QtGui
from PyQt5.QtGui import * 
from PyQt5.QtCore import * 
import sys
import requests
import pandas as pd
import pyqtgraph as pg
import numpy as np
from io import StringIO, BytesIO
import colorsys 
import os
import shutil
from openpyxl import load_workbook
import re
import xlrd
import time
import json
pd.set_option('future.no_silent_downcasting', True)

class PopDataFetcher():

    POP2000_STATE_PATH = 'https://www2.census.gov/programs-surveys/popest/datasets/2000-2009/state/asrh/sc-est2009-alldata5-all.csv'
    POP2010_STATE_PATH = 'https://www2.census.gov/programs-surveys/popest/datasets/2010-2020/state/asrh/SC-EST2020-ALLDATA5.csv'
    POP2020_STATE_PATH = 'https://www2.census.gov/programs-surveys/popest/datasets/2020-2024/state/asrh/sc-est2024-alldata5.csv'

    def __init__(self):
        self.dataDict = {}

        self.loadDemographics()

    def loadDemographics(self):
        df001yUS = pd.read_excel('0009y1.xlsx', skiprows = 0, skipfooter = 0)
        df101yUS = pd.read_excel('1020y1.xlsx', skiprows = 0, skipfooter = 0)
        df201yUS = pd.read_excel('2024y1.xlsx', skiprows = 0, skipfooter = 0)

        self.dataDict['0009y1'] = df001yUS
        self.dataDict['1020y1'] = df101yUS
        self.dataDict['2024y1'] = df201yUS

    def loadDemographics_(self):
        #df001y = pd.read_csv(self.POP2000_STATE_PATH)
        #df101y = pd.read_csv(self.POP2010_STATE_PATH)
        #df201y = pd.read_csv(self.POP2020_STATE_PATH)

        df001y = pd.read_excel('df001y.xlsx', skiprows = 0, skipfooter = 0)
        df101y = pd.read_excel('df101y.xlsx', skiprows = 0, skipfooter = 0)
        df201y = pd.read_excel('df201y.xlsx', skiprows = 0, skipfooter = 0)

        self.dataDict['0009y1s'] = df001y
        self.dataDict['1020y1s'] = df101y
        self.dataDict['2024y1s'] = df201y

        groups = ['SUMLEV', 'SEX', 'ORIGIN', 'RACE', 'AGE']
        df001yUS = df001y.groupby(groups)[[c for c in df001y.columns if not c.isalpha()]].sum().reset_index()
        df101yUS = df101y.groupby(groups)[[c for c in df101y.columns if not c.isalpha()]].sum().reset_index()
        df201yUS = df201y.groupby(groups)[[c for c in df201y.columns if not c.isalpha()]].sum().reset_index()

        self.dataDict['0009y1'] = df001yUS
        self.dataDict['1020y1'] = df101yUS
        self.dataDict['2024y1'] = df201yUS

        df001yUS.to_excel('0009y1.xlsx', index=False)
        df101yUS.to_excel('1020y1.xlsx', index=False)
        df201yUS.to_excel('2024y1.xlsx', index=False)

    def getPopulation(self, year, age, sex = 0, origin = 0, race = range(1,7)):
        if year < 2000:
            return float("NaN")
        elif year < 2010:
            df = self.dataDict['0009y1']
        elif year < 2020:
            df = self.dataDict['1020y1']
        else:
            df = self.dataDict['2024y1']
        yearCol = 'POPESTIMATE' + str(year)
        return df.loc[(df['AGE'] >= age[0]) & (df['AGE'] <= age[1]) & (df['SEX'] == sex) & (df['ORIGIN'] == origin) & \
            (df['RACE'].isin(race)), yearCol].sum()


    def getAgeDist(self, year, start_age, end_age):
        if start_age > end_age or end_age > 85:
            raise Exception(f'Illegal age: {start_age}, {end_age}')
        df = self.dataDict['0010y1']
        pops = df.loc[(df['MONTH'] == 7) & (df['YEAR'] == year) & (df['AGE'] >= start_age) & (df['AGE'] <= end_age), 'TOT_POP']
        return pops / pops.sum()

class UCRDataFetcher():

    DATA_PATH = r"ucr\data"
    CIUS_STRING = 'Crime in the United States: National volume and rate'
    OKLESC_STRING = 'Offenses Known to Law Enforcement by State by City'
    ARE_STRING = 'Arrests by Race and Ethnicity'
    AS_STRING = 'Arrests by State'
    POP_STRING = 'US Census Population Estimations'
    MV_STRING = 'Murder Victims by Age vs. Sex and Race'
    MO_STRING = 'Murder Offenders by Age vs. Sex and Race'
    MVO_STRING = 'Murder (single victim/single offender) victim vs offender descriptions'
    CLR_STRING = 'Offenses Cleared by Arrest or Exceptional Means, by Region'
    MW_STRING = 'Murder Victims by Weapon Used'
    CP_STRING = 'Offenses by Population Group'
    
    def __init__(self):
        self.years = [int(y) for y in self.get_directories_in_path(self.DATA_PATH)]
        self.dataDict = {self.CIUS_STRING : {},
                         self.OKLESC_STRING : {},
                         self.ARE_STRING : {},
                         self.AS_STRING : {},
                         self.POP_STRING : {},
                         self.MV_STRING : {},
                         self.MO_STRING : {},
                         self.MVO_STRING : {},
                         self.CLR_STRING : {},
                         self.MW_STRING : {},
                         self.CP_STRING : {}}
        self.metaDict = {}
        self.noteDict = {self.AS_STRING : "NOTE:  Because the number of agencies submitting arrest data varies from year to year, users are cautioned about making direct comparisons between arrest totals and those published in previous years' editions of Crime in the United States. Further, arrest figures may vary widely from state to state because some Part II crimes are not considered crimes in some states."}

        with open('states.json', 'r') as file:
            states = json.load(file)
        self.statesDict = {v.capitalize() : k for k,v in states.items()}

        self.popFetcher = PopDataFetcher()
        self.loadPopulationData()

    def getTablePath(self, year, tableNr, partName = '', tableLetter = '', fullName = ''):
        root_dir = f'{self.DATA_PATH}\\{year}'
        if len(fullName) > 0:
            search_string = fullName
        elif year < 2011:
            if isinstance(tableNr, int):
                search_string = f'{year % 100}tbl{tableNr:02d}{tableLetter.lower()}.xls'
            else:
                search_string = f'{year % 100}tbl{tableNr}{tableLetter.lower()}.xls'
        elif year == 2011 and tableNr < 8:
            search_string = f'table-{tableNr}'
        else:
            search_string = f'Table_{tableNr}{tableLetter}_'
        tablePaths = self.find_files_with_string_in_name(root_dir, search_string)
        tablePaths = [p for p in tablePaths if 'MACOSX' not in p]
        if len(tablePaths) > 1 and len(partName) > 0:
            tablePaths = [p for p in tablePaths if partName in p]
        if len(tablePaths) > 1:
            print(tablePaths)
            raise Exception("More than one candidate found for table during search")
        elif len(tablePaths) == 0:
            return None
        return tablePaths[0]

    def getBasePath(self, year):
        if year == 2007:
            basePath = 'documents\\'
        elif year == 2008:
            basePath = 'CIUS2008datatables\\'
        elif year == 2009:
            basePath = 'Data Tables\\'
        elif year == 2010:
            basePath = 'Excel\\'
        elif year == 2011:
            basePath = 'CIUS2011datatables\\'
        elif year == 2012:
            basePath = 'Tables\\'
        elif year == 2019:
            basePath = 'Data Tables\\'
        else:
            basePath = ''
        return basePath

    def loadTable(self, tableName):
        if tableName == self.CIUS_STRING and len(self.dataDict[self.CIUS_STRING]) == 0:
            self.loadTable1Data()
        elif tableName == self.OKLESC_STRING and len(self.dataDict[self.OKLESC_STRING]) == 0:
            self.loadTable8Data()
        elif tableName == self.CP_STRING and len(self.dataDict[self.CP_STRING]) == 0:
            self.loadTable16Data()
        elif tableName == self.CLR_STRING and len(self.dataDict[self.CLR_STRING]) == 0:
            self.loadTable26Data()
        elif tableName == self.ARE_STRING and len(self.dataDict[self.ARE_STRING]) == 0:
            self.loadTable43Data()
        elif tableName == self.AS_STRING and len(self.dataDict[self.AS_STRING]) == 0:
            self.loadTable69Data()
        elif tableName == self.MV_STRING and len(self.dataDict[self.MV_STRING]) == 0:
            self.loadHTable2and3Data(self.MV_STRING)
        elif tableName == self.MO_STRING and len(self.dataDict[self.MO_STRING]) == 0:
            self.loadHTable2and3Data(self.MO_STRING, 3)
        elif tableName == self.MVO_STRING and len(self.dataDict[self.MVO_STRING]) == 0:
            self.loadHTable6Data()
        elif tableName == self.MW_STRING and len(self.dataDict[self.MW_STRING]) == 0:
            self.loadHTable8Data()

    def getNotification(self, tableName):
        return self.noteDict.get(tableName, '')

    def loadPopulationData(self):
        ageGrpDict = {
                      'All ages' : (0,99),
                      'Under 18' : (0,17),  
                      'Under 22' : (0,21),
                      '18 and over' : (18,99),
                      '19 to 24' : (19,24),
                      '25 to 34' : (25,34),
                      '35 to 44' : (35,44),
                      '45 to 54' : (45,54),
                      '55 to 64' : (55,64),
                      '65 and over' : (65,99)
                     }
        sexDict = {
                   'Total' : 0,
                   'Male' : 1,
                   'Female' : 2
                   }
        ethnDict = {
                   'Total' : 0,
                   'Not Hispanic' : 1,
                   'Hispanic' : 2
                   }
        raceDict = {
                   'White' : 1,
                   'Black or African American' : 2,
                   'American Indian and Alaska Native' : 3,
                   'Asian' : 4,
                   'Native Hawaiian and Other Pacific Islander' : 5
        }
        for ageGrp, age in ageGrpDict.items():
            for sex, sexVal in sexDict.items():
                df = pd.DataFrame()
                df['Year'] = range(2000,2025)
                for ethn, ethnVal in ethnDict.items():
                    df[ethn] = df['Year'].apply(self.popFetcher.getPopulation, \
                        age = age, sex = sexVal, origin = ethnVal, race = range(1,6))
                for race, raceVal in raceDict.items():
                    df[race] = df['Year'].apply(self.popFetcher.getPopulation, \
                        age = age, sex = sexVal, origin = ethnVal, race = [raceVal])
                df = df.set_index('Year')


                self.deep_update(self.dataDict, \
                                  {self.POP_STRING : 
                                  {ageGrp :
                                  {sex :
                                  {'Population estimate' : df
                                  }}}})
                self.deep_update(self.metaDict, \
                                  {self.POP_STRING : 
                                  {ageGrp :
                                  {sex :
                                  {'Population estimate' : {}
                                  }}}})

    def loadTable1Data(self):
        earliestYear = min(self.years)
        filePathMin = self.getTablePath(earliestYear, 1)
        dfMin = pd.read_excel(filePathMin, skiprows = 3, skipfooter = 0)
        dfMin = dfMin.iloc[4:,:]
        dfMin = self.cleanEmptyCells(dfMin)
        dfMin,_ = self.dropFooter(dfMin)
        dfMin = self.cleanColumnNames(dfMin)
        dfMin[['Year', 'Population']] = dfMin['Population'].str.split('-', expand=True)
        dfMin['Year'] = range(earliestYear-19, earliestYear + 1)
        dfMin = self.dropEmptyColumns(dfMin)
        dfMin['Rape (revised definition)'] = float("NaN")
        dfMin = dfMin.rename(columns={'Forcible rape': 'Rape (legacy definition)'})
        dfMin = dfMin.rename(columns={'Murder and non- negligent man- slaughter': 'Murder and nonnegligent manslaughter'})
        dfMin = dfMin[dfMin['Year'] < 1987]
        # Remove the comma thousands separator
        dfMin['Population'] = dfMin['Population'].str.replace(',', '')
        # Convert the cleaned string to a numeric type, coercing errors to NaN
        dfMin['Population'] = pd.to_numeric(dfMin['Population'], errors='coerce')
        dfMin = dfMin.drop(columns=['Crime Index total'])

        for c in dfMin.columns:
            if c in ['Year', 'Population']:
                continue
            dfMin[c + ' rate'] = 100000 * dfMin[c] / dfMin['Population']

        dfMin = dfMin.rename(columns={'Violent Crime': 'Violent crime total'})
        dfMin = dfMin.rename(columns={'Property crime': 'Property crime total'})
        dfMin = dfMin.rename(columns={'Violent Crime rate': 'Violent crime rate total'})
        dfMin = dfMin.rename(columns={'Property crime rate': 'Property crime rate total'})

        filePathMed = self.getTablePath(2006, 1)
        dfMed = pd.read_excel(filePathMed, skiprows = 2, skipfooter = 0)
        dfMed = self.cleanEmptyCells(dfMed)
        dfMed,_ = self.dropFooter(dfMed)
        dfMed['Year'] = range(1987, 2007)
        dfMed = self.dropEmptyColumns(dfMed)
        dfMed.insert(loc=8, column='Rape (revised definition)', value=float("NaN"))
        dfMed.insert(loc=9, column='Rape (revised definition) rate', value=float("NaN"))

        latestYear = max(self.years)
        filePath2024 = self.getTablePath(latestYear, 1, 'Crime_in_the_U')
        dfMax = pd.read_excel(filePath2024, skiprows = 3, skipfooter = 0)
        dfMax = self.cleanEmptyCells(dfMax)
        dfMax,_ = self.dropFooter(dfMax)
        dfMax = dfMax[dfMax['Year'].astype(int) > 2006]
        dfMax['Year'] = range(2007, latestYear+1)
        dfMax = self.dropEmptyColumns(dfMax)

        dfMed = self.cleanColumnNames(dfMed)
        dfMax = self.cleanColumnNames(dfMax)

        dfMed = dfMed.rename(columns={'Forcible rape': 'Rape (legacy definition)'})
        dfMed = dfMed.rename(columns={'Forcible rape rate': 'Rape (legacy definition) rate'})
        dfMed = dfMed.rename(columns={'Violent crime': 'Violent crime total'})
        dfMed = dfMed.rename(columns={'Property crime': 'Property crime total'})
        dfMed = dfMed.rename(columns={'Violent crime rate': 'Violent crime rate total'})
        dfMed = dfMed.rename(columns={'Property crime rate': 'Property crime rate total'})

        dfMax = dfMax.rename(columns={'Murder andnonnegligent manslaughter': 'Murder and nonnegligent manslaughter'})
        dfMax = dfMax.rename(columns={'Violentcrime': 'Violent crime'})
        dfMax = dfMax.rename(columns={'Violent crime': 'Violent crime total'})
        dfMax = dfMax.rename(columns={'Property crime': 'Property crime total'})
        dfMax = dfMax.rename(columns={'Violent crime rate': 'Violent crime rate total'})
        dfMax = dfMax.rename(columns={'Property crime rate': 'Property crime rate total'})
        dfMax = dfMax.rename(columns = {'Larceny- theft' : 'Larceny-theft'})
        dfMax = dfMax.rename(columns = {'Larceny- theft rate' : 'Larceny-theft rate'})

        df = pd.concat([dfMin,dfMed,dfMax])
        df = self.cleanEmptyCells(df)

        df.set_index('Year', inplace = True)
        df = df.astype(float)

        #df.to_excel('df.xlsx', index=True)

        dfRates, dfVolumes = self.seperateRates(df)

        self.dataDict[self.CIUS_STRING] = \
                                  {'All crime' :
                                  {'All groups' :
                                  {'Rate per 100,000 Inhabitants' : dfRates,
                                   'Volume' : dfVolumes}}}

        popDf = dfVolumes.copy()
        for crime in dfRates.columns:
            popDf[crime] = dfVolumes['Population']

        self.deep_update(self.metaDict, \
                                  {self.CIUS_STRING : 
                                  {'All crime' :
                                  {'All groups' :
                                  {'Rate per 100,000 Inhabitants' : {'Population' : popDf, 'Volume' : dfVolumes},
                                   'Volume' : {'Population' : popDf}}}}})


    def loadTable8Data(self):
        wholeData = pd.DataFrame()
        wholeData.flags.allows_duplicate_labels = False
        footerDict = {}
        for year in self.years: 
            footerDict[year] = {}    
            if year == 2004:
                filePath = self.getTablePath(year, 8, '', 'a') 
            elif year < 2006:
                filePath = self.getTablePath(year, 8)
            else: 
                filePath = self.getTablePath(year, 8, 'Offenses_Known_to_Law_Enforcement_by_State_by_City_') 
            if filePath is None:
                if len(wholeData) > 0:
                    df = pd.DataFrame()
                    df['City'] = wholeData['City'].unique()
                    df['Year'] = year
                    wholeData = pd.concat([wholeData, df])
                continue
            df = pd.read_excel(filePath, skiprows = 3 + 2*(year == 2020) + (year == 2003), skipfooter = 0)
            df = self.cleanEmptyCells(df)
            df = df.replace(r'^\s*-\s*$', 0, regex=True)
            df = self.cleanColumnNames(df)

            if year > 2004:
                df,footer = self.dropFooter(df)

                nextStateIdx = df['State'].first_valid_index()
                while nextStateIdx is not None:
                    lastStateIdx = nextStateIdx
                    nextStateIdx = df.iloc[(lastStateIdx + 1):len(df),0].first_valid_index()
                    if nextStateIdx is None:
                        df.iloc[lastStateIdx:len(df),0] = df.iloc[lastStateIdx,0].capitalize()
                    else:
                        df.iloc[lastStateIdx:nextStateIdx,0] = df.iloc[lastStateIdx,0].capitalize()

            else:
                df = df.replace(r'^\s+Jurisdiction$', "Las Vegas", regex=True)
                if year == 1999:
                    df = df.replace('Las Vegas Metropolitan Police Department', float("NaN"))
                isStateOrFooter = df.iloc[:, 1:].isnull().all(axis=1) & ~df.isnull().all(axis=1)
                isCity = ~df.iloc[:, 1:].isnull().all(axis=1)
                firstFooterIndex = df.loc[isStateOrFooter,'City by state'].str[0].str.isdigit().idxmax()
                isStates = isStateOrFooter
                isStates[isStates.index >= firstFooterIndex] = False
                states = df.loc[isStates,'City by state']
                citiesStartIdxs = isStates[isStates].index + 1
                citiesEndIdxs = isStates[isStates].index - 1
                citiesEndIdxs = list(citiesEndIdxs[1:])
                citiesEndIdxs += [firstFooterIndex - 1]
                for state, startIdx, endIndx in zip(states, citiesStartIdxs, citiesEndIdxs):
                    df.loc[startIdx:(endIndx+1),'State'] = state.capitalize()
                footer = df.loc[firstFooterIndex:,:]
                if year == 1999:
                    df = df.drop(columns=['Burglary'])
                    df = df.rename(columns={'Burglary.1' : 'Burglary'})
                if year < 2002:
                    df = df.drop(columns=['Crime Index total','Modified Crime Index total'])
                elif year < 2003:
                    df = df.drop(columns=['Crime Index','Modified Crime Index'])
                df = df.rename(columns={'Murder and non-negligent man- slaughter': 'Murder and nonnegligent manslaughter'})
                df = df.rename(columns={'City by state': 'City'})
                df = df[~df['Population'].isna()]
            if 'Unnamed: 13' in df.columns:
                df = df.drop(columns=['Unnamed: 13'])
            df = df.rename(columns={'Murder and non-negligent man-slaughter': 'Murder and nonnegligent manslaughter'}) 

            df['State'] = self.cleanRows(df['State'])

            for n in range(1,20):
                footNote = footer.loc[footer['State'].str.startswith(str(n), na=False) & ~footer['State'].str.startswith(str(n+9), na=False),'State']
                if len(footNote) == 1:
                    footNote = footNote.iloc[0]
                elif len(footNote) > 1:
                    raise Exception(footNote)
                if 'previous years' in footNote:
                    footerDict[year][n] = {}
                    citiesNoted = df.loc[df['City'].str.endswith(str(n), na=False),['State','City']]
                    citiesNoted['City'] = self.cleanRows(citiesNoted['City'])
                    footerDict[year][n]['cities'] = self.joinCitiesAndStates(citiesNoted['City'], citiesNoted['State'])
                    footerDict[year][n]['note'] = footNote.replace(f'{n} ','')
                elif 'submitting rape data' in footNote:
                    footerDict[year][n] = {}
                    citiesNoted = df.loc[df['City'].str.endswith(str(n), na=False),['State','City']]
                    citiesNoted['City'] = self.cleanRows(citiesNoted['City'])
                    footerDict[year][n]['cities'] = self.joinCitiesAndStates(citiesNoted['City'], citiesNoted['State'])
                    footerDict[year][n]['rape'] = footNote.replace(f'{n} ','')
                elif 'revised' in footNote and year < 2017:
                    footerDict[year][n] = {}
                    citiesNoted = df.loc[~df['Rape (revised definition)'].isna(),['State','City']]
                    citiesNoted['City'] = self.cleanRows(citiesNoted['City'])
                    footerDict[year][n]['cities'] = self.joinCitiesAndStates(citiesNoted['City'], citiesNoted['State'])
                    footerDict[year][n]['rape'] = footNote.replace(f'{n} ','')
                elif 'legacy' in footNote and year < 2017:
                    footerDict[year][n] = {}
                    citiesNoted = df.loc[~df['Rape (legacy definition)'].isna(),['State','City']]
                    citiesNoted['City'] = self.cleanRows(citiesNoted['City'])
                    footerDict[year][n]['cities'] = self.joinCitiesAndStates(citiesNoted['City'], citiesNoted['State'])
                    footerDict[year][n]['rape'] = footNote.replace(f'{n} ','')
            
            df = self.dropEmptyColumns(df)
            df['City'] = self.cleanRows(df['City'])
            df['City'] = df['City'].str.replace(r'^Miami$','Miami-Dade', regex=True)
            df['City'] = df['City'].str.replace('Nashville Metropolitan','Nashville')
            df['City'] = df['City'].str.replace('Louisville Metro','Louisville')
            df['City'] = df['City'].str.replace('Metropolitan Nashville Police Department','Nashville')
            df['City'] = df['City'].str.replace('Las Vegas Metropolitan Police Department','Las Vegas')
            df['City'] = self.joinCitiesAndStates(df['City'], df['State'])
            df = df.rename(columns = {'Forcible rape' : 'Rape'})
            df = df.rename(columns = {'Violent crime' : 'Violent Crime Total'})
            df = df.rename(columns = {'Property crime' : 'Property Crime Total'})
            df = df.rename(columns = {'Larceny- theft' : 'Larceny-theft'})
            df['Year'] = year
            wholeData = pd.concat([wholeData,df])
        wholeData = self.aggregateColumns(wholeData, 'Rape', 'Rape (revised definition)', 'Rape (legacy definition)')
        wholeData = wholeData.drop(columns = ['Rape (revised definition)','Rape (legacy definition)'])

        citiesToInclude = []
        cities = wholeData.loc[wholeData['Population'] > 220000,'City'].unique()
        for city in cities:
            if wholeData.loc[wholeData['City'] == city, 'Population'].to_numpy()[0] >= 250000:
                citiesToInclude.append(city)
        wholeData = wholeData[wholeData['City'].isin(citiesToInclude)]

        for year in wholeData['Year'].unique():
            missingCities = [c for c in citiesToInclude if c not in wholeData.loc[wholeData['Year'] == year, 'City'].values]
            missingDf = pd.DataFrame(columns = wholeData.columns)
            missingDf['City'] = missingCities
            missingDf['Year'] = year
            wholeData = pd.concat([wholeData,missingDf])
        wholeData = wholeData.sort_values(by=['Year', 'City'])
        wholeData = wholeData.reindex()
        #wholeData.to_excel('test.xlsx', index=False)
        
        
        crimes = [c for c in wholeData.columns if c not in ['State','City','Population','Year']]
        for crime in crimes:
            volumeDf = pd.DataFrame()
            volumeDf['Year'] = self.years
            volumeDf.set_index('Year', inplace = True)
            rateDf = volumeDf.copy()
            for idx,city in enumerate(citiesToInclude):
                volumeDf[city] = float("NaN")
                rateDf[city] = float("NaN")
                cityData = wholeData.loc[wholeData['City'] == city, ['Year',crime]]
                cityData.set_index('Year', inplace = True)
                volumeDf.loc[cityData.index, city] = cityData[crime]
                population = wholeData.loc[wholeData['City'] == city,['Year','Population']]
                population.set_index('Year', inplace = True)
                rateDf.loc[population.index, city] = 100000 * cityData[crime] / population['Population']

            self.deep_update(self.dataDict, \
                                  {self.OKLESC_STRING : 
                                  {crime :
                                  {'All groups' :
                                  {'Rate per 100,000 Inhabitants' : rateDf,
                                   'Volume' : volumeDf}}}})
        
        notesDf = pd.DataFrame()
        for year, yearDict in footerDict.items():
            for noteNr, noteDict in yearDict.items():
                df = pd.DataFrame()
                df['City'] = noteDict['cities'][noteDict['cities'].isin(citiesToInclude)]
                df['Year'] = year
                noteType = [key for key in noteDict.keys() if key != 'cities'][0]
                df['NoteType'] = noteType
                noteStr = noteDict[noteType].replace('shown in this column ', '')
                noteStr = re.sub(r'\s*See Data Declaration for further explanation.\s*','',noteStr)
                noteStr = re.sub(r'\s*See the data declaration for further explanation.\s*','',noteStr)
                df['Note'] = noteStr
                notesDf = pd.concat([notesDf, df])

        #notesDf.to_excel('notesDf.xlsx', index=False)

        popDf = pd.DataFrame()
        popDf['Year'] = self.years
        popDf.set_index('Year', inplace = True)
        for city in citiesToInclude:
            population = wholeData.loc[wholeData['City'] == city,['Year','Population']]
            population.set_index('Year', inplace = True)
            popDf.loc[population.index, city] = population['Population']

        for crime in crimes:
            nDf = pd.DataFrame()
            nDf['Year'] = self.years
            nDf.set_index('Year', inplace = True)
            volumeDf = self.dataDict[self.OKLESC_STRING][crime]['All groups']['Volume']
            for city in citiesToInclude:
                nDf[city] = float("NaN")
                nDf[city] = nDf[city].astype(object)
                cityNotesDf = notesDf[notesDf['City'] == city]
                if crime != 'Rape':
                    cityNotesDf = cityNotesDf[cityNotesDf['NoteType'] == 'note']
                if len(cityNotesDf) == 0:
                    continue
                yearsWithEntries = volumeDf[~volumeDf[city].isna()].index
                for year in yearsWithEntries:
                    yearNotes = cityNotesDf.loc[cityNotesDf['Year'] == year, 'Note']
                    noteStr = ''
                    for note in yearNotes:
                        noteStr += f'\nNOTE: {note}'
                    if len(noteStr) > 0:
                        nDf.loc[year,city] = noteStr

            self.deep_update(self.metaDict, \
                                  {self.OKLESC_STRING : 
                                  {crime :
                                  {'All groups' :
                                  {'Rate per 100,000 Inhabitants' : {'Notes' : nDf, 'Population' : popDf, 'Volume' : volumeDf, 'Important' : nDf},
                                   'Volume' : {'Notes' : nDf, 'Population' : popDf, 'Important' : nDf}}}}})
    
    def loadTable16Data(self):
        wholeData = pd.DataFrame()
        wholeData.flags.allows_duplicate_labels = False
        footerDict = {}
        for year in self.years: 
            footerDict[year] = {}    
            if year < 2011:
                filePath = self.getTablePath(year, 16)
            else: 
                filePath = self.getTablePath(year, 16, '','','Table_16_Rate_') 
            if filePath is None:
                if len(wholeData) > 0:
                    df = pd.DataFrame(columns = wholeData.columns)
                    df['Population group'] = wholeData.loc[wholeData['Year'] == year - 1,'Population group']
                    df['Year'] = year
                    wholeData = pd.concat([wholeData, df])
                continue
            df = pd.read_excel(filePath, skiprows = 3 + 2*(year == 2020), skipfooter = 0, sheet_name=0 + (year==2008))
            df = self.cleanEmptyCells(df)
            #.to_excel('df.xlsx',index=False) 
            df = self.cleanColumnNames(df)
            df = self.dropEmptyColumns(df)
            df = df.rename(columns={f'{year} estimated population' : 'Population', 'Forcible rape' : 'Rape (legacy definition)'})
            df = df.rename(columns={'Unnamed: 0' : 'Population group', 'Rape (revised definition)' : 'Rape'})
            df = df.rename(columns={'Murder and Nonnegligent Homicide' : 'Murder and nonnegligent manslaughter', \
                'Aggravated Assault' : 'Aggravated assault', 'Motor Vehicle Theft' : 'Motor vehicle theft'})
            df = df[~df['Population group'].isna()]
            df['Population group'] = self.cleanRows(df['Population group'])
            df = df[~df['Population group'].str.contains('Population Group')]
            df = df.replace(r'^\s*-+\s*$', np.nan, regex=True)
            df['Population group'] = df['Population group'].str.replace(':','')
            df['Population group'] = df['Population group'].str.replace('GROUP I (all cities 250,000 and over)','GROUP I (250,000 and over)')
            df['Population group'] = df['Population group'].str.replace('1,000,000 or over (Group I subset)','1,000,000 and over (Group I subset)')
            df['Population group'] = df['Population group'].str.replace('500,000 thru 999,999 (Group I subset)','500,000 to 999,999 (Group I subset)')
            df['Population group'] = df['Population group'].str.replace('250,000 thru 499,999 (Group I subset)','250,000 to 499,999 (Group I subset)')
            df['Population group'] = df['Population group'].str.replace('SUBURBAN AREAS','SUBURBAN AREA')
            df = df.rename(columns={'Violent crime' : 'Violent crime total', 'Property crime' : 'Property crime total'})
            df = df.drop(columns=[c for c in df.columns if 'Unnamed' in c])
            df,footer = self.dropFooter(df) 
            df['Year'] = year
            wholeData = pd.concat([wholeData,df])
        #wholeData.to_excel('wholeData.xlsx',index=False)
        crimeCols = [c for c in wholeData.columns if c not in ['Population group','Population','Number of agencies']]
        popGrps = wholeData['Population group'].unique()
        for pGrp in popGrps:
            volumeKnownDf = wholeData.loc[wholeData['Population group'] == pGrp,crimeCols]
            volumeKnownDf = volumeKnownDf.set_index('Year')
            popDf = volumeKnownDf.copy()
            agenciesDf = volumeKnownDf.copy()
            popSeries = wholeData.loc[wholeData['Population group'] == pGrp,'Population'].reindex()
            agSeries = wholeData.loc[wholeData['Population group'] == pGrp,'Number of agencies'].reindex()
            for crime in crimeCols:
                if crime == 'Year':
                    continue
                popDf[crime] = popSeries.values
                agenciesDf[crime] = agSeries.values
            rateKnownDf = 100000 * volumeKnownDf / popDf
            self.deep_update(self.dataDict, \
                                  {self.CP_STRING : 
                                  {'Population group by Crime' :
                                  {pGrp :
                                  {'Volume' : volumeKnownDf,
                                   'Rate per 100,000 inhabitants' : rateKnownDf}}}})

            self.deep_update(self.metaDict, \
                                  {self.CP_STRING : 
                                  {'Population group by Crime' :
                                  {pGrp :
                                  {
                                  'Volume' : {'Population' : popDf, 'Agencies' : agenciesDf},
                                   'Rate per 100,000 inhabitants' : {'Volume' : volumeKnownDf, 
                                                                     'Population' : popDf, 'Agencies' : agenciesDf}}}}})
        
        years = wholeData['Year'].unique()
        popDf = pd.DataFrame(columns=popGrps, index=years)
        agenciesDf = pd.DataFrame(columns=popGrps, index=years)
        for pGrp in popGrps:
            popSeries = wholeData.loc[wholeData['Population group'] == pGrp,'Population'].reindex()
            agSeries = wholeData.loc[wholeData['Population group'] == pGrp,'Number of agencies'].reindex()
            popDf[pGrp] = popSeries.values
            agenciesDf[pGrp] = agSeries.values
        for crime in crimeCols:
            if crime == 'Year':
                continue
            offSeries = wholeData.loc[:, crime]
            volumeKnownDf = pd.DataFrame(offSeries.values.reshape(len(years), len(popGrps)), columns=popGrps)
            volumeKnownDf['Year'] = years
            volumeKnownDf = volumeKnownDf.set_index('Year')
            rateKnownDf = 100000 * volumeKnownDf / popDf

            self.deep_update(self.dataDict, \
                                  {self.CP_STRING : 
                                  {'Crime by population group' :
                                  {crime :
                                  {'Volume' : volumeKnownDf,
                                   'Rate per 100,000 inhabitants' : rateKnownDf}}}})

            self.deep_update(self.metaDict, \
                                  {self.CP_STRING : 
                                  {'Crime by population group' :
                                  {crime :
                                  {
                                  'Volume' : {'Population' : popDf, 'Agencies' : agenciesDf},
                                   'Rate per 100,000 inhabitants' : {'Volume' : volumeKnownDf, 
                                                                     'Population' : popDf, 'Agencies' : agenciesDf}}}}})

    
    def loadTable26Data(self):
        wholeData = pd.DataFrame()
        wholeData.flags.allows_duplicate_labels = False
        footerDict = {}
        tableNr = 26
        for year in self.years: 
            footerDict[year] = {}    
            
            if year < 2011:
                filePath = self.getTablePath(year, tableNr)
            else: 
                filePath = self.getTablePath(year, tableNr, 'Table_26_Percent_of_Offenses_Cleared_') 
            if filePath is None:
                if len(wholeData) > 0:
                    df = pd.DataFrame(columns = wholeData.columns)
                    df['Region'] = wholeData.loc[wholeData['Year'] == year - 1,'Region']
                    df['Variable'] = wholeData.loc[wholeData['Year'] == year - 1,'Variable']
                    df['Year'] = year
                    wholeData = pd.concat([wholeData, df])
                continue
            
            df = pd.read_excel(filePath, skiprows = 3 + 2*(year == 2020), skipfooter = 0)
            df = self.cleanEmptyCells(df)
            df = self.dropEmptyColumns(df)
            df = self.cleanColumnNames(df)
            df = df.rename(columns={'Motor Vehicle Theft' : 'Motor vehicle theft', 'Unnamed: 2' : 'Violent crime', \
                'Unnamed: 7' : 'Property crime'})
            df = df.rename(columns={'Geographic Region/Division' : 'Region', 'Larceny- theft' : 'Larceny-theft'})
            df = df.rename(columns={'Unnamed: 1' : 'Variable', 'Geographic region/division' : 'Region'})
            df = df.rename(columns={'Region/geographic division' : 'Region', 'Forcible rape' : 'Rape (legacy definition)'})
            df = df.rename(columns={'Rape (revised definition)' : 'Rape', \
                'Murder and Nonnegligent Homicide' : 'Murder and nonnegligent manslaughter'})
            df = df.rename(columns={f'{year} estimated population' : 'Population', 'Aggravated Assault' : 'Aggravated assault'})
            df = df.rename(columns={'Violent crime' : 'Violent crime total', 'Property crime' : 'Property crime total'})
            df,footer = self.dropFooter(df)  
            emptyIdxs = df.loc[df['Region'].isna()].index
            fullIdxs = df.loc[~df['Region'].isna()].index
            df.loc[emptyIdxs,['Region','Number of agencies','Population']] = \
                df.loc[fullIdxs,['Region','Number of agencies','Population']].values
            df['Region'] = self.cleanRows(df['Region'])
            df['Variable'] = self.cleanRows(df['Variable'])
            df['Region'] = df['Region'].str.replace(':','')
            df['Region'] = df['Region'].str.title()
            df['Region'] = df['Region'].str.replace('New England','New England (NE)')
            df['Region'] = df['Region'].str.replace('Middle Atlantic','Middle Atlantic (NE)')
            df['Region'] = df['Region'].str.replace('East North Central','East North Central (MW)')
            df['Region'] = df['Region'].str.replace('West North Central','West North Central (MW)')
            df['Region'] = df['Region'].str.replace('South Atlantic','South Atlantic (S)')
            df['Region'] = df['Region'].str.replace('East South Central','East South Central (S)')
            df['Region'] = df['Region'].str.replace('West South Central','West South Central (S)')
            df['Region'] = df['Region'].str.replace('Mountain','Mountain (W)')
            df['Region'] = df['Region'].str.replace('Pacific','Pacific (W)')
            df['Variable'] = df['Variable'].str.replace('Percent cleared by arrest','Percent cleared')
            df['Year'] = year
            wholeData = pd.concat([wholeData,df])
        #wholeData.to_excel('wholeData.xlsx',index=False) 
        crimeCols = [c for c in wholeData.columns if c not in ['Region','Variable','Number of agencies','Population']]
        regions = wholeData['Region'].unique()
        for region in regions:
            volumeKnownDf = wholeData.loc[(wholeData['Region'] == region) & (wholeData['Variable'] == 'Offenses known'),crimeCols]
            percentClrdDf = wholeData.loc[(wholeData['Region'] == region) & (wholeData['Variable'] == 'Percent cleared'),crimeCols]
            volumeKnownDf = volumeKnownDf.set_index('Year')
            percentClrdDf = percentClrdDf.set_index('Year')
            volumeClrdDf = (volumeKnownDf * percentClrdDf) / 100
            volumeClrdDf = volumeClrdDf.round()
            #volumeKnownDf.to_excel('volumeKnownDf.xlsx',index=True)
            #percentClrdDf.to_excel('percentClrdDf.xlsx',index=True)
            #volumeClrdDf.to_excel('volumeClrdDf.xlsx',index=True)
            popDf = volumeKnownDf.copy()
            agenciesDf = volumeKnownDf.copy()
            popSeries = wholeData.loc[(wholeData['Region'] == region) & \
                (wholeData['Variable'] == 'Percent cleared'),'Population'].reindex()
            agSeries = wholeData.loc[(wholeData['Region'] == region) & \
                (wholeData['Variable'] == 'Percent cleared'),'Number of agencies'].reindex()
            for crime in crimeCols:
                if crime == 'Year':
                    continue
                popDf[crime] = popSeries.values
                agenciesDf[crime] = agSeries.values
            rateKnownDf = 100000 * volumeKnownDf / popDf
            rateClrdDf = 100000 * volumeClrdDf / popDf
            #percDf.to_excel('percDf.xlsx',index=True)
            self.deep_update(self.dataDict, \
                                  {self.CLR_STRING : 
                                  {'Region by Crime' :
                                  {region :
                                  {'Offenses known (volume)' : volumeKnownDf,
                                   'Offenses known (rate per 100,000 inhabitants)' : rateKnownDf,
                                   'Offenses cleared (volume)' : volumeClrdDf,
                                   'Percentage cleared (%)' : percentClrdDf}}}})

            self.deep_update(self.metaDict, \
                                  {self.CLR_STRING : 
                                  {'Region by Crime' :
                                  {region :
                                  {
                                  'Offenses known (volume)' : {'Population' : popDf, 'Agencies' : agenciesDf},
                                   'Offenses known (rate per 100,000 inhabitants)' : {'Volume' : volumeKnownDf, 'Population' : popDf, 'Agencies' : agenciesDf},
                                   'Offenses cleared (volume)' : {'Population' : popDf, 'Agencies' : agenciesDf},
                                   'Percentage cleared (%)' : {'Volume' : volumeClrdDf, 'Population' : popDf, 'Agencies' : agenciesDf}}}}})
        years = wholeData['Year'].unique()
        popDf = pd.DataFrame(columns=regions, index=years)
        agenciesDf = pd.DataFrame(columns=regions, index=years)
        for region in regions:
            popSeries = wholeData.loc[(wholeData['Region'] == region) & \
                (wholeData['Variable'] == 'Percent cleared'),'Population'].reindex()
            agSeries = wholeData.loc[(wholeData['Region'] == region) & \
                (wholeData['Variable'] == 'Percent cleared'),'Number of agencies'].reindex()
            popDf[region] = popSeries.values
            agenciesDf[region] = agSeries.values
        for crime in crimeCols:
            if crime == 'Year':
                continue
            offSeries = wholeData.loc[wholeData['Variable'] == 'Offenses known', crime]
            clrdSeries = wholeData.loc[wholeData['Variable'] == 'Percent cleared', crime]
            volumeKnownDf = pd.DataFrame(offSeries.values.reshape(len(years), len(regions)), columns=regions)
            volumeKnownDf['Year'] = years
            percentClrdDf = pd.DataFrame(clrdSeries.values.reshape(len(years), len(regions)), columns=regions)
            percentClrdDf['Year'] = years
            volumeKnownDf = volumeKnownDf.set_index('Year')
            percentClrdDf = percentClrdDf.set_index('Year')
            volumeClrdDf = (volumeKnownDf * percentClrdDf) / 100
            volumeClrdDf = volumeClrdDf.round()
            #volumeKnownDf.to_excel('volumeKnownDf.xlsx',index=True)
            #percentClrdDf.to_excel('percentClrdDf.xlsx',index=True)
            #volumeClrdDf.to_excel('volumeClrdDf.xlsx',index=True)
            rateKnownDf = 100000 * volumeKnownDf / popDf
            rateClrdDf = 100000 * volumeClrdDf / popDf
            #percDf.to_excel('percDf.xlsx',index=True)
            self.deep_update(self.dataDict, \
                                  {self.CLR_STRING : 
                                  {'Crime by Region' :
                                  {crime :
                                  {'Offenses known (volume)' : volumeKnownDf,
                                   'Offenses known (rate per 100,000 inhabitants)' : rateKnownDf,
                                   'Offenses cleared (volume)' : volumeClrdDf,
                                   'Percentage cleared (%)' : percentClrdDf}}}})

            self.deep_update(self.metaDict, \
                                  {self.CLR_STRING : 
                                  {'Crime by Region' :
                                  {crime :
                                  {
                                  'Offenses known (volume)' : {'Population' : popDf, 'Agencies' : agenciesDf},
                                   'Offenses known (rate per 100,000 inhabitants)' : {'Volume' : volumeKnownDf, 'Population' : popDf, 'Agencies' : agenciesDf},
                                   'Offenses cleared (volume)' : {'Population' : popDf, 'Agencies' : agenciesDf},
                                   'Percentage cleared (%)' : {'Volume' : volumeClrdDf, 'Population' : popDf, 'Agencies' : agenciesDf}}}}})

    def loadTable43Data(self):
        wholeData = pd.DataFrame()
        for year in self.years:
            for letter in ['A','B','C']:
                if year == 2004:
                    filePath = self.getTablePath(year, 43, '', letter + 'adj')
                else:
                    filePath = self.getTablePath(year, 43, '', letter) 
                if filePath is None:
                    if len(wholeData) == 0:
                        continue
                    df['Year'] = year
                    df['Letter'] = letter
                    df[[c for c in df.columns if c not in ['Offense charged','Year','Letter']]] = float("NaN")
                    wholeData = pd.concat([wholeData,df])
                    continue
                if letter == 'A':
                    if filePath.endswith('.xls'):
                        workbook = xlrd.open_workbook(filePath)
                        sheet_by_index = workbook.sheet_by_index(0)
                        cell_value = sheet_by_index.cell_value(rowx=3, colx=0)
                    else:
                        workbook = load_workbook(filePath)
                        sheet = workbook.active  # or workbook['Sheet1']
                        cell_value = sheet['A4'].value
                    numbers = re.findall(r'\b\d{1,3}(?:,\d{3})*(?:\\.\\d+)?(?!\d)', cell_value)
                    pop = int(numbers[1].replace(',',''))
                    agencies = int(numbers[0].replace(',',''))
                          
                df = pd.read_excel(filePath, skiprows = 5 + (year>2012), skipfooter = 0)
                df = self.cleanEmptyCells(df)
                df,_ = self.dropFooter(df)
                df = self.dropEmptyColumns(df)
                df = df.rename(columns = {'Total.2' : 'Total.3'})
                df = df.rename(columns = {'Total2' : 'Total.2'})
                df = self.cleanColumnNames(df)
                if year == 2013:
                    df = df.rename(columns = {'American Indian or Alaskan Native' : 'American Indian or Alaska Native.1'})
                else:
                    df = df.rename(columns = {'American Indian or Alaskan Native' : 'American Indian or Alaska Native'})
                    df = df.rename(columns = {'American Indian or Alaskan Native.1' : 'American Indian or Alaska Native.1'})
                df = df.rename(columns = {'Black' : 'Black or African American'})
                df = df.rename(columns = {'Black.1' : 'Black or African American.1'})
                df = df.rename(columns = {'Unnamed: 0' : 'Offense charged'})
                df['Offense charged'] = self.cleanRows(df['Offense charged'])
                df['Year'] = year
                df['Letter'] = letter
                df['Population'] = pop
                df['Agencies'] = agencies
                wholeData = pd.concat([wholeData,df])
        wholeData = self.cleanEmptyCells(wholeData)
        #wholeData.to_excel('test.xlsx', index=False)
        wholeData.set_index('Year', inplace = True)
        wholeData.replace("*", 0, inplace=True)
        crimes = df['Offense charged'].unique()
        groupDict = {'A' : 'Total arrests', 'B' : 'Arrests under 18', 'C' : 'Arrests 18 or over'}
        for letter,group in groupDict.items():
            for crime in crimes:
                if letter == 'C' and crime in ['Curfew and loitering law violations','Runaways']:
                    continue
                df = wholeData[(wholeData['Letter'] == letter) & (wholeData['Offense charged'] == crime)]
                #df.to_excel('df.xlsx', index=True)
                df = df.drop(columns = ['Letter','Offense charged'])
                df = df.astype(float)
                agencies = df['Agencies']
                df.drop(columns = ['Agencies'], inplace = True)
                dfVolumes = df[[c for c in df.columns if c[-2] != '.']]
                dfRates = dfVolumes.copy()
                for c in dfRates.columns:
                    dfRates[c] = 100000 * dfRates[c] / dfVolumes['Population']
                dfRates.drop(columns = ['Population'], inplace = True)
                dfPercent = df[[c for c in df.columns if c.endswith('.1')]] 
                dfPercent = dfPercent.drop(columns = ['Total.1']) 
                dfPercent = dfPercent.rename(columns={c : c.replace('.1','') for c in dfPercent.columns})
                
                self.deep_update(self.dataDict, \
                                  {self.ARE_STRING : 
                                  {crime :
                                  {group :
                                  {'Rate per 100,000 Inhabitants' : dfRates,
                                   'Volume' : dfVolumes,
                                   'Percentages (%)' : dfPercent}}}})

                dfPop = dfVolumes.copy()
                dfAge = dfVolumes.copy()
                for series in dfPop.columns:
                    dfPop[series] = dfPop['Population']
                    dfAge[series] = agencies

                self.deep_update(self.metaDict, \
                                  {self.ARE_STRING : 
                                  {crime :
                                  {group :
                                  {
                                  'Rate per 100,000 Inhabitants' : {'Population' : dfPop, 'Volume' : dfVolumes, 'Agencies' : dfAge},
                                  'Volume' : {'Population' : dfPop, 'Agencies' : dfAge},
                                  'Percentages (%)' : {'Population' : dfPop, 'Volume' : dfVolumes, 'Agencies' : dfAge}}}}})
            

    def loadTable69Data(self):
        wholeData = pd.DataFrame()
        wholeData.flags.allows_duplicate_labels = False
        footerDict = {}
        for year in self.years: 
            footerDict[year] = {}    
            if year == 2004:
                filePath = self.getTablePath(year, 69, '', 'adj') 
            elif year == 2000:
                filePath = self.getTablePath(year, 69, '', '', 'rtbl69_00.xls')
            elif year < 2011:
                filePath = self.getTablePath(year, 69)
            else: 
                filePath = self.getTablePath(year, 69, 'Table_69_Arrest_by_State_') 
            if filePath is None:
                if len(wholeData) > 0:
                    df = pd.DataFrame(columns = wholeData.columns)
                    df['State'] = wholeData.loc[wholeData['Year'] == year - 1,'State']
                    df['Group'] = wholeData.loc[wholeData['Year'] == year - 1,'Group']
                    df['Year'] = year
                    wholeData = pd.concat([wholeData, df])
                continue
            df = pd.read_excel(filePath, skiprows = 3 + (year < 2005) + (year == 2002), skipfooter = 0)
            df = self.cleanEmptyCells(df)
            df = self.cleanColumnNames(df)
            df = self.dropEmptyColumns(df)
            df = df.replace(r'^\s*-+\s*$', np.nan, regex=True)

            stateTexts = df.loc[df['State'].str.contains(r'^[A-Z]+\s*[A-Z]+\s*[A-Z]+\d*\s*(\:\s*)?(\d+\,*\s*\d*)?(\,\d+)?(\s+\d+\s+agenc(y|ies)(;|:)\s*)?$', na=False),'State']
            statesTextMatches = stateTexts.str.extract(r'^([A-Z]+\s*[A-Z]+\s*[A-Z]+\d*)\s*(\:\s*)?(\d+\,*\s*\d*)?(\,\d+)?(\s+(\d+)\s+agenc(y|ies)(;|:)\s*)?$')
                

            if year > 2004:
                df,footer = self.dropFooter(df)
                df = df.rename(columns={'Unnamed: 1' : 'Group'})
                df.columns = df.columns.str.replace(r'\d+ estimated population', 'Population', regex=True)
                df.loc[df['State'].isna(),'State'] = df.loc[~df['State'].isna(),'State'].values
                populations = df.loc[df['Group'] == 'Under 18','Population']
                agencies = df.loc[df['Group'] == 'Under 18','Number of agencies']
                missingPopIdxs = populations[populations.isna() | (populations == 0)].index
                df.loc[missingPopIdxs,[c for c in df.columns if c not in ['State','Group']]] = float("NaN")
                df.loc[missingPopIdxs.values + 1,[c for c in df.columns if c not in ['State','Group']]] = float("NaN")
                df.loc[populations.index.values + 1,'Population'] = populations.values
                df.loc[populations.index.values + 1,'Number of agencies'] = agencies.values
            else:
                popTexts = df.loc[df['State'].str.contains(r'opulation\s+\d+', na=False),'State']
                populations = popTexts.str.extract(r'opulation\s*(\d{1,3}(?:,\d{3})*(?:\.\d+)?)')[0]
                populations = pd.to_numeric(populations.str.replace(',',''))
                states = statesTextMatches[0]
                missingStateIdxs = list(set(states.index).difference(set(populations.index.values - 1)))
                numberOfAgencies = statesTextMatches[5]
                missingStates = states[missingStateIdxs]
                states = states.drop(labels=missingStateIdxs)
                numberOfAgencies = pd.to_numeric(numberOfAgencies[~numberOfAgencies.isna()])
                missingPopIdxs = set(missingStateIdxs).intersection(set(numberOfAgencies.index.values))
                numberOfAgencies = numberOfAgencies.drop(labels=missingPopIdxs)
                df = df.drop([i + 1 for i in missingPopIdxs])
                df = df.drop([i + 2 for i in missingPopIdxs])
                under18Idxs = df[df['State'].str.contains('nder 18', na=False)].index
                totalIdxs = df[df['State'].str.contains('all ages', na=False)].index

                df.loc[under18Idxs,'Group'] = df.loc[under18Idxs,'State'].values
                df.loc[totalIdxs,'Group'] = df.loc[totalIdxs,'State'].values
                df.loc[under18Idxs,'State'] = states.values
                df.loc[totalIdxs,'State'] = states.values
                df.loc[under18Idxs,'Population'] = populations.values
                df.loc[totalIdxs,'Population'] = populations.values
                df.loc[under18Idxs,'Number of agencies'] = numberOfAgencies.values
                df.loc[totalIdxs,'Number of agencies'] = numberOfAgencies.values

                df = df.dropna(how='all')
                df = df.drop(states.index)
                df = df.drop(popTexts.index)
                df = df.drop(missingStates.index)
                df,footer = self.dropFooter(df)
                missingStatesDf = pd.DataFrame(columns=df.columns)
                missingStatesDf['State'] = missingStates.repeat(2).values
                grpSeries = pd.Series([df.iloc[i,df.columns.get_loc("Group")] for i in range(0, 2*len(missingStates))])
                missingStatesDf['Group'] = grpSeries.values
                df = pd.concat([df,missingStatesDf])

                df.columns = df.columns.str.replace(r'\d+', '', regex=True)
                df.columns = df.columns.str.replace('- ', '')
                df.columns = df.columns.str.replace('-', '')
                df = df.rename(columns={'Larcenytheft' : 'Larceny-theft'})
                df = df.rename(columns={'Crime Index total' : 'Crime Index'})
                if 'Crime Index' in df.columns:
                    df = df.drop(columns=['Crime Index'])
                
            df = df.rename(columns={'Larceny- theft' : 'Larceny-theft'})
            df = df.rename(columns = {'Forcible rape' : 'Rape'})
            df = df.rename(columns = {'Violent crime' : 'Violent Crime Total'})
            df = df.rename(columns = {'Property crime' : 'Property Crime Total'})
            df.columns = df.columns.str.replace('forcible rape', 'rape')
            df['Year'] = year

            statesWithNotes = statesTextMatches[0].str.cat(statesTextMatches[2], na_rep='').str.cat(statesTextMatches[3], na_rep='')
            statesWithNotesMatches = statesWithNotes.str.extract(r'^([A-Z]+\s*[A-Z]+\s*[A-Z]+)(\d+)?\s*\,?\s*(\d+)?s*\,?\s*(\d+)?')
            for n in range(1,20):
                footNoteSeries = footer.loc[footer['State'].str.startswith(str(n), na=False) & ~footer['State'].str.startswith(str(n+9), na=False),'State']
                nextFootNoteSeries = footer.loc[footer['State'].str.startswith(str(n+1), na=False) & ~footer['State'].str.startswith(str(n+10), na=False),'State']
                if len(footNoteSeries) == 1:
                    footNote = 'NOTE: ' + footNoteSeries.iloc[0].replace(str(n),'')
                    footNote = footNote.replace('  ',' ')
                    nextIdx = len(footer) if len(nextFootNoteSeries) == 0 else nextFootNoteSeries.index[0]
                    for i in range(footNoteSeries.index[0] + 1, nextIdx):
                        nextline = footer.loc[i,'State']
                        if nextline.startswith('NOTE'):
                            break
                        footNote += f'\n{nextline}'
                    containsNote1 = pd.to_numeric(statesWithNotesMatches[1]) == n
                    containsNote2 = pd.to_numeric(statesWithNotesMatches[2]) == n
                    containsNote3 = pd.to_numeric(statesWithNotesMatches[3]) == n
                    statesNoted = statesWithNotesMatches.loc[containsNote1 | containsNote2 | containsNote3, 0]
                    statesNoted = statesNoted[statesNoted != 'DISTRICT OF COLUMBIA']
                    if len(statesNoted) > 0:
                        footerDict[year][n] = {'note' : footNote, 'states' : statesNoted}
                    else:
                        footerDict[year][n] = {}
                elif len(footNoteSeries) > 1:
                    raise Exception(footNote)

            df['State'] = self.cleanRows(df['State'])
            df = df[df['State'] != 'DISTRICT OF COLUMBIA']
            if len(wholeData) > 0:
                missingStateNames = [s for s in wholeData['State'].unique() if s not in df['State'].values]
                if len(missingStateNames) > 0:
                    missingStatesDf = pd.DataFrame(columns=df.columns)
                    missingStatesDf['State'] = pd.Series(missingStateNames).repeat(2).values
                    grpSeries = pd.Series([df.iloc[i,df.columns.get_loc("Group")] for i in range(0, 2*len(missingStateNames))])
                    missingStatesDf['Group'] = grpSeries.values
                    missingStatesDf['Year'] = year
                    df = pd.concat([df,missingStatesDf])
            df = df.sort_values(by=['State','Group'])
            wholeData = pd.concat([wholeData,df])
        wholeData = wholeData.reindex()
        wholeData['Group'] = wholeData['Group'].str.strip()
        wholeData['State'] = wholeData['State'].str.title()
        #wholeData.to_excel('test.xlsx', index=False)
        
        crimes = [c for c in wholeData.columns if c not in ['State','Population','Year','Group','Number of agencies']]
        years = wholeData['Year'].unique()

        notesDf = pd.DataFrame(columns=wholeData['State'].unique(), index=years)
        importantDf = notesDf.copy()
        for year, yearDict in footerDict.items():
            for noteNr, noteDict in yearDict.items():
                if (len(noteDict) == 0) or ('further explanation.' in noteDict['note']) or \
                ('Miccosukee and Seminole' in noteDict['note']) or ('NOTE: See' in noteDict['note']):
                    continue
                for state in [s.title() for s in noteDict['states']]:
                    if ('Limited' in noteDict['note']) or ('unable to supply' in noteDict['note']) or \
                    ('is the only' in noteDict['note']) or ('to previous years' in noteDict['note']) or \
                    ('include only those' in noteDict['note']):
                        importantDf.loc[year, state] = 'X'
                    if pd.isna(notesDf.loc[year, state]):
                        notesDf.loc[year, state] = f'\n{noteDict['note']}'
                    else:
                        notesDf.loc[year, state] = f'{notesDf.loc[year, state]}\n{noteDict['note']}'
                
        #notesDf.to_excel('notesDf.xlsx', index=True)  

        for group in wholeData['Group'].unique():
            states = wholeData.loc[(wholeData['Year'] == wholeData.loc[0,'Year'].values[0]) & (wholeData['Group'] == group), 'State']
            numberOfStates = len(states)
            populationSeries = wholeData.loc[wholeData['Group'] == group,'Population']
            populationDf = pd.DataFrame.from_records(populationSeries.values.reshape(len(years), numberOfStates), columns=states.values)
            populationDf.index = years
            agenciesSeries = wholeData.loc[wholeData['Group'] == group,'Number of agencies']
            agencyDf = pd.DataFrame.from_records(agenciesSeries.values.reshape(len(years), numberOfStates), columns=states.values)
            agencyDf.index = years
            for crime in crimes:
                series = wholeData.loc[wholeData['Group'] == group, crime]
                volumeDf = pd.DataFrame(series.values.reshape(len(years), numberOfStates), columns=states.values)
                volumeDf.index = years
                rateDf = 100000 * volumeDf / populationDf

                self.deep_update(self.dataDict, \
                                      {self.AS_STRING : 
                                      {crime :
                                      {group :
                                      {'Rate per 100,000 Inhabitants' : rateDf,
                                       'Volume' : volumeDf}}}})

                self.deep_update(self.metaDict, \
                                  {self.AS_STRING : 
                                  {crime :
                                  {group :
                                  {'Rate per 100,000 Inhabitants' : 
                                              {'Notes' : notesDf, 'Population' : populationDf, 'Volume' : volumeDf, 'Agencies' : agencyDf, 'Important' : importantDf},
                                   'Volume' : {'Notes' : notesDf, 'Population' : populationDf, 'Agencies' : agencyDf, 'Important' : importantDf}}}}})
        
    def loadHTable2and3Data(self, tableString, tableNr = 2):
        wholeData = pd.DataFrame()
        wholeData.flags.allows_duplicate_labels = False
        footerDict = {}
        for year in self.years: 
            footerDict[year] = {}    
            if year == 2003:
                filePath = self.getTablePath(year, '2', '', '-4' if tableNr == 2 else '-5') 
            elif year == 2004:
                filePath = self.getTablePath(year, '2', '', '-4a' if tableNr == 2 else '-5a')
            elif year < 2003:
                filePath = self.getTablePath(year, '2', '', '-5' if tableNr == 2 else '-6')
            elif year == 2005:
                filePath = self.getTablePath(year, tableNr, '', '', f'shrtab{tableNr}')
            elif year == 2006:
                filePath = self.getTablePath(year, tableNr, '', '', f'shrtable{tableNr}')
            elif year == 2007:
                filePath = self.getTablePath(year, tableNr, '', '', f'shrtbl{tableNr}')
            elif year < 2011:
                filePath = self.getTablePath(year, tableNr, '', '', f'shrtbl0{tableNr}')
            elif year == 2020:
                filePath = self.getTablePath(year, tableNr, '', '', f'Expanded_Homicide_Data_Table_0{tableNr}_Murder_') 
            else: 
                filePath = self.getTablePath(year, tableNr, '', '', f'Expanded_Homicide_Data_Table_{tableNr if year != 2016 else tableNr - 1 - (tableNr == 2)}_Murder_')
            if filePath is None:
                if len(wholeData) > 0:
                    df = pd.DataFrame(columns = wholeData.columns)
                    df['Age'] = wholeData.loc[wholeData['Year'] == year - 1,'Age']
                    df['Year'] = year
                    wholeData = pd.concat([wholeData, df])
                continue
            
            df = pd.read_excel(filePath, skiprows = 4 + (year == 2001) + (year == 2020), skipfooter = 0)
            df = self.cleanEmptyCells(df)
            df = self.dropEmptyColumns(df)
            df = self.cleanColumnNames(df)
            df = df.replace(r'^\s*-+\s*$', 0, regex=True)
            df = df.dropna(how='all')
            df,footer = self.dropFooter(df)
            df = df.drop(columns=['Unknown'])
            df = df.rename(columns={'Unknown.1' : 'Unknown Race', 'Other' : 'Other Race', 'Black' : 'Black or African American'})
            df = df.rename(columns={'Unnamed: 0' : 'Age', 'Unnamed: 1' : 'Total', 'Unnamed: 2' : 'Total'})
            df = df.rename(columns={'Unknown.2' : 'Unknown Ethnicity'})
            df['Age'] = df['Age'].str.strip()
            df['Age'] = df['Age'].str.replace(r'Percent distribution\d','Percent distribution', regex=True)
            df['Age'] = df['Age'].str.replace(r'Under 18\d','Under 18', regex=True)
            df['Age'] = df['Age'].str.replace(r'Under 22\d','Under 22', regex=True)
            df['Age'] = df['Age'].str.replace(r'18 and over\d','18 and over', regex=True)

            ageList = ['1 to 4','5 to 8','9 to 12','13 to 16','30 to 34','35 to 39','40 to 44','45 to 49','50 to 54', \
                        '55 to 59','60 to 64','65 to 69','70 to 74','75 and over']
            ageGrpDict = {
                '1 to 8' : (1,8), '9 to 16' : (9,16), '30 to 39' : (30,39), '40 to 49' : (40,49), '50 to 59' : (50,59),
                '60 to 69' : (60,69), '70 and over' : (70,99)
            }
            startGrpIdxs = df[df['Age'].isin(ageList[::2])].index
            endGrpIdxs = df[df['Age'].isin(ageList[1::2])].index
            newGrps = pd.DataFrame(columns=df.columns)
            newGrps['Age'] = ageGrpDict.keys()
            startGrps = df.loc[startGrpIdxs,:].select_dtypes(include='number')
            startGrps.index = range(0,len(startGrpIdxs))
            endGrps = df.loc[endGrpIdxs,:].select_dtypes(include='number')
            endGrps.index = range(0,len(endGrpIdxs))
            newGrps.iloc[:,1:] = startGrps + endGrps
            df = df.drop(startGrpIdxs)
            df = df.drop(endGrpIdxs)
            df = pd.concat([df,newGrps])
            if year == 2020:
                raceCols = [c for c in df.columns if c != 'Age']
                df.loc[df['Age'] == 'Percent distribution', raceCols] *= 100
            df['Year'] = year
            df = self.dropEmptyColumns(df)
            #df.to_excel('df.xlsx', index=False)
            wholeData = pd.concat([wholeData,df])
        wholeData = wholeData.reindex()
        #wholeData.to_excel('test.xlsx', index=False)
        
        ageGrpDict = {
                '1 to 8' : (1,8), '9 to 16' : (9,16), '30 to 39' : (30,39), '40 to 49' : (40,49), '50 to 59' : (50,59),
                '60 to 69' : (60,69), '70 and over' : (70,99), 'Percent distribution' : (0,99), 'Under 18' : (0,17),
                'Under 22' : (0,21), 'Total' : (0,99), '18 and over' : (18,99), 'Infant (under 1)' : (0,0), 'Unknown' : (0,99),
                '17 to 19' : (17,19), '20 to 24': (20,24), '25 to 29' : (25,29)
            }
        raceGrpDict = {
                'White' : [1], 'Black or African American' : [2], 'Other Race' : [3,4,5], 'Unknown Race' : range(1,6), \
                'Total' : range(1,6), 'Unknown Ethnicity' : range(1,6)
        }
        for ageGrp in [a for a in wholeData['Age'].unique() if a != 'Percent distribution']:
            raceCols = [c for c in wholeData.columns if c not in ['Age','Year']]
            volumeDf = wholeData[wholeData['Age'] == ageGrp]
            volumeDf = volumeDf.drop(columns=['Age'])
            demDf = volumeDf.copy()
            popDf = volumeDf.copy()
            for raceGrp in raceCols:
                if raceGrp in ['Male','Female']:
                    sex = 1 + (raceGrp == 'Female')
                    origin = 0
                    race = range(1,6)
                elif 'Hispanic' in raceGrp:
                    sex = 0
                    origin = 1 + (raceGrp == 'Hispanic or Latino')
                    race = range(1,6)
                else:
                    race = raceGrpDict[raceGrp]
                    sex = 0
                    origin = 0

                demDf[raceGrp] = demDf['Year'].apply(self.popFetcher.getPopulation, \
                            age = ageGrpDict[ageGrp], sex = sex, origin = origin, race = race)
                popDf[raceGrp] = demDf['Year'].apply(self.popFetcher.getPopulation, \
                            age = (0,99), sex = 0, origin = 0, race = range(1,6))
            volumeDf = volumeDf.set_index('Year')
            demDf = demDf.set_index('Year')
            popDf = popDf.set_index('Year')
            rateDf = 100000 * volumeDf / popDf
            drateDf = 100000 * volumeDf / demDf

            self.deep_update(self.dataDict, \
                                  {tableString : 
                                  {'Age by Sex and Race' :
                                  {ageGrp :
                                  {'Rate per 100,000 Members of Population at Risk' : drateDf,
                                   'Rate per 100,000 Inhabitants' : rateDf,
                                   'Volume' : volumeDf}}}})

            self.deep_update(self.metaDict, \
                                  {tableString : 
                                  {'Age by Sex and Race' :
                                  {ageGrp :
                                  {
                                  'Rate per 100,000 Members of Population at Risk' : {'Demographic' : demDf,
                                                                                    'Population' : popDf, 'Volume' : volumeDf},
                                  'Rate per 100,000 Inhabitants' : {'Demographic' : demDf, 'Population' : popDf,
                                                                                    'Volume' : volumeDf},
                                  'Volume' : {'Demographic' : demDf, 'Population' : popDf}}}}})
        percDf = wholeData[wholeData['Age'] == 'Percent distribution'].drop(columns=['Age'])
        percDf = percDf.set_index('Year')
        #popDf.to_excel('popDf.xlsx', index=True)
        #rateDf.to_excel('rateDf.xlsx', index=True)
        #drateDf.to_excel('drateDf.xlsx', index=True)
        self.dataDict[tableString]['Age by Sex and Race']['Total']['Percentages (%)'] = percDf  
        self.metaDict[tableString]['Age by Sex and Race']['Total']['Percentages (%)'] = \
            self.metaDict[tableString]['Age by Sex and Race']['Total']['Rate per 100,000 Inhabitants']

        years = wholeData['Year'].unique()
        raceCols = [c for c in wholeData.columns if c not in ['Age','Year']]
        ageGrps = [a for a in wholeData['Age'].unique() if a != 'Percent distribution']
        numberOfGroups = len(ageGrps)
        for raceGrp in raceCols:
            series = wholeData.loc[wholeData['Age'].isin(ageGrps), raceGrp]
            volumeDf = pd.DataFrame(series.values.reshape(len(years), numberOfGroups), columns=ageGrps)
            volumeDf['Year'] = years
            demDf = volumeDf.copy()
            popDf = volumeDf.copy()
            if raceGrp in ['Male','Female']:
                sex = 1 + (raceGrp == 'Female')
                origin = 0
                race = range(1,6)
            elif 'Hispanic' in raceGrp:
                sex = 0
                origin = 1 + (raceGrp == 'Hispanic or Latino')
                race = range(1,6)
            else:
                race = raceGrpDict[raceGrp]
                sex = 0
                origin = 0
            for ageGrp in ageGrps:               

                demDf[ageGrp] = demDf['Year'].apply(self.popFetcher.getPopulation, \
                            age = ageGrpDict[ageGrp], sex = sex, origin = origin, race = race)
                popDf[ageGrp] = demDf['Year'].apply(self.popFetcher.getPopulation, \
                            age = (0,99), sex = 0, origin = 0, race = range(1,6))
            
            volumeDf = volumeDf.set_index('Year')
            demDf = demDf.set_index('Year')
            popDf = popDf.set_index('Year')
            rateDf = 100000 * volumeDf / popDf
            drateDf = 100000 * volumeDf / demDf

            self.deep_update(self.dataDict, \
                                  {tableString : 
                                  {'Sex and Race by Age' :
                                  {raceGrp :
                                  {'Rate per 100,000 Members of Population at Risk' : drateDf,
                                   'Rate per 100,000 Inhabitants' : rateDf,
                                   'Volume' : volumeDf}}}})

            self.deep_update(self.metaDict, \
                                  {tableString : 
                                  {'Sex and Race by Age' :
                                  {raceGrp :
                                  {
                                  'Rate per 100,000 Members of Population at Risk' : {'Demographic' : demDf,
                                                                                    'Population' : popDf, 'Volume' : volumeDf},
                                  'Rate per 100,000 Inhabitants' : {'Demographic' : demDf, 'Population' : popDf,
                                                                                    'Volume' : volumeDf},
                                  'Volume' : {'Demographic' : demDf, 'Population' : popDf}}}}})

    def loadHTable6Data(self, tableNr = 6):
        wholeData = pd.DataFrame()
        wholeData.flags.allows_duplicate_labels = False
        footerDict = {}
        for year in self.years: 
            footerDict[year] = {}    
            if year == 2003:
                filePath = self.getTablePath(year, '2', '', '-7') 
            elif year == 2004:
                filePath = self.getTablePath(year, '2', '', '-7a')
            elif year < 2003:
                filePath = self.getTablePath(year, '2', '', '-8')
            elif year == 2005:
                filePath = self.getTablePath(year, tableNr, '', '', f'shrtab{tableNr-1}')
            elif year == 2006:
                filePath = self.getTablePath(year, tableNr, '', '', f'shrtable{tableNr-1}')
            elif year == 2007:
                filePath = self.getTablePath(year, tableNr, '', '', f'shrtbl{tableNr-1}')
            elif year < 2011:
                filePath = self.getTablePath(year, tableNr, '', '', f'shrtbl0{tableNr}')
            elif year == 2020:
                filePath = self.getTablePath(year, tableNr, '', '', f'Expanded_Homicide_Data_Table_0{tableNr}_') 
            else: 
                filePath = self.getTablePath(year, tableNr, '', '', f'Expanded_Homicide_Data_Table_{tableNr if year != 2016 else 3}_')
            if filePath is None:
                if len(wholeData) > 0:
                    df = pd.DataFrame(columns = wholeData.columns)
                    df['Race of victim'] = wholeData.loc[wholeData['Year'] == year - 1,'Race of victim']
                    df['Year'] = year
                    wholeData = pd.concat([wholeData, df])
                continue
            df = pd.read_excel(filePath, skiprows = 5 + (year == 2020), skipfooter = 0)
            df = self.cleanEmptyCells(df)
            df = self.dropEmptyColumns(df)
            df = self.cleanColumnNames(df)
            df = df.replace(r'^\s*-+\s*$', 0, regex=True)
            if 'Unnamed: 12' in df.columns:
                df = df.drop(columns=['Unnamed: 12'])
            if 'Unnamed: 13' in df.columns:
                df = df.drop(columns=['Unnamed: 13'])
            df = df.rename(columns={'Unknown' : 'Unknown Race', 'Other' : 'Other Race', 'Black' : 'Black or African American'})
            df = df.rename(columns={'Unnamed: 0' : 'Race of victim', 'Unnamed: 1' : 'Total', 'Unnamed: 2' : 'Total'})
            df = df.rename(columns={'Unknown.1' : 'Unknown Sex'})
            df = df.rename(columns={'Unknown.2' : 'Unknown Ethnicity'})
            df = df.loc[~df['Race of victim'].isna(),:]
            df,footer = self.dropFooter(df)
            #df.to_excel('df.xlsx', index=False)
            df['Race of victim'] = self.cleanRows(df['Race of victim'])
            df['Race of victim'] = df['Race of victim'].str.capitalize()
            df.columns = df.columns.str.capitalize()
            df['Race of victim'] = df['Race of victim'].str.replace('Unknown race','Unknown race victims')
            df['Race of victim'] = df['Race of victim'].str.replace('Unknown sex','Unknown sex victims')
            df['Race of victim'] = df['Race of victim'].str.replace(r'^Other race$','Other race victims',regex=True)
            df['Race of victim'] = df['Race of victim'].str.replace('Hispanic or latino','Hispanic or latino victims')
            df['Race of victim'] = df['Race of victim'].str.replace('Not hispanic or latino','Not hispanic or latino victims')
            if year == 2003:
                unknownIdxs = df.loc[df['Race of victim'] == 'Unknown',:].index
                df.loc[unknownIdxs[0], 'Race of victim'] = 'Unknown race victims'
                df.loc[unknownIdxs[1], 'Race of victim'] = 'Unknown sex victims'
            df['Race of victim'] = df['Race of victim'].str.replace(r'^Unknown$','Unknown ethnicity victims',regex=True)
            df['Race of victim'] = df['Race of victim'].str.replace(r'^Female$','Female victims',regex=True)
            df['Race of victim'] = df['Race of victim'].str.replace(r'^Male$','Male victims',regex=True)
            df['Race of victim'] = df['Race of victim'].str.replace(r'^White$','White victims',regex=True)
            df['Race of victim'] = df['Race of victim'].str.replace(r'^Other$','Other race victims',regex=True)
            df['Race of victim'] = df['Race of victim'].str.replace(r'^Black victims$','Black or african american victims',regex=True)
            df['Race of victim'] = df['Race of victim'].str.replace(r'^Black$','Black or african american',regex=True)
            df['Race of victim'] = df['Race of victim'].str.replace(r'^Black or african american$', \
                'Black or african american victims',regex=True)
            df = df.loc[df['Race of victim'] != 'Sex of victim',:]
            df = df.loc[df['Race of victim'] != 'Ethnicity of victim',:]
            if 'Unknown ethnicity victims' not in df['Race of victim'].values:
                newRows = pd.DataFrame({ 'Race of victim' : ['Hispanic or latino victims', 'Not hispanic or latino victims', \
                    'Unknown ethnicity victims']})
                df = pd.concat([df,newRows])
            raceCols = [c for c in df.columns if c != 'Race of victim']
            totalVictims = df.loc[df['Race of victim'].isin(['Male victims', 'Female victims', 'Unknown sex victims']), raceCols].sum()
            newRow = pd.DataFrame([['Total victims'] + list(totalVictims.values)], columns = df.columns)
            df = pd.concat([df, newRow])
            df['Year'] = year
            df = df.add_suffix(' offenders')
            df = df.rename(columns={'Race of victim offenders' : 'Race of victim', \
                'Year offenders' : 'Year'})
            wholeData = pd.concat([wholeData,df])
        wholeData = wholeData.reindex()
        #wholeData.to_excel('test.xlsx', index=False)

        raceGrpDict = {
                'White' : [1], 'Black or african american' : [2], 'Other race' : [3,4,5], 'Unknown race' : range(1,6), \
                'Total' : range(1,6), 'Unknown ethnicity' : range(1,6), 'Unknown sex' : range(1,6)
        }
        years = wholeData['Year'].unique()
        raceCols = [c for c in wholeData.columns if c not in ['Race of victim','Year']]
        raceRows = wholeData['Race of victim'].unique()
        for vRaceGrp in raceRows:
            volumeDf = wholeData[wholeData['Race of victim'] == vRaceGrp]
            volumeDf = volumeDf.drop(columns=['Race of victim'])
            demDf = volumeDf.copy()
            popDf = volumeDf.copy()
            percDf = volumeDf.copy()
            percDf = percDf.reindex()
            for raceGrp in [c.replace(' offenders','') for c in raceCols]:
                if raceGrp in ['Male','Female']:
                    sex = 1 + (raceGrp == 'Female')
                    origin = 0
                    race = range(1,6)
                elif 'ispanic' in raceGrp:
                    sex = 0
                    origin = 1 + (raceGrp == 'Hispanic or latino')
                    race = range(1,6)
                else:
                    race = raceGrpDict[raceGrp]
                    sex = 0
                    origin = 0

                demDf[raceGrp + ' offenders'] = demDf['Year'].apply(self.popFetcher.getPopulation, \
                            age = (0,99), sex = sex, origin = origin, race = race)
                popDf[raceGrp + ' offenders'] = popDf['Year'].apply(self.popFetcher.getPopulation, \
                            age = (0,99), sex = 0, origin = 0, race = range(1,6))
                if raceGrp != 'Total':
                    percDf[raceGrp + ' offenders'] = 100 * percDf[raceGrp + ' offenders'] / percDf['Total offenders']
            percDf['Total offenders'] = 100 * percDf['Total offenders'] / percDf['Total offenders']
            volumeDf = volumeDf.set_index('Year')
            demDf = demDf.set_index('Year')
            popDf = popDf.set_index('Year')
            percDf = percDf.set_index('Year')
            rateDf = 100000 * volumeDf / popDf
            drateDf = 100000 * volumeDf / demDf
            #percDf.to_excel('percDf.xlsx',index=True)
            self.deep_update(self.dataDict, \
                                  {self.MVO_STRING : 
                                  {'Victim Description by Offender Description' :
                                  {vRaceGrp :
                                  {'Rate per 100,000 Members of Population at Risk' : drateDf,
                                   'Rate per 100,000 Inhabitants' : rateDf,
                                   'Volume' : volumeDf,
                                   'Percentages (%)' : percDf}}}})

            self.deep_update(self.metaDict, \
                                  {self.MVO_STRING : 
                                  {'Victim Description by Offender Description' :
                                  {vRaceGrp :
                                  {
                                  'Rate per 100,000 Members of Population at Risk' : {'Demographic' : demDf,
                                                                                    'Population' : popDf, 'Volume' : volumeDf},
                                  'Rate per 100,000 Inhabitants' : {'Demographic' : demDf, 'Population' : popDf,
                                                                                    'Volume' : volumeDf},
                                  'Volume' : {'Demographic' : demDf, 'Population' : popDf},
                                  'Percentages (%)' : {'Demographic' : demDf, 'Population' : popDf, 'Volume' : volumeDf}}}}})
        

        for oRaceGrp in raceCols:
            series = wholeData[oRaceGrp]
            volumeDf = pd.DataFrame(series.values.reshape(len(years), len(raceRows)), columns=raceRows)
            volumeDf['Year'] = years
            demDf = volumeDf.copy()
            popDf = volumeDf.copy()
            percDf = volumeDf.copy()
            percDf = percDf.reindex()
            for raceGrp in [c.replace(' victims','') for c in raceRows]:
                if raceGrp in ['Male','Female']:
                    sex = 1 + (raceGrp == 'Female')
                    origin = 0
                    race = range(1,6)
                elif 'ispanic' in raceGrp:
                    sex = 0
                    origin = 1 + (raceGrp == 'Hispanic or latino')
                    race = range(1,6)
                else:
                    race = raceGrpDict[raceGrp]
                    sex = 0
                    origin = 0

                demDf[raceGrp + ' victims'] = demDf['Year'].apply(self.popFetcher.getPopulation, \
                            age = (0,99), sex = sex, origin = origin, race = race)
                popDf[raceGrp + ' victims'] = popDf['Year'].apply(self.popFetcher.getPopulation, \
                            age = (0,99), sex = 0, origin = 0, race = range(1,6))
                if raceGrp != 'Total':
                    percDf[raceGrp + ' victims'] = 100 * percDf[raceGrp + ' victims'] / percDf['Total victims']

            percDf['Total victims'] = 100 * percDf['Total victims'] / percDf['Total victims']   
            volumeDf = volumeDf.set_index('Year')
            demDf = demDf.set_index('Year')
            popDf = popDf.set_index('Year')
            percDf = percDf.set_index('Year')
            rateDf = 100000 * volumeDf / popDf
            drateDf = 100000 * volumeDf / demDf

            self.deep_update(self.dataDict, \
                                  {self.MVO_STRING : 
                                  {'Offender Description by Victim Description' :
                                  {oRaceGrp :
                                  {'Rate per 100,000 Members of Population at Risk' : drateDf,
                                   'Rate per 100,000 Inhabitants' : rateDf,
                                   'Volume' : volumeDf,
                                   'Percentages (%)' : percDf}}}})

            self.deep_update(self.metaDict, \
                                  {self.MVO_STRING : 
                                  {'Offender Description by Victim Description' :
                                  {oRaceGrp :
                                  {
                                  'Rate per 100,000 Members of Population at Risk' : {'Demographic' : demDf,
                                                                                    'Population' : popDf, 'Volume' : volumeDf},
                                  'Rate per 100,000 Inhabitants' : {'Demographic' : demDf, 'Population' : popDf,
                                                                                    'Volume' : volumeDf},
                                  'Volume' : {'Demographic' : demDf, 'Population' : popDf},
                                  'Percentages (%)' : {'Demographic' : demDf, 'Population' : popDf, 'Volume' : volumeDf}}}}})
        
    def loadHTable8Data(self):
        wholeData = pd.DataFrame()
        wholeData.flags.allows_duplicate_labels = False
        footerDict = {}
        tableNr = 8
        for year in self.years: 
            footerDict[year] = {}     
            if year == 2004:
                filePath = self.getTablePath(year, '2', '', '-9a')
            elif year < 2004:
                filePath = self.getTablePath(year, '2', '', '-10')
            elif year == 2005:
                filePath = self.getTablePath(year, tableNr, '', '', f'shrtab{tableNr-1}')
            elif year == 2006:
                filePath = self.getTablePath(year, tableNr, '', '', f'shrtable{tableNr-1}')
            elif year == 2007:
                filePath = self.getTablePath(year, tableNr, '', '', f'shrtbl{tableNr-1}')
            elif year < 2011:
                filePath = self.getTablePath(year, tableNr, '', '', f'shrtbl0{tableNr}')
            elif year == 2020:
                filePath = self.getTablePath(year, tableNr, '', '', f'Expanded_Homicide_Data_Table_0{tableNr}_') 
            else: 
                filePath = self.getTablePath(year, tableNr, '', '', f'Expanded_Homicide_Data_Table_{tableNr if year != 2016 else tableNr-4}_')
            if filePath is None:
                continue
            df = pd.read_excel(filePath, skiprows = 3 + (year == 2020) + (year == 2001), skipfooter = 0)
            df = self.cleanEmptyCells(df)
            df = self.dropEmptyColumns(df)
            #df = self.cleanColumnNames(df)
            df = df.replace(r'^\s*-+\s*$', 0, regex=True)
            if 'Unnamed: 0' in df.columns:
                df = df.rename(columns={'Unnamed: 0' : 'Weapon'})
            if 'Unnamed: 1' in df.columns:
                df = df.rename(columns={'Unnamed: 0' : year-4})
            if 'Unnamed: 2' in df.columns:
                df = df.rename(columns={'Unnamed: 0' : year-3})
            if 'Unnamed: 3' in df.columns:
                df = df.rename(columns={'Unnamed: 0' : year-2})
            if 'Unnamed: 4' in df.columns:
                df = df.rename(columns={'Unnamed: 0' : year-1})
            if 'Unnamed: 5' in df.columns:
                df = df.rename(columns={'Unnamed: 0' : year})
            df = df.rename(columns={'Weapons' : 'Weapon', '20011' : 2001})
            df.columns = [int(c) if c != 'Weapon' else c for c in df.columns]
            df = df.loc[~df['Weapon'].isna(),:]
            df['Weapon'] = self.cleanRows(df['Weapon'])
            df['Weapon'] = df['Weapon'].str.replace(':','')
            df['Weapon'] = df['Weapon'].str.replace('Total fireams','Total firearms')
            hasFirearms = df['Weapon'].str.contains('Firearms')
            df.loc[hasFirearms,'Weapon'] = 'Firearms, type not stated'
            #df['Weapon'] = df['Weapon'].str.replace('Other weapons or Firearms, not stated','Firearms, type not stated')
            #df['Weapon'] = df['Weapon'].str.replace('Other weapons or Firearms, type not stated','Firearms, type not stated')
            emptyKnives = df.loc[df['Weapon'].str.contains('Knives') & df.iloc[:,1].isna(),'Weapon']
            if len(emptyKnives) == 1:
                onlyInstrument = df.loc[df['Weapon'].str.contains('instruments'),'Weapon']
                df.loc[onlyInstrument.index,'Weapon'] = emptyKnives.values + ' ' + onlyInstrument.values
                df = df.drop(emptyKnives.index)
            elif len(emptyKnives) > 1:
                raise Exception(emptyKnives)
            emptyBlunt = df.loc[df['Weapon'].str.contains('Blunt') & df.iloc[:,1].isna(),'Weapon']
            if len(emptyBlunt) == 1:
                onlyHammers = df.loc[df['Weapon'].str.contains('hammers'),'Weapon']
                df.loc[onlyHammers.index,'Weapon'] = emptyBlunt.values + ' ' + onlyHammers.values
                df = df.drop(emptyBlunt.index)
            elif len(emptyBlunt) > 1:
                raise Exception(emptyBlunt)
            emptyPersonal = df.loc[df['Weapon'].str.contains('Personal') & df.iloc[:,1].isna(),'Weapon']
            if len(emptyPersonal) == 1:
                onlyFeet = df.loc[df['Weapon'].str.contains('feet'),'Weapon']
                df.loc[onlyFeet.index,'Weapon'] = emptyPersonal.values + ' ' + onlyFeet.values
                df = df.drop(emptyPersonal.index)
            elif len(emptyPersonal) > 1:
                raise Exception(emptyPersonal)
            emptyWeapons = df.loc[df['Weapon'].str.contains('Other weapons') & df.iloc[:,1].isna(),'Weapon']
            if len(emptyWeapons) == 1:
                onlyStated = df.loc[df['Weapon'].str.contains('not stated'),'Weapon']
                df.loc[onlyStated.index,'Weapon'] = emptyWeapons.values + ' ' + onlyStated.values
                df = df.drop(emptyWeapons.index)
            elif len(emptyWeapons) > 1:
                raise Exception(emptyWeapons)
            df,footer = self.dropFooter(df)
            df = df.set_index('Weapon')
            #df.to_excel('df.xlsx',index=False)
            if len(wholeData) == 0:
                wholeData = df.T
            else:
                df = df.drop(columns=[c for c in df.columns if c in wholeData.index])
                wholeData = pd.concat([wholeData, df.T])
        wholeData.loc[wholeData['Firearms, type not stated'].isna(),'Firearms, type not stated'] = \
            wholeData.loc[wholeData['Firearms, type not stated'].isna(),'Other weapons or Firearms, type not stated']
        wholeData = wholeData.drop(columns=['Other weapons or Firearms, type not stated'])
        column_to_move = wholeData.pop('Firearms, type not stated')  # Remove 'colC' and store it
        wholeData.insert(6, 'Firearms, type not stated', column_to_move) # Insert 'colC' at index 0
        #wholeData.to_excel('wholeData.xlsx',index=True)
        
        percDf = wholeData.copy()
        nonTotals = [c for c in wholeData if c != 'Total']
        print(percDf[nonTotals].shape,len(percDf['Total']))
        percDf.loc[:,nonTotals] = 100 * wholeData.loc[:,nonTotals].div(wholeData['Total'], axis=0)
        percDf['Total'] = 100 * percDf['Total'] / percDf['Total']  
        percDf.to_excel('percDf.xlsx',index=True)
        self.dataDict[self.MW_STRING] = \
                                  {'All crime' :
                                  {'All groups' :
                                  {'Percentages (%)' : percDf,
                                   'Volume' : wholeData}}}

        self.deep_update(self.metaDict, \
                                  {self.MW_STRING : 
                                  {'All crime' :
                                  {'All groups' :
                                  {'Percentages (%)' : {'Volume' : wholeData},
                                   'Volume' : {}}}}})         

    def joinCitiesAndStates(self, cities, states):
        return cities + ', ' + states.apply(lambda s: self.statesDict[s])

    def aggregateColumns(self, df, targetCol, sourceCol1, sourceCol2 = None):
        if sourceCol2 is None:
            df.loc[df[targetCol].isna(), targetCol] = df.loc[df[targetCol].isna(), sourceCol1]
        else:
            df.loc[df[targetCol].isna(), targetCol] = df.loc[df[targetCol].isna(), [sourceCol1, sourceCol2]].max(axis = 1)
        return df

    def cleanEmptyCells(self, df):
        df = df.replace(r'^\s*$', np.nan, regex=True)
        df = df.replace(to_replace=[None], value=np.nan)
        return df

    def dropEmptyColumns(self, df):
        return df.dropna(axis='columns', how='all')

    def dropFooter(self, df):
        empties = df.iloc[:, 1:].isnull().all(axis=1)
        emptyIdxs = empties[empties].index
        if len(emptyIdxs) == 0:
            return df, pd.DataFrame(columns=df.columns)
        first_footer_index = emptyIdxs[0]
        footer = df.loc[first_footer_index:,:]
        intLoc = df.index.get_loc(first_footer_index)
        df = df.iloc[:intLoc,:]
        return df, footer

    def seperateRates(self, df):
        rateCols = [c for c in df.columns if ' rate' in c]
        dfRates = df[rateCols]
        dfVolumes = df.drop(columns=rateCols)
        dfRates = dfRates.rename(columns={c : c.replace(' rate','') for c in rateCols})
        return dfRates, dfVolumes

    def cleanColumnNames(self, df):
        for cname in df.columns:
            try:
                newcName = self.cleanName(cname)
            except Exception as e:
                # Catch-all for any other type of exception
                print(f"An unexpected error occurred: {e}")
                print(cname)
                raise Exception(e)
            df = df.rename(columns={cname : newcName})
        return df

    def cleanRows(self, column):
        return column.apply(self.cleanName)

    def cleanName(self, name):
        if 'Unnamed' in name:
            return name
        newName = name.replace('\n',' ')
        newName = re.sub(r"\s+", " ", newName)
        newName = newName.strip()
        newName = re.sub(r"\s*\d+(,\s*\d+)+$", "", newName)
        if newName[-1].isdigit() and not bool(re.search(r'\.\d+$', newName)):
            newName = self.remove_trailing_numbers(newName)
        return newName

    def remove_trailing_numbers(self, s):
        if not s:  # Handle empty string case
            return s

        index = len(s) - 1
        while index >= 0 and s[index].isdigit():
            index -= 1
        return s[:index + 1]

    def get_directories_in_path(self, directory_path):
        """
        Returns a list of all immediate subdirectories within the given directory.

        Args:
            directory_path (str): The path to the directory to search.

        Returns:
            list: A list of strings, where each string is the name of a subdirectory.
        """
        directories = []
        try:
            # List all entries in the directory
            for entry in os.listdir(directory_path):
                full_path = os.path.join(directory_path, entry)
                # Check if the entry is a directory
                if os.path.isdir(full_path):
                    directories.append(entry)
        except FileNotFoundError:
            print(f"Error: Directory not found at '{directory_path}'")
        except Exception as e:
            print(f"An error occurred: {e}")
        return directories

    def find_files_with_string_in_name(self, root_dir, search_string):
        """
        Searches a directory and its subdirectories for files whose names
        contain a specific string.

        Args:
            root_dir (str): The starting directory for the search.
            search_string (str): The string to search for within filenames.

        Returns:
            list: A list of full paths to files whose names contain the search string.
        """
        found_files = []
        for dirpath, _, filenames in os.walk(root_dir):
            for filename in filenames:
                if search_string in filename:
                    full_path = os.path.join(dirpath, filename)
                    found_files.append(full_path)
        return found_files

    def deep_update(self, target_dict, source_dict):
        """
        Recursively updates a nested dictionary.
        Values that are dictionaries are merged, others are overwritten.
        """
        for key, value in source_dict.items():
            if key in target_dict and isinstance(target_dict[key], dict) and isinstance(value, dict):
                # If both values are dictionaries, recurse
                self.deep_update(target_dict[key], value)
            else:
                # Otherwise, overwrite or add the value
                target_dict[key] = value

    def getMinAndMaxYears(self, tableName):
        minYear = None
        maxYear = None
        for var, varDict in self.dataDict[tableName].items():
            for group, groupDict in varDict.items():
                for measure, data in groupDict.items():
                    if minYear is None:
                        minYear = data.index.min()
                    else:
                        minYear = min(minYear, data.index.min())
                    if maxYear is None:
                        maxYear = data.index.max()
                    else:
                        maxYear = max(maxYear, data.index.max())
        return minYear, maxYear 

    def getSeries(self, tableName, varName, groupName, measureName):
        return list(self.dataDict[tableName][varName][groupName][measureName].columns)

class CheckListDialog(QDialog):

    def __init__(
        self,
        parent=None,
        name='Check List Dialog',
        stringlist=[],
        checkedItems = None,
        icon=None
        ):
        super().__init__(parent)

        self.name = name
        self.icon = icon
        self.model = QStandardItemModel()
        self.listView = QListView()
        self.checkedItems = checkedItems
        self.stringlist = stringlist

        for string in self.stringlist:
            item = QStandardItem(string)
            item.setCheckable(True)
            check = \
                (Qt.Checked if (self.checkedItems is not None and string in self.checkedItems) else Qt.Unchecked)
            item.setCheckState(check)
            self.model.appendRow(item)

        self.listView.setModel(self.model)

        self.okButton = QPushButton('OK')
        self.cancelButton = QPushButton('Cancel')
        self.selectButton = QPushButton('Select All')
        self.unselectButton = QPushButton('Unselect All')

        hbox = QHBoxLayout()
        hbox.addStretch(1)
        hbox.addWidget(self.okButton)
        hbox.addWidget(self.cancelButton)
        hbox.addWidget(self.selectButton)
        hbox.addWidget(self.unselectButton)

        vbox = QVBoxLayout(self)
        vbox.addWidget(self.listView)
        vbox.addStretch(1)
        vbox.addLayout(hbox)

        self.setWindowTitle(self.name)
        if self.icon:
            self.setWindowIcon(self.icon)

        self.okButton.clicked.connect(self.onAccepted)
        self.cancelButton.clicked.connect(self.onRejected)
        self.selectButton.clicked.connect(self.select)
        self.unselectButton.clicked.connect(self.unselect)

    def onAccepted(self):
        self.choices = [self.model.item(i).text() for i in
                        range(self.model.rowCount())
                        if self.model.item(i).checkState()
                        == Qt.Checked]
        self.accept()

    def onRejected(self):
        self.choices = list(set(self.checkedItems).intersection(set(self.stringlist)))
        self.reject()

    def select(self):
        for i in range(self.model.rowCount()):
            item = self.model.item(i)
            item.setCheckState(Qt.Checked)

    def unselect(self):
        for i in range(self.model.rowCount()):
            item = self.model.item(i)
            item.setCheckState(Qt.Unchecked)

class UCRWindow(QWidget):

    def __init__(self, mainMWindow):
        super().__init__()
        self.setWindowTitle("Crime python GUI - UCR viewing application")
        self.resize(1500, 700)
        self.hide()

        self.fetcher = UCRDataFetcher()

        self.mainMenuWindow = mainMWindow

        self.seriesDict = {}

        self.currentData = pd.DataFrame()

        # Create an outer layout
        outerLayout = QHBoxLayout()

        # Create year range selection
        yearLayout = QHBoxLayout()
        yearLayout.addWidget(QLabel("Year minimum:"))
        self.year_range_min = QComboBox()
        yearLayout.addWidget(self.year_range_min)
        yearLayout.addWidget(QLabel("Year maximum:"))
        self.year_range_max = QComboBox()
        yearLayout.addWidget(self.year_range_max)

        # Create the plot layout
        rightLayout = QVBoxLayout()
        self.graphWidget = pg.PlotWidget()
        rightLayout.addWidget(self.graphWidget)
        # follow mouse movement on graph
        self.proxy = pg.SignalProxy(self.graphWidget.scene().sigMouseMoved, rateLimit=60, slot=self.mouseMoved)
        self.view_box = self.graphWidget.getViewBox()

        # Create the sidebar on the left
        self.tableBox = QComboBox()
        self.tableBox.addItems(self.fetcher.dataDict.keys())
        self.tableBox.setCurrentIndex(-1)

        self.variableBox = QComboBox()
        
        self.groupBox = QComboBox()
        
        self.measureBox = QComboBox()
         
        leftLayout = QVBoxLayout()
        leftLayout.addWidget(QLabel("Select crime statistic:"))
        leftLayout.addWidget(self.tableBox)
        leftLayout.addWidget(QLabel("Select variable/relationship:"))
        leftLayout.addWidget(self.variableBox)
        leftLayout.addWidget(QLabel("Select group/variable:"))
        leftLayout.addWidget(self.groupBox)
        leftLayout.addWidget(QLabel("Select measure:"))
        leftLayout.addWidget(self.measureBox)
        leftLayout.addLayout(yearLayout)

        self.tableBox.activated.connect(self.selectTable)
        self.variableBox.activated.connect(self.selectVariable)
        self.groupBox.activated.connect(self.selectGroup)
        self.measureBox.activated.connect(self.selectMeasure)
        self.year_range_min.activated.connect(self.updatePlot)
        self.year_range_max.activated.connect(self.updatePlot)

        # Create the series selection
        self.series_button = QPushButton("Select data series")
        self.series_button.clicked.connect(self.open_series_dialog) # Connect button to a slot
        leftLayout.addWidget(self.series_button)

        # Create notifications checkbox
        self.notes_check = QCheckBox("Notifications on/off:")
        self.notes_check.stateChanged.connect(self.updatePlot)
        leftLayout.addWidget(self.notes_check)

        # UCR path
        link = "https://www.fbi.gov/how-we-can-help-you/more-fbi-services-and-information/ucr"
        self.weblink = QLabel(f'Data: <a href="{link}">About the UCR data</a>')
        self.weblink.setOpenExternalLinks(True)
        leftLayout.addWidget(self.weblink)

        # Data path
        link = "https://www.fbi.gov/how-we-can-help-you/more-fbi-services-and-information/ucr/publications"
        self.datalink = QLabel(f'Data: <a href="{link}">Data is from here</a>')
        self.datalink.setOpenExternalLinks(True)
        leftLayout.addWidget(self.datalink)

        self.backButton = QPushButton("Go back to main menu")
        self.backButton.clicked.connect(self.goBackToMenu) # Connect button to a slot
        leftLayout.addWidget(self.backButton)

        # Notifications
        self.notesLabel = QLabel('')
        self.notesLabel.setWordWrap(True)
        leftLayout.addWidget(self.notesLabel)

        leftLayout.addStretch()

        # Nest the inner layouts into the outer layout
        outerLayout.addLayout(leftLayout)
        outerLayout.addLayout(rightLayout)
        # Set the Window's main layout
        self.setLayout(outerLayout)

    def closeEvent(self, event):
        # Optional: Ask for confirmation before closing
        reply = QMessageBox.question(self, 'Message',
                                     "Are you sure to quit?", QMessageBox.Yes |
                                     QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            event.accept()  # Accept the close event, allowing the window to close
            print("Application is closing.")
            # Perform any cleanup actions here
            sys.exit(0) # Explicitly exit the application
        else:
            event.ignore()  # Ignore the close event, keeping the window open

    def goBackToMenu(self):
        self.mainMenuWindow.show()
        self.hide()

    def selectTable(self):
        self.variableBox.clear()
        self.groupBox.clear()
        self.measureBox.clear()
        self.notesLabel.setText(self.fetcher.getNotification(self.tableBox.currentText()))
        QApplication.setOverrideCursor(Qt.WaitCursor)
        self.fetcher.loadTable(self.tableBox.currentText())
        QApplication.restoreOverrideCursor()
        self.variableBox.addItems(self.fetcher.dataDict[self.tableBox.currentText()].keys())
        self.variableBox.setCurrentIndex(-1)
        self.updateYears()
        self.updatePlot()

    def selectVariable(self):
        currentTable = self.tableBox.currentText()
        currentVariable = self.variableBox.currentText()
        currentGroup = self.groupBox.currentText()
        currentMeasure = self.measureBox.currentText()
        nextGroups = self.fetcher.dataDict[currentTable][currentVariable].keys()
        self.groupBox.clear()
        self.measureBox.clear()
        self.groupBox.addItems(nextGroups)
        if currentGroup in nextGroups:
            self.groupBox.setCurrentText(currentGroup)
            nextMeasures = self.fetcher.dataDict[currentTable][currentVariable][currentGroup].keys()
            self.measureBox.addItems(nextMeasures)
            if currentMeasure in nextMeasures:
                self.measureBox.setCurrentText(currentMeasure)
            else:
                self.measureBox.setCurrentIndex(-1)
        else:
            self.groupBox.setCurrentIndex(-1)
        self.updateSeries()
        self.updatePlot()

    def selectGroup(self):
        currentTable = self.tableBox.currentText()
        currentVariable = self.variableBox.currentText()
        currentGroup = self.groupBox.currentText()
        currentMeasure = self.measureBox.currentText()
        nextMeasures = self.fetcher.dataDict[currentTable][currentVariable][currentGroup].keys()
        self.measureBox.clear()
        self.measureBox.addItems(nextMeasures)
        if currentMeasure in nextMeasures:
            self.measureBox.setCurrentText(currentMeasure)
        else:
            self.measureBox.setCurrentIndex(-1)
        self.updateSeries()
        self.updatePlot()

    def selectMeasure(self):
        self.updateSeries()
        self.updatePlot()

    def updateYears(self):
        currentMinYear = self.year_range_min.currentText()
        currentMaxYear = self.year_range_max.currentText()
        self.year_range_min.clear()
        self.year_range_max.clear()
        minYear, maxYear = self.fetcher.getMinAndMaxYears(self.tableBox.currentText())
        sorted_years = [str(y) for y in range(minYear, maxYear + 1)]
        self.year_range_min.addItems(sorted_years[:-1])
        self.year_range_max.addItems(sorted_years[1:])
        self.year_range_min.setCurrentText(sorted_years[0])
        self.year_range_max.setCurrentText(sorted_years[-1])

    def updateSeries(self):
        table = self.tableBox.currentText()
        variable = self.variableBox.currentText()
        group = self.groupBox.currentText()
        measure = self.measureBox.currentText()
        if '' in [table,variable,group,measure]:
            return
        series = self.fetcher.getSeries(table, variable, group, measure)
        for s in series:
            if s not in self.seriesDict.keys():
                self.seriesDict[s] = {'number' : len(self.seriesDict.keys()),
                                             'checked' : len(series) < 10}
        delete = []
        for s in self.seriesDict.keys():
            if s not in series:
                delete.append(s)
        for d in delete:
            del self.seriesDict[d] 

    def getData(self):
        table = self.tableBox.currentText()
        variable = self.variableBox.currentText()
        group = self.groupBox.currentText()
        measure = self.measureBox.currentText()
        if '' in [table, variable, group, measure]:
            return pd.DataFrame()
        return self.fetcher.dataDict[table][variable][group][measure]

    def getMetaData(self):
        table = self.tableBox.currentText()
        variable = self.variableBox.currentText()
        group = self.groupBox.currentText()
        measure = self.measureBox.currentText()
        if '' in [table, variable, group, measure]:
            return pd.DataFrame()
        return self.fetcher.metaDict[table][variable][group][measure]

    def getExtraToolTips(self, series, year):
        mDict = self.getMetaData()
        text = ''
        if 'Volume' in mDict.keys():
            text = f'{text}\nVolume: {mDict['Volume'].loc[year,series]:,.0f}'
        if 'Population' in mDict.keys():
            text = f'{text}\nPopulation: {mDict['Population'].loc[year,series]:,.0f}'
        if 'Demographic' in mDict.keys():
            text = f'{text}\nDemographic population: {mDict['Demographic'].loc[year,series]:,.0f}'
        if 'Agencies' in mDict.keys():
            text = f'{text}\nNumber of agencies reporting: {mDict['Agencies'].loc[year,series]:,.0f}'
        if 'Notes' in mDict.keys() and isinstance(mDict['Notes'].loc[year,series], str):
            text = f'{text}{mDict['Notes'].loc[year,series]}'
        return text
        
    def updatePlot(self):
        self.graphWidget.clear()
        self.graphWidget.addLegend()
        df = self.getData()
        if (len(df) == 0) or (int(self.year_range_min.currentText()) >= int(self.year_range_max.currentText())):
            return
        checkedSeries = [c for c in self.seriesDict.keys() if self.seriesDict[c]['checked']]
        self.currentData = df[[s for s in df.columns if s in checkedSeries]]
        # Set x-axis label
        self.graphWidget.getAxis('bottom').setLabel(text="Year")
        # Set y-axis label with styling
        self.graphWidget.getAxis('left').setLabel(text=self.measureBox.currentText())
        # Set title
        minYear, maxYear = self.fetcher.getMinAndMaxYears(self.tableBox.currentText())
        self.graphWidget.setTitle(f'{self.tableBox.currentText()} {minYear}-{maxYear}')
        for series in checkedSeries:
            if len(self.currentData) > 0:
                # color scheme
                seriesNumber = self.seriesDict[series]['number']
                rgb = self.generate_distinct_color(seriesNumber, df.shape[1])
                linestyle = self.generate_line_style(seriesNumber)
                # plot
                pen_style = pg.mkPen(rgb, width=2, style=linestyle)
                year = self.currentData.index.to_numpy().astype(float)
                value = self.currentData[series].to_numpy().astype(float)
                self.graphWidget.plot(year, value, pen=pen_style, name=series)
                self.updateNotes(series)
        # Get the ViewBox
        view_box = self.graphWidget.getPlotItem().vb

        # Set the X-axis range from 0 to 6
        view_box.setXRange(int(self.year_range_min.currentText()), int(self.year_range_max.currentText()))

    def updateNotes(self, series):
        if self.notes_check.isChecked():
            mDict = self.getMetaData()
            if 'Important' not in mDict.keys():
                return

            # Create a ScatterPlotItem
            scatter_item = pg.ScatterPlotItem(
                size=12,
                symbol='o',
                brush=pg.mkBrush(255, 255, 255, 200)
            )
            years = mDict['Important'].index[~mDict['Important'][series].isna()]
            values = self.currentData[series][~mDict['Important'][series].isna()]

            # Set the data for the scatter plot
            scatter_item.setData(years, values)

            # Add the scatter plot item to the plot Window
            self.graphWidget.addItem(scatter_item)

    def mouseMoved(self, e):
        pos = e[0]
        all_plot_items = self.graphWidget.listDataItems()       
        if self.graphWidget.sceneBoundingRect().contains(pos) and (len(self.currentData) > 0):
            mousePoint = self.graphWidget.getPlotItem().vb.mapSceneToView(pos)
            closestYear = round(mousePoint.x())
            yearData = self.currentData[self.currentData.index == closestYear]
            self.graphWidget.setToolTip('')
            if (self.currentData.shape[1] > 0) and (len(yearData) > 0):
                closestSeries = yearData.columns.to_series().iloc[(yearData-mousePoint.y()).abs().to_numpy().argsort()[0,0]]
                closestValue = yearData[[closestSeries]].iloc[0].to_numpy()[0]
                x_range, y_range = self.view_box.viewRange()
                if (np.abs(mousePoint.x() - closestYear) < 0.025*np.abs(x_range[0] - x_range[1])) and (np.abs(mousePoint.y() - closestValue) < 0.025*np.abs(y_range[0] - y_range[1])):
                    p = self.graphWidget.mapToGlobal(QPoint(int(pos.x()),int(pos.y()))) 
                    if closestValue.is_integer() or isinstance(closestValue, int):
                        tooltip_text = f"{closestSeries}: {closestYear}\nValue: {closestValue:,.0f}"
                    elif np.abs(closestValue) < 1:                   
                        tooltip_text = f"{closestSeries}: {closestYear}\nValue: {closestValue:.2f}"
                    elif np.abs(closestValue) < 20:
                        tooltip_text = f"{closestSeries}: {closestYear}\nValue: {closestValue:.1f}"
                    else:
                        tooltip_text = f"{closestSeries}: {closestYear}\nValue: {closestValue:,.1f}" 
                    tooltip_text = f'{tooltip_text}{self.getExtraToolTips(closestSeries,closestYear)}'
                    #QToolTip.showText(p, tooltip_text)
                    for item in all_plot_items:
                        if item.name() is None:
                            continue
                        if item.name() == closestSeries:
                            pen_style = pg.mkPen('w', width=4)
                            item.setPen(pen_style)
                            self.graphWidget.setToolTip(tooltip_text)
                        else:
                            pen_style = self.getPen(item.name())
                            item.setPen(pen_style)
                else:
                    for item in all_plot_items:
                        if item.name() is None:
                            continue
                        pen_style = self.getPen(item.name())
                        item.setPen(pen_style)
                    #QToolTip.hideText()
                
    def getPen(self, level, lineWidth=2):
        # color scheme
        df = self.getData()
        levelNumber = self.seriesDict[level]['number']
        rgb = self.generate_distinct_color(levelNumber, df.shape[1])
        linestyle = self.generate_line_style(levelNumber)
        pen_style = pg.mkPen(rgb, width=lineWidth, style=linestyle)
        return pen_style

    def open_series_dialog(self):
        df = self.getData()
        if len(df) == 0:
            return
        availableSeries = list(self.getData().columns)
        checkedSeries = [c for c in self.seriesDict.keys() if self.seriesDict[c]['checked']]
        dialog = CheckListDialog(self, 'Series', availableSeries, checkedItems=checkedSeries) # Pass 'self' (the main Window) as parent
        dialog.exec_() # Use exec_() for a modal dialog, or show() for non-modal
        addedSeries = [c for c in dialog.choices if c not in checkedSeries]
        deletedSeries = [c for c in checkedSeries if c not in dialog.choices and c in availableSeries]
        for c in addedSeries:
            self.seriesDict[c]['checked'] = True
        for c in deletedSeries:
            self.seriesDict[c]['checked'] = False
        # check for data
        self.updatePlot()

    def generate_distinct_color(self, step, num_steps): 
        # Generate a color in HSL 
        hue = step / num_steps  # evenly spaced hues 
        lightness = 0.5  # fixed lightness 
        saturation = 1.0  # full saturation 
        rgb = colorsys.hls_to_rgb(hue, lightness, saturation) 
        # Convert from [0, 1] to [0, 255] and round 
        rgb = tuple(int(c * 255) for c in rgb)  
        return rgb 

    def generate_line_style(self, step):
        # line style
        if step % 5 == 0:
            linestyle = Qt.SolidLine
        elif step % 5 == 1:
            linestyle = Qt.DashLine
        elif step % 5 == 2:
            linestyle = Qt.DotLine
        elif step % 5 == 3:
            linestyle = Qt.DashDotLine
        elif step % 5 == 4:
            linestyle = Qt.DashDotDotLine
        return linestyle

#if __name__ == "__main__":
#    app = QApplication(sys.argv)
#    window = UCRWindow()
#    window.show()
#    sys.exit(app.exec())