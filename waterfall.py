import xlwings as xw
import pandas as pd
import numpy as np

class AssumptionObject:
    def __init__(self):
        self.region_name = ''
        self.dataRowStart = ''
        self.df = ''
        self.filtered_df = ''
        self.colEmpty = []
        self.isConstant = False

class AssumptionRegion:
    def __init__(self):
        self.region = []

class Formulas:
    def __init__(self):
        self.matrix = [([None] * 13) for _ in range(len(dateValues))]

class PipelineCreate:
    def __init__(self):
        self.values = []

wb = xw.Book('VBA and Python Interview Exercise v1.xlsx')
wb.app.calculation = 'automatic'

wsOutput = wb.sheets['Output']
wsAssumptions = wb.sheets['Assumptions']

dimensionsTitles = wsOutput.range("A2:H2").value
dimensionsValues = wsOutput.range("A3:H15").value
dateValues = wsOutput.range("I2:AR2").value

aRegion = AssumptionRegion()
formulas = Formulas()
create = PipelineCreate()

def getAssumptionsDF():
    numRowsAssumptions = wsAssumptions.used_range.rows.count
    currentRow = wsAssumptions.range("B1")

    while currentRow.row < numRowsAssumptions:
        if currentRow.value != None:
            currentRow = buildAssumptionsDF(currentRow)
        else:
            currentRow = currentRow.end('down')

def buildAssumptionsDF(currentRow):
    aObj = AssumptionObject()
    currentRegion = currentRow.current_region

    regionTitles = wsAssumptions.range((currentRow.row+1, currentRow.column),(currentRow.row+1, currentRow.column+currentRegion.columns.count-1))
    regionData = wsAssumptions.range((currentRow.row+2, currentRow.column),(currentRow.row+currentRegion.rows.count, currentRow.column+currentRegion.columns.count-1))

    aObj.region_name = currentRow.value
    aObj.dataRowStart = currentRow.row+2
    aObj.df = pd.DataFrame(np.array(regionData.value), columns=regionTitles.value)

    dimensionsToCheck = dimensionsTitles[:7]

    for title in dimensionsToCheck:
        is_empty = aObj.df[title].isnull().all()
        aObj.colEmpty.append(is_empty)
        
    if aObj.df['Timeframe'].isnull().all() or aObj.region_name == "Sales Cycle":
        aObj.isConstant = True
    
    aRegion.region.append(aObj)

    currentRow = wsAssumptions.range((currentRow.row+currentRegion.rows.count, currentRow.column))

    return currentRow


def GetProductSplitConstant(aObj):
    filtered_df = aObj.df

    for i in range(0, len(dimensionsTitles)-1):
        if aObj.colEmpty[i] != True:
            filtered_df = filtered_df.loc[filtered_df[dimensionsTitles[i]] == dimensionsValues[0][i]]

    row = filtered_df.index.to_list()
    wsRange = wsAssumptions.range((row[0] + aObj.dataRowStart, 11)).get_address(True, True, include_sheetname=True)
    return wsRange

    
def GetSalesCycleConstants(aObj):
    filtered_df = aObj.df

    for i in range(0, len(dimensionsTitles)-1):
        if aObj.colEmpty[i] != True:
            filtered_df = filtered_df.loc[filtered_df[dimensionsTitles[i]] == dimensionsValues[0][i]]

    row = filtered_df.index.to_list()
    
    for j in range(0, len(row)):
        row[j] = wsAssumptions.range((row[j] + aObj.dataRowStart, 11)).get_address(True, True, include_sheetname=True)
    return row

def cycleRegions():
    for i in range(0, len(aRegion.region)):
        aObj = aRegion.region[i]

        if aObj.isConstant != True:
            parseRegion(aObj)
    return

def parseRegion(aObj):
    filtered_df = aObj.df

    for i in range(0, len(dimensionsTitles)-1):
        if aObj.colEmpty[i] != True:
            filtered_df = filtered_df.loc[filtered_df[dimensionsTitles[i]] == dimensionsValues[0][i]]

    aObj.filtered_df = filtered_df
    return

def runThroughDates():
    refMatrix = []
    for i in range(0, len(dateValues)):
        currentDate = dateValues[i]

        if currentDate.year < 2027:
            refMatrix.append(getCellReferenceForDate(currentDate))
        else:
            break

    return refMatrix

def getCellReferenceForDate(currentDate):
    refArray = [None] * 6
    for j in range(0, len(aRegion.region)):
        aObj = aRegion.region[j]
        
        if aObj.isConstant != True:
            if aObj.df['Timeframe'].dtype == 'object':
                dateVal = convertDateToQuarter(currentDate) 
            else:
                dateVal = currentDate

            refArray[j] = getCellReference(dateVal, aObj)

    return refArray

def getCellReference(dateVal, aObj):
    date_df = aObj.filtered_df.loc[aObj.filtered_df['Timeframe'] == dateVal]

    row = date_df.index.to_list()
    wsRange = wsAssumptions.range((row[0] + aObj.dataRowStart, 11)).get_address(True, True, include_sheetname=True)


    return wsRange

def convertDateToQuarter(currentDate):
    yearNum = currentDate.year
    monthNum = currentDate.month

    if monthNum < 4:
        quarterNum = str(yearNum) + '-Q1'
    elif monthNum > 3 and monthNum < 7:
        quarterNum = str(yearNum) + '-Q2'
    elif monthNum > 5 and monthNum < 10:
        quarterNum = str(yearNum) + '-Q3'
    elif monthNum > 9:
        quarterNum = str(yearNum) + '-Q4'

    return quarterNum

def buildCellCatalogue(refMatrix, salesCycleConsts, prodSplitConst):
    catalogue = [([None] * len(salesCycleConsts)) for _ in range(len(refMatrix))]

    print(refMatrix[0][0] + '*' + refMatrix[0][1])
    for n in range(0, len(refMatrix)):
        create.values.append(refMatrix[n][0] + '*' + refMatrix[n][1])
        for m in range(0, len(salesCycleConsts)):
            nonConsts = prodSplitConst

            for p in range(0, len(refMatrix[n])):
                if refMatrix[n][p] != None:
                    nonConsts += '*' + refMatrix[n][p]

            catalogue[n][m] = '(' + salesCycleConsts[m] + '*' + nonConsts  + ')'

    return catalogue

def builFormulas(catalogue):
    maxval = len(dateValues)
    for x in range(0, len(catalogue)):
        cellRef = catalogue[x]
        colStart = x
        cycleCatalogue(cellRef, colStart, maxval)
    return

def cycleCatalogue(cellRef, colStart, maxCol):
    for y in range(0, len(cellRef)):
        arrayOffset = y
        colNum = colStart + y

        if colNum > maxCol-1:
            break

        insertReferences(cellRef, colNum, arrayOffset)
    return

def insertReferences(cellRef, colNum, arrayOffset):
    for z in range(0, 13):
        if z + arrayOffset > len(cellRef)-1:
            break

        currentValue = formulas.matrix[colNum][z]

        if currentValue == None:
            formulas.matrix[colNum][z] = cellRef[z + arrayOffset]
        else:
            formulas.matrix[colNum][z] = currentValue + '+' + cellRef[z + arrayOffset]
    return

def BuildWaterfall():
    for x in range(0, len(dateValues)):
        for y in range(0, 13):
            wsOutput.range((3+y,9+x)).value = '=' + formulas.matrix[x][y]
            
    return

def BuildPipelineCreate():
    for x in range(0, len(create.values)):
        wsOutput.range((17,9+x)).color = (11, 48, 64)
        wsOutput.range((17,9+x)).value = 'Create'
        wsOutput.range((17,9+x)).font.color = (255, 255, 255)
        wsOutput.range((18,9+x)).value = '=' + create.values[x]
        wsOutput.range((18,9+x)).number_format = '#,##0'
    return

getAssumptionsDF()
prodSplitConst = GetProductSplitConstant(aRegion.region[2])
salesCycleConsts = GetSalesCycleConstants(aRegion.region[5])

cycleRegions()
refMatrix = runThroughDates()
catalogue = buildCellCatalogue(refMatrix, salesCycleConsts, prodSplitConst)
builFormulas(catalogue)
print(len(create.values))
BuildWaterfall()
BuildPipelineCreate()