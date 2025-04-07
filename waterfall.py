import xlwings as xw
import pandas as pd
import numpy as np
import copy

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

class PipelineWaterfall:
    def __init__(self):
        self.aRegion = AssumptionRegion()
        # self.formulas = Formulas()
        self.wb = ''
        self.wsOutput = ''
        self.wsAssumptions = ''
        self.dimensionsTitles = ''
        self.dimensionsValues = ''
        self.dateValues = ''
        self.catalogue = ''
        self.output = {
            'dimension_titles': [],
            'dimensions': [],
            'value_titles': [],
            'values': []
        }

    def buildWaterFall(self, filename, path):
        self.wb = xw.Book(path + '/' + filename)
        # self.wb = xw.Book('uploads\VBA_and_Python_Interview_Exercise_v1.xlsx')
        self.wb.app.calculation = 'automatic'
        self.assignValuesFromFile()

        # print(self.dimensionsValues)
        self.getAssumptionsDF()
        prodSplitConst = self.GetProductSplitConstant(self.aRegion.region[2])
        salesCycleConsts = self.GetSalesCycleConstants(self.aRegion.region[5])

        self.cycleRegions()
        refMatrix = self.runThroughDates()
        self.catalogue = self.buildCellCatalogue(refMatrix, salesCycleConsts, prodSplitConst)
        self.getCellPositions()

        self.output['dimension_titles'] = self.dimensionsTitles
        self.output['dimensions'] = self.dimensionsValues
        self.output['value_titles'] = self.dateValues
        self.output['values'] = self.catalogue

        return self.output

    def assignValuesFromFile(self):
        self.wsOutput = self.wb.sheets['Output']
        self.wsAssumptions = self.wb.sheets['Assumptions']

        self.dimensionsTitles = self.wsOutput.range("A2:H2").value
        self.dimensionsValues = self.wsOutput.range("A3:H21").value
        self.dateValues = self.wsOutput.range("I2:AR2").value

    def getAssumptionsDF(self):
        """
        Jump across rows to find each new "region" which 
        contains a datasheet needed for calculation

        Args:
            None
        Returns:
            None 
        """
        numRowsAssumptions = self.wsAssumptions.used_range.rows.count
        currentRow = self.wsAssumptions.range("B1")

        while currentRow.row < numRowsAssumptions:
            if currentRow.value != None:
                currentRow = self.buildAssumptionsDF(currentRow)
            else:
                currentRow = currentRow.end('down')

    def buildAssumptionsDF(self, currentRow):
        """
        Gets the current region and builds the AssumptionObject class that will house 
        necessary information of each region. This includes region name, the row the region
        starts on, a dataframe representing the region data, and whether the data is constant

        Args:
            currentRow (range object): The range position of the first row of the region
        Returns:
            currentRow (range object): Updated range position that is below the last row 
            of the current region
        """
        aObj = AssumptionObject()
        currentRegion = currentRow.current_region

        regionTitles = self.wsAssumptions.range((currentRow.row+1, currentRow.column),(currentRow.row+1, currentRow.column+currentRegion.columns.count-1))
        regionData = self.wsAssumptions.range((currentRow.row+2, currentRow.column),(currentRow.row+currentRegion.rows.count, currentRow.column+currentRegion.columns.count-1))

        aObj.region_name = currentRow.value
        aObj.dataRowStart = currentRow.row+2
        aObj.df = pd.DataFrame(np.array(regionData.value), columns=regionTitles.value)

        dimensionsToCheck = self.dimensionsTitles[:7]

        for title in dimensionsToCheck:
            is_empty = aObj.df[title].isnull().all()
            aObj.colEmpty.append(is_empty)
            
        if aObj.df['Timeframe'].isnull().all() or aObj.region_name == "Sales Cycle":
            aObj.isConstant = True
        
        self.aRegion.region.append(aObj)

        currentRow = self.wsAssumptions.range((currentRow.row+currentRegion.rows.count, currentRow.column))

        return currentRow


    def GetProductSplitConstant(self, aObj):
        """
        Gets the constant value of the "Product Split" region

        Args:
            aObj (AssumptionObject class): The object that contains the data for this region
        Returns:
            wsRange (range object): the range object that represents the address of the 
            constant value
        """
        filtered_df = aObj.df

        for i in range(0, len(self.dimensionsTitles)-1):
            if aObj.colEmpty[i] != True:
                filtered_df = filtered_df.loc[filtered_df[self.dimensionsTitles[i]] == self.dimensionsValues[0][i]]

        row = filtered_df.index.to_list()
        wsRange = self.wsAssumptions.range((row[0] + aObj.dataRowStart, 11)).get_address(True, True, include_sheetname=False)
        wsValue = self.wsAssumptions.range((row[0] + aObj.dataRowStart, 11)).value

        return {'address': wsRange, 'value': round(wsValue, 3), 'sheet': "Assumptions", 'table': aObj.region_name}

        
    def GetSalesCycleConstants(self, aObj):
        """
        Gets the constant values of the "Sales Cycle" region

        Args:
            aObj (AssumptionObject class): The object that contains the data for this region
        Returns:
            wsRange (range object): the range object that represents the addresses of the 
            constant values
        """
        filtered_df = aObj.df

        for i in range(0, len(self.dimensionsTitles)-1):
            if aObj.colEmpty[i] != True:
                filtered_df = filtered_df.loc[filtered_df[self.dimensionsTitles[i]] == self.dimensionsValues[0][i]]

        row = filtered_df.index.to_list()
        
        for j in range(0, len(row)):
            wsRange = self.wsAssumptions.range((row[j] + aObj.dataRowStart, 11)).get_address(True, True, include_sheetname=False)
            wsValue = self.wsAssumptions.range((row[j] + aObj.dataRowStart, 11)).value

            row[j] = {'address': wsRange, 'value': round(wsValue, 4), 'sheet': "Assumptions", 'table': aObj.region_name}
        return row

    def cycleRegions(self):
        """
        Cycles through each AssumptionObject within the AssumptionRegion class' region list

        Args:
            None
        Returns:
            None
        """
        for i in range(0, len(self.aRegion.region)):
            aObj = self.aRegion.region[i]

            if aObj.isConstant != True:
                self.parseRegion(aObj)
        return

    def parseRegion(self, aObj):
        """
        Parses through the data in the AssumptionObject's dataframe, removing any rows that do not
        match with the expected dimensions from the Output worksheet.

        Args:
            aObj (AssumptionObject class): The object that contains the data for a particular region
        Returns:
            None
        """
        filtered_df = aObj.df

        for i in range(0, len(self.dimensionsTitles)-1):
            if aObj.colEmpty[i] != True:
                filtered_df = filtered_df.loc[filtered_df[self.dimensionsTitles[i]] == self.dimensionsValues[0][i]]

        aObj.filtered_df = filtered_df
        return

    def runThroughDates(self):
        """
        Runs through each date from the associated column titles in the Output worksheet.
        Creates the "refMatrix" list that stores the range addresses for each value of each 
        region for the associated date.
        
        Args:
            None
        Returns:
            refMatrix (multidimensional list): The list that contains that values from each 
            region for each date.
        """
        refMatrix = []
        for i in range(0, len(self.dateValues)):
            currentDate = self.dateValues[i]

            if currentDate.year < 2027:
                refMatrix.append(self.getCellReferenceForDate(currentDate))
            else:
                break

        return refMatrix

    def getCellReferenceForDate(self, currentDate):
        """
        Loops through each region and gets the range address for each value associated 
        with that date
        
        Args:
            currentDate (date object): The date being used to get the range address for 
            each value
        Returns:
            refArray (list): The list that contains that values from each 
            region for each date.
        """
        refArray = [None] * 6
        for j in range(0, len(self.aRegion.region)):
            aObj = self.aRegion.region[j]
            
            if aObj.isConstant != True:
                if aObj.df['Timeframe'].dtype == 'object':
                    dateVal = self.convertDateToQuarter(currentDate) 
                else:
                    dateVal = currentDate

                refArray[j] = self.getCellReference(dateVal, aObj)

        return refArray

    def getCellReference(self, dateVal, aObj):
        """
        Gets the range object from the filtered dataframe from the associated
        AssumptionObject using the date value.
        
        Args:
            dateVal (date object | string): The date to be searched
            aObj (AssumptionObject class): The region object whose dataframe is used to 
            retrive the range object for the specified date.
        Returns:
            wsRange (range object): The range object of the associated value for the 
            specified date.
        """
        date_df = aObj.filtered_df.loc[aObj.filtered_df['Timeframe'] == dateVal]

        row = date_df.index.to_list()
        wsRange = self.wsAssumptions.range((row[0] + aObj.dataRowStart, 11)).get_address(True, True, include_sheetname=False)
        wsValue = self.wsAssumptions.range((row[0] + aObj.dataRowStart, 11)).value

        return {'address': wsRange, 'value': round(wsValue, 4), 'sheet': "Assumptions", 'table': aObj.region_name}

    def convertDateToQuarter(self, currentDate):
        """
        Converts the date object to a string that represents the financial quarter and year
        (01/01/2025) => 2025-Q1
        
        Args:
            currentDate (date object): The date to be converted
        Returns:
            quarterNum (string): The quarter and year from the given currentDate
        """
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

    def buildCellCatalogue(self, refMatrix, salesCycleConsts, prodSplitConst):
        """
        Builds the cell catalogue. The cell catalogue is a multi-dimensional list where
        each value stored is a string that represents a single formula for the given dimensions
        in the Output workSheet. 
        
        Args:
            refMatrix (multidimensional list): The list that contains that values from each 
            region for each date.
            salesCycleConsts (range object): The range objects for the const values in the 
            "Sales Cycle" region from "in month" to "24 months".
            prodSplitConst (range object): The range object for the const value in the 
            "Product Split" region.
        Returns:
            catalogue (multi-dimensional list): The single formula for each combination of region
            value addresses for each date.
        """
        catalogue = [([None] * len(salesCycleConsts)) for _ in range(len(refMatrix))]

        for n in range(0, len(refMatrix)):
            cellRef = []
            calcValue = 1
            for k in range(0, len(refMatrix[n])):
                if refMatrix[n][k] != None:
                    cellData = refMatrix[n][k]
                    calcValue = calcValue * cellData['value']
                    cellRef.append(cellData)
                    
            # print(calcValue)
            for m in range(0, len(salesCycleConsts)):
                catArray = copy.deepcopy(cellRef)
                catValue = calcValue

                # print(cellRef)
                # print()
                # print(catArray)
                # print()

                catArray.append(prodSplitConst)
                catArray.append(salesCycleConsts[m])

                catValue = catValue * prodSplitConst['value'] * salesCycleConsts[m]['value']

                catalogue[n][m] = {'cell_id': str(n)+str(m) ,'cell_ref': catArray, 'calc_value': catValue, 'column': None, 'row': None}
                # print({'cell_ref': catArray, 'calc_value': catValue, 'column': None, 'row': None})

        return catalogue

    def getCellPositions(self):
        """
        Cycles through each sublist in the catalogue multi-dimensional list. This sublist is
        represented by the variable "cellRef".
        
        Args:
            catalogue (multi-dimensional list): The single formula for each combination of region
            value addresses for each date.
        Returns:
            None
        """
        maxval = len(self.dateValues)
        print(maxval)
        for x in range(0, len(self.catalogue)):
            # cellRef = self.catalogue[x]
            colStart = x
            self.cycleCatalogue(colStart, maxval)
        return

    def cycleCatalogue(self, colStart, maxCol):
        """
        Loops through each cellRef sublist from the catalogue multi-dimensional list. Creates 
        variables "colNum" and "arrayOffset" needed for correctly building the waterfall. 

        |   Date 1   |   Date 2                | ...
        ---------------------------------------
        | cellRef 1  | cellRef 2  + cellRef 1a | ...
        ---------------------------------------
        | cellRef 1a | cellRef 2a + cellRef 1b | ...
        ---------------------------------------
        | cellRef 1b | cellRef 2b + ...        | 
        
        Args:
            cellRef (list): The sublist of the catalogue that contains the single formula for
            date and dimension.
            colStart (integer): The starting column for the respective date column in the 
            Output worksheet
            maxCol: The maximum column value to be used, in this case it's the column representing
            december 2026.
        Returns:
            None
        """
        for y in range(0, 19):
            arrayOffset = y
            colNum = colStart + y

            if colNum > maxCol-1:
                break

            self.catalogue[colStart][arrayOffset]['column'] = colNum
            self.catalogue[colStart][arrayOffset]['row'] = arrayOffset
            
            # print(self.catalogue[colNum][arrayOffset])
            # print()
            # self.insertReferences(cellRef, colNum, arrayOffset)
        return

    def insertReferences(self, cellRef, colNum, arrayOffset):
        """
        Goes down each column row (y) for the respective dates (colNum) and adds the formula string
        to the previous formula. These are stored in the PipelineCreate object "formulas.matrix"

            | colNum =  1 | colNum =  2     | ...
        -------|------------------------------------
        y = 0 | cellRef 1   | cellRef 1a + ...| ...
        -------|------------------------------------
        y = 1 | cellRef 1a  | cellRef 1b + ...| ...
        -------|------------------------------------
        y = 2 | cellRef 1b  | cellRef 1c + ...| 
        
        Args:
            cellRef (list): The sublist of the catalogue that contains the single formula for
            each date and dimension.
            colNum (integer): The column offset from the starting column
            maxCol: The offset value to start inserting the values from the cellRef list.
        Returns:
            None
        """
        for z in range(0, 13):
            if z + arrayOffset > len(cellRef)-1:
                break

            currentValue = self.formulas.matrix[colNum][z]

            if currentValue == None:
                self.formulas.matrix[colNum][z] = cellRef[z + arrayOffset]
            else:
                self.formulas.matrix[colNum][z] = currentValue + '+' + cellRef[z + arrayOffset]
        return

    def BuildWaterfall(self):
        """
        Inserts the full formula from formulas.matrix into the corresponding cell in 
        the output worksheet.
        
        Args:
            None
        Returns:
            None
        """
        for x in range(0, len(self.dateValues)):
            for y in range(0, 13):
                self.wsOutput.range((3+y,9+x)).value = '=' + self.formulas.matrix[x][y]
        return

    # def BuildPipelineCreate(self, refMatrix, prodSplitConst):
    #     """
    #     Inserts the pipeline create formulas (headcount x productivity) the corresponds
    #     to the dimensions for that date.
        
    #     Args:
    #         None
    #     Returns:
    #         None
    #     """

    #     calcMatrix = []

    #     # self.wsOutput.range((17,8),(21,8)).color = (11, 48, 64)
    #     # self.wsOutput.range((17,8),(21,8)).font.color = (255, 255, 255)
    #     # self.wsOutput.range((17,8)).value = 'Create'
    #     # self.wsOutput.range((18,8)).value = 'Source Split'
    #     # self.wsOutput.range((19,8)).value = 'Deal Type'
    #     # self.wsOutput.range((20,8)).value = 'Product Split'
    #     # self.wsOutput.range((21,8)).value = 'Calculated'
        
    #     for x in range(0, len(refMatrix)):
    #         calcArray = []
    #         # self.wsOutput.range((16,9+x)).color = (221, 221, 221)
    #         # self.wsOutput.range((21,9+x)).color = (235, 235, 235)
    #         for y in range(0, 4):
    #             currentCell = self.wsOutput.range((17+y,9+x))
    #             # if y == 0:
    #             #     currentCell.value = '=(' + refMatrix[x][0] + '*' + refMatrix[x][1] + ')'
    #             #     currentCell.number_format = '#,##0'
    #             # elif y == 3:
    #             #     currentCell.value = '=' + prodSplitConst
    #             #     currentCell.number_format = '0.0%'
    #             # else:
    #             #     currentCell.value = '=' + refMatrix[x][y+2] 
    #             #     currentCell.number_format = '0.0%'
                
    #             calcArray.append(currentCell.get_address(False, False, include_sheetname=False))
            
    #         calcFormula = ''
    #         for z in range(0, len(calcArray)): 
    #             if z == 0:
    #                 calcFormula = calcArray[z]
    #             else:
    #                 calcFormula += '*' + calcArray[z]
                
    #         calculatedCell = self.wsOutput.range((21,9+x))
    #         calculatedCell.value = '=(' + calcFormula + ')'
    #         calculatedCell.number_format = '#,##0'
    #         calcMatrix.append(calculatedCell.get_address(False, False, include_sheetname=False))
        
    #     return calcMatrix


# getAssumptionsDF()
# prodSplitConst = GetProductSplitConstant(aRegion.region[2])
# salesCycleConsts = GetSalesCycleConstants(aRegion.region[5])

# cycleRegions()
# refMatrix = runThroughDates()
# calcMatrix = BuildPipelineCreate(refMatrix, prodSplitConst)
# catalogue = buildCellCatalogue(calcMatrix, salesCycleConsts, prodSplitConst)
# builFormulas(catalogue)
# BuildWaterfall()

# pipeline = PipelineWaterfall()
# pipeline.buildWaterFall('VBA_and_Python_Interview_Exercise_v1.xlsx','uploads')