# activeDocumentSheets.py
# LGPL license; Copyright (C) 2018 Uri Benchetrit

import FreeCAD
from itertools import product
import re
from utils import Utils

class RequestParameters:
    """
    Holds the parameters for a request of an action request on the properties 
    of selected cells in the associated spreadsheet.

    Each spreadsheet in the document is associated with an instance of RequestParameters.

    Actual request parameters can be set by using a UI Form (like in this Macro with 
    SheetPropertiesActionsForm), or by using other means.

    Attributes:
        targetSpreadsheet           -- target spreadsheet for actions (e.g., set, clear) on cells properties
        context                     -- context of this script
        initData()                  -- data initializing method for this instance
        hasValidHeaders             -- True if the headers of the target spreadsheet are valid, False otherwise
        invalidHeadersReason        -- reason for invalid headers
        hasValidPropertiesData      -- True if the properties data of the target spreadsheet are valid, False otherwise
        invalidPropertiesDataReason -- reason for invalid properties data
        headersRowNumber            -- row number of the headers
        mandatoryHeaders            -- list of mandatory headers
        headersToLocMap             -- dictionary of {header name : header location} pairs
        headersToColumnMap          -- dictionary of {header name : header column} pairs containing only headers that were found
        dataRowsRanges              -- list of continuous continuous ranges of rows having source data
    """

    MAX_SEARCH_COL = 100    # Max value = 27*26=702='ZZ'
    MAX_SEARCH_ROW = 100    # Max value = 128^2=16,384
    END_DATA_HINT = 5       # Min number of consecutive empty lines 
                            # indicating end row of properties source data

    def __init__(self, sheet, context):
        self.targetSpreadsheet = sheet
        self.context = context
        self.initData()

    def initData(self):
        self.hasValidHeaders = False
        self.invalidHeadersReason = ''
        self.hasValidPropertiesData = False
        self.invalidPropertiesDataReason = ''

        self.headersRowNumber = None
        self.mandatoryHeaders = [self.context.HEADER_ALIAS, self.context.HEADER_VALUE]
        self.headersToLocMap = {self.context.HEADER_UNITS:'', self.context.HEADER_ALIAS:'', self.context.HEADER_VALUE:''}
        self.headersToColumnMap = {}
        self.dataRowsRanges = []    # list of continuous continuous ranges of rows having source data

        # search for the headers in the associated sheet
        self.findSheetHeaders()

        if self.hasValidHeaders:
            self.initHeadersToColumnMap()
            self.dataRowsRanges = self.findDataRowsRanges()

            if Utils.isEmpty(self.dataRowsRanges):
                self.hasValidPropertiesData = False
                self.invalidPropertiesDataReason = 'No usable data rows for property setting were provided in \'{0}\' sheet\n'.  \
                                                    format(self.targetSpreadsheet.Label)
            else:
                self.hasValidPropertiesData = True

    def findSheetHeaders(self):
        """"""
        result = False

        searchedHeadersCount = len(self.headersToLocMap)
        uniqueHeadersFound = 0

        for row, col in product(range(1, self.MAX_SEARCH_ROW), range(1, self.MAX_SEARCH_COL)):
                cellLoc = Utils.colNumberToColName(col) + str(row)
                cellContent = self.targetSpreadsheet.getContents(cellLoc)
                # check if the current cell content matches any of the possible headers
                for header in self.headersToLocMap.keys():
                    if cellContent.lower() == header.lower():
                        # make sure the header that was found is not a duplicate
                        if self.headersToLocMap[header] == '':
                            self.headersToLocMap[header] = cellLoc                          
                            # set headers row number only when the first header is found
                            if uniqueHeadersFound == 0:
                                self.headersRowNumber = row
                            uniqueHeadersFound += 1
                            result = True
                        else:
                            self.hasValidHeaders = False
                            self.invalidHeadersReason = 'Found a duplicated header: ' + header
                            return False

                # stop searching if all the possible headers were found
                if uniqueHeadersFound == searchedHeadersCount:
                    break

        # validate search results with rules additional to those applied during the search for the headers in the spreadsheet 
        if not self.validateHeaders():
            result = False

        return result

    def initHeadersToColumnMap(self):
        """compose a dictionary of {header name : header column} pairs containing only headers that were found"""

        for header in self.headersToLocMap:
            headerLoc = self.headersToLocMap[header]
            if headerLoc != '':
                # extract the header column name from the header location string (e.g., 'AB' from 'AB27')
                col = re.findall('^[A-Z]+', headerLoc)[0]
                self.headersToColumnMap.update({header: col})

    def findDataRowsRanges(self):
        """
        Returns a list of continuous usable data rows ranges

        Returns:
            :return (list): List of dictionaries {'From': None, 'To': None} for each continuous data rows range, 
                            or empty list [] if none has been found
        """

        dataRowsRanges = []                         # list of continuous continuous ranges of rows having source data
        rangeFrom = self.headersRowNumber + 1       # initial guess. empty rows after headers row are possible
        rangeTo = self.MAX_SEARCH_ROW

        done = False
        while not done:
            nextDataRowsRange = {'From': None, 'To': None}

            # find the start of the next range of data rows
            leadingNoneDataRowsCount = self.countLeadingNoneDataRows(rangeFrom, rangeTo)
            if leadingNoneDataRowsCount is None:
                # EOF reached and all rows between rangeFrom and rangeTo are none data rows
                nextDataRowsRange = None
                break
            else:
                # zero or more leading none data rows were found
                rangeFrom += leadingNoneDataRowsCount
                nextDataRowsRange['From'] = rangeFrom

            # find the end of the next range of data rows
            firstNoneDataRow = self.findFirstNoneDataRow(rangeFrom, rangeTo)
            if firstNoneDataRow is None:
                # EOF reached and all rows between nextRangeFrom and rangeTo are data rows 
                # (i.e., no none data rows were found in the given range)
                nextDataRowsRange['To'] = rangeTo
                done = True     # don't break before finalizing this iteration
            else:
                nextDataRowsRange['To'] = firstNoneDataRow - 1

            # finalize the current iteration
            dataRowsRanges.append(nextDataRowsRange)
            rangeFrom = firstNoneDataRow

        return dataRowsRanges

    def countLeadingNoneDataRows(self, fromRow, toRow):
        """counts the leading continuous empty rows in the given range"""
        
        noneDataRowCount = 0

        for row in range(fromRow, toRow):
            if self.isValidDataRow(row):
                # end of none data rows block found
                break

            # a none data row found
            noneDataRowCount += 1

            if noneDataRowCount == self.END_DATA_HINT:
                # min consecutive empty rows reached (i.e., a hint for end of properties source data)
                return None

        return noneDataRowCount

    def findFirstNoneDataRow(self, fromRow, toRow):
        """Returns the row number of the first row in the given range that has no data for property setting"""

        for row in range(fromRow, toRow):
            if not self.isValidDataRow(row):
                return row
 
        return None     # end of range reached with no empty data rows

    def isValidDataRow(self, row):
        """
        Validates a property data source row

        Notes:
            - a row having at least one valid content as data source for property setting is considered a valid row.
            - the column under the HEADER_VALUE header is ignored in the following process (i.e., it may or may not have values). 

        Args:
            :param row (int): Number of the row to be validated in the target spreadsheet

        Returns:
            :return (bool): True if valid, False otherwise.
        """

        result = False

        for header in self.headersToColumnMap:
            # iterate only over the property data sources columns headers (i.e., skip the value header)
            if header == self.context.HEADER_VALUE:
                continue

            dataCellLocation = self.headersToColumnMap[header] + str(row)
            cellContent = self.targetSpreadsheet.getContents(dataCellLocation)
            if cellContent == '':
                continue

            # prepare the validation function associated with the given header
            validationFunc, settingFunc = self.getPropertiesValidationAndSettingFunctions(header)

            # the property data cell has a value, if the value is valid consider the entire row as valid
            if validationFunc(cellContent):
                result = True
                break   # the row has relevant data. no need to iterate any further inside the inspected row.
            else:   
                # invalid value for a property data cell. keep trying other data columns of the inspected row
                pass

        return result

    def getPropertiesValidationAndSettingFunctions(self, header):
        """
        Returns header dependent Validation and Setting functions for properties

        Args:
            :param header (header_type_constant): The header of the column to which the provided property data source cell content belongs.
                                        The validation depends on the type of the header.
            :type header_type_constant: The property data source column header type constant defined by ActiveDocumentSheets 
                                         (e.g., HEADER_UNITS, HEADER_ALIAS)
        Returns:
            :return (tuple): References to Validation and Setting functions.
        """
        settingFuncName, validationFuncName = self.context.headerToFunctionsMap[header]
        settingFunc = getattr(self.targetSpreadsheet, settingFuncName)
        validationFunc = getattr(self, validationFuncName)

        return validationFunc, settingFunc

    def validateHeaders(self):
        """validates the headers as stored in the provided request parameters"""

        result = True
        oneFailedRuleReasons = []
        allFailedRulesReasons = []
        oneFailedRuleSeparator = ', '
        allFailedRulesSeparator = '\n\t'

        # Rule #1: all mandatory headers must exist
        for header in self.mandatoryHeaders:
            if self.headersToLocMap[header] == '':
                oneFailedRuleReasons.append(header)
                # look for all missing headers, don't stop on the first missing one.
                result = False

        if not result:
            # compose the reason string
            if len(oneFailedRuleReasons) == 1:
                reasonPrefix = 'Mandatory header is missing: '
            else:
                reasonPrefix = 'Mandatory headers are missing: '
            allFailedRulesReasons.append(reasonPrefix + oneFailedRuleSeparator.join(oneFailedRuleReasons))

        # Rule #2: headers row number has to be set
        if self.headersRowNumber is None:
            allFailedRulesReasons.append('Headers row number is missing')
            result = False

        # Rule #3: all headers must be on the same row
        for headerLoc in self.headersToLocMap.values():
            if headerLoc == '':
                continue
            # extract the header row number from the header location string (e.g., 27 from 'AB27')
            row = int(re.findall('[\d]+', headerLoc)[0])
            if row != self.headersRowNumber:
                allFailedRulesReasons.append('Not all the headers are on the same row')
                result = False

        # mark the hasValidHeaders flag appropriately and update the reason
        self.hasValidHeaders = result
        if not Utils.isEmpty(allFailedRulesReasons):
            self.invalidHeadersReason = allFailedRulesSeparator.join(allFailedRulesReasons)

        return result

    def validateUnits(self, units):
        try:
            FreeCAD.Units.parseQuantity(units)
            return True
        except IOError:
            return False

    def validateAlias(self, alias):
        # REVISIT: implement a true validation. 
        # temporarily we just count the number of words and accept only a single word as a valid alias.
        tokens = alias.split()
        return len(tokens) == 1

