# sheetPropertiesActions.py
# LGPL license; Copyright (C) 2018 Uri Benchetrit

import FreeCAD as App
from .utils import Utils

class SheetPropertiesActions:
    """
    Performs the requested action on the properties of selected cells in a
    selected spreadsheet as indicated by the provided RequestParameters.

    Attributes:
        readAndSetProperties()  -- set the properties of the cells in the column
                                   having HEADER_VALUE header based on the data
                                   of the respective cells in the columns having the
                                   data source headers (e.g., HEADER_UNITS, HEADER_ALIAS)
        clearProperties()       -- set the properties of the cells in the column
                                   having HEADER_VALUE header
    """

    def __init__(self, requestParams):
        self.requestParams = requestParams
        self.sheet = self.requestParams.targetSpreadsheet

    def readAndSetProperties(self, dataRowsRanges):
        """Sets the properties of the value column based on the data source cells"""

        # expecting a valid dataRowsRanges
        if Utils.isEmpty(dataRowsRanges):
            print('readAndSetProperties(): Internal Error: a valid dataRowsRanges is expected')
            return

        for dataRowsRange in dataRowsRanges:
            for row in range(dataRowsRange['From'], dataRowsRange['To'] + 1):
                # the cell location of the target cell for property setting
                # needs to be updated only once per row
                valueCellLocation = \
                    self.requestParams.headersToColumnMap[self.requestParams.context.HEADER_VALUE] \
                    + str(row)
                for header in self.requestParams.headersToColumnMap:
                    # iterate only over the property data headers (i.e., skip the value header)
                    if header == self.requestParams.context.HEADER_VALUE:
                        continue

                    dataCellLocation = self.requestParams.headersToColumnMap[header] + str(row)
                    cellContent = self.sheet.getContents(dataCellLocation)
                    if cellContent == '':
                        continue

                    # prepare the validation function associated with the given header
                    validationFunc, settingFunc = \
                        self.requestParams.getPropertiesValidationAndSettingFunctions(header)

                    # the property data cell has a value, validate and if valid
                    # use it to set the respective property
                    if validationFunc(cellContent):
                        settingFunc(valueCellLocation, cellContent)
                    else:
                        print('Ignoring invalid {0} \'{1}\' found at: {2}' \
                              .format(header, cellContent, dataCellLocation))

        App.ActiveDocument.recompute()

    def clearProperties(self, dataRowsRanges):
        """Clears the properties of the value column for the give range"""

        # expecting a valid dataRowsRanges
        if Utils.isEmpty(dataRowsRanges):
            print('readAndSetProperties(): Internal Error: a valid dataRowsRanges is expected')
            return

        rangeFrom = dataRowsRanges[0]['From']    # get 'From' value from the first tuple in the list
        rangeTo = dataRowsRanges[-1]['To']       # get 'To' value from the last tuple in the list

        for row in range(rangeFrom, rangeTo + 1):
            # the cell location of the target cell for property setting
            # needs to be updated only once per row
            valueCellLocation = \
                self.requestParams.headersToColumnMap[self.requestParams.context.HEADER_VALUE] \
                + str(row)
            for header in self.requestParams.headersToColumnMap:
                # iterate only over the property data headers (i.e., skip the value header)
                if header == self.requestParams.context.HEADER_VALUE:
                    continue

                # prepare and execute the setting function associated with the given header
                validationFunc, settingFunc = \
                    self.requestParams.getPropertiesValidationAndSettingFunctions(header)
                settingFunc(valueCellLocation, '')

        App.ActiveDocument.recompute()
