# activeDocumentSheets.py
# LGPL license; Copyright (C) 2018 Uri Benchetrit

import FreeCAD as App
import FreeCADGui
import Spreadsheet
from .utils import Utils
from .requestParameters import RequestParameters
from .preconditionError import PreconditionError

class ActiveDocumentSheets:
    """
    Context of this script. Common to all sheets.

    Attributes:
        HEADER_UNITS                -- constant string defining the expected string for this header
        HEADER_ALIAS                -- constant string defining the expected string for this header
        HEADER_VALUE                -- constant string defining the expected string for this header
        sheetToRequestParamsMap     -- maps spreadsheet reference to request params reference
        sheetLabelToSheetMap        -- maps spreadsheet label to spreadsheet reference
        headerToFunctionsMap        -- maps headers to respective setting and validation functions
        getSheets()                 -- returns all the spreadsheet included in the active document
        getSelectedSheet()          -- returns the spreadsheet found in the active document
    """

    # Constants
    HEADER_UNITS = 'Units'
    HEADER_ALIAS = 'Alias'
    HEADER_VALUE = 'Value'

    # Useful maps
    sheetToRequestParamsMap = {}
    sheetLabelToSheetMap = {}
    # for each cell property data header associate a tuple in the format of:
    #   Header name: (setting method name, validation method name)
    # it is assumed that the setting methods belong to a 'Spreadsheet::Sheet' object,
    # and the validation method belong to a 'RequestParameters' object.
    headerToFunctionsMap = {HEADER_UNITS: ('setDisplayUnit', 'validateUnits'),
                            HEADER_ALIAS: ('setAlias', 'validateAlias')}

    def __init__(self):
        # Check preconditions
        if App.ActiveDocument is None:
            raise PreconditionError('There is no active document')

        self.activeDocument = App.ActiveDocument

        sheets = self.getSheets()
        if Utils.isEmpty(sheets):
            raise PreconditionError('No spreadsheets were found in the active document')

        # Initialize useful maps
        for sheet in sheets:
            # Associate an instance of RequestParameters to each known sheet.
            # This way, the request parameters can be cached for each sheet.
            self.sheetToRequestParamsMap.update({sheet: RequestParameters(sheet, self)})
            self.sheetLabelToSheetMap.update({sheet.Label: sheet})

    def getSheets(self):
        """Returns the spreadsheet found in the active document"""

        return App.ActiveDocument.findObjects('Spreadsheet::Sheet')

    def getSelectedSheet(self):
        """Returns the first selected spreadsheet in the active document tree view"""

        result = None

        sel = FreeCADGui.Selection.getSelection()
        # Note: getSelection() returns an empty list if the selected object
        #       does not belong to the active document
        if (not Utils.isEmpty(sel)) and sel[0].isDerivedFrom('Spreadsheet::Sheet'):
            result = sel[0]

        return result
