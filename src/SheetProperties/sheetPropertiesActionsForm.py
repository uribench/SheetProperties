# sheetPropertiesActionsForm.py
# LGPL license; Copyright (C) 2018 Uri Benchetrit

import re
import json
from .mainFormUI import MainFormUI
from .sheetPropertiesActions import SheetPropertiesActions
from .treeViewSelectionObserver import TreeViewSelectionObserver
import FreeCADGui
from PySide import QtCore

class SheetPropertiesActionsForm(MainFormUI):
    """
    A form interactively collecting the request parameters for actions on the
    properties of selected cells in a selected spreadsheet.

    This class provides the functional logic for a concrete UI (the base class for this class)
    that is defined elsewhere.

    When confirmed, the properties are set according to the current request
    parameters.
    """

    # status message types
    STATUS_INFO = 0
    STATUS_ERROR = 1

    def __init__(self, context):
        self.context = context
        self.targetSpreadsheet = None   # selected target spreadsheet
        self.requestParams = None       # request params associated with
                                        # the selected target spreadsheet

        super(SheetPropertiesActionsForm, self).__init__()
        self.initForm()

    def initForm(self):
        """"""
        # do this before self.connectSignalHandlingMethods()
        # see details inside the method implementation
        self.initTargetSheetSelector()

        self.connectSignalHandlingMethods()

        # Install a selection observer
        self.treeViewSelectionObserver = TreeViewSelectionObserver(self)
        FreeCADGui.Selection.addObserver(self.treeViewSelectionObserver)

        # sync selections in the right direction with a particular order (tree view gets priority)
        if self.context.getSelectedSheet() is not None:
            self.syncComboBoxFromTreeViewSelection()
            # the next condition deals with unique initial state in which the
            # first sheet is selected in the tree view, and the combo box is
            # already showing the first sheet as the selected one by default
            # (even without sync of combo box from the tree view selection).
            # in this case, the self.selectSheetComboBox.currentIndexChanged
            # signal is not fired and self.handleTargetSpreadsheetChanged()
            # is not called from that signal handler.
            if self.selectSheetComboBox.currentIndex() == 0:
                self.handleTargetSpreadsheetChanged()
        else:
            self.syncTreeViewFromComboBoxSelection(self.selectSheetComboBox.currentText())
            self.handleTargetSpreadsheetChanged()

        # make the window visible
        self.show()

    def initTargetSheetSelector(self):
        """
        Populates the pop-up menu for selecting the target spreadsheet from
        all the spreadsheets included in the active document.

        Note that if signal handling methods for this ComboBx were installed
        before this point, then if the signals are not blocked,
        adding items to a ComboBox also triggers the currentIndexChanged signals,
        and this may be undesired before completing initialization.
        """
        for sheet in self.context.getSheets():
            self.selectSheetComboBox.addItem(sheet.Label)

    def connectSignalHandlingMethods(self):
        """Connects signal handling methods for widgets of the main dialog"""
        self.selectSheetComboBox.currentIndexChanged.connect(self.onSelectSheetComboBoxCurrentIndexChanged)
        self.selectSheetComboBox.activated[str].connect(self.onSelectSheetComboBoxActivated)
        self.AutoTargetRowsRangeRadioButton.clicked.connect(self.onTargetRowsRangeModeChanged)
        self.CustomTargetRowsRangeRadioButton.clicked.connect(self.onTargetRowsRangeModeChanged)
        self.rangeFromRowSpinBox.valueChanged[str].connect(self.onRangeFromRowSpinBoxValueChanged)
        self.rangeToRowSpinBox.valueChanged[str].connect(self.onRangeToRowSpinBoxValueChanged)
        self.setPropertiesPushButton.clicked.connect(self.onSetProperties)
        self.clearPropertiesPushButton.clicked.connect(self.onClearProperties)
        self.statusRefreshPushButton.clicked.connect(self.onRefreshStatus)
        self.dismissPushButton.clicked.connect(self.onDismiss)


    def syncComboBoxFromTreeViewSelection(self):
        """Sync the selection in the pop-up menu with the selection in the Tree View"""

        # when the first selected item in the active document tree view is a spreadsheet,
        # set it as the default in the combo box (i.e., target spreadsheet selection pop-up menu)
        selectedSheet = self.context.getSelectedSheet()
        comboBoxItems = [self.selectSheetComboBox.itemText(i) for i in range(self.selectSheetComboBox.count())]
        if selectedSheet is not None and selectedSheet.Label in comboBoxItems:
            # self.targetSpreadsheet has to be set before setting a new index for
            # selectSheetComboBox combo box, as the last fires a self.selectSheetComboBox.currentIndexChanged
            # signal and its handler rely on current self.targetSpreadsheet.
            # alternatively, we could block the currentIndexChanged signals until we are done here.
            self.targetSpreadsheet = selectedSheet
            self.selectSheetComboBox.setCurrentIndex(comboBoxItems.index(selectedSheet.Label))

    def syncTreeViewFromComboBoxSelection(self, selectedText):
        """Sync the selection in the Tree View with the selection in the pop-up menu"""

        if selectedText in self.context.sheetLabelToSheetMap:
            self.targetSpreadsheet = self.context.sheetLabelToSheetMap[selectedText]
        else:
            print('syncTreeViewFromComboBoxSelection(): Internal Error: \'{0}\' is unknown'.format(selectedText))
            self.close()

        # update the selection in the tree view to reflect the selection in the ComboBox
        FreeCADGui.Selection.clearSelection()
        FreeCADGui.Selection.addSelection(self.targetSpreadsheet)

    def appendStatus(self, statusMessage, statusMessageType=STATUS_INFO):
        # preserve current text color
        previousTextColor = self.statusTextContent.textColor()

        if statusMessageType == self.STATUS_ERROR:
            self.statusTextContent.setTextColor(QtCore.Qt.red)

        self.statusTextContent.append(str(statusMessage))
        self.statusTextContent.update()

        # restore previous text color
        self.statusTextContent.setTextColor(previousTextColor)

        # show the same text also on the 'Report View' panel of FreeCAD
        print(statusMessage)

    def clearStatus(self):
        self.statusTextContent.clear()
        self.update()

    def enableCustomDataRowsRangeSetting(self, enable):
        self.rangeFromRowSpinBox.setEnabled(enable)
        self.rangeToRowSpinBox.setEnabled(enable)

    def getCustomDataRowsRange(self):
        customDataRowsRange = {'From': None, 'To': None}
        customDataRowsRange['From'] = self.rangeFromRowSpinBox.value()
        customDataRowsRange['To'] = self.rangeToRowSpinBox.value()

        return [customDataRowsRange]

    def displayStatusMessage(self):
        # clear the current content of the status display then print the new content
        self.clearStatus()
        if self.requestParams.hasValidHeaders:
            statusMessage = 'Valid headers found for sheet \'{0}\' at:'.format(self.targetSpreadsheet.Label)
            self.appendStatus(statusMessage)
            statusMessage = '\t' + str(self.requestParams.headersToLocMap) + '\n'
            self.appendStatus(statusMessage)
        else:
            statusMessage = 'Invalid headers for sheet \'{0}\':'.format(self.targetSpreadsheet.Label)
            self.appendStatus(statusMessage, self.STATUS_ERROR)
            statusMessage = '\t' + self.requestParams.invalidHeadersReason
            self.appendStatus(statusMessage, self.STATUS_ERROR)
            statusMessage = '\nSheet \'{0}\' cannot be used'.format(self.targetSpreadsheet.Label)
            self.appendStatus(statusMessage, self.STATUS_ERROR)
            statusMessage = 'Fix the problem or select a valid sheet'
            self.appendStatus(statusMessage, self.STATUS_ERROR)

        if self.requestParams.hasValidPropertiesData:
            statusMessage = 'Valid data rows found for sheet \'{0}\''.format(self.targetSpreadsheet.Label)
            self.appendStatus(statusMessage)
        else:
            statusMessage = self.requestParams.invalidPropertiesDataReason
            self.appendStatus(statusMessage, self.STATUS_ERROR)

    def displayDataRowsRanges(self):
        if self.requestParams.hasValidPropertiesData:
            # a dictionary does not track the insertion order, and iterating over it produces
            # the values in an arbitrary order.
            # sort the ranges that were found and replace the double-quotes with single-quotes
            # to make it easier to the eyes.
            sortedRanges = json.dumps(self.requestParams.dataRowsRanges, sort_keys=True)
            sortedRanges = re.sub('"', '\'', sortedRanges)
            self.AutoTargetRowsRangeTextContent.setPlainText(sortedRanges)
        else:
            self.AutoTargetRowsRangeTextContent.clear()

    def setValueAndMinimumForCustomRowsRange(self):
        # prepare initial and min/max values for the custom range spin boxes
        # assume no valid headers were found, set initial and minimum values
        # for the custom range spin boxes to be 1
        rangeFromRowSpinBoxMinimum = 1
        rangeToRowSpinBoxMinimum = 1
        rangeFromRowSpinBoxMaximum = self.requestParams.MAX_SEARCH_ROW
        rangeToRowSpinBoxMaximum = self.requestParams.MAX_SEARCH_ROW
        rangeFromRowSpinBoxValue = 1
        rangeToRowSpinBoxValue = 1
        if self.requestParams.hasValidHeaders:
            # both rangeFromRowSpinBox and rangeToRowSpinBox must have values higher than the headers row number
            rangeFromRowSpinBoxMinimum = self.requestParams.headersRowNumber + 1
            rangeToRowSpinBoxMinimum = self.requestParams.headersRowNumber + 1
            if self.requestParams.hasValidPropertiesData:
                # bound the data rows ranges that were found (by taking th extreme values)
                rangeFromRowSpinBoxValue = self.requestParams.dataRowsRanges[0]['From']
                rangeToRowSpinBoxValue = self.requestParams.dataRowsRanges[-1]['To']

        # set min/max for the custom range spin boxes
        self.rangeFromRowSpinBox.setMinimum(rangeFromRowSpinBoxMinimum)
        self.rangeToRowSpinBox.setMinimum(rangeToRowSpinBoxMinimum)
        self.rangeFromRowSpinBox.setMaximum(rangeFromRowSpinBoxMaximum)
        self.rangeToRowSpinBox.setMaximum(rangeToRowSpinBoxMaximum)

        # set initial values based on the ranges found automatically
        self.rangeFromRowSpinBox.setValue(rangeFromRowSpinBoxValue)
        self.rangeToRowSpinBox.setValue(rangeToRowSpinBoxValue)

    def setDefaultRowsRangeSettingMode(self):
        if self.requestParams.hasValidHeaders:
            self.AutoTargetRowsRangeRadioButton.setEnabled(True)
            self.CustomTargetRowsRangeRadioButton.setEnabled(True)
            if self.requestParams.hasValidPropertiesData:
                self.AutoTargetRowsRangeRadioButton.setChecked(True)
                self.enableCustomDataRowsRangeSetting(False)
            else:
                self.CustomTargetRowsRangeRadioButton.setChecked(True)
                self.enableCustomDataRowsRangeSetting(True)
        else:
            # the sheet is unusable. nothing is allowed.
            self.enableCustomDataRowsRangeSetting(False)
            self.AutoTargetRowsRangeRadioButton.setEnabled(False)
            self.CustomTargetRowsRangeRadioButton.setEnabled(False)


    def setActionsAvailability(self):

        # enable the 'set' properties action based on the validity of the target sheet
        if self.requestParams.hasValidHeaders and self.requestParams.hasValidPropertiesData:
            self.setPropertiesPushButton.setEnabled(True)
        else:
            self.setPropertiesPushButton.setEnabled(False)

        # enable the 'clear' properties action based on the validity of the target sheet
        if self.requestParams.hasValidHeaders:
            self.clearPropertiesPushButton.setEnabled(True)
        else:
            self.clearPropertiesPushButton.setEnabled(False)

    def handleTargetSpreadsheetChanged(self):
        """
        Called after a new target spreadsheet has been selected using
        either the pop-up menu or the tree view.

        This is the central location for setting the SheetPropertiesActionsForm based on the selected concrete target sheet.
        This includes for instance, updating the displayed status messages and displayed valid data ranges, recalculate the
        default values for the custom target data rows range, and enable and disable widgets.
        """

        # expecting the targetSpreadsheet attribute of SheetPropertiesActionsForm to be current.
        if self.targetSpreadsheet is None:
            print('handleTargetSpreadsheetChanged(): Internal Error: a valid target spreadsheet is expected')
            return

        # for convenience, keep a local copy of the reference to the current requestParams in SheetPropertiesActionsForm
        self.requestParams = self.context.sheetToRequestParamsMap[self.targetSpreadsheet]

        # perform operations based on the validity of the headers and data rows of the selected target sheet
        self.displayStatusMessage()
        self.displayDataRowsRanges()
        self.setValueAndMinimumForCustomRowsRange()
        self.setDefaultRowsRangeSettingMode()
        self.setActionsAvailability()

    def onSelectSheetComboBoxCurrentIndexChanged(self, selectedIndex):
        # sync and handle selected target sheet
        self.syncTreeViewFromComboBoxSelection(self.selectSheetComboBox.itemText(selectedIndex))
        self.handleTargetSpreadsheetChanged()

    def onSelectSheetComboBoxActivated(self, selectedText):
        # ignore this signal. the onSelectSheetComboBoxCurrentIndexChanged() is taking care of selection changes.
        pass

    def onTargetRowsRangeModeChanged(self):
        # enable the custom range settings spin boxes only in custom mode
        if self.targetRowsRangeRadioButtonsGroup.checkedButton().text() == 'Custom':
            self.enableCustomDataRowsRangeSetting(True)
            # we allow unconditionally using the 'Clear' action in this state
            self.clearPropertiesPushButton.setEnabled(True)
        else:
            self.enableCustomDataRowsRangeSetting(False)
            if self.requestParams.hasValidPropertiesData:
                self.clearPropertiesPushButton.setEnabled(True)
            else:
                # we allow selecting 'Auto' in this state, but the 'Clear' action has to be disabled
                self.clearPropertiesPushButton.setEnabled(False)

    def onRangeFromRowSpinBoxValueChanged(self):
        rangeTo = self.rangeToRowSpinBox.value()
        if self.rangeFromRowSpinBox.value() > rangeTo:
            self.rangeFromRowSpinBox.setValue(rangeTo)

    def onRangeToRowSpinBoxValueChanged(self):
        rangeFrom = self.rangeFromRowSpinBox.value()
        if self.rangeToRowSpinBox.value() < rangeFrom:
            self.rangeToRowSpinBox.setValue(rangeFrom)

    def onDismiss(self):
        self.close()

    def onSetProperties(self):

        # expecting valid headers for the selected target sheet
        if not self.requestParams.hasValidPropertiesData:
            print('onSetProperties(): Internal Error. The \'Set\' action button was supposed to be disabled')
            return

        # perform the actual cells properties setting based on the relevant request parameters
        sheetPropertyActions = SheetPropertiesActions(self.requestParams)
        if self.targetRowsRangeRadioButtonsGroup.checkedButton().text() == 'Auto':
            sheetPropertyActions.readAndSetProperties(self.requestParams.dataRowsRanges)
        else:
            sheetPropertyActions.readAndSetProperties(self.getCustomDataRowsRange())

    def onClearProperties(self):

        # expecting valid headers for the selected target sheet
        if not self.requestParams.hasValidHeaders:
            print('onClearProperties(): Internal Error. The \'Clear\' action button was supposed to be disabled')
            return

        # clear the properties of the target cells based on the relevant request parameters
        sheetPropertyActions = SheetPropertiesActions(self.requestParams)
        if self.targetRowsRangeRadioButtonsGroup.checkedButton().text() == 'Auto':
            sheetPropertyActions.clearProperties(self.requestParams.dataRowsRanges)
        else:
            sheetPropertyActions.clearProperties(self.getCustomDataRowsRange())

    def onRefreshStatus(self):
        # REVISIT: Improve implementation to show the updated status of the selected sheet
        print('onRefreshStatus() entered')
        self.clearStatus()

    def onSetSelection(self, doc):
        """Called by the selection observer when a new selection is done in the tree view"""

        if doc == self.context.activeDocument.Name:
            # Sync the selection in the tree view with the pop-up menu
            self.syncComboBoxFromTreeViewSelection()
        else:
            # we could set the parent document of the new selection as the active document as follows:
            #App.setActiveDocument(doc)
            # but, at least for now, we will suppress the signal and leave it to the user to decide
            statusMessage = 'The selected sheet does not belong to the active document \'{0}\''.format(self.context.activeDocument.Name)
            self.appendStatus(statusMessage, self.STATUS_ERROR)
            statusMessage = 'The parent document \'{0}\' has to be activated first'.format(doc)
            self.appendStatus(statusMessage, self.STATUS_ERROR)

    def closeEvent(self, event):
        """
        Called when the user closes the window or when the code calls QWidget.close()
        to close a widget programmatically
        """

        print('Form Closed')

        # -----------------------------------------------------------------------
        # Cleanup
        # -----------------------------------------------------------------------

        # Uninstall the selection observer
        FreeCADGui.Selection.removeObserver(self.treeViewSelectionObserver)

        return self
