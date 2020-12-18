# mainFormUI.py
# LGPL license; Copyright (C) 2018 Uri Benchetrit

from PySide import QtGui, QtCore

class MainFormUI(QtGui.QMainWindow):
    """
    UI definitions of the central widget for the main form
    """
    def __init__(self):
        super(MainFormUI, self).__init__()
        self.initUI()

    def initUI(self):
        self.defineWindow()
        self.defineTargetSheetBox()
        self.defineTargetRowsRangeBox()
        self.defineActionsBox()
        self.defineStatusBox()
        self.defineDialogDismiss()

    def defineWindow(self):
        """Defines the window of the main dialog"""
        self.setGeometry(250, 250, 400, 515)    # xLoc,yLoc,width,height
        self.setWindowTitle("Sheet Properties Actions")
        self.setWindowFlags(QtCore.Qt.MSWindowsFixedSizeDialogHint | QtCore.Qt.WindowStaysOnTopHint)
        font = QtGui.QFont()
        font.setStyleStrategy(QtGui.QFont.PreferAntialias)
        font.setPixelSize(12)
        font.setFamily('Verdana')
        self.setFont(font)

    def defineTargetSheetBox(self):
        """
        Defines the target spread sheet selection group

        Attributes:
            selectSheetComboBox -- QtGui.QComboBox to be populated and connected
        """
        targetSheetGroupBox = QtGui.QGroupBox('Target Spreadsheet:', self)
        targetSheetGroupBox.setGeometry(10, 10, 380, 50)    # xLoc,yLoc,width,height
        # pop-up menu for selecting the target spreadsheet
        # (will be populated by the consuming object)
        self.selectSheetComboBox = QtGui.QComboBox(self)
        self.selectSheetComboBox.setMinimumSize(300, 23)        # width,height
        self.selectSheetComboBox.setSizeAdjustPolicy(QtGui.QComboBox.AdjustToContents)
        targetSheetGroupBoxLayout = QtGui.QVBoxLayout()
        targetSheetGroupBoxLayout.addWidget(self.selectSheetComboBox)
        targetSheetGroupBox.setLayout(targetSheetGroupBoxLayout)

    def defineTargetRowsRangeBox(self):
        """
        Defines the target rows range group

        Attributes:
            rangeFromSpinBox -- QtGui.QSpinBox to be connected
            rangeToSpinBox   -- QtGui.QSpinBox to be connected
        """
        targetRowsRangeGroupBox = QtGui.QGroupBox('Target Rows Range:', self)
        targetRowsRangeGroupBox.setGeometry(10, 70, 380, 130)    # xLoc,yLoc,width,height

        # mode (custom/auto) radio buttons
        self.AutoTargetRowsRangeRadioButton = QtGui.QRadioButton('Auto')
        self.CustomTargetRowsRangeRadioButton = QtGui.QRadioButton('Custom')
        self.targetRowsRangeRadioButtonsGroup = QtGui.QButtonGroup()
        self.targetRowsRangeRadioButtonsGroup.addButton(self.AutoTargetRowsRangeRadioButton)
        self.targetRowsRangeRadioButtonsGroup.addButton(self.CustomTargetRowsRangeRadioButton)
        targetRowsRangeRadioButtonsLayout = QtGui.QHBoxLayout()
        targetRowsRangeRadioButtonsLayout.addWidget(self.AutoTargetRowsRangeRadioButton)
        targetRowsRangeRadioButtonsLayout.addWidget(self.CustomTargetRowsRangeRadioButton)

        # custom 'from' row
        rangeFromRowLabel = QtGui.QLabel("From:", self)
        self.rangeFromRowSpinBox = QtGui.QSpinBox()
        rangeFromRowLayout = QtGui.QHBoxLayout()
        rangeFromRowLayout.addStretch(1)    # align to the right
        rangeFromRowLayout.addWidget(rangeFromRowLabel)
        rangeFromRowLayout.addWidget(self.rangeFromRowSpinBox)

        # custom 'to' row
        rangeToRowLabel = QtGui.QLabel("To:", self)
        self.rangeToRowSpinBox = QtGui.QSpinBox()
        rangeToRowLayout = QtGui.QHBoxLayout()
        rangeToRowLayout.addWidget(rangeToRowLabel)
        rangeToRowLayout.addWidget(self.rangeToRowSpinBox)

        # custom 'from' row and 'to' row combined layout
        customTargetRowsRangeLayout = QtGui.QHBoxLayout()
        customTargetRowsRangeLayout.addLayout(rangeFromRowLayout)
        customTargetRowsRangeLayout.addLayout(rangeToRowLayout)

        # mode radio buttons and custom range layout
        modeAndCustomRangeLayout = QtGui.QHBoxLayout()
        modeAndCustomRangeLayout.addLayout(targetRowsRangeRadioButtonsLayout)
        modeAndCustomRangeLayout.addLayout(customTargetRowsRangeLayout)

        # auto discovered rows range display area
        self.AutoTargetRowsRangeTextContent = QtGui.QTextEdit()
        self.AutoTargetRowsRangeTextContent.setReadOnly(True)
        # Tab stop width in pixels (default: 80 pixels)
        self.AutoTargetRowsRangeTextContent.setTabStopWidth(20)

        # combined layout
        targetRowsRangeLayout = QtGui.QVBoxLayout()
        targetRowsRangeLayout.addLayout(modeAndCustomRangeLayout)
        targetRowsRangeLayout.addWidget(self.AutoTargetRowsRangeTextContent)
        targetRowsRangeGroupBox.setLayout(targetRowsRangeLayout)

    def defineActionsBox(self):
        """
        Defines the actions selection group

        Attributes:
            setPropertiesPushButton   -- QtGui.QPushButton to be connected
            clearPropertiesPushButton -- QtGui.QPushButton to be connected
        """
        actionsGroupBox = QtGui.QGroupBox('Actions:', self)
        actionsGroupBox.setGeometry(10, 210, 380, 50)   # xLoc,yLoc,width,height
        self.setPropertiesPushButton = QtGui.QPushButton('&Set', self)
        self.setPropertiesPushButton.setMinimumSize(81, 23)        # width,height
        self.setPropertiesPushButton.setSizePolicy(QtGui.QSizePolicy.Fixed,
                                                   QtGui.QSizePolicy.Fixed)
        self.clearPropertiesPushButton = QtGui.QPushButton('&Clear', self)
        self.clearPropertiesPushButton.setMinimumSize(81, 23)      # width,height
        self.clearPropertiesPushButton.setSizePolicy(QtGui.QSizePolicy.Fixed,
                                                     QtGui.QSizePolicy.Fixed)
        actionsGroupBoxLayout = QtGui.QHBoxLayout()
        actionsGroupBoxLayout.addWidget(self.setPropertiesPushButton)
        actionsGroupBoxLayout.addWidget(self.clearPropertiesPushButton)
        actionsGroupBox.setLayout(actionsGroupBoxLayout)

    def defineStatusBox(self):
        """
        Defines the status group

        Attributes:
            statusTextContent       -- QtGui.QTextEdit to be populated
            statusRefreshPushButton -- QtGui.QPushButton to be connected
        """
        statusGroupBox = QtGui.QGroupBox('Status:', self)
        statusGroupBox.setGeometry(10, 270, 380, 200)   # xLoc,yLoc,width,height
        self.statusTextContent = QtGui.QTextEdit()
        self.statusTextContent.setReadOnly(True)
        # Tab stop width in pixels (default: 80 pixels)
        self.statusTextContent.setTabStopWidth(40)
        self.statusRefreshPushButton = QtGui.QPushButton('&Refresh', self)
        self.statusRefreshPushButton.setMinimumSize(81, 23)        # width,height
        self.statusRefreshPushButton.setSizePolicy(QtGui.QSizePolicy.Fixed, QtGui.QSizePolicy.Fixed)
        statusGroupBoxLayout = QtGui.QVBoxLayout()
        statusGroupBoxLayout.addWidget(self.statusTextContent)
        statusGroupBoxLayout.addWidget(self.statusRefreshPushButton)
        statusGroupBox.setLayout(statusGroupBoxLayout)

    def defineDialogDismiss(self):
        """
        Defines the dialog dismiss widget

        Attributes:
            dismissPushButton -- QtGui.QPushButton to be connected
        """
        self.dismissPushButton = QtGui.QPushButton('&Dismiss', self)
        self.dismissPushButton.setMinimumSize(81, 23)        # width,height
        self.dismissPushButton.setSizePolicy(QtGui.QSizePolicy.Fixed, QtGui.QSizePolicy.Fixed)
        self.dismissPushButton.move(160, 480)
