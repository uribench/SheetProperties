# mainFormUI.py
# LGPL license; Copyright (C) 2018 Uri Benchetrit

from PySide import QtGui, QtCore

class MainFormUI(QtGui.QMainWindow):
    """
    UI definitions of the central widget for the main form

    Attributes:
        selectSheetComboBox         -- QtGui.QComboBox to be populated and connected to signal handlers by the consumer
        rangeFromSpinBox            -- QtGui.QSpinBox to be connected to signal handlers by the consumer
        rangeToSpinBox              -- QtGui.QSpinBox to be connected to signal handlers by the consumer
        setPropertiesPushButton     -- QtGui.QPushButton to be connected to signal handlers by the consumer
        clearPropertiesPushButton   -- QtGui.QPushButton to be connected to signal handlers by the consumer
        statusTextContent           -- QtGui.QTextEdit to be populated by the consumer
        statusRefreshPushButton     -- QtGui.QPushButton to be connected to signal handlers by the consumer
        dismissPushButton           -- QtGui.QPushButton to be connected to signal handlers by the consumer
    """

    def __init__(self):
        super(MainFormUI, self).__init__()
        self.initUI()

    def initUI(self):
        # define window
        self.setGeometry(250, 250, 400, 515)    # xLoc,yLoc,width,height
        self.setWindowTitle("Sheet Properties Actions")
        self.setWindowFlags(QtCore.Qt.MSWindowsFixedSizeDialogHint | QtCore.Qt.WindowStaysOnTopHint)
        font = QtGui.QFont()
        font.setStyleStrategy(QtGui.QFont.PreferAntialias)
        font.setPixelSize(12)
        font.setFamily('Verdana')
        self.setFont(font)

        # target spread sheet selection group
        targetSheetGroupBox = QtGui.QGroupBox('Target Spreadsheet:', self)
        targetSheetGroupBox.setGeometry(10, 10, 380, 50)    # xLoc,yLoc,width,height
        # pop-up menu for selecting the target spreadsheet (will be populated by the consuming object)
        self.selectSheetComboBox = QtGui.QComboBox(self)
        self.selectSheetComboBox.setMinimumSize(300,23)        # width,height
        self.selectSheetComboBox.setSizeAdjustPolicy(QtGui.QComboBox.AdjustToContents)
        targetSheetGroupBoxLayout = QtGui.QVBoxLayout()
        targetSheetGroupBoxLayout.addWidget(self.selectSheetComboBox)
        targetSheetGroupBox.setLayout(targetSheetGroupBoxLayout)

        # target rows range group
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
        self.AutoTargetRowsRangeTextContent.setTabStopWidth(20)      # Tab stop width in pixels (default: 80 pixels)
        # combined layout
        targetRowsRangeLayout = QtGui.QVBoxLayout()
        targetRowsRangeLayout.addLayout(modeAndCustomRangeLayout)
        targetRowsRangeLayout.addWidget(self.AutoTargetRowsRangeTextContent)
        targetRowsRangeGroupBox.setLayout(targetRowsRangeLayout)

        # actions selection group
        actionsGroupBox = QtGui.QGroupBox('Actions:', self)
        actionsGroupBox.setGeometry(10, 210, 380, 50)   # xLoc,yLoc,width,height
        self.setPropertiesPushButton = QtGui.QPushButton('&Set', self)
        self.setPropertiesPushButton.setMinimumSize(81,23)        # width,height
        self.setPropertiesPushButton.setSizePolicy(QtGui.QSizePolicy.Fixed, QtGui.QSizePolicy.Fixed)
        self.clearPropertiesPushButton = QtGui.QPushButton('&Clear', self)
        self.clearPropertiesPushButton.setMinimumSize(81,23)      # width,height
        self.clearPropertiesPushButton.setSizePolicy(QtGui.QSizePolicy.Fixed, QtGui.QSizePolicy.Fixed)
        actionsGroupBoxLayout = QtGui.QHBoxLayout()
        actionsGroupBoxLayout.addWidget(self.setPropertiesPushButton)
        actionsGroupBoxLayout.addWidget(self.clearPropertiesPushButton)
        actionsGroupBox.setLayout(actionsGroupBoxLayout)

        # status group
        statusGroupBox = QtGui.QGroupBox('Status:', self)
        statusGroupBox.setGeometry(10, 270, 380, 200)   # xLoc,yLoc,width,height
        self.statusTextContent = QtGui.QTextEdit()
        self.statusTextContent.setReadOnly(True)
        self.statusTextContent.setTabStopWidth(40)      # Tab stop width in pixels (default: 80 pixels)
        self.statusRefreshPushButton = QtGui.QPushButton('&Refresh', self)
        self.statusRefreshPushButton.setMinimumSize(81,23)        # width,height
        self.statusRefreshPushButton.setSizePolicy(QtGui.QSizePolicy.Fixed, QtGui.QSizePolicy.Fixed)
        statusGroupBoxLayout = QtGui.QVBoxLayout()
        statusGroupBoxLayout.addWidget(self.statusTextContent)
        statusGroupBoxLayout.addWidget(self.statusRefreshPushButton)
        statusGroupBox.setLayout(statusGroupBoxLayout)

        # 'dismissPushButton' button
        self.dismissPushButton = QtGui.QPushButton('&Dismiss', self)
        self.dismissPushButton.setMinimumSize(81,23)        # width,height
        self.dismissPushButton.setSizePolicy(QtGui.QSizePolicy.Fixed, QtGui.QSizePolicy.Fixed)
        self.dismissPushButton.move(160, 480)
        #dismissPushButtonLayout = QtGui.QHBoxLayout()
        #dismissPushButtonLayout.addWidget(self.dismissPushButton)

        # arrange the layout of the main form UI with all its elements
        #mainFormUILayout = QtGui.QHBoxLayout(self)
        #mainFormUILayout.addWidget(targetSheetGroupBox)
        #mainFormUILayout.addWidget(targetRowsRangeGroupBox)
        #mainFormUILayout.addWidget(actionsGroupBox)
        #mainFormUILayout.addWidget(statusGroupBox)
        #mainFormUILayout.addLayout(dismissPushButtonLayout)
        #self.setLayout(mainFormUILayout)
