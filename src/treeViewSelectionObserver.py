# treeViewSelectionObserver.py
# LGPL license; Copyright (C) 2018 Uri Benchetrit

import FreeCADGui
from utils import Utils

class TreeViewSelectionObserver:
    """Installable Observer for selections in the Tree View"""

    def __init__(self, subscriber):
        self.subscriber = subscriber

    def setSelection(self, doc):
        """called by the installed selection observer when a selection is made in the tree-view (aka, ComboView)"""

        # simple workaround to the unexplained behavior of getting two signals of setSelection when switching selection 
        # between objects that belong to two different parent documents.
        # it relies on the fact that getSelection() returns an empty list if the selected object does not belong to the active document
        sel = FreeCADGui.Selection.getSelection()
        if (not Utils.isEmpty(sel)):
            # true signal. delegate handling to subscriber normally.
            self.subscriber.onSetSelection(doc)
        else:
            # fake signal. signal suppressed.
            pass
