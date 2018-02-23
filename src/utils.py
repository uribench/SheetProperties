# utils.py
# LGPL license; Copyright (C) 2018 Uri Benchetrit

import string

class Utils:
    """Utilities"""

    @staticmethod
    def isEmpty(anyStruct):
        """Checks if a built-in structure is empty"""
        return False if anyStruct else True

    @staticmethod
    def colNumberToColName(colNumber):
        """Converts a 1-based column number to Excel style column name (i.e., 1 to 'A' up to 702 to 'ZZ')"""
        alphabetList = string.ascii_uppercase
        alphabetListLen = len(alphabetList)
        d, m = divmod(colNumber-1, alphabetListLen)
        return alphabetList[d-1] + alphabetList[m] if (colNumber>alphabetListLen) else alphabetList[colNumber-1]
