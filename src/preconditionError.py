# preconditionError.py
# LGPL license; Copyright (C) 2018 Uri Benchetrit

class PreconditionError(Exception):
    """Custom Exception raised for precondition violation.

    Attributes:
        reason -- explanation of the error
    """
    def __init__(self, reason):
        self.reason = reason
        super(PreconditionError, self).__init__()
    def __str__(self):
        return repr(self.reason)
