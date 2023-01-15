class OleFileIONotClosed(RuntimeWarning):
    """
    Warning type used when OleFileIO is destructed but has open file handle.
    """
    def __init__(self, stack_of_open=None):
        super(OleFileIONotClosed, self).__init__()
        self.stack_of_open = stack_of_open

    def __str__(self):
        msg = 'Deleting OleFileIO instance with open file handle. ' \
              'You should ensure that OleFileIO is never deleted ' \
              'without calling close() first. Consider using '\
              '"with OleFileIO(...) as ole: ...".'
        if self.stack_of_open:
            return ''.join([msg, '\n', 'Stacktrace of open() call:\n'] +
                           self.stack_of_open.format())
        else:
            return msg
