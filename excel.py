from concretizations import ExcelManOpenPyXl
class Excel():
    '''
    NO.1: Excel provide some common method for excel options,
    NO.2: Excel choice appropriate concretization Class base on parameters
    '''

    lines_len = 0
    columns_len = 0
    instance = None

    def __init__(self, filepath, has_title=False):
        suffix = filepath.split('.')[-1]
        if suffix == 'xlsx':
            self.instance = ExcelManOpenPyXl(filepath, has_title=has_title)
