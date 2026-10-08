"""Correct openpyxl calibration solution; never supplied to measured agents."""
import sys
from pathlib import Path
from suite import TASKS_BY_ID, solve

if __name__=='__main__':
    solve(TASKS_BY_ID[sys.argv[1]],Path('book.xlsx'))
    print('EVAL_STATUS: SUCCESS')
