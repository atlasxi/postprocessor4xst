#!/usr/bin/env -S uv run --script
"""
! how to prcess and classify data
1. sort FD and DD
2. sort speed
3. sort by torque(ascended)
! how to darw some interesting pictures
1. Torque-Speed Characteristic Curve
2. Efficiency Map
3. Id / Iq Tracking
"""
import re
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from enum import Enum

class Loglevel(Enum):
    errors_interupts = 0
    warning_as_errors = 1
    ignore_errors = 2
    ignore_warnings = 3


txt_file_list = []
def recursive_find_all_txt()->list[str]:
    root = Path("./")
    for path in root.rglob("*.txt"):
        txt_file_list.append(path)
    return txt_file_list

log_info: list[dict[Loglevel, str]] = [] 
def is_a_valid_format(txtfile)->list[dict[Loglevel, str]]:
    # warning or errors in log file
    with open(txtfile) as buffer:
        #check here
        pass

    return log_info
    

def main():
    recursive_find_all_txt()

if __name__ == "__main__":
    main()
