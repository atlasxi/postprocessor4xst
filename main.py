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

"""
# Data type below:

Num:2,
BusVol:752.8,
BusCurrent:3.7,
PhaseCurrent:43.2,
PhaseVol:30.8,
MotorSpeed:498,
TorqDmd:711,
TorqFedBck:711,
IdDmd:-617,
IdBck:-770,
IqDmd:954,
IqBck:916,
VdDmd:-97,
VqDmd:246,
SpeedSet:0,
BusVolSet:0,
WT_BusVol:750.08,
WT_BusCurrent:4.47,
WT_PhaseVol:33.66,
WT_PhaseCurrent:60.85,
WT_Torque:54.01,
WT_Speed:500.06,
WT_MotorEfficiency:92.73,
WT_ControllerEfficiency:91.69,
WT_SystemEfficiency:85.02,
WT_AcPower:3.05,
WT_DcPower:3.33,
WT_MechanicalPower:2.83
"""

import re
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from enum import Enum

class Loglevel(Enum):
    interupts = 0
    error = 1
    warning = 2


def recursive_find_all_txt(excluded_dirs:list[str])->list[str]:
    txt_file_list = []
    root = Path("./")
    for path in root.rglob("*.txt"):
        if any(part in excluded_dirs for part in path.parts):
            continue
        txt_file_list.append(path)
    return txt_file_list

def is_valid_format(file_list:list[str])->list[dict[Loglevel, str]]:
    # warning or errors in log file
    log_info: list[dict[Loglevel, str]] = [] 
    key_list = None
    row_dict = {}
    for file in file_list:
        with open(file, 'r', encoding="utf-8") as buffer:
            for _line_num, value in enumerate(buffer,1):
                # check key-value in var->value
                line_str = value.strip()
                # print(_line_num)
                # print(value)
                if not line_str:
                    msg = {Loglevel.warning: f"{file} Line {_line_num}" }
                    log_info.append(msg)
                    continue
                units = line_str.split(",")
                for unit in units:
                    if not unit:
                        log_info.append({Loglevel.error: f"{file} Line{_line_num} pair failed"})
                        continue
                    if len(unit.split(":")) !=2:
                        log_info.append({Loglevel.error: f"{file} Line{_line_num} pair failed"})
                        continue
                    pair = unit.split(":")
                    try:
                        float(pair[1])
                    except ValueError:
                        log_info.append({Loglevel.error: f"value({pair[1]}) not valid"})
                    row_dict[pair[0]] = pair[1]
                current_keys = set(row_dict.keys())
                if key_list is None:
                    key_list = set(row_dict.keys())
                elif set(row_dict.keys())!=key_list:
                        missing = key_list - current_keys
                        extra   = current_keys - key_list 
                        log_info.append({ Loglevel.error: f"{file}:{_line_num} row key not reflect " f"(missing={sorted(missing)}, extra={sorted(extra)})" })
                        key_list.clear()
                        key_list = current_keys
    return log_info

def fuzzy_comp(a:float,b:float,delta:float)->bool:
    if(abs(a-b)<delta):
        return True
    return False
    


def main():
    excluded_dirs = [".git", "__pycache__", "venv", ".venv", "output"]
    file_list = recursive_find_all_txt(excluded_dirs)
    logs = is_valid_format(file_list)
    for idx, log_entry in enumerate(logs, 1):
        for level, message in log_entry.items():
            if level == Loglevel.warning:
                prefix = "[WARNING]"
            elif level == Loglevel.error:
                prefix = "[ERROR]  "
            else:
                prefix = "[INTERUPT]  "
                
            print(f"{idx:3d}. {prefix} {message}")

if __name__ == "__main__":
    main()
