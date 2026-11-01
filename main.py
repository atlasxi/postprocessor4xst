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
import pandas as pd

class Loglevel(Enum):
    interupts = 0
    error = 1
    warning = 2
log_info: list[dict[Loglevel, str]] = [] 

def recursive_find_all_txt(excluded_dirs:list[str])->list[str]:
    txt_file_list = []
    root = Path("./")
    for path in root.rglob("*.txt"):
        if any(part in excluded_dirs for part in path.parts):
            continue
        txt_file_list.append(path)
    return txt_file_list

def check_txt_format(file_list:list[str])->list[dict[Loglevel, str]]:
    # warning or errors in log file
    # TODO: remove all file with fatal error
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

def convert_txt_into_list(file_list:list[str]):
    data:dict[str,list] = {}
    for txt in file_list:
        with open(txt, "r",encoding="utf-8") as opened_txt:
            # read the txt into a dict
            for line in opened_txt:
                line = line.strip("\r\n")
                if not line:
                    continue
                for unit in line.split(","):
                    unit = unit.strip()
                    if not unit or ":" not in unit:
                        continue
                    key, value = unit.split(":", 1)
                    data.setdefault(key.strip(), []).append(value.strip())
    return data

def is_valid_data(data: dict[str,list])->bool:
    data_len = None
    for _,list in data.items():
        if data_len == None:
            data_len = len(list)
        elif data_len != len(list):
            # TODO: log msg to add
            log_info.append({Loglevel.interupts: "data row is not aligned"})
            return False
    return True

def sort_all_pds_data(pd_data: pd.DataFrame, chunk_size: int, s_interval_sort_col: str)->pd.DataFrame:
    # TODO: sort not work well in sorting. consider one TorqDmd with more data. in a spec speed get more than one data.
    pd_data = pd_data.sort_values(by=["WT_Speed"], ascending=[True])
    sorted_data = []
    for start in range(0, len(pd_data),chunk_size):
        block = pd_data.iloc[start:start + chunk_size]
        block_sorted = block.sort_values(by=s_interval_sort_col)
        sorted_data.append(block_sorted)
    result = pd.concat(sorted_data).reset_index(drop=True)
    pd_data.loc[:, :] = result.values
    return result

def strip_repeated_pds_data(pd_data: pd.DataFrame)->pd.DataFrame:
    pass

def fuzzy_comp(a:float,b:float,delta:float)->bool:
    if(abs(a-b)<delta):
        return True
    return False
    


def main():
    excluded_dirs = [".git", "__pycache__", "venv", ".venv", "output"]
    file_list = recursive_find_all_txt(excluded_dirs)
    data = convert_txt_into_list(file_list)
    if is_valid_data(data) == False:
        return False
    for key,value in data.items():
        print(f"{key}->{len(value)}")
    pd_data = pd.DataFrame(data)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_colwidth', None)
    pd.set_option('display.width', 1000)
    sort_all_pds_data(pd_data, 41, "TorqDmd")
    print(pd_data.drop([
    # "Num",
    "BusVol",
    "BusCurrent",
    "PhaseCurrent",
    "PhaseVol",
    # "MotorSpeed",
    # "TorqDmd",
    "TorqFedBck",
    "IdDmd",
    "IdBck",
    "IqDmd",
    "IqBck",
    "VdDmd",
    "VqDmd",
    "SpeedSet",
    "BusVolSet",
    "WT_BusVol",
    "WT_BusCurrent",
    "WT_PhaseVol",
    "WT_PhaseCurrent",
    # "WT_Torque",
    # "WT_Speed",
    "WT_MotorEfficiency",
    "WT_ControllerEfficiency",
    "WT_SystemEfficiency",
    "WT_AcPower",
    "WT_DcPower",
    "WT_MechanicalPower"
    ], axis=1))

if __name__ == "__main__":
    main()
