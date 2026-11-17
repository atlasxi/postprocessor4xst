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
import sys
import pandas as pd
import io
from enum import Enum
from pathlib import Path
from enum import Enum

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter


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

def sort_all_pds_data(pd_data: pd.DataFrame, s_interval_sort_col: str, rpm_list: list[int], rpm_delta: float):
    # TODO: sort not work well in sorting. consider one TorqDmd with more data. in a spec speed get more than one data.
    '''
    1. classify by WT_Torq
    2. if TorqDmd > 0, DD, else FD
    # 2. if DC_power > AC_power, DD mode .else FD mode
    '''
    rpm_list = sorted(list(set(rpm_list)))
    print(f"rpm_list:{rpm_list}")
    temp_dict:dict[int, pd.DataFrame] = {}
    pd_data["WT_Speed"] = pd.to_numeric(pd_data["WT_Speed"])
    pd_data["WT_Torque"] = pd.to_numeric(pd_data["WT_Torque"])
    pd_data[s_interval_sort_col] = pd.to_numeric(pd_data[s_interval_sort_col])
    for rpm in rpm_list:
        mask = pd_data["WT_Speed"].between(rpm-rpm_delta, rpm+rpm_delta)
        block = pd_data[mask].copy()
        if block.empty:
            print(f"warning: {rpm}rpm empty")
            continue
        pos_block = block[block[s_interval_sort_col] > 0].sort_values(s_interval_sort_col, key=abs)
        neg_block = block[block[s_interval_sort_col] < 0].sort_values(s_interval_sort_col, key=abs)
        zero_block = block[block[s_interval_sort_col] == 0].head(1)
        dd = pd.concat([zero_block, pos_block]).reset_index(drop=True)
        fd = pd.concat([zero_block, neg_block]).reset_index(drop=True)
        temp_dict[rpm] = pd.concat([zero_block, pos_block, zero_block, neg_block]).reset_index(drop=True)
    current_row = 0
    for rpm, sub_pd_data in temp_dict.items():
        print(f"{rpm}->{len(sub_pd_data)}")
        # if current_row == 0:
        #     current_row = sub_pd_data.shape
        # else:
        #     if current_row != sub_pd_data.shape:
        #         log_info.append({Loglevel.interupts: "sub_pd_data shape not match"})
        #         return;    


    # TODO:  drop keyword will be deprecated in future
    # pd_data = pd.concat([temp_dict[rpm] for rpm in sorted(temp_dict.keys())]).reset_index(drop=True)
    sorted_df = pd.concat([temp_dict[rpm] for rpm in sorted(temp_dict.keys()) if rpm in temp_dict]).reset_index(drop=True)
    pd_data.drop(pd_data.index, inplace=True)
    for col in sorted_df.columns:
        pd_data[col] = sorted_df[col]
    return temp_dict

def strip_repeated_pds_data(data_dict: dict[int, pd.DataFrame],weight_list:dict[str,list[int]], density: list[int])->pd.DataFrame:
    pass

def add_weight_col_2data(data_dict: dict[int, pd.DataFrame], required_col: list[str]=[""])->dict(int,pd.DataFrame): 
    """
    WT_Queue & Queue
    WT_BusVol & BusVol
    WT_BusCurrent & BusCurrent
    WT_PhaseVol & PhaseVol
    WT_PhaseCurrent & PhaseCurrent
    WT_Speed & MotorSpeed
    9550*WT_MechanicalPower/WT_ & WT_Queue
    WT_ControllerEfficiency & WT_AcPower/WT_DcPower
    WT_MotorEfficiency & WT_MotorEfficiency/WT_AcPower
    WT_SystemEfficiency & WT_MechanicalPower/WT_DcPower
    Bound efficiency > 0% & <100%
    TorqDmd & (3/2)p[phi_f*i_q+(Ld-Lq)i_d*i_q] 
    // phi_f->Permanent magnet flux linkage
    """
    return data_dict

def fuzzy_comp(a:float,b:float,delta:float)->bool:
    if(abs(a-b)<delta):
        return True
    return False
    
def display_loginfo(logs: list[dict[Loglevel, str]], stream="stdout"):
    close_after = False
    if stream == "stdout":
        out = sys.stdout
    elif stream == "stderr":
        out = sys.stderr
    elif hasattr(stream, "write") and not getattr(stream, "closed", False):
        if getattr(stream, "writable", lambda: True)() is False:
            raise ValueError(f"{stream!r} not writable")
        out = stream
    else:
        out = open(stream, "w", encoding="utf-8")
        close_after = True
    try:
        for entry in logs:
            for key, msg in entry.items():
                out.write(f"[{key}]:{msg}\n")
    finally:
        if close_after:
            out.close()

def main():
    excluded_dirs = [".git", "__pycache__", "venv", ".venv", "output"]
    file_list = recursive_find_all_txt(excluded_dirs)
    data = convert_txt_into_list(file_list)
    if is_valid_data(data) == False:
        print("data invalid")
        return False
    for key,value in data.items():
        print(f"{key}->{len(value)}")
    pd_data = pd.DataFrame(data)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_colwidth', None)
    pd.set_option('display.width', 1000)
    sort_all_pds_data(pd_data, "WT_Torque",list(range(500, 10001, 500)), 10)
    # if pd_data is None or pd_data.empty:
    #     log_info.append({Loglevel.interupts:"sort_all_pds_data failed"})
    #     return
    # else:
    keep_cols = ["Num", "MotorSpeed", "TorqDmd", "WT_Torque", "WT_Speed"]
    print(pd_data[keep_cols]) 
    display_loginfo(log_info, stream="output")

if __name__ == "__main__":
    main()
