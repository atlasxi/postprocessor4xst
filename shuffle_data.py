#!/usr/bin/env -S uv run --script

import os
import pathlib 
import random
import sys

def get_file_list(root: str)->dict[str, int]:
    root_dir = pathlib.Path(root)
    txt_files = root_dir.rglob("*.txt")
    pairs = {}
    for file in txt_files:
        non_empty_contents = [ line for line in file.read_text(encoding="utf-8").splitlines(keepends=True) if line.strip() ]
        pairs[file.absolute()] = len(non_empty_contents)
    return pairs

def read_all_file_line_in_list(file_names: list[str])->list[str]:
    return [ line for name in file_names if (p := pathlib.Path(name)).is_file() for line in p.read_text(encoding="utf-8").splitlines(keepends=True) ]

def make_a_random_copy(pairs: dict[str, int], all_lines: list[str], dest:str)->bool:
    dest_path = pathlib.Path(dest)
    dest_path.mkdir(parents=True, exist_ok=True)
    random.shuffle(all_lines)
    for file, line_count in pairs.items():
        file_path = dest_path / pathlib.Path(file).name
        buffer = all_lines[-line_count:]
        del all_lines[-line_count:]
        with open(file_path.__str__(), "w", encoding='utf-8') as file:
            file.writelines(buffer)

if __name__=="__main__":
    selected_dir = sys.argv[1:]
    print(selected_dir)
    for single_dir in selected_dir:
        pairs = get_file_list(single_dir)
        file_names = list(pairs)
        all_lines = read_all_file_line_in_list(file_names)
        lines_counts = len(all_lines)
        for key, value in pairs.items():
            print(f"{pathlib.Path(key).name}->{value}")
            lines_counts = lines_counts - value
        if lines_counts != 0:
            print("some errors occured")
        make_a_random_copy(pairs, all_lines, single_dir+"_random")
