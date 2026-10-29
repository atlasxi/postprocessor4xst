#!/usr/bin/env -S uv run --script
import re
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

PATTERN = "*.txt"


# ---------- 纯函数：解析与转换 ----------

def parse_line(line: str) -> dict:
    """把一行 'Num:1,BusVol:752.7,...' 解析成 dict。"""
    pairs = []
    for item in line.split(","):
        if ":" not in item:
            continue
        key, value = item.split(":", 1)
        pairs.append((key.strip(), value.strip()))
    return dict(pairs)


def parse_text(text: str) -> tuple[list[str], list[dict]]:
    """解析整个文本，返回 (表头顺序, 行字典列表)。"""
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    dicts = [parse_line(ln) for ln in lines]
    dicts = [d for d in dicts if d]
    headers = list(dicts[0].keys()) if dicts else []
    return headers, dicts


def to_number(s: str):
    """字符串转 int/float，失败原样返回。"""
    if s == "":
        return s
    try:
        if re.fullmatch(r"[+-]?\d+", s):
            return int(s)
        return float(s)
    except ValueError:
        return s


def to_rows(headers: list[str], dicts: list[dict]) -> list[list]:
    """按表头顺序把字典列表转成二维表，并做数值转换。"""
    return [[to_number(d.get(h, "")) for h in headers] for d in dicts]


# ---------- 副作用：文件 IO 与 Excel 写入 ----------

def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def convert_file(path: Path) -> Path:
    """读取一个 txt，写出同名 xlsx，返回输出路径。"""
    headers, dicts = parse_text(read_text(path))
    rows = to_rows(headers, dicts)
    out_path = path.with_suffix(".xlsx")
    write_xlsx(out_path, headers, rows)
    return out_path


def write_xlsx(out_path: Path, headers: list[str], rows: list[list]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = out_path.stem

    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in rows:
        ws.append(row)

    ws.freeze_panes = "A2"

    for col_idx, header in enumerate(headers, start=1):
        max_len = max(
            [len(header)]
            + [len(str(r[col_idx - 1])) for r in rows if r[col_idx - 1] is not None]
        )
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 30)

    wb.save(out_path)


# ---------- 编排 ----------

def find_inputs(base: Path) -> list[Path]:
    return sorted(base.glob(PATTERN))


def main() -> None:
    base = Path.cwd()
    inputs = find_inputs(base)

    if not inputs:
        print(f"当前目录未找到匹配 {PATTERN} 的文件：{base}")
        return

    for path in inputs:
        out_path = convert_file(path)
        print(f"{path.name} -> {out_path.name}")

    print(f"完成，共转换 {len(inputs)} 个文件。")


if __name__ == "__main__":
    main()
