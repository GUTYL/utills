"""
jsonl_to_excel.py

将 JSONL 文件转换为 Excel (.xlsx) 文件。

用法:
    python jsonl_to_excel.py input.jsonl output.xlsx
"""

import sys
import json
from pathlib import Path
from typing import Any, Dict, List

import pandas as pd


def flatten_obj(obj: Dict[str, Any], parent_key: str = "", sep: str = ".") -> Dict[str, Any]:
    out: Dict[str, Any] = {}
    for k, v in obj.items():
        new_key = f"{parent_key}{sep}{k}" if parent_key else k
        if isinstance(v, dict):
            out.update(flatten_obj(v, new_key, sep=sep))
        else:
            out[new_key] = v
    return out


def load_jsonl(path: Path) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    with path.open("r", encoding="utf-8") as f:
        for line_num, line in enumerate(f, start=1):
            text = line.strip()
            if not text:
                continue
            try:
                obj = json.loads(text)
            except json.JSONDecodeError as e:
                raise ValueError(f"第 {line_num} 行不是合法 JSON: {e}")
            if not isinstance(obj, dict):
                raise ValueError(f"第 {line_num} 行不是 JSON 对象: {text}")
            rows.append(flatten_obj(obj))
    return rows


def save_excel(rows: List[Dict[str, Any]], out_path: Path):
    df = pd.DataFrame(rows)
    df.to_excel(out_path, index=False)
    print(f"已输出: {out_path} ({len(rows)} 行, {len(df.columns)} 列)")


def main(argv=None):
    if argv is None:
        argv = sys.argv
    if len(argv) != 3:
        print("用法: python jsonl_to_excel.py input.jsonl output.xlsx")
        sys.exit(1)

    input_path = Path(argv[1])
    output_path = Path(argv[2])

    if not input_path.exists():
        print(f"输入文件不存在: {input_path}")
        sys.exit(2)

    rows = load_jsonl(input_path)
    if not rows:
        print("输入文件没有记录")
        sys.exit(3)

    save_excel(rows, output_path)


if __name__ == "__main__":
    main()
