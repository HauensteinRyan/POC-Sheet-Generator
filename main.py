"""
CLI entry point:
  python main.py input.docx [output.xlsx]

If output path is omitted, saves next to the input file with .xlsx extension.
"""

import sys
import os
from parser import parse_doc
from writer import write_xlsx


def main():
    if len(sys.argv) < 2:
        print("Usage: python main.py <input.docx> [output.xlsx]")
        sys.exit(1)

    in_path = sys.argv[1]
    if not os.path.isfile(in_path):
        print(f"Error: file not found: {in_path}")
        sys.exit(1)

    if len(sys.argv) >= 3:
        out_path = sys.argv[2]
    else:
        base = os.path.splitext(in_path)[0]
        out_path = base + "_output.xlsx"

    rows = parse_doc(in_path)
    write_xlsx(rows, out_path)


if __name__ == "__main__":
    main()
