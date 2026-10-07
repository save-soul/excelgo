"""用 openpyxl 加载 xlsx，失败时以非零退出码报出原因。

存在的意义：openpyxl 是独立的第三方实现，它加载失败即证明产物确实是坏文件。
本脚本刻意让异常冒泡（不做静默兜底），退出码非 0 即加载失败。
"""
import sys

import openpyxl


def main(path):
    try:
        wb = openpyxl.load_workbook(path)
    except Exception as exc:  # noqa: BLE001 - 任何异常都意味着文件坏了
        print(f"LOAD_FAIL {type(exc).__name__}: {exc}")
        return 1
    print("LOAD_OK " + ",".join(wb.sheetnames))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
