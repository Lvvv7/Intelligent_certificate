# 设置虚拟打印机的状态
import win32print, win32con, pywintypes

PRINTER_CONTROL_PAUSE  = 1
PRINTER_CONTROL_RESUME = 2

def pause_printer(printer_name: str, pause: bool = True) -> None:
    """暂停或恢复指定打印机"""
    try:
        h = win32print.OpenPrinter(
                printer_name,
                {"DesiredAccess": win32print.PRINTER_ACCESS_ADMINISTER}
            )
        cmd = PRINTER_CONTROL_PAUSE if pause else PRINTER_CONTROL_RESUME
        win32print.SetPrinter(h, 0, None, cmd)
        win32print.ClosePrinter(h)
        print("已暂停" if pause else "已恢复")
    except pywintypes.error as e:
        print("操作失败：", e)

if __name__ == "__main__":
    pause_printer("TestPrinter", False)  # True：暂停  False：恢复