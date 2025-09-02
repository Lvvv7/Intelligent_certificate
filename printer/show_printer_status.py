# 输出当前连接的所有打印机的状态
import win32print

# 官方文档里常见的状态位
PRINTER_STATUS_PAUSED          = 0x00000001
PRINTER_STATUS_ERROR           = 0x00000002
PRINTER_STATUS_PENDING_DELETION= 0x00000004
PRINTER_STATUS_PAPER_JAM       = 0x00000008
PRINTER_STATUS_PAPER_OUT       = 0x00000010
PRINTER_STATUS_MANUAL_FEED     = 0x00000020
PRINTER_STATUS_PAPER_PROBLEM   = 0x00000040
PRINTER_STATUS_OFFLINE         = 0x00000080
PRINTER_STATUS_IO_ACTIVE       = 0x00000100
PRINTER_STATUS_BUSY            = 0x00000200
PRINTER_STATUS_PRINTING        = 0x00000400
PRINTER_STATUS_OUTPUT_BIN_FULL = 0x00000800
PRINTER_STATUS_NOT_AVAILABLE   = 0x00001000
PRINTER_STATUS_WAITING         = 0x00002000
PRINTER_STATUS_PROCESSING      = 0x00004000
PRINTER_STATUS_INITIALIZING    = 0x00008000
PRINTER_STATUS_WARMING_UP      = 0x00010000
PRINTER_STATUS_TONER_LOW       = 0x00020000
PRINTER_STATUS_NO_TONER        = 0x00040000
PRINTER_STATUS_PAGE_PUNT       = 0x00080000
PRINTER_STATUS_USER_INTERVENTION=0x00100000
PRINTER_STATUS_OUT_OF_MEMORY   = 0x00200000
PRINTER_STATUS_DOOR_OPEN       = 0x00400000
PRINTER_STATUS_SERVER_UNKNOWN  = 0x00800000
PRINTER_STATUS_POWER_SAVE      = 0x01000000

STATUS_MAP = {
    PRINTER_STATUS_PAUSED          : "已暂停",
    PRINTER_STATUS_ERROR           : "发生错误",
    PRINTER_STATUS_PENDING_DELETION: "将被删除",
    PRINTER_STATUS_PAPER_JAM       : "卡纸",
    PRINTER_STATUS_PAPER_OUT       : "缺纸",
    PRINTER_STATUS_MANUAL_FEED     : "手动送纸",
    PRINTER_STATUS_PAPER_PROBLEM   : "纸张异常",
    PRINTER_STATUS_OFFLINE         : "脱机",
    PRINTER_STATUS_IO_ACTIVE       : "I/O 活跃",
    PRINTER_STATUS_BUSY            : "忙碌",
    PRINTER_STATUS_PRINTING        : "正在打印",
    PRINTER_STATUS_OUTPUT_BIN_FULL : "出纸槽满",
    PRINTER_STATUS_NOT_AVAILABLE   : "不可用",
    PRINTER_STATUS_WAITING         : "等待",
    PRINTER_STATUS_PROCESSING      : "正在处理",
    PRINTER_STATUS_INITIALIZING    : "初始化中",
    PRINTER_STATUS_WARMING_UP      : "预热中",
    PRINTER_STATUS_TONER_LOW       : "碳粉不足",
    PRINTER_STATUS_NO_TONER        : "无碳粉",
    PRINTER_STATUS_PAGE_PUNT       : "页被跳过",
    PRINTER_STATUS_USER_INTERVENTION:"需要用户干预",
    PRINTER_STATUS_OUT_OF_MEMORY   : "内存不足",
    PRINTER_STATUS_DOOR_OPEN       : "盖子打开",
    PRINTER_STATUS_SERVER_UNKNOWN  : "服务器未知",
    PRINTER_STATUS_POWER_SAVE      : "节能模式",
}

def get_printer_status(printer_name: str) -> str:
    try:
        hPrinter = win32print.OpenPrinter(printer_name)
        status = win32print.GetPrinter(hPrinter, 2)["Status"]
        win32print.ClosePrinter(hPrinter)
    except Exception as e:
        return f"打开打印机失败：{e}"

    if status == 0:
        return "就绪"

    # 可能有多个位同时置位，保留所有匹配到的描述
    desc_list = [desc for flag, desc in STATUS_MAP.items() if status & flag]
    return " | ".join(desc_list) if desc_list else f"未知状态(0x{status:X})"

def get_printer_status_verbose(printer_name: str):
    try:
        hPrinter = win32print.OpenPrinter(printer_name)
        info = win32print.GetPrinter(hPrinter, 2)
        status = info["Status"]
        win32print.ClosePrinter(hPrinter)
    except Exception as e:
        return {"name": printer_name, "raw": None, "error": str(e)}

    if status == 0:
        return {"name": printer_name, "raw": 0, "states": ["就绪"]}

    states = [desc for flag, desc in STATUS_MAP.items() if status & flag]
    return {"name": printer_name, "raw": status, "states": states or [f"未知(0x{status:X})"]}

if __name__ == "__main__":
    # 列出所有本地打印机
    # printers = win32print.EnumPrinters(
    #     win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    # )
    # for p in printers:
    #     name = p[2]
    #     print(f"{name}: {get_printer_status(name)}")
    # 打印详细结果
    # for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS):
    #     print(get_printer_status_verbose(p[2]))
    import time, sys
    last_cache = {}
    while True:
        for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS):
            name = p[2]
            now = get_printer_status_verbose(name)
            if last_cache.get(name) != now:
                print(now)
                last_cache[name] = now
        time.sleep(2)