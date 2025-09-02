# 测试调用某台打印机执行打印任务
import os, sys, urllib.request, zipfile, tempfile, subprocess, win32print, time
from pathlib import Path

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

# 自动加上系统路径
base_path = os.path.dirname(os.path.abspath(__file__))
PDFTO_PRINTER_EXE = os.path.join(base_path, "PDFtoPrinter.exe")
# D:\captcha-recognizer\downloads
EXTRACT_PATH = r"D:\captcha-recognizer\downloads"

def ensure_pdftoprinter() -> str:
    """如果本地不存在 PDFtoPrinter.exe，就自动下载并返回绝对路径"""
    if os.path.isfile(PDFTO_PRINTER_EXE):
        return PDFTO_PRINTER_EXE

    return False

def print_document(printer_name: str, pdf_folder: str) -> None:  # 改成文件夹中的所有文件而非单个文件
        # 1. 检查 PDF 文件夹是否存在
        print(base_path)
        if not os.path.isdir(pdf_folder):
            print("PDF 文件夹不存在")
            return {"success": False, "message": "PDF 文件夹不存在"}
        print("PDF 文件夹存在")

        # 2. 检查打印机状态
        status = get_printer_status(printer_name)
        if status != "就绪":
            print(f"打印机状态异常：{status}")
            return {"success": False, "message": f"打印机状态异常：{status}"}
        else:
            print("打印机状态正常")

        # 3. 检查 PDFtoPrinter
        exe = ensure_pdftoprinter()
        print(f"使用打印程序：{exe}")
        # 4. 下发打印任务
        for pdf_file in Path(pdf_folder).rglob("*.pdf"):
            cmd = [exe, str(pdf_file), printer_name]
            try:
                subprocess.run(cmd, check=True, capture_output=True)
                print(f"已发送打印任务：{pdf_file}")
            except subprocess.CalledProcessError as e:
                print(f"打印任务失败：{e.stderr.decode(errors='ignore')}")
                return {"success": False, "message": f"打印任务失败：{e.stderr.decode(errors='ignore')}"}

        # 5. 轮询直到完成或出错
        while True:
            status = get_printer_status(printer_name)
            if status == "就绪":
                return {"success": True, "message": "打印完成"}
            elif status == "正在打印":
                time.sleep(0.5)
                continue
            else:
                return {"success": False, "message": f"打印异常：{status}"}


# ---------- 3. 使用示例 ----------
if __name__ == "__main__":
    printer = "TestPrinter"   # ← 改成你实际的打印机名称
    pdf = r"doc.pdf"  # ← 改成你实际的 PDF 文件路径

    status = get_printer_status(printer)
    if status == "就绪":
        print_document(printer, EXTRACT_PATH)
    else:
        print("打印机异常，请联系管理员！")


    