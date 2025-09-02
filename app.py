from flask import Flask, request, jsonify
from flask_cors import CORS
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.action_chains import ActionChains
import time, random, base64, io, os, threading, configparser
from PIL import Image
from captcha_recognizer.slider import SliderV2
import logging
from selenium.common.exceptions import TimeoutException
import zipfile
import shutil
from pathlib import Path
import win32print
import subprocess

app = Flask(__name__)
CORS(app)  # 启用跨域支持

# 获取应用程序的根目录
def get_resource_path(relative_path):
    base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# 加载配置文件
config = configparser.ConfigParser()
config_file = get_resource_path('config.ini')
if os.path.exists(config_file):
    config.read(config_file, encoding='utf-8')
else:
    print("警告：配置文件 config.ini 不存在，使用默认配置")

# 配置日志
log_dir = get_resource_path(config.get('DEFAULT', 'LOG_DIR', fallback='logs'))
os.makedirs(log_dir, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(log_dir, 'app.log'), encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# 从配置文件读取参数
IMG_DIR = get_resource_path(config.get('DEFAULT', 'IMG_DIR', fallback=r'test-image'))
EDGE_DRIVER_PATH = get_resource_path(config.get('DEFAULT', 'EDGE_DRIVER_PATH', fallback=r'browser_driver\msedgedriver.exe'))
MAX_RETRY = config.getint('DEFAULT', 'MAX_RETRY', fallback=5)
SESSION_TIMEOUT = config.getint('DEFAULT', 'SESSION_TIMEOUT', fallback=1800)
HEADLESS = config.getboolean('DEFAULT', 'HEADLESS', fallback=True)
EXTRACT_PATH = get_resource_path(config.get('DEFAULT', 'EXTRACT_PATH', fallback=r'extract'))
DOWNLOAD_DIR = get_resource_path(config.get('DEFAULT', 'DOWNLOAD_DIR', fallback=r'downloads'))
PRINTER_NAME = config.get('PRINTER', 'PRINTER_NAME', fallback="TestPrinter")
PDFTO_PRINTER_EXE = get_resource_path(config.get('PRINTER', 'PDFTO_PRINTER_EXE', fallback=r'printer\PDFtoPrinter.exe'))

# 打印机状态码
# ---------- 状态常量 ----------
PRINTER_STATUS_PAUSED           = 0x00000001
PRINTER_STATUS_ERROR            = 0x00000002
PRINTER_STATUS_PENDING_DELETION = 0x00000004
PRINTER_STATUS_PAPER_JAM        = 0x00000008
PRINTER_STATUS_PAPER_OUT        = 0x00000010
PRINTER_STATUS_MANUAL_FEED      = 0x00000020
PRINTER_STATUS_PAPER_PROBLEM    = 0x00000040
PRINTER_STATUS_OFFLINE          = 0x00000080
PRINTER_STATUS_IO_ACTIVE        = 0x00000100
PRINTER_STATUS_BUSY             = 0x00000200
PRINTER_STATUS_PRINTING         = 0x00000400
PRINTER_STATUS_OUTPUT_BIN_FULL  = 0x00000800
PRINTER_STATUS_NOT_AVAILABLE    = 0x00001000
PRINTER_STATUS_WAITING          = 0x00002000
PRINTER_STATUS_PROCESSING       = 0x00004000
PRINTER_STATUS_INITIALIZING     = 0x00008000
PRINTER_STATUS_WARMING_UP       = 0x00010000
PRINTER_STATUS_TONER_LOW        = 0x00020000
PRINTER_STATUS_NO_TONER         = 0x00040000
PRINTER_STATUS_PAGE_PUNT        = 0x00080000
PRINTER_STATUS_USER_INTERVENTION = 0x00100000
PRINTER_STATUS_OUT_OF_MEMORY    = 0x00200000
PRINTER_STATUS_DOOR_OPEN        = 0x00400000
PRINTER_STATUS_SERVER_UNKNOWN   = 0x00800000
PRINTER_STATUS_POWER_SAVE       = 0x01000000

STATUS_MAP = {
    PRINTER_STATUS_PAUSED:           "已暂停",
    PRINTER_STATUS_ERROR:            "发生错误",
    PRINTER_STATUS_PENDING_DELETION: "将被删除",
    PRINTER_STATUS_PAPER_JAM:        "卡纸",
    PRINTER_STATUS_PAPER_OUT:        "缺纸",
    PRINTER_STATUS_MANUAL_FEED:      "手动送纸",
    PRINTER_STATUS_PAPER_PROBLEM:    "纸张异常",
    PRINTER_STATUS_OFFLINE:          "脱机",
    PRINTER_STATUS_IO_ACTIVE:        "I/O 活跃",
    PRINTER_STATUS_BUSY:             "忙碌",
    PRINTER_STATUS_PRINTING:         "正在打印",
    PRINTER_STATUS_OUTPUT_BIN_FULL:  "出纸槽满",
    PRINTER_STATUS_NOT_AVAILABLE:    "不可用",
    PRINTER_STATUS_WAITING:          "等待",
    PRINTER_STATUS_PROCESSING:       "正在处理",
    PRINTER_STATUS_INITIALIZING:     "初始化中",
    PRINTER_STATUS_WARMING_UP:       "预热中",
    PRINTER_STATUS_TONER_LOW:        "碳粉不足",
    PRINTER_STATUS_NO_TONER:         "无碳粉",
    PRINTER_STATUS_PAGE_PUNT:        "页被跳过",
    PRINTER_STATUS_USER_INTERVENTION: "需要用户干预",
    PRINTER_STATUS_OUT_OF_MEMORY:    "内存不足",
    PRINTER_STATUS_DOOR_OPEN:        "盖子打开",
    PRINTER_STATUS_SERVER_UNKNOWN:   "服务器未知",
    PRINTER_STATUS_POWER_SAVE:       "节能模式",
}
# 存储登录状态和结果的全局变量
login_status = {
    'is_processing': False,
    'success': False,
    'message': '',
    'last_login_time': None,
    'error_type': '',
    'user_type': '',
    'document_type': ''
}

class CertificateAutomation:
    def __init__(self):
        self.driver = None
        self.wait = None
        self.document_set = {
            '1': '食品经营许可证',
            '2': '法人身份证',
            '3': '组织机构代码证',
            '4': '税务登记证',
            '5': '开户许可证',
            '6': '社保登记证',
            '7': '其他',
            '8': '无',
            '9': '个人身份证',
            '10': '护照',
            '11': '驾驶证',
            '12': '户口本',
            '13': '军官证',
            '14': '港澳通行证',
            '15': '台湾通行证',
            '16': '其他',
            '17': '无',
            '18': '营业执照',
            '19': '法人身份证',
            '20': '组织机构代码证',
            '21': '税务登记证',
            '22': '开户许可证',
            '23': '社保登记证',
            '24': '其他',
            '25': '无',
            '26': '个人身份证',
            '27': '护照',
            '28': '驾驶证',
            '29': '户口本',
            '30': '军官证',
            '31': '港澳通行证',
            '32': '台湾通行证',
            '33': '其他',
            '34': '无',
            '35': '营业执照',
            '36': '法人身份证',
            '37': '组织机构代码证',
            '38': '税务登记证',
            '39': '开户许可证',
            '40': '社保登记证'
        }
        self.document_url = {
            '1': "https://tyrz.zwfw.gxzf.gov.cn/am/auth/login?service=initService&goto=aHR0cHM6Ly90eXJ6Lnp3ZncuZ3h6Zi5nb3YuY24vYW0vb2F1dGgyL2F1dGhvcml6ZT9jbGllbnRfaWQ9bmV3Z3h6d2Z3JmNsaWVudF9zZWNyZXQ9MTExMTExJnNjb3BlPWFsbCZyZXNwb25zZV90eXBlPWNvZGUmc2VydmljZT1pbml0U2VydmljZSZyZWRpcmVjdF91cmk9aHR0cHMlM0ElMkYlMkZ6d2Z3Lmd4emYuZ292LmNuJTJGZXBvcnRhbGFwcGx5JTJGcG9ydGxldCUyRmF1dGhVc2VyTG9naW4lMkZvYXV0aDJVcmwlM0ZqdW1wUGF0aCUzRGFIUjBjSE02THk5NmQyWjNMbWQ0ZW1ZdVoyOTJMbU51TDJKaGJuTm9hUzlwYm1SbGVDOCUzRA==",
            '2': 'https://example.com/doc2',
            '3': 'https://example.com/doc3',
            '4': 'https://example.com/doc4',
            '5': 'https://example.com/doc5',
            '6': 'https://example.com/doc6',
            '7': 'https://example.com/doc7',
            '8': 'https://example.com/doc8',
            '9': 'https://example.com/doc9',
            '10': 'https://example.com/doc10',
            '11': 'https://example.com/doc11',
            '12': 'https://example.com/doc12',
            '13': 'https://example.com/doc13',
            '14': 'https://example.com/doc14',
            '15': 'https://example.com/doc15',
            '16': 'https://example.com/doc16',
            '17': 'https://example.com/doc17',
            '18': 'https://example.com/doc18',
            '19': 'https://example.com/doc19',
            '20': 'https://example.com/doc20',
            '21': 'https://example.com/doc21',
            '22': 'https://example.com/doc22',
            '23': 'https://example.com/doc23',
            '24': 'https://example.com/doc24',
            '25': 'https://example.com/doc25',
            '26': 'https://example.com/doc26',
            '27': 'https://example.com/doc27',
            '28': 'https://example.com/doc28',
            '29': 'https://example.com/doc29',
            '30': 'https://example.com/doc30',
            '31': 'https://example.com/doc31',
            '32': 'https://example.com/doc32',
            '33': 'https://example.com/doc33',
            '34': 'https://example.com/doc34',
            '35': 'https://example.com/doc35',
            '36': 'https://example.com/doc36',
            '37': 'https://example.com/doc37',
            '38': 'https://example.com/doc38',
            '39': 'https://example.com/doc39',
            '40': 'https://example.com/doc40'
        }

    def setup_driver(self):
        """初始化浏览器驱动"""
        options = webdriver.EdgeOptions()
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        options.add_argument("--window-size=1280,1024")
        # if HEADLESS:
        #     options.add_argument("--headless")  # 无头模式，适合服务器运行


        prefs = {
            "download.default_directory": DOWNLOAD_DIR,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        options.add_experimental_option("prefs", prefs)

        service = Service(EDGE_DRIVER_PATH)
        self.driver = webdriver.Edge(service=service, options=options)
        self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        self.wait = WebDriverWait(self.driver, 20)
        
    def fill_legal_login(self, username, password):
        """填写法人登录信息"""
        try:
            # 切法人登录
            legal_login_tab = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//span[text()='法人登录']"))
            )
            if login_status['user_type'] == 'corporate':
                logger.info("切换到法人登录")
                legal_login_tab.click()
                time.sleep(1.5)
            else:
                logger.info("当前为个人登录，无需切换")

            username_field = self.wait.until(EC.presence_of_element_located((By.ID, 'legal_login_name')))
            password_field = self.wait.until(EC.presence_of_element_located((By.ID, 'legal_pswd')))

            username_field.clear()
            username_field.send_keys(username)
            time.sleep(1.3)
            password_field.clear()
            password_field.send_keys(password)
            time.sleep(1.3)
            
            logger.info("账号和密码输入完成")
            return True
        except Exception as e:
            logger.error(f"填写登录信息失败: {str(e)}")
            return False
    
    def get_drag_distance_with_retry(self, web_image_width, max_retry=None):
        """获取滑块拖动距离，识别失败会自动刷新图片重试"""
        if max_retry is None:
            max_retry = MAX_RETRY
            
        for attempt in range(1, max_retry + 1):
            try:
                captcha_img = self.wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "#mpanel2 .backImg"))
                )
                src_data = captcha_img.get_attribute("src")
                if not src_data.startswith("data:image"):
                    raise RuntimeError("验证码图片src异常")

                bg_b64 = src_data.split("base64,")[1]
                bg_bytes = base64.b64decode(bg_b64)

                # 保存图片到本地临时文件
                tt = time.time()
                img_name = f'{tt}_image.png'
                img_abs_path = os.path.join(IMG_DIR, img_name)
                
                # 确保目录存在
                os.makedirs(IMG_DIR, exist_ok=True)
                
                with open(img_abs_path, "wb") as f:
                    f.write(bg_bytes)

                logger.info(f"第{attempt}次尝试：调用本地模型识别缺口位置...")
                box, _ = SliderV2().identify(source=img_abs_path, show=False)

                if not box:
                    raise RuntimeError("未能识别出缺口位置")

                raw_x = float(box[0])
                logger.info(f"识别出的原始缺口X坐标: {raw_x}")

                # 计算缩放
                with Image.open(io.BytesIO(bg_bytes)) as img:
                    orig_w = float(img.width)

                scale = web_image_width / orig_w if orig_w else 1.0
                initial_slider_x = 12
                distance = (raw_x - initial_slider_x) * scale

                logger.info(f"网页图片宽度: {web_image_width}, 原始图片宽度: {orig_w}")
                logger.info(f"缩放比例: {scale:.4f}, 拖动距离: {distance:.2f}")

                return max(1, int(distance))

            except RuntimeError as e:
                logger.error(f"识别失败: {e}")
                if attempt < max_retry:
                    logger.info("点击刷新图片按钮重试...")
                    try:
                        refresh_btn = self.wait.until(
                            EC.element_to_be_clickable((By.XPATH, '//*[@id="mpanel2"]/div[1]/div/div/i'))
                        )
                        refresh_btn.click()
                        time.sleep(2)
                    except Exception as refresh_e:
                        logger.error(f"刷新验证码失败: {refresh_e}")
                else:
                    raise RuntimeError("达到最大重试次数，仍未识别出缺口位置")
    
    def generate_human_like_track(self, distance):
        """生成类人的拖动轨迹"""
        track, current = [], 0.0
        mid = distance * random.uniform(0.6, 0.8)
        t, v = 0.2, 0.0
        while current < distance:
            a = random.uniform(2, 4) if current < mid else -random.uniform(3, 5)
            v0 = v
            v = max(v0 + a * t, 0)
            move = v0 * t + 0.5 * a * (t ** 2)
            move = max(1, move)
            if current + move > distance: 
                move = distance - current
            current += move
            track.append(int(round(move)))
        return track
    
    def solve_slider_captcha(self):
        """解决滑块验证码"""
        try:
            # 先找到滑块并按住 → 背景图才会加载
            slider_button = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//div[@id='mpanel2']//div[contains(@class,'verify-move-block')]"))
            )
            action = ActionChains(self.driver)
            action.move_to_element(slider_button).click_and_hold(slider_button).perform()
            time.sleep(1)  # 等背景图渲染

            # 背景图出现
            captcha_element = self.wait.until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, "#mpanel2 .backImg"))
            )
            web_image_width = captcha_element.size['width']
            logger.info(f"网页验证码图片宽度: {web_image_width}")

            # 识别并算拖动距离
            drag_distance = self.get_drag_distance_with_retry(web_image_width, max_retry=5)

            # 继续拖动
            track = self.generate_human_like_track(drag_distance)
            logger.info(f"开始拖动滑块，轨迹步数: {len(track)}，总距离：{drag_distance}")

            for move in track:
                action.move_by_offset(xoffset=move, yoffset=random.uniform(-1, 1)).perform()
                time.sleep(random.uniform(0.01, 0.03))

            action.release().perform()
            logger.info("滑块拖动完成")
            return True
            
        except Exception as e:
            logger.error(f"解决滑块验证码失败: {str(e)}")
            return False

    # 文件解压函数
    def extract_zip_file(self, src_dir: str, dst_dir: str, enc='gbk'):
        src_path = Path(src_dir)
        dst_path = Path(dst_dir)
        dst_path.mkdir(parents=True, exist_ok=True)

        for zf in src_path.rglob("*.zip"):
            target_dir = dst_path / zf.stem
            counter = 1
            while target_dir.exists():
                target_dir = dst_path / f"{zf.stem}_{counter}"
                counter += 1

            try:
                with zipfile.ZipFile(zf, 'r') as zip_ref:
                    for info in zip_ref.infolist():
                        # 1. 用 GBK 解码文件名
                        name = info.filename.encode('cp437').decode(enc)
                        target_file = target_dir / name
                        target_file.parent.mkdir(parents=True, exist_ok=True)
                        # 2. 写文件
                        if info.is_dir():
                            target_file.mkdir(exist_ok=True)
                        else:
                            with zip_ref.open(info) as src, open(target_file, 'wb') as dst:
                                dst.write(src.read())
                logger.info(f"解压完成：{zf} → {target_dir}")
                for item in src_path.iterdir():
                    if item.is_file():
                        item.unlink()
                    elif item.is_dir():
                        shutil.rmtree(item)
                logger.info("已清空 src_dir 中的所有文件/子目录")
            except Exception as e:
                logger.error(f"解压失败：{zf}，原因：{e}")

    def _get_printer_status(self,printer_name: str) -> str:
        try:
            h_printer = win32print.OpenPrinter(printer_name)
            status = win32print.GetPrinter(h_printer, 2)["Status"]
            win32print.ClosePrinter(h_printer)
        except Exception as e:
            return f"打开打印机失败：{e}"

        if status == 0:
            return "就绪"

        desc_list = [desc for flag, desc in STATUS_MAP.items() if status & flag]
        return " | ".join(desc_list) if desc_list else f"未知状态(0x{status:X})"

    def  _ensure_pdftoprinter(self):
        print(PDFTO_PRINTER_EXE)
        if os.path.isfile(PDFTO_PRINTER_EXE):
            logger.info("打印程序存在")
            return PDFTO_PRINTER_EXE
        logger.info("打印程序不存在")
        return None

    # 打印机打印函数
    def print_document(self, printer_name: str, pdf_folder: str) -> None:
        # 1. 检查 PDF 文件夹是否存在
        if not os.path.isdir(pdf_folder):
            logger.error("PDF 文件夹不存在")
            return {"success": False, "message": "PDF 文件夹不存在"}
        logger.info("PDF 文件夹存在")

        # 2. 检查打印机状态
        status = self._get_printer_status(printer_name)
        if status != "就绪":
            login_status['error_type'] = 'printer_error'
            logger.error(f"打印机状态异常：{status}")
            return {"success": False, "message": f"打印机状态异常：{status}"}
        logger.info("打印机状态正常")

        # 3. 检查 PDFtoPrinter
        exe = self._ensure_pdftoprinter()
        # 4. 下发打印任务
        for pdf_file in Path(pdf_folder).rglob("*.pdf"):
            logger.info(f"开始执行打印")
            cmd = [exe, str(pdf_file), printer_name]
            try:
                subprocess.run(cmd, check=True, capture_output=True)
                logger.info(f"已发送打印任务：{pdf_file}")
            except subprocess.CalledProcessError as e:
                logger.error(f"打印任务失败：{e.stderr.decode(errors='ignore')}")
                return {"success": False, "message": f"打印任务失败：{e.stderr.decode(errors='ignore')}"}

        # 5. 轮询直到完成或出错
        while True:
            status = self._get_printer_status(printer_name)
            if status == "就绪":
                return {"success": True, "message": "打印完成"}
            elif status == "正在打印":
                time.sleep(0.5)
                continue
            else:
                return {"success": False, "message": f"打印异常：{status}"}

    def login_and_check_status(self, username, password):
        """执行登录并检查证件状态"""
        try:
            # 初始化浏览器
            self.setup_driver()
            
            # 打开登录页面
            LOGIN_URL = self.document_url[login_status['document_type']]
            # LOGIN_URL = "https://tyrz.zwfw.gxzf.gov.cn/am/auth/login?service=initService&goto=aHR0cHM6Ly90eXJ6Lnp3ZncuZ3h6Zi5nb3YuY24vYW0vb2F1dGgyL2F1dGhvcml6ZT9jbGllbnRfaWQ9bmV3Z3h6d2Z3JmNsaWVudF9zZWNyZXQ9MTExMTExJnNjb3BlPWFsbCZyZXNwb25zZV90eXBlPWNvZGUmc2VydmljZT1pbml0U2VydmljZSZyZWRpcmVjdF91cmk9aHR0cHMlM0ElMkYlMkZ6d2Z3Lmd4emYuZ292LmNuJTJGZXBvcnRhbGFwcGx5JTJGcG9ydGxldCUyRmF1dGhVc2VyTG9naW4lMkZvYXV0aDJVcmwlM0ZqdW1wUGF0aCUzRGFIUjBjSE02THk5NmQyWjNMbWQ0ZW1ZdVoyOTJMbU51TDJKaGJuTm9hUzlwYm1SbGVDOCUzRA=="

            self.driver.get(LOGIN_URL)
            logger.info("网页已打开，等待元素加载...")
            
            # 填写登录信息
            if not self.fill_legal_login(username, password):
                login_status['error_type'] = 'time_error'
                return False, "填写登录信息失败"
            
            # 解决滑块验证码
            if not self.solve_slider_captcha():
                login_status['error_type'] = 'time_error'
                return False, "验证码识别失败"
            
            
            # 登录重试逻辑
            max_login_attempts = 3
            for attempt in range(max_login_attempts):
                old_url = self.driver.current_url
                logger.info(f"尝试登录，第 {attempt + 1} 次")
                
                # 点击登录按钮
                login_btn = self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="form_lists"]/div[1]/div[2]/button'))
                )
                time.sleep(1)
                login_btn.click()
                
                try:
                    # 等待 URL 变化（1 秒内）
                    WebDriverWait(self.driver, 1).until(lambda d: d.current_url != old_url)
                    logger.info("登录成功，已跳转到下一页")
                    break  # 登录成功，跳出重试循环
                    
                except TimeoutException:
                    # 1 秒后仍在登录页 ⇒ 出现了错误提示
                    try:
                        error_tip = self.driver.find_element(By.CSS_SELECTOR, ".err_tip .err_text")
                        error_text = error_tip.text.strip()
                        logger.info(f"登录失败：{error_text}")
                        
                        if error_text == "用户名或密码不正确":
                            login_status['error_type'] = 'username_or_password_error'
                            return False, "用户名或密码不正确"
                        
                        elif error_text in ["请输入统一社会信用代码", "请进行滑块验证"]:
                            # 验证码相关错误，可以重试
                            if attempt < max_login_attempts - 1:  # 不是最后一次尝试
                                logger.info(f"验证码错误，准备重试 (剩余 {max_login_attempts - attempt - 1} 次)")
                                try:
                                    # 刷新验证码
                                    # refresh_btn = self.wait.until(
                                    #     EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[2]/div/div[1]/div[2]/div[3]/div/div[1]/div/div/i'))
                                    # )
                                    # refresh_btn.click()
                                    # time.sleep(1)
                                    
                                    # 重新解决滑块验证码
                                    if not self.solve_slider_captcha():
                                        logger.error("重试时验证码识别失败")
                                        if attempt == max_login_attempts - 1:  # 最后一次尝试
                                            return False, "验证码识别失败"
                                        continue  # 继续下一次重试
                                except Exception as refresh_e:
                                    logger.error(f"刷新验证码失败: {refresh_e}")
                                    if attempt == max_login_attempts - 1:  # 最后一次尝试
                                        return False, "刷新验证码失败"
                            else:
                                # 已是最后一次尝试
                                login_status['error_type'] = 'time_error'
                                return False, f"多次重试后仍然登录失败: {error_text}"
                        else:
                            # 其他类型错误，不重试
                            login_status['error_type'] = 'time_error'
                            return False, f"登录失败: {error_text}"
                            
                    except Exception as e:
                        logger.error(f"获取错误信息失败: {e}")
                        if attempt == max_login_attempts - 1:  # 最后一次尝试
                            return False, "登录失败，无法获取错误信息"
            
            # 检查最终是否登录成功
            if "login" in self.driver.current_url:
                login_status['error_type'] = 'time_error'
                return False, "登录超时或失败"

            time.sleep(2)
            
            # 检查登录结果
            if "login" not in self.driver.current_url:            # 登录成功,成功跳转到下一个页面
                logger.info(f"登录成功! 当前URL: {self.driver.current_url}")
                
                # 导航到证件页面
                self.driver.get("https://zhjg.scjdglj.gxzf.gov.cn:10001/TopFDOAS/topic/homePage.action?currentLink=foodOp")
                time.sleep(2)
                self.driver.get("https://zhjg.scjdglj.gxzf.gov.cn:10001/TopFDOAS/topic/homePage.action?currentLink=foodOp")
                time.sleep(2)
                # 点击相关tab
                my_button = self.wait.until(
                    EC.element_to_be_clickable((By.ID, "tab-second"))
                )
                ActionChains(self.driver).click(my_button).perform()
                time.sleep(2)

                # 检查证件状态
                try:
                    ele = WebDriverWait(self.driver, 20).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "div.tni-status.tni-status__success"))
                    )
                    
                    text = WebDriverWait(self.driver, 20).until(
                        lambda d: ele.get_attribute("textContent").strip()
                    )
                    text = text.strip()
                    logger.info(f"证件状态：{text}")

                    if text == "准予":
                        # 点击更多按钮进行打印
                        more_btn = self.wait.until(
                            EC.element_to_be_clickable(
                                (By.XPATH, '/html/body/div[1]/div[2]/div/div[1]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[3]/table/tbody/tr/td[7]/div/div/div/button')
                            )
                        )
                        more_btn.click()
                        time.sleep(2)
                        print_btn = self.wait.until(
                            EC.element_to_be_clickable(
                                (By.XPATH, '/html/body/ul/li[2]/button')
                            )
                        )
                        print_btn.click()
                        time.sleep(2)
                        # 如果文件夹非空，解压文件夹中下载的文件
                        print(EXTRACT_PATH)
                        print(DOWNLOAD_DIR)
                        if os.listdir(DOWNLOAD_DIR):
                            # 先清空目标文件夹
                            if os.path.exists(EXTRACT_PATH):
                                shutil.rmtree(EXTRACT_PATH)
                            os.makedirs(EXTRACT_PATH, exist_ok=True)
                            

                            self.extract_zip_file(DOWNLOAD_DIR, EXTRACT_PATH)

                        self.print_document(PRINTER_NAME, EXTRACT_PATH)  # 打印文件夹中所有的 PDF 文件
                        logger.info("证件打印成功")
                        login_status['error_type'] = ''
                        return True, "证件打印成功"
                    else:
                        login_status['error_type'] = 'invalid_certificate_status'
                        return False, f"证件状态不符合打印条件，当前状态：{text}"
                        
                except Exception as status_e:
                    logger.error(f"检查证件状态失败: {str(status_e)}")
                    login_status['error_type'] = 'time_error'
                    return False, "无法获取证件状态"

            else:  # 登录失败，页面未跳转
                try:
                    err = self.driver.find_element(By.CLASS_NAME, 'layui-layer-content').text
                    login_status['error_type'] = 'time_error'
                    return False, f"登录失败: {err}"
                except:
                    return False, "登录失败，未找到具体错误信息"
                finally:
                    login_status['error_type'] = 'time_error'
        except Exception as e:
            logger.error(f"自动化流程失败: {str(e)}")
            login_status['error_type'] = 'time_error'
            return False, f"系统错误: {str(e)}"
        finally:
            if self.driver:
                self.driver.quit()

def background_login_task(username, password):
    """后台执行登录任务"""
    global login_status
    
    login_status['is_processing'] = True
    login_status['success'] = False
    login_status['message'] = '正在处理中...'
    
    try:
        automation = CertificateAutomation()
        success, message = automation.login_and_check_status(username, password)
        
        login_status['success'] = success
        login_status['message'] = message
        login_status['last_login_time'] = time.time()
        
    except Exception as e:
        login_status['success'] = False
        login_status['message'] = f"处理过程中发生错误: {str(e)}"
    finally:
        login_status['is_processing'] = False

@app.route('/api/document_type', methods=['POST'])
def document_type():
    """文档类型接口"""
    try:
        # 验证请求数据
        if not request.is_json:
            return jsonify({'error': '请求必须是JSON格式'}), 400
        
        data = request.get_json()
        if not data or 'user_type' not in data or 'document_type' not in data:
            return jsonify({'error': '缺少必要参数：user_type 和 document_type'}), 400

        user_type = data['user_type']
        document_type = data['document_type']
        # 将数据保存到全局变量中，方便后续使用
        login_status['user_type'] = user_type
        login_status['document_type'] = document_type

        if user_type not in ['corporate', 'individual']:
            return jsonify({'error': 'user_type参数值无效，必须是corporate或individual'}), 400

        if document_type not in ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10",
                                  "11", "12", "13", "14", "15", "16", "17", "18", "19", "20",
                                  "21", "22", "23", "24", "25", "26", "27", "28", "29", "30",
                                  "31", "32", "33", "34", "35", "36", "37", "38", "39", "40"]:
            return jsonify({'error': 'document_type参数值无效，必须是1到40之间的数字'}), 400

        return jsonify({
            'message': f'证件类型已设置为: {document_type}'
        }), 200
        
    except Exception as e:
        logger.error(f"文档类型接口错误: {str(e)}")
        return jsonify({'error': f'服务器内部错误: {str(e)}'}), 500

@app.route('/api/corporate_login', methods=['POST'])
def corporate_login():
    """法人登录接口"""
    try:
        # 验证请求数据
        if not request.is_json:
            return jsonify({'error': '请求必须是JSON格式','error_type': login_status['error_type']}), 400

        data = request.get_json()
        if not data or 'username' not in data or 'password' not in data:
            return jsonify({'error': '缺少必要参数：username 和 password','error_type': login_status['error_type']}), 400

        username = data['username']
        password = data['password']
        
        # 验证参数不为空
        if not username or not password:
            return jsonify({'error': 'username 和 password 不能为空','error_type': login_status['error_type']}), 400

        # 检查是否有任务正在处理
        if login_status['is_processing']:
            return jsonify({'message': '有任务正在处理中，请稍后再试'}), 429
        
        # 启动后台任务
        login_status['user_type'] = 'corporate'
        thread = threading.Thread(target=background_login_task, args=(username, password))
        thread.daemon = True
        thread.start()
        
        return jsonify({
            'message': '登录请求已接收，正在后台处理',
            'status': 'processing'
        }), 200
        
    except Exception as e:
        logger.error(f"登录接口错误: {str(e)}")
        return jsonify({'error': f'服务器内部错误: {str(e)}'}), 500

@app.route('/api/individual_login', methods=['POST'])
def individual_login():
    """个人登录接口"""
    try:
        # 验证请求数据
        if not request.is_json:
            return jsonify({'error': '请求必须是JSON格式'}), 400
        
        data = request.get_json()
        if not data or 'username' not in data or 'password' not in data:
            return jsonify({'error': '缺少必要参数：username 和 password'}), 400
        
        username = data['username']
        password = data['password']
        
        # 验证参数不为空
        if not username or not password:
            return jsonify({'error': 'username 和 password 不能为空'}), 400
        
        # 检查是否有任务正在处理
        if login_status['is_processing']:
            return jsonify({'message': '有任务正在处理中，请稍后再试'}), 429
        
        # 启动后台任务
        login_status['user_type'] = 'individual'
        if login_status['document_type'] is None:
            return jsonify({'error': '请先设置document_type'}), 400
        thread = threading.Thread(target=background_login_task, args=(username, password))
        thread.daemon = True
        thread.start()
        
        return jsonify({
            'message': '登录请求已接收，正在后台处理',
            'status': 'processing'
        }), 200
        
    except Exception as e:
        logger.error(f"登录接口错误: {str(e)}")
        return jsonify({'error': f'服务器内部错误: {str(e)}'}), 500

@app.route('/api/print_status', methods=['GET'])
def check_print_status():
    """打印状态查询接口"""
    try:
        global login_status
        
        if login_status['is_processing']:
            return jsonify({
                'success': False,
                'msg': '正在处理中，请稍后查询'
            }), 204

        # 如果从未执行过登录
        if login_status['last_login_time'] is None:
            return jsonify({
                'success': False,
                'msg': '尚未执行登录操作'
            }), 410

        # 检查任务是否太久之前执行的（超过配置的超时时间认为过期）
        if time.time() - login_status['last_login_time'] > SESSION_TIMEOUT:
            return jsonify({
                'success': False,
                'msg': '登录状态已过期，请重新执行登录'
            }), 410

        return jsonify({
            'success': login_status['success'],
            'msg': login_status['message'],
            'error_type': login_status['error_type']
        }), 200

    except Exception as e:
        logger.error(f"状态查询接口错误: {str(e)}")
        return jsonify({
            'success': False,
            'msg': f'查询状态时发生错误: {str(e)}'
        }), 500

# 清除extract数据
@app.route('/api/clear_data', methods=['GET'])
def clear_data():
    try:
        # 清除提取的数据
        if os.path.exists(EXTRACT_PATH):
            shutil.rmtree(EXTRACT_PATH)
        os.makedirs(EXTRACT_PATH, exist_ok=True)

        return jsonify({'message': '提取数据已清除'}), 200
    except Exception as e:
        logger.error(f"清除数据接口错误: {str(e)}")
        return jsonify({'error': f'服务器内部错误: {str(e)}'}), 500

@app.errorhandler(404)
def not_found(error):
    return jsonify({'error': '接口不存在'}), 404

@app.errorhandler(500)
def internal_error(error):
    return jsonify({'error': '服务器内部错误'}), 500

if __name__ == '__main__':
    # 确保必要目录存在
    os.makedirs(IMG_DIR, exist_ok=True)
    os.makedirs(log_dir, exist_ok=True)
    
    # 从配置文件读取Flask运行参数
    host = config.get('DEFAULT', 'HOST', fallback='0.0.0.0')
    port = config.getint('DEFAULT', 'PORT', fallback=8848)
    debug = config.getboolean('DEFAULT', 'DEBUG', fallback=True)
    
    logger.info(f"启动Flask应用: http://{host}:{port}")
    app.run(host=host, port=port, debug=debug)
