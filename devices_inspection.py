#!/usr/bin/env python3
# -*- coding: UTF-8 -*-

"""
project: devices_inspection
file: devices_inspection_new_v11.py
author: Robin
date: 2026-09-18
version: v11
updated: 2026-09-18
description:
    网络设备自动化巡检脚本。
    本版本以 v10 为基线进行小幅修正，
    在不减少原有功能的前提下，优化文件输入、配置校验、日志记录、
    并发调度、异常处理和程序结构；进一步精确过滤 Paramiko SSH banner
    底层 traceback，并保证最终分隔线与总结同时写入 Terminal 和 00_inspection.log。

usage:
    1. 将 Excel 设备信息文件放在本脚本所在目录。
    2. Excel 第一个 Sheet 保存设备登录信息，第二个 Sheet 保存巡检命令。
    3. 运行：python devices_inspection_new_v11.py
    4. 按提示输入 Excel 文件名；可省略 .xlsx。
    5. 直接按 Enter 使用默认文件名 info.xlsx。
    6. 巡检结果保存到本次运行独立创建的时间目录中。
    7. Excel 配置或 Sheet 结构错误时，程序等待用户按 Enter 后再退出。

dependencies:
    Python >= 3.10
    netmiko
    pandas
    msoffcrypto-tool
    openpyxl
"""

import getpass
import logging
import os
import sys
import threading

if os.name == "nt":
    import msvcrt  # Windows 平台保留该模块兼容能力；当前密码输入统一使用 getpass。
import time
from concurrent.futures import Future, ThreadPoolExecutor, as_completed
from contextlib import contextmanager
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Any

import msoffcrypto
import pandas as pd
from netmiko import ConnectHandler, exceptions
from paramiko import transport as paramiko_transport


# ==============================
# 全局配置
# ==============================

DEFAULT_INFO_NAME = "info.xlsx"  # 用户直接按 Enter 时使用的默认 Excel 文件名。
INFO_EXTENSION = ".xlsx"  # Excel 设备信息文件允许省略的文件扩展名。
MAX_WORKERS = 200  # 最大并发巡检任务数量；网络巡检主要受 I/O 等待影响。
CONNECTION_TIMEOUT = 40  # Netmiko 建立设备连接时使用的 TCP/连接超时时间，单位为秒。
COMMAND_READ_TIMEOUT = 120  # 单条巡检命令等待设备返回数据的最长时间，单位为秒。
PASSWORD_MAX_RETRY = 3  # 加密 Excel 密码最多允许输入的次数。

if getattr(sys, "frozen", False):
    # PyInstaller 打包后的 EXE：使用实际 EXE 文件所在目录，避免指向 _MEI 临时解压目录。
    SCRIPT_DIR = Path(sys.executable).resolve().parent
else:
    # 直接运行 Python 脚本：使用当前 .py 文件所在目录。
    SCRIPT_DIR = Path(__file__).resolve().parent


# ==============================
# 自定义异常和数据结构
# ==============================


class PasswordRequiredError(Exception):
    """表示加密 Excel 必须提供非空密码。"""


class ExcelSheetError(Exception):
    """表示 Excel 缺少程序要求的 Sheet 结构。"""


@dataclass
class InspectionResult:
    """
    保存单台设备的最终巡检结果。

    统一返回结果可以让主线程直接统计成功、失败和耗时，
    不需要再通过读取日志文件的行数来推断异常设备数量。
    """

    host: str
    success: bool
    duration: float
    error_type: str = ""
    error_message: str = ""


# ==============================
# 文件名处理
# ==============================


def normalize_filename(
    user_input: str,
    default_name: str,
    extension: str,
) -> str:
    """
    规范化用户输入的文件名。

    支持以下输入方式：
        1. 直接按 Enter：使用默认文件名。
        2. 只输入文件名：自动追加指定扩展名。
        3. 输入完整文件名：不重复追加扩展名。

    注意：
        本函数只负责“文件名字符串规范化”，不负责判断文件是否存在。
        文件存在性检查由调用方负责，以保持函数职责单一。
    """
    value = user_input.strip()  # 删除用户输入文件名前后的空白字符。

    if not value:  # 用户直接按 Enter，说明没有提供文件名。
        return default_name  # 返回调用方指定的默认文件名。

    if value.lower().endswith(extension.lower()):  # 不区分大小写检查是否已经包含正确扩展名。
        return value  # 已经存在正确扩展名时直接返回，避免出现 .xlsx.xlsx。

    return f"{value}{extension}"  # 用户只输入文件名时自动追加扩展名。


def prompt_for_info_file() -> Path:
    """
    循环提示用户输入 Excel 设备信息文件名，并检查文件是否存在。

    用户可以输入：
        info
        info.xlsx

    直接按 Enter 时使用默认文件名 info.xlsx。
    如果文件不存在，则继续提示，直到找到有效文件。
    """
    while True:
        info_input = input(
            f"请输入设备信息文件名（默认：{DEFAULT_INFO_NAME}）："
        )  # 获取用户输入的 Excel 文件名。

        info_name = normalize_filename(
            info_input,
            DEFAULT_INFO_NAME,
            INFO_EXTENSION,
        )  # 将用户输入转换为标准文件名。

        info_path = SCRIPT_DIR / info_name  # 使用脚本所在目录构造 Excel 文件的完整路径。

        if info_path.is_file():  # 确认路径存在且指向普通文件。
            return info_path  # 返回已经验证过的文件路径。

        print(f"[ERROR] Excel 文件不存在：{info_path}")  # 告知用户当前文件路径不存在。
        print("请重新输入。\n")  # 引导用户重新输入。


# ==============================
# Paramiko 底层异常过滤
# ==============================


_ORIGINAL_THREADING_EXCEPTHOOK = threading.excepthook  # 保存 Python 原始线程异常处理器，避免覆盖其他未预期异常的处理行为。
_ORIGINAL_PARAMIKO_TRANSPORT_LOG = paramiko_transport.Transport._log  # 保存 Paramiko Transport 原始日志方法，便于只过滤目标异常。


class ParamikoBannerLogFilter(logging.Filter):
    """
    过滤 Paramiko SSH protocol banner 异常对应的日志记录。

    Paramiko 在 SSH Transport 后台线程中处理连接。当目标地址虽然建立了 TCP
    连接，但对端没有正常返回 SSH banner 时，Paramiko 会在后台线程中记录
    ``Error reading SSH protocol banner``，并且通过 logging 的 ``exc_info`` 输出
    完整 traceback。

    本过滤器只识别这一类明确的 banner 异常，其他 Paramiko 日志不主动过滤。
    """

    BANNER_ERROR_TEXT = "Error reading SSH protocol banner"  # Paramiko SSH banner 异常的固定关键字。

    def filter(self, record: logging.LogRecord) -> bool:
        """
        判断日志记录是否应继续输出。

        Returns:
            bool: False 表示过滤该条 banner 异常日志；True 表示正常输出。
        """
        message = record.getMessage()  # 获取 logging record 最终格式化后的消息文本。
        return self.BANNER_ERROR_TEXT not in message  # 只过滤明确包含 banner 错误关键字的日志。


def install_paramiko_log_filter() -> None:
    """
    安装 Paramiko banner 异常的多层过滤。

    为什么 v8 的 logger filter 仍可能无法覆盖所有实际输出路径：
        某些 Paramiko/运行环境组合会直接在 Transport 的内部日志调用路径中
        产生异常输出，单纯给 logger 添加 Filter 并不能保证所有路径都经过
        用户程序预期的 Handler。

    因此 v9 增加一层更靠近 Paramiko 调用源头的精确过滤：
        Transport._log() 收到 banner 异常消息时直接丢弃。

    该处理只针对 ``Error reading SSH protocol banner``，不会关闭整个 Paramiko
    日志系统，也不会修改全局 sys.stderr。
    """

    paramiko_logger = logging.getLogger("paramiko.transport")  # 获取 Paramiko Transport 专用 logger。
    banner_filter = ParamikoBannerLogFilter()  # 创建用于 logger 和 handler 的 banner 过滤器。

    paramiko_logger.addFilter(banner_filter)  # 保留 v8 的 logger 层过滤，兼容正常 logging 路径。

    for handler in paramiko_logger.handlers:
        handler.addFilter(banner_filter)  # 同时过滤现有 Handler，覆盖 logger 直接挂载 Handler 的场景。

    def filtered_transport_log(
        transport: paramiko_transport.Transport,
        level: int,
        message: str,
        *args: Any,
        **kwargs: Any,
    ) -> Any:
        """
        包装 Paramiko Transport._log，只拦截 SSH banner 异常。
        """
        rendered_message = str(message)  # 将 Paramiko 日志消息转换为字符串进行关键字判断。

        if ParamikoBannerLogFilter.BANNER_ERROR_TEXT in rendered_message:
            return None  # 在 Paramiko 产生 logging record 之前直接丢弃 banner 异常，阻止 traceback 输出。

        return _ORIGINAL_PARAMIKO_TRANSPORT_LOG(
            transport,
            level,
            message,
            *args,
            **kwargs,
        )  # 其他 Paramiko Transport 日志完全保持原有行为。

    paramiko_transport.Transport._log = filtered_transport_log  # 安装精确的 Transport 日志过滤包装器。


def install_thread_exception_hook() -> None:
    """
    安装线程级异常过滤器，作为 Paramiko 后台异常处理的最后一道保护。

    Paramiko 正常情况下会自行捕获 Transport 内部异常并通过 logging 输出，
    v9 已在 _log() 层进行精确过滤。这里仍保留线程异常 hook，用于极少数
    Paramiko 版本/异常路径直接把 SSH banner 异常作为未处理线程异常抛出的情况。
    其他异常继续交给 Python 原始 excepthook，不隐藏真正的程序 Bug。
    """

    def thread_exception_hook(args: threading.ExceptHookArgs) -> None:
        exception = args.exc_value  # 获取后台线程实际抛出的异常对象。
        exception_type = type(exception).__name__  # 获取异常类名，例如 SSHException。

        if (
            exception_type == "SSHException"
            and ParamikoBannerLogFilter.BANNER_ERROR_TEXT in str(exception)
        ):
            return  # 已知 SSH banner 异常不输出完整 traceback。

        _ORIGINAL_THREADING_EXCEPTHOOK(args)  # 其他线程异常继续使用 Python 原始处理机制。

    threading.excepthook = thread_exception_hook  # 注册线程级异常过滤器。


# ==============================
# Excel 读取
# ==============================


def is_encrypted(info_path: Path) -> bool:
    """
    判断 Excel 文件是否受 Office 密码保护。

    如果文件无法被 msoffcrypto 正常识别，则交给后续 Excel 读取流程处理。
    这里只负责判断“是否为 Office 加密文件”，不负责解密。
    """
    try:
        with info_path.open("rb") as file:
            return msoffcrypto.OfficeFile(file).is_encrypted()
    except Exception:
        return False


def prompt_for_password() -> str:
    """
    隐式读取 Excel 密码。

    使用标准库 getpass 获取密码：
        - 输入密码时不显示任何字符；
        - 不显示 *；
        - 不把密码写入日志；
        - 输入完成按 Enter 后返回真实密码。

    说明：
        PyCharm、Windows Terminal、VS Code 等运行环境对密码输入的终端回显
        处理方式可能不同，因此这里不自行操作 Windows 控制台回显，而使用
        Python 标准库提供的隐式密码输入机制。
    """
    return getpass.getpass(prompt="")  # 输入过程中完全不显示字符，按 Enter 后返回密码。




def read_encrypted_file(
    info_path: Path,
    max_retry: int = PASSWORD_MAX_RETRY,
) -> tuple[list[dict[str, Any]], dict[str, list[Any]]]:
    """
    解密并读取受密码保护的 Excel 文件。

    处理流程明确拆成四步：
        1. 提示用户输入密码。
        2. 加载密码并解密 Excel 到内存。
        3. 验证解密结果确实可以被 pandas 读取。
        4. 读取两个 Sheet。

    密码不会写入日志，也不会写入任何结果文件。
    """
    for attempt in range(1, max_retry + 1):
        try:
            print(
                "Excel 文件已加密，请输入密码：",
                end="",
                flush=True,
            )  # 先明确显示提示，避免用户误以为程序没有响应。

            password = prompt_for_password()  # 隐式读取密码，输入过程中完全不显示任何字符。

            if not password:
                raise PasswordRequiredError(
                    "文件受密码保护，必须提供密码！"
                )  # 空密码直接判定为无效输入，不进入解密流程。

            print(
                "正在解密 Excel，请稍候...",
                flush=True,
            )  # 解密可能需要一定时间，明确告诉用户程序正在工作而不是卡住。

            decrypted_data = BytesIO()  # 只在内存中保存解密后的 Excel，不生成明文临时文件。

            with info_path.open("rb") as file:
                office_file = msoffcrypto.OfficeFile(file)  # 创建 Office 加密文件对象。
                office_file.load_key(password=password)  # 使用密码派生并加载解密密钥。
                office_file.decrypt(decrypted_data)  # 执行实际解密操作。

            decrypted_data.seek(0)  # 将内存文件指针移动到开头，准备交给 pandas。

            print(
                "Excel 解密完成，正在读取设备信息...",
                flush=True,
            )  # 明确告诉用户解密已经完成，当前进入 Excel 读取阶段。

            devices_info, cmds_info = read_excel_data(
                decrypted_data
            )  # 使用统一的 Excel 读取函数处理两个 Sheet。

            print(
                "Excel 读取完成。",
                flush=True,
            )  # 两个 Sheet 均成功读取后再提示完成。

            return devices_info, cmds_info

        except ExcelSheetError:
            raise  # Excel Sheet 结构错误交由 main() 统一提示并等待 Enter 后退出。

        except RuntimeError:
            raise  # Excel 读取错误交由 main() 统一提示并等待 Enter 后退出。

        except PasswordRequiredError:
            remaining = max_retry - attempt  # 计算剩余密码输入次数。

            if remaining:
                print(
                    f"密码错误或未输入密码，请重新输入！"
                    f"（剩余尝试次数：{remaining}）"
                )  # 空密码允许重新输入。
            else:
                print("超过最大密码尝试次数，程序退出。")
                sys.exit(1)

        except Exception as error:
            error_type = type(error).__name__  # 保存底层异常类型，便于判断解密失败原因。
            raw_message = format_exception_detail(error)  # 提取简短原始异常信息。

            # msoffcrypto 不同版本/不同 Office 加密格式在错误密码时可能抛出不同异常，
            # 因此不能只依赖某一个固定异常类名。
            password_error_keywords = (
                "invalid key",
                "wrong password",
                "bad password",
                "password",
                "key",
            )
            is_password_error = (
                error_type == "InvalidKeyError"
                or any(
                    keyword in raw_message.lower()
                    for keyword in password_error_keywords
                )
            )

            if is_password_error:
                remaining = max_retry - attempt  # 计算剩余密码输入次数。

                if remaining:
                    print(
                        f"密码错误，请重新输入！"
                        f"（剩余尝试次数：{remaining}）"
                    )  # 密码/密钥错误允许重新输入。
                    continue

                print("超过最大密码尝试次数，程序退出。")
                sys.exit(1)

            print(
                f"Excel 解密/读取失败：{error_type}"
                f"{f' - {raw_message}' if raw_message else ''}"
            )
            sys.exit(1)

    raise RuntimeError("Excel 密码处理流程异常结束。")




def read_excel_data(
    excel_source: Any,
) -> tuple[list[dict[str, Any]], dict[str, list[Any]]]:
    """
    从 Excel 文件或内存二进制流中读取第一个和第二个 Sheet。

    第一个 Sheet：
        转换为“每行一个字典”的设备列表。

    第二个 Sheet：
        转换为“列名 -> 所有行值列表”的命令字典。

    keep_default_na=False 用于设备信息 Sheet，避免 pandas 把空字符串
    自动转换成 NaN，从而影响用户名、密码、secret 等字段的判断。
    """
    try:
        devices_dataframe = pd.read_excel(
            excel_source,
            sheet_name=0,
            dtype=str,
            keep_default_na=False,
        )  # 第一张表保存设备登录信息。

        if hasattr(excel_source, "seek"):
            excel_source.seek(0)  # 第二次读取同一个内存文件前，把文件指针重新定位到开头。

        cmds_dataframe = pd.read_excel(
            excel_source,
            sheet_name=1,
            dtype=str,
            keep_default_na=False,
        )  # 第二张表保存不同 device_type 对应的巡检命令。

    except ValueError as error:
        raise ExcelSheetError(
            "Excel 文件缺少必要的 Sheet 信息。"
        ) from error
    except Exception as error:
        raise RuntimeError(
            f"读取 Excel 文件失败：{error}"
        ) from error

    devices_dict = devices_dataframe.to_dict("records")  # 将设备信息表转换为设备字典列表。
    cmds_dict = cmds_dataframe.to_dict("list")  # 将命令表转换为“设备类型 -> 命令列表”。

    return devices_dict, cmds_dict


def read_info(
    info_path: Path,
) -> tuple[list[dict[str, Any]], dict[str, list[Any]]]:
    """
    根据 Excel 是否加密选择对应读取方式。
    """
    if is_encrypted(info_path):
        return read_encrypted_file(info_path)  # 加密文件先解密，再读取两个 Sheet。

    return read_excel_data(info_path)  # 未加密文件直接交给 pandas 读取。


# ==============================
# 配置校验
# ==============================


def validate_info(
    devices_info: list[dict[str, Any]],
    cmds_info: dict[str, list[Any]],
) -> bool:
    """
    在连接任何设备之前检查 Excel 配置的基本完整性。

    必须字段：
        host
        device_type
        username
        password

    secret 明确允许为空，因为部分设备不需要 Enable/特权模式密码。

    巡检命令中的空值也允许存在：
        Excel 第二个 Sheet 为了让不同设备类型的命令列对齐，
        某些设备类型可能没有对应命令，因此空单元格属于合法情况。

    返回：
        True  - 校验通过。
        False - 发现配置错误。
    """
    has_error = False  # 统一记录是否发现任何配置错误。

    if not devices_info:
        print("[ERROR] 设备信息 Sheet 为空。")
        return False

    required_columns = {"host", "ip", "device_type", "username", "password"}  # 定义真正必须存在的设备字段；port 和 secret 均为可选字段。
    actual_columns = set(devices_info[0].keys())  # 获取设备信息表的列名。

    missing_columns = required_columns - actual_columns  # 找出 Excel 中缺失的必要列。
    if missing_columns:
        print(
            f"[ERROR] 设备信息 Sheet 缺少必要字段："
            f"{', '.join(sorted(missing_columns))}"
        )
        has_error = True

    for index, device in enumerate(devices_info, start=2):
        host = str(device.get("host", "")).strip()  # 设备名称/标识用于错误提示和结果文件命名。
        display_host = host or f"第 {index} 行"  # host 为空时仍给用户一个明确的行号。

        for field in required_columns:
            value = device.get(field, "")
            if value is None or not str(value).strip():
                print(
                    f"[ERROR] 设备 {display_host} 的 {field} 字段为空。"
                )
                has_error = True

        # secret 是可选字段，因此这里绝不能把空值视为配置错误。
        # 没有 secret 的设备会在建立连接时直接省略该参数。
        _ = device.get("secret", "")

        device_type = str(device.get("device_type", "")).strip()
        if device_type and device_type not in cmds_info:
            print(
                f"[ERROR] 设备 {display_host} 的 device_type "
                f"“{device_type}”在巡检命令 Sheet 中不存在。"
            )
            has_error = True

    if not has_error:
        print(
            f"{time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())} "
            "[INFO] Excel 配置预检查通过。"
        )

    return not has_error


# ==============================
# 日志系统
# ==============================


def create_run_directory() -> Path:
    """
    为本次巡检创建独立结果目录。

    不再删除已经存在的历史结果目录，避免不同巡检任务之间发生误删。
    目录名精确到秒；如果同一秒已经存在，则自动增加后缀保证唯一。
    """
    timestamp = time.strftime(
        "%Y-%m-%d_%H-%M-%S",
        time.localtime(),
    )  # 使用本地时间生成本次巡检的基础目录名。

    run_dir = SCRIPT_DIR / timestamp  # 将结果目录放在脚本所在目录下。

    if not run_dir.exists():
        run_dir.mkdir(parents=True)  # 首次使用该时间目录时直接创建。
        return run_dir

    counter = 2  # 如果同一秒已有目录，则从 _2 开始寻找可用目录。
    while True:
        candidate = SCRIPT_DIR / f"{timestamp}_{counter}"
        if not candidate.exists():
            candidate.mkdir(parents=True)
            return candidate
        counter += 1


def setup_logging(run_dir: Path) -> logging.Logger:
    """
    创建本次巡检的统一程序日志。

    00_inspection.log 专门保存程序运行过程中的 INFO/WARNING/ERROR 等日志。
    DEBUG 级别用于开发/调试过程中的巡检命令提示，本版本不写入该文件。
    文件名以 00 开头，使其在常规文件管理器排序中稳定排在设备结果文件之前。
    """
    logger = logging.getLogger("devices_inspection")  # 使用固定 logger 名称，避免重复创建多个实例。
    logger.setLevel(logging.DEBUG)  # Logger 保留全部级别，具体输出级别由 Handler 决定。
    logger.handlers.clear()  # 防止重复运行/重复初始化时出现重复日志。
    logger.propagate = False  # 禁止日志继续传播到 root logger，避免重复输出。

    log_file = run_dir / "00_inspection.log"  # 本次运行的统一程序日志文件。

    file_handler = logging.FileHandler(
        log_file,
        mode="w",
        encoding="utf-8",
    )  # 将完整日志写入结果目录。

    console_handler = logging.StreamHandler(sys.stdout)  # 将用户需要看到的运行信息输出到终端。

    formatter = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )  # 统一日志时间、级别和消息格式。

    file_handler.setLevel(logging.INFO)  # 文件只保存 INFO 及以上级别，不记录 DEBUG 巡检命令日志。
    console_handler.setLevel(logging.INFO)  # 控制台默认只显示 INFO 及以上，避免 DEBUG 信息干扰用户。

    file_handler.setFormatter(formatter)  # 给文件 Handler 应用统一格式。
    console_handler.setFormatter(formatter)  # 给控制台 Handler 应用统一格式。

    logger.addHandler(file_handler)  # 将文件 Handler 加入 logger。
    logger.addHandler(console_handler)  # 将控制台 Handler 加入 logger。

    return logger


# ==============================
# 异常转换
# ==============================


def classify_connection_error(error: Exception) -> tuple[str, str]:
    """
    将 Netmiko/Paramiko 连接异常转换成用户容易理解的中文提示。

    返回：
        error_type    - Python/Netmiko 异常类名。
        error_message - 面向用户的简短中文说明。

    原始异常类名会保留在日志中，便于进一步定位问题。
    """
    error_type = type(error).__name__  # 获取实际异常类名，例如 NetmikoTimeoutException。

    if isinstance(error, exceptions.NetmikoAuthenticationException):
        return error_type, "用户名或密码认证失败"

    if isinstance(error, exceptions.NetmikoTimeoutException):
        return error_type, "管理地址或端口不可达"

    if isinstance(error, exceptions.ReadTimeout):
        return error_type, "设备响应超时"

    if isinstance(error, ConnectionRefusedError):
        return error_type, "远程登录协议或服务拒绝连接"

    if isinstance(error, TimeoutError):
        return error_type, "Telnet 连接超时"

    if isinstance(error, AttributeError):
        return error_type, "设备管理地址等连接参数缺失"

    if isinstance(error, ValueError):
        return error_type, "Enable/特权模式密码认证失败或连接参数无效"

    return error_type, "未知连接错误"


def format_exception_detail(error: Exception) -> str:
    """
    提取原始异常中的关键文字，并限制日志长度。

    这里不记录完整 traceback，避免日志被底层 Paramiko 堆栈刷屏；
    同时保留异常类名和简短原始信息，方便判断真实故障类型。
    """
    raw_message = " ".join(str(error).split())  # 将多行异常信息压缩成一行，便于日志阅读。
    if len(raw_message) > 300:
        raw_message = f"{raw_message[:300]}..."  # 防止底层库返回异常长的信息。
    return raw_message


# ==============================
# 单设备巡检
# ==============================


def build_connection_params(
    login_info: dict[str, Any],
) -> dict[str, Any]:
    """
    根据 Excel 中的一行设备登录信息构造 Netmiko 连接参数。

    secret 是可选字段：
        有值 -> 加入连接参数。
        空值 -> 完全不传给 Netmiko。

    这样可以兼容不需要 Enable/特权密码的设备。
    """
    # Excel 中的 host 是设备名称，仅用于日志、结果文件名等标识场景。
    # Netmiko 的 host 参数必须使用 Excel 中的 ip，即设备实际的管理地址。
    connection_params: dict[str, Any] = {
        "device_type": str(login_info["device_type"]).strip(),
        "host": str(login_info["ip"]).strip(),
        "username": str(login_info["username"]),
        "password": str(login_info["password"]),
        "conn_timeout": CONNECTION_TIMEOUT,
    }  # 构造 Netmiko 最基本的连接参数。

    port = login_info.get("port", "")  # 读取可选管理端口。
    if str(port).strip():
        connection_params["port"] = int(str(port).strip())  # 如果填写端口，则转换为整数传给 Netmiko。

    secret = login_info.get("secret", "")  # secret 可以为空。
    if str(secret).strip():
        connection_params["secret"] = str(secret)  # 只有实际填写 secret 时才传递该参数。

    return connection_params


def write_device_log(
    run_dir: Path,
    host: str,
    device_type: str,
    command_results: list[tuple[str, str]],
) -> None:
    """
    将单台设备的巡检结果写入独立日志文件。

    command_results 中每个元素分别为：
        command：实际执行的巡检命令。
        output：设备返回结果或超时提示。
    """
    safe_host = host.replace("/", "_").replace("\\", "_")  # 防止 host 中的路径分隔符破坏日志文件路径。
    device_log_path = run_dir / f"{safe_host}.log"  # 每台设备对应一个独立结果文件。

    with device_log_path.open("w", encoding="utf-8") as device_log_file:
        device_log_file.write(
            f"{'=' * 10} Local Time {'=' * 10}\n\n"
        )  # 写入巡检开始时间的标题。
        device_log_file.write(
            f"{time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())}\n\n"
        )  # 写入本次设备巡检的开始时间。
        device_log_file.write(f"Device: {host}\n")
        device_log_file.write(f"Device Type: {device_type}\n\n")

        for command, output in command_results:
            device_log_file.write(
                f"{'=' * 10} {command} {'=' * 10}\n\n"
            )  # 为每条巡检命令增加明显的分隔标题。
            device_log_file.write(f"{output}\n\n")  # 保存设备实际返回结果或超时提示。


def inspection(
    login_info: dict[str, Any],
    cmds_info: dict[str, list[Any]],
    run_dir: Path,
    logger: logging.Logger,
) -> InspectionResult:
    """
    执行单台设备的完整巡检流程。

    流程：
        构造连接参数
        -> 建立 Netmiko 连接
        -> 执行该设备类型对应的巡检命令
        -> 保存设备结果
        -> 返回统一 InspectionResult。
    """
    start_time = time.perf_counter()  # 使用高精度计时器记录单台设备的执行时间。
    host = str(login_info.get("host", "")).strip()  # 获取设备标识，供日志和结果统计使用。
    ssh = None  # 先初始化连接对象，确保 finally 可以安全释放。

    try:
        connection_params = build_connection_params(login_info)  # 将 Excel 数据转换成 Netmiko 参数。

        logger.info("设备 %s 开始连接。", host)

        try:
            # 直接建立 Netmiko 连接。
            # v9 已在 Paramiko Transport._log() 层精确过滤 SSH banner 底层 traceback。
            ssh = ConnectHandler(**connection_params)
        except Exception as error:
            error_type, error_message = classify_connection_error(error)
            raw_message = format_exception_detail(error)

            logger.error(
                "设备 %s 连接失败：%s。",
                host,
                error_message,
            )  # 控制台和日志首先给出直接、易懂的中文故障提示。

            logger.error(
                "设备 %s 原始异常：%s%s",
                host,
                error_type,
                f" - {raw_message}" if raw_message else "",
            )  # 同时保留异常类名和简短原始信息，便于技术排查。

            return InspectionResult(
                host=host,
                success=False,
                duration=time.perf_counter() - start_time,
                error_type=error_type,
                error_message=error_message,
            )

        device_type = str(login_info.get("device_type", "")).strip()
        commands = cmds_info.get(device_type, [])  # 根据设备类型获取对应命令列表。
        command_results: list[tuple[str, str]] = []  # 保存每条命令及其返回结果。

        logger.info("设备 %s 开始巡检。", host)

        for command in commands:
            if not isinstance(command, str):
                continue  # 防御性处理，非字符串内容不是有效巡检命令。

            command = command.strip()
            if not command:
                continue  # 第二个 Sheet 中允许存在空值，用于对齐不同设备类型的命令，因此必须跳过。

            logger.debug("设备 %s 执行命令：%s", host, command)

            try:
                show = ssh.send_command(
                    command,
                    read_timeout=COMMAND_READ_TIMEOUT,
                )  # 执行单条巡检命令，并设置最长读取等待时间。
            except exceptions.ReadTimeout:
                show = f"命令 {command} 执行超时！"  # 将命令超时转换为可读结果，而不是让整台设备任务失败。
                logger.warning("设备 %s 命令 %s 执行超时。", host, command)

            command_results.append((command, show))  # 保存该命令的最终结果。

        write_device_log(
            run_dir,
            host,
            device_type,
            command_results,
        )  # 所有命令完成后一次性写入该设备的巡检结果文件。

        duration = time.perf_counter() - start_time  # 计算单台设备从连接开始到巡检结束的总耗时。
        logger.info(
            "设备 %s 巡检完成，用时 %.1f 秒。",
            host,
            duration,
        )

        return InspectionResult(
            host=host,
            success=True,
            duration=duration,
        )

    except Exception as error:
        error_type = type(error).__name__  # 捕获巡检流程中未预期的异常，并保留真实异常类型。
        raw_message = format_exception_detail(error)

        logger.exception(
            "设备 %s 巡检过程中发生未处理异常：%s%s",
            host,
            error_type,
            f" - {raw_message}" if raw_message else "",
        )  # 对非预期异常记录 traceback 到日志文件，便于开发和故障定位。

        return InspectionResult(
            host=host,
            success=False,
            duration=time.perf_counter() - start_time,
            error_type=error_type,
            error_message="巡检过程中发生未处理异常",
        )

    finally:
        if ssh is not None:
            try:
                ssh.disconnect()  # 无论巡检成功还是失败，都主动关闭 Netmiko 连接释放网络资源。
            except Exception as error:
                logger.warning(
                    "设备 %s 关闭连接时发生异常：%s。",
                    host,
                    format_exception_detail(error),
                )


# ==============================
# 主程序
# ==============================


def main() -> None:
    """
    程序主入口。

    主线程负责：
        文件输入
        -> Excel 读取
        -> 配置校验
        -> 创建运行目录
        -> 初始化日志
        -> 提交并发任务
        -> 收集结果
        -> 输出最终统计。

    这样设备工作线程只负责设备巡检，不负责全局状态管理。
    """
    total_start = time.perf_counter()  # 记录整个程序从开始到最终统计的总耗时。

    install_paramiko_log_filter()  # 在建立设备连接前过滤 Paramiko SSH banner 底层日志。
    install_thread_exception_hook()  # 同时保留线程级异常过滤器，处理可能直接冒泡的 Paramiko 异常。

    info_path = prompt_for_info_file()  # 获取用户确认存在的 Excel 文件路径。

    try:
        devices_info, cmds_info = read_info(info_path)  # 读取设备信息和巡检命令。
    except ExcelSheetError as error:
        print(f"[ERROR] {error}")  # 输出 Excel Sheet 结构错误。
        input("输入 Enter 退出！")  # 给用户时间查看 CMD 回显后再退出。
        return
    except RuntimeError as error:
        print(f"[ERROR] {error}")  # 输出 Excel 读取错误。
        input("输入 Enter 退出！")  # 给用户时间查看 CMD 回显后再退出。
        return

    if not validate_info(devices_info, cmds_info):
        print("\n[ERROR] Excel 配置存在错误，未连接任何设备，程序退出。")
        input("输入 Enter 退出！")
        return

    run_dir = create_run_directory()  # 配置检查通过后才创建本次巡检结果目录。
    logger = setup_logging(run_dir)  # 初始化本次巡检的统一日志系统。

    logger.info("巡检开始。")
    logger.info("设备信息文件：%s", info_path.name)
    logger.info("本次共发现 %d 台设备。", len(devices_info))

    results: list[InspectionResult] = []  # 保存所有设备最终返回的结构化结果。
    future_to_host: dict[Future[InspectionResult], str] = {}  # 建立 Future 与设备名的对应关系，便于异常定位。

    with ThreadPoolExecutor(
        max_workers=MAX_WORKERS,
        thread_name_prefix="Inspection",
    ) as executor:
        for device_info in devices_info:
            future = executor.submit(
                inspection,
                device_info,
                cmds_info,
                run_dir,
                logger,
            )  # 将每台设备作为独立任务提交到线程池。
            future_to_host[future] = str(device_info.get("host", "")).strip()

        for future in as_completed(future_to_host):
            host = future_to_host[future]
            try:
                result = future.result()  # 获取该设备线程最终返回的 InspectionResult。
                results.append(result)
            except Exception as error:
                # 正常情况下 inspection() 已经捕获异常；这里是线程池层面的最后一道保护。
                logger.exception(
                    "设备 %s 的任务线程发生未处理异常：%s",
                    host,
                    format_exception_detail(error),
                )
                results.append(
                    InspectionResult(
                        host=host,
                        success=False,
                        duration=0.0,
                        error_type=type(error).__name__,
                        error_message="任务线程发生未处理异常",
                    )
                )

    failed_results = [
        result for result in results if not result.success
    ]  # 从结构化结果中直接统计失败设备，不再依赖日志文件行数。

    success_count = len(results) - len(failed_results)  # 成功设备数量等于总设备数减去失败设备数。
    total_duration = time.perf_counter() - total_start  # 计算整个批量巡检的总耗时。

    logger.info("<" * 40)  # 通过统一日志系统输出最终分隔线，使 Terminal 与 00_inspection.log 完全一致。

    logger.info(
        "巡检完成，共巡检 %d 台设备，%d 台成功，%d 台异常，共用时 %.1f 秒。",
        len(results),
        success_count,
        len(failed_results),
        total_duration,
    )  # 通过统一日志格式同时输出到控制台和 00_inspection.log。

    input("输入 Enter 退出！")


if __name__ == "__main__":
    main()
