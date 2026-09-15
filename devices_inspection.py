#!/usr/bin/env python3
# -*- coding: UTF-8 -*-

"""
devices_inspection.py —— 网络设备自动化巡检脚本

功能简介：
    本脚本通过读取 Excel 格式的 info 配置文件，自动批量登录网络设备并执行预设
    的巡检命令，将巡检结果保存为本地日志文件，适用于日常网络运维的自动化巡检场景。

主要功能：
    - 支持读取加密/未加密的 Excel 配置文件（info.xlsx），获取设备登录信息与巡检命令；
    - 通过 Netmiko 库以 SSH/Telnet 方式登录设备，支持多种设备类型；
    - 采用多线程并发巡检，最大并发线程数为 200（可配置），显著提升巡检效率；
    - 自动按当天日期创建结果目录，每台设备生成独立的巡检日志文件；
    - 对登录失败（超时、认证失败、协议错误等）的设备统一记录至 01log.log 文件；
    - 巡检完成后 CLI 输出汇总信息，包含巡检设备总数、异常数及总耗时。

依赖库：
    netmiko, pandas, msoffcrypto, openpyxl

使用方式：
    1. 在脚本同目录下准备 info.xlsx（Sheet1：设备登录信息，Sheet2：巡检命令）；
    2. 直接运行脚本：python devices_inspection.py；
    3. 按提示输入 info 文件名（默认为 info.xlsx）；
    4. 巡检结果将保存至当天日期命名的文件夹中。

         Author: Robin
  Creation Date: 2023-12-25
"""

import os
import sys
import time
import getpass
import threading
import msoffcrypto
import pandas as pd
from io import BytesIO
from netmiko import ConnectHandler
from netmiko import exceptions
from contextlib import contextmanager
from pathlib import Path

DEFAULT_INFO_NAME = "info.xlsx"  # info 文件默认文件名
INFO_EXTENSION = ".xlsx"  # nfo 文件扩展名
SCRIPT_DIR = Path(__file__).resolve().parent  # 获取当前 Python 脚本所在目录
LOCAL_TIME = time.strftime('%Y-%m-%d_%H-%M-%S', time.localtime())  # 读取当前日期和时间
LOCK = threading.Lock()  # 线程锁实例化
POOL = threading.BoundedSemaphore(200)  # 最大线程控制


# 自定义异常类，用于处理输入密码为None情况
class PasswordRequiredError(Exception):
    """文件受密码保护，必须提供密码"""
    pass


@contextmanager
def suppress_stderr():
    """
    临时屏蔽 stderr 输出（仅用于抑制 Paramiko 在 SSH 连接失败时
    输出的 'Error reading SSH protocol banner' 等底层 traceback）。

    说明：
    - Paramiko 的 Transport 在独立线程中会直接向 stderr 打印异常，
      try/except 无法捕获这些输出。
    - 本方法仅用于“连接 / enable 阶段”，不影响异常本身的抛出。
    - Netmiko 的异常类型和业务逻辑仍然正常生效。
    - 禁止将此上下文扩大到命令执行阶段，否则会影响调试。
    """
    with open(os.devnull, 'w') as devnull:
        old_stderr = sys.stderr
        sys.stderr = devnull
        try:
            yield
        finally:
            sys.stderr = old_stderr


# 规范化文件名
def normalize_filename(
        user_input: str,
        default_name: str,
        extension: str,
) -> str:
    """
    规范化用户输入的文件名。

    支持：
        带后缀：info.xlsx
        不带后缀：info

    如果直接回车，使用默认文件名；
    如果用户已经输入正确扩展名，则不重复添加扩展名。
    """
    value = user_input.strip()  # 删除用户输入文件名前后的空白字符

    if not value:  # 判断用户是否直接按下Enter
        return default_name  # 空输入时返回默认文件名

    if value.lower().endswith(extension.lower()):  # 判断是否已经包含正确的扩展名
        return value  # 已包含扩展名时直接返回

    return f"{value}{extension}"  # 未包含扩展名时自动追加扩展名


def prompt_for_info_file() -> Path:
    """
    让用户输入 info 文件名，并检查文件是否存在。

    用户可以输入：
        info
        info.xlsx

    如果输入的文件不存在，则重新提示用户输入文件名。
    """
    while True:  # 持续提示用户，直到找到有效的 info 文件
        info_input = input(  # 获取用户输入的 info 文件名
            f"\n请输入info文件名（默认：{DEFAULT_INFO_NAME}）："
        )

        info_name = normalize_filename(  # 规范化用户输入的文件名
            info_input,
            DEFAULT_INFO_NAME,
            INFO_EXTENSION,
        )

        info_path = SCRIPT_DIR / info_name  # 将脚本目录与规范化后的文件名拼接

        if info_path.is_file():  # 检查 info 文件是否真实存在
            return info_path  # 返回有效的 info 文件路径

        print(f"\n[ERROR] info文件不存在：{info_path}")  # 提示用户文件不存在
        print("请重新输入。")  # 提示用户重新输入文件名




# 判断info文件是否被加密，使用不同的读取方式
def read_info(info_path: Path):
    if is_encrypted(info_path):
        return read_encrypted_file(info_path)  # 读取被加密info文件
    else:
        return read_unencrypted_file(info_path)  # 读取未加密info文件


# 检测info文件是否被加密
def is_encrypted(info_file: str) -> bool:
    try:
        with open(info_file, "rb") as f:
            return msoffcrypto.OfficeFile(f).is_encrypted()  # 检测info文件是否被加密
    except Exception:
        return False


# 读取被加密info文件
def read_encrypted_file(info_file: str, max_retry: int = 3) -> pd.DataFrame:
    retry_count = 0  # 初始化重试计数器，用于记录用户尝试输入密码的次数
    while retry_count < max_retry:  # 当重试次数小于最大允许重试次数时，继续循环
        try:
            password = getpass.getpass("\ninfo文件被加密，请输入密码：") or None  # 提示用户输入密码，隐式输入。如果用户直接按Enter键，password将为None
            if not password:  # 如果用户没有输入密码
                raise PasswordRequiredError("文件受密码保护，必须提供密码！")  # 抛出自定义异常，提示用户必须提供密码

            # 解密文件
            decrypted_data = BytesIO()  # 创建一个BytesIO对象，用于在内存中存储解密后的文件内容
            # BytesIO是一个内存中的二进制流，可以像文件一样进行读写操作
            with open(info_file, "rb") as f:  # 以二进制只读模式打开加密的info文件
                office_file = msoffcrypto.OfficeFile(f)  # 使用msoffcrypto库创建一个OfficeFile对象，表示加密的Office文件
                office_file.load_key(password=password)  # 使用用户提供的密码加载解密密钥
                office_file.decrypt(decrypted_data)  # 解密文件内容，并将解密后的数据写入decrypted_data对象中
            decrypted_data.seek(0)  # 将decrypted_data的指针重置到起始位置，以便后续读取操作
            # 由于解密后的数据已经写入decrypted_data，需要将指针重置到开头，以便后续读取

            # 读取解密后的文件
            devices_dataframe = pd.read_excel(decrypted_data, sheet_name=0, dtype=str, keep_default_na=False)
            cmds_dataframe = pd.read_excel(decrypted_data, sheet_name=1, dtype=str)

        except FileNotFoundError:  # 如果没有配置info文件或info文件名错误
            print(f'\n没有找到info文件！\n')  # 提示用户没有找到info文件或info文件名错误
            input('输入Enter退出！')  # 提示用户按Enter键退出
            sys.exit(1)  # 异常退出
        except ValueError:  # 捕获异常信息
            print(f'\ninfo文件缺失子表格信息！\n')  # 代表info文件缺失子表格信息
            input('输入Enter退出！')  # 提示用户按Enter键退出
            sys.exit(1)  # 异常退出
        except (msoffcrypto.exceptions.InvalidKeyError, PasswordRequiredError) as e:
            retry_count += 1
            if retry_count < max_retry:
                print(f"\n密码错误，请重新输入！（剩余尝试次数：{max_retry - retry_count}）")
            else:
                input("\n超过最大尝试次数，输入Enter退出！")
                sys.exit(1)
        except Exception as e:
            print(f"\n解密失败：{str(e)}")
            sys.exit(1)
        else:
            devices_dict = devices_dataframe.to_dict('records')  # 将DataFrame转换成字典
            # "records"参数规定外层为列表，内层以列标题为key，以此列的行内容为value的字典
            # 若有多列，代表字典内有多个key:value对；若有多行，每行为一个字典

            cmds_dict = cmds_dataframe.to_dict('list')  # 将DataFrame转换成字典
            # "list"参数规定外层为字典，列标题为key，列下所有行内容以list形式为value的字典
            # 若有多列，代表字典内有多个key:value对

            return devices_dict, cmds_dict


# 读取未加密info文件
def read_unencrypted_file(info_file: str) -> pd.DataFrame:
    try:
        devices_dataframe = pd.read_excel(info_file, sheet_name=0, dtype=str, keep_default_na=False)
        cmds_dataframe = pd.read_excel(info_file, sheet_name=1, dtype=str)
    except FileNotFoundError:  # 如果没有配置info文件或info文件名错误
        print(f'\n没有找到info文件！\n')  # 代表没有找到info文件或info文件名错误
        input('输入Enter退出！')  # 提示用户按Enter键退出
        sys.exit(1)  # 异常退出
    except ValueError:  # 捕获异常信息
        print(f'\ninfo文件缺失子表格信息！\n')  # 代表info文件缺失子表格信息
        input('输入Enter退出！')  # 提示用户按Enter键退出
        sys.exit(1)  # 异常退出
    else:
        devices_dict = devices_dataframe.to_dict('records')  # 将DataFrame转换成字典
        # "records"参数规定外层为列表，内层以列标题为key，以此列的行内容为value的字典
        # 若有多列，代表字典内有多个key:value对；若有多行，每行为一个字典

        cmds_dict = cmds_dataframe.to_dict('list')  # 将DataFrame转换成字典
        # "list"参数规定外层为字典，列标题为key，列下所有行内容以list形式为value的字典
        # 若有多列，代表字典内有多个key:value对

        return devices_dict, cmds_dict


# 巡检
def inspection(login_info, cmds_dict):
    # 使用传入的设备登录信息和巡检命令，登录设备依次输入巡检命令，如果设备登录出现异常，生成01log文件记录。
    t11 = time.time()  # 子线程执行计时起始点
    ssh = None  # 初始化ssh对象

    try:  # 尝试登录设备
        # 仅在连接阶段使用 suppress_stderr，
        # 用于屏蔽 Paramiko 在目标不可达/非 SSH 服务时输出的噪音 traceback
        with suppress_stderr():
            ssh = ConnectHandler(**login_info)  # 使用设备登录信息，SSH登录设备
            # ssh.enable()  # 进入设备Enable模式
    except Exception as ssh_error:  # 登录设备出现异常
        with LOCK:  # 线程锁
            match type(ssh_error).__name__:  # 匹配异常名称
                case 'AttributeError':  # 异常名称为：AttributeError
                    print(f'设备 {login_info["host"]} 缺少设备管理地址！')  # CMD输出提示信息
                    with open(os.path.join(os.getcwd(), LOCAL_TIME, '01log.log'), 'a', encoding='utf-8') as log:
                        log.write(f'设备 {login_info["host"]} 缺少设备管理地址！\n')  # 记录到log文件
                case 'NetmikoTimeoutException':
                    print(f'设备 {login_info["host"]} 管理地址或端口不可达！')
                    with open(os.path.join(os.getcwd(), LOCAL_TIME, '01log.log'), 'a', encoding='utf-8') as log:
                        log.write(f'设备 {login_info["host"]} 管理地址或端口不可达！\n')
                case 'NetmikoAuthenticationException':
                    print(f'设备 {login_info["host"]} 用户名或密码认证失败！')
                    with open(os.path.join(os.getcwd(), LOCAL_TIME, '01log.log'), 'a', encoding='utf-8') as log:
                        log.write(f'设备 {login_info["host"]} 用户名或密码认证失败！\n')
                case 'ValueError':
                    print(f'设备 {login_info["host"]} Enable密码认证失败！')
                    with open(os.path.join(os.getcwd(), LOCAL_TIME, '01log.log'), 'a', encoding='utf-8') as log:
                        log.write(f'设备 {login_info["host"]} Enable密码认证失败！\n')
                case 'TimeoutError':
                    print(f'设备 {login_info["host"]} Telnet连接超时！')
                    with open(os.path.join(os.getcwd(), LOCAL_TIME, '01log.log'), 'a', encoding='utf-8') as log:
                        log.write(f'设备 {login_info["host"]} Telnet连接超时！\n')
                case 'ReadTimeout':
                    print(f'设备 {login_info["host"]} Enable密码认证失败！')
                    with open(os.path.join(os.getcwd(), LOCAL_TIME, '01log.log'), 'a', encoding='utf-8') as log:
                        log.write(f'设备 {login_info["host"]} Enable密码认证失败！\n')
                case 'ConnectionRefusedError':
                    print(f'设备 {login_info["host"]} 远程登录协议错误！')
                    with open(os.path.join(os.getcwd(), LOCAL_TIME, '01log.log'), 'a', encoding='utf-8') as log:
                        log.write(f'设备 {login_info["host"]} 远程登录协议错误！\n')
                case _:
                    print(f'设备 {login_info["host"]} 未知错误！{type(ssh_error).__name__}: {str(ssh_error)}')
                    with open(os.path.join(os.getcwd(), LOCAL_TIME, '01log.log'), 'a', encoding='utf-8') as log:
                        log.write(f'设备 {login_info["host"]} 未知错误！{type(ssh_error).__name__}: {str(ssh_error)}\n')
    else:  # 如果登录正常，开始巡检
        with open(os.path.join(os.getcwd(), LOCAL_TIME, login_info['host'] + '.log'), 'w', encoding='utf-8') as device_log_file:
            # 创建当前设备的巡检信息记录文件
            with LOCK:  # 线程锁
                print(f'设备 {login_info["host"]} 正在巡检...')  # 打印当前设备正在巡检提示信息
            device_log_file.write('=' * 10 + ' ' + 'Local Time' + ' ' + '=' * 10 + '\n\n')  # 写入当前巡检时间
            device_log_file.write(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + '\n\n')  # 写入当前巡检时间
            for cmd in cmds_dict[login_info['device_type']]:  # 从所有设备类型巡检命令中找到与当前设备类型匹配的命令列表，遍历所有巡检命令
                if type(cmd) is str:  # 判断读取的命令是否为字符串
                    device_log_file.write('=' * 10 + ' ' + cmd + ' ' + '=' * 10 + '\n\n')  # 写入当前巡检命令分行符，至巡检信息记录文件
                    try:  # 尝试执行当前巡检命令，获取结果，并设置最长等待时间
                        show = ssh.send_command(cmd, read_timeout=120)
                    except exceptions.ReadTimeout:  # 如果等待时间依然超时，捕获异常并提示、记录
                        print(f'设备 {login_info["host"]} 命令 {cmd} 执行超时！')  # cmd输出命令执行超时提示
                        show = f'命令 {cmd} 执行超时！'  # 赋值结果，在巡检记录log文件中提示此命令执行超时
                    finally:  # 最终将结果写入巡检信息记录文件
                        device_log_file.write(show + '\n\n')  # 写入当前巡检命令的结果，至巡检信息记录文件
        t12 = time.time()  # 子线程执行计时结束点
        with LOCK:  # 线程锁
            print(f'设备 {login_info["host"]} 巡检完成，用时 {round(t12 - t11, 1)} 秒。')  # 打印子线程执行时长
    finally:  # 最后结束SSH连接释放线程
        if ssh is not None:  # 判断ssh对象是否被正确赋值，赋值成功不为None，即SSH连接已建立，需要关闭连接
            ssh.disconnect()  # 关闭SSH连接
        POOL.release()  # 最大线程限制，释放一个线程


if __name__ == '__main__':
    t1 = time.time()  # 程序执行计时起始点
    threading_list = []  # 创建一个线程列表，准备存放所有线程
    INFO_PATH = prompt_for_info_file()  # 在函数定义完成后获取并验证 info 文件路径
    devices_info, cmds_info = read_info(INFO_PATH)  # 读取info文件，获取设备登录信息和命令信息

    print(f'\n巡检开始...')  # 提示巡检开始
    print(f'\n' + '>' * 40 + '\n')  # 打印一行“>”，隔开巡检提示信息

    if not os.path.exists(LOCAL_TIME):  # 判断是否存在当前日期的文件夹，判断当天是否执行过巡检
        os.makedirs(LOCAL_TIME)  # 如果没有，创建当天日期文件夹
    else:  # 如果有，则清空文件夹内所有文件
        for filename in os.listdir(LOCAL_TIME):  # 遍历当前日期文件夹内所有文件
            file_path = os.path.join(LOCAL_TIME, filename)  # 构建当前文件的完整路径
            os.remove(file_path)  # 删除当前文件

    for device_info in devices_info:  # 遍历所有设备登录信息
        updated_device_info = device_info.copy()  # 创建一个更新后的设备登录信息字典，用于传参
        updated_device_info["conn_timeout"] = 40  # 更新设备登录信息字典，设置TCP连接超时时间
        pre_device = threading.Thread(target=inspection, args=(updated_device_info, cmds_info), name=device_info['host'] + '_Thread')
        # 创建一个线程，执行inspection函数，传入当前遍历的设备登录信息和所有设备类型巡检命令，并定义线程名称
        threading_list.append(pre_device)  # 将当前创建的线程追加进线程列表
        POOL.acquire()  # 从最大线程限制，获取一个线程令牌
        pre_device.start()  # 开启这个线程

    for _ in threading_list:  # 遍历所有创建的线程
        _.join()  # 等待所有线程的结束

    try:  # 尝试打开01log文件
        with open(os.path.join(os.getcwd(), LOCAL_TIME, '01log.log'), 'r', encoding='utf-8') as log_file:
            file_lines = len(log_file.readlines())  # 读取01log文件共有多少行。有多少行，代表出现了多少个设备登录异常
    except FileNotFoundError:  # 如果找不到01log文件
        file_lines = 0  # 证明本次巡检没有出现巡检异常情况
    t2 = time.time()  # 程序执行计时结束点
    print(f'\n' + '<' * 40 + '\n')  # 打印一行“<”，隔开巡检报告信息
    print(f'巡检完成，共巡检 {len(threading_list)} 台设备，{file_lines} 台异常，共用时 {round(t2 - t1, 1)} 秒。\n')  # 打印巡检报告
    input('输入Enter退出！')  # 提示用户按Enter键退出
