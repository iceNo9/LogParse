import json
import os
import re
import csv
import sys
import logging
from enum import Enum, auto
from typing import List
import tkinter as tk
from tkinter import messagebox
from pathlib import Path
from datetime import datetime
import pandas as pd
import chardet

# 获取程序运行时的工作目录
if getattr(sys, 'frozen', False):  # 检查是否为打包环境
    executable_path = sys.executable
else:
    # 非打包环境，获取当前脚本的路径
    executable_path = __file__

# 获取可执行文件所在的目录
work_dir = Path(executable_path).parent.resolve()

# 读取当前路径下config.json文件
config_file_path = work_dir / 'config.csv'
name_file_path = work_dir / 'name.csv'
log_folder_path = work_dir / 'log'


class CommandType(Enum):
    Invalid = auto()
    Main = auto()
    Sub = auto()


class CommandData:
    def __init__(self, head: str, send: list[str], return_values: list[str]):
        self.head = head
        self.send = send
        self.return_values = return_values

    def __eq__(self, other):
        if not isinstance(other, CommandData):
            return False
        return self.send == other.send and self.return_values == other.return_values

    def copy(self):
        return CommandData(self.head, self.send.copy(), self.return_values.copy())


# 获取当前日期，格式化为 'yyyymmdd' 格式
current_date = datetime.now().strftime("%Y%m%d")
log_file_path = os.path.join(log_folder_path, f"{current_date}.log")

# 确保 log 文件夹存在，若不存在则创建
if not os.path.exists(log_folder_path):
    os.mkdir(log_folder_path)

# 设置日志记录器
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)  # 设置日志级别，可根据需要调整

# 创建一个文件处理器，指定日志文件路径和模式（追加模式 'a'）
file_handler = logging.FileHandler(log_file_path, mode='a')
file_handler.setLevel(logging.DEBUG)  # 文件处理器的日志级别与记录器相同

# 定义日志输出格式
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s')
file_handler.setFormatter(formatter)

# 将文件处理器添加到日志记录器
logger.addHandler(file_handler)


def create_command(head: str):
    return CommandData(head=head, send=[head], return_values=[])


def judge_command(rv_str):
    # 定义正则模式
    cmd_keyword_pattern = re.compile(r':CMD:')
    hex_format_pattern = re.compile(r'\[0x[a-fA-F0-9]{2}\]->\[0x[a-fA-F0-9]{2}\]')

    # 检查是否包含关键词和格式
    contains_cmd_keyword = bool(cmd_keyword_pattern.search(rv_str))
    hex_match = hex_format_pattern.search(rv_str)  # Save the match object if found

    if contains_cmd_keyword and hex_match:
        return CommandType.Main, hex_match
    elif hex_match:
        return CommandType.Sub, hex_match
    else:
        return CommandType.Invalid, None


def judge_commands(rv_command: CommandData, rv_keys_dict, rv_commands: List[CommandData]):
    for key_list_name in ["matchkey", "changematchkey"]:
        key_list = rv_keys_dict.get(key_list_name)
        if key_list and rv_command.head in key_list:
            if key_list_name == "matchkey":
                rv_commands.append(rv_command)
                repeat_commands[int(rv_command.head, 16)] = rv_command
                break
            elif key_list_name == "changematchkey":
                existing_command = repeat_commands[int(rv_command.head, 16)]
                if not existing_command == rv_command:
                    repeat_commands[int(rv_command.head, 16)] = rv_command
                    rv_commands.append(rv_command)
                    break


def parse_commands(rv_data_lines, rv_keys_dict):
    at_commands = []
    current_main_command = None

    for line in rv_data_lines:
        try:
            command_type, hex_match = judge_command(line)

            if hex_match is not None:
                hex_send, hex_return = hex_match.group().split('->')
                hex_send_value = hex_send[1:-1]  # Remove the outer brackets
                hex_return_value = hex_return[1:-1]  # Remove the outer brackets

                if command_type == CommandType.Main:
                    if current_main_command is not None:
                        judge_commands(current_main_command, rv_keys_dict, at_commands)
                    current_main_command = CommandData(head=hex_send_value, send=[hex_send_value],
                                                       return_values=[hex_return_value])
                elif command_type == CommandType.Sub:
                    if current_main_command is not None:
                        current_main_command.send.append(hex_send_value)
                        current_main_command.return_values.append(hex_return_value)
                    else:
                        # 处理 `current_main_command` 未初始化的情况
                        logger.warning("当前主命令未初始化，跳过子命令处理")
                        logger.critical(f"源日志:{line}")

        except Exception as e:
            logger.critical(f"发生错误:{e.__class__.__name__}: {str(e)}")
            logger.critical(f"错误源日志:{line}")
            popup_error(f"发生错误，已保存到日志{log_file_path}")
            sys.exit()

    if current_main_command is not None:
        judge_commands(current_main_command, rv_keys_dict, at_commands)

    return at_commands


def write_commands_to_csv(rv_commands, rv_path, rv_name_dict):
    # 根据路径创建或打开CSV文件
    with open(rv_path, 'w', newline='', encoding='utf-8') as output_file:
        writer = csv.writer(output_file)

        # 写入表头
        writer.writerow(["中文名称", "Head", "Send", "Return_Values", "Hex_Format"])

        for command in rv_commands:
            # 获取对应的中文名称
            chinese_name = rv_name_dict.get(command.head)
            if chinese_name is None:
                chinese_name = "该值不存在或者为空，请修改"

            # 构建Hex_Format行
            hex_format_line = f"CMD:[{command.send[0]}]->[{command.return_values[0]}]"
            sub_hex_format_lines = []
            for i in range(1, min(len(command.send), len(command.return_values))):
                sub_hex_format_lines.append(f"\n   :[{command.send[i]}]->[{command.return_values[i]}]")

            # 将子命令行连接起来
            hex_format_content = hex_format_line + ''.join(sub_hex_format_lines).rstrip('\n')

            # 写入CSV行
            writer.writerow([chinese_name, command.head, ','.join(command.send), ','.join(command.return_values),
                             hex_format_content])

    popup_error(f"操作完成")


def write_commands_to_file(rv_commands, output_file_path):
    with open(output_file_path, "w") as output_file:
        for command in rv_commands:
            if len(command.send) > 0 and len(command.return_values) > 0:
                output_file.write(f"CMD:[{command.send[0]}]->[{command.return_values[0]}]\n")

                for i in range(1, min(len(command.send), len(command.return_values))):
                    output_file.write(f"   :[{command.send[i]}]->[{command.return_values[i]}]\n")




def process_input(input_path: str, rv_keys_dict, rv_names_dict, out_subfolder=""):
    if os.path.isfile(input_path):  # 判断输入路径是否为文件
        if input_path.endswith(('.txt', '.log')):  # 判断是否为文本文件或日志文件
            at_commands = parse_commands(read_file_lines(input_path), rv_keys_dict)

            # 创建output子目录（如果需要的话）
            output_root = os.path.join(os.path.dirname(input_path), out_subfolder)
            os.makedirs(output_root, exist_ok=True)

            base_filename = os.path.splitext(os.path.basename(input_path))[0]
            output_file = os.path.join(output_root, f"{base_filename}_parse.csv")
            write_commands_to_csv(at_commands, output_file, rv_names_dict)
        else:
            popup_error(f"输入的不是文本文件：{input_path}, 目前仅支持.txt或.log")
    elif os.path.isdir(input_path):  # 判断输入路径是否为目录
        for root, dirs, files in os.walk(input_path):
            for file in files:
                if file.endswith('.log'):  # 遍历目录及子目录下的文本文件或日志文件
                    txt_file_path = os.path.join(root, file)
                    at_commands = parse_commands(read_file_lines(txt_file_path), rv_keys_dict)

                    # 创建output子目录（如果需要的话）
                    output_root = os.path.join(root, out_subfolder)
                    os.makedirs(output_root, exist_ok=True)

                    base_filename = os.path.splitext(file)[0]
                    output_file = os.path.join(output_root, f"parse_{base_filename}.csv")
                    write_commands_to_csv(at_commands, output_file, rv_names_dict)
    else:
        popup_error(f"输入的既不是文件也不是目录：{input_path}")


def read_file_lines(rv_file_path: str) -> list[str]:
    with open(rv_file_path, 'r') as file:
        return file.readlines()


def popup_error(message: str):
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口

    messagebox.showerror("错误", message)  # 显示错误弹窗
    root.destroy()  # 销毁隐藏的主窗口


# with open("test.log", "r") as file:
#     lines = file.readlines()




default_name_data = {
    "0x01": "查询引擎基本状态",
    "0x02": "查询可修复异常",
    "0x03": "查询机器盖打开详细",
    "0x08": "查询机器卡纸位置详细1",
    "0x09": "查询机器卡纸位置详细2",
    "0x0D": "查询打印错误信息",
    "0x0E": "查询打印等待信息",
    "0x0F": "查询打印等待详细",
    "0x11": "查询引擎省能详细",
    "0x12": "查询引擎复位详细",
    "0x13": "查询警告信息1",
    "0x14": "查询警告信息2",
    "0x1B": "查询耗材要求信息",
    "0x1C": "查询引擎要求信息",
    "0x1D": "查询进纸盒安装情况",
    "0x20": "查询进纸盒变化信息",
    "0x21": "查询标准进纸盒变化详细",
    "0x22": "查询标准进纸盒纸尺寸变化详细",
    "0x23": "查询多功能进纸盒变化详细",
    "0x25": "查询选配进纸盒1变化详细",
    "0x26": "查询选配进纸盒1纸尺寸详细",
    "0x27": "查询选配进纸盒2变化详细",
    "0x28": "查询选配进纸盒2纸尺寸详细",
    "0x30": "查询打印管理",
    "0x31": "查询每一页A面的开始结束信息",
    "0x32": "查询每一页B面的开始结束信息",
    "0x33": "查询每一页的给纸信息",
    "0x37": "查询耗材寿命变化信息",
    "0x3D": "查询执行型参数操作状态",
    "0x40": "查询某些特殊机能是否可以执行",
    "0x41": "查询异常后CTL需补发的页数",
    "0x51": "握手命令",
    "0x52": "打印准备要求",
    "0x53": "指定打印色模式/给/排纸盒",
    "0x54": "指定打印解像度",
    "0x55": "指定打印纸种纸厚",
    "0x56": "指定打印纸尺寸",
    "0x57": "指定打印自定义纸尺寸的长和宽",
    "0x58": "指定尺寸不一致检知功能是否开启",
    "0x59": "指定打印模式",
    "0x5A": "指定画像处理级别",
    "0x5B": "指定打印PPM规格",
    "0x5C": "指定打印品质",
    "0x5E": "指定打印浓度等级",
    "0x65": "打印执行命令",
    "0x66": "指示引擎迁移到省能模式",
    "0x67": "指示引擎退出省能模式",
    "0x68": "指示引擎迁移到诊断模式",
    "0x69": "指示引擎退出诊断模式",
    "0x70": "清零打印错误关联状态位",
    "0x71": "清零卡纸关联状态位",
    "0x72": "清零尺寸不一致关联状态位",
    "0x82": "CTL画像准备完了通知",
    "0x83": "设定CHAR型MP参数值",
    "0x84": "设定SHORT型MP参数值",
    "0x85": "设定LONG型MP参数值",
    "0x86": "设定STRING型MP参数值",
    "0x87": "设定EXECUTION型MP参数值",
    "0x88": "传送机器信息",
    "0x89": "传送机器当前日期",
    "0x8A": "传送Dot值",
    "0x8C": "传送打印机信息",
    "0x99": "要求取消打印",
    "0x9A": "要求通信关闭",
    "0x9B": "预约引擎复位",
    "0x9C": "要求引擎复位",
    "0x9D": "要求引擎色彩校正开始",
    "0x9E": "要求引擎色彩校正结束",
    "0x9F": "要求引擎色彩校正类型实施",
    "0xA0": "要求引擎耗材特殊检测执行",
    "0xA1": "预约引擎耗材特殊检测执行",
    "0xA2": "要求引擎碳粉消耗检测执行",
    "0xB5": "要求引擎重启实施",
    "0xB6": "色彩校正参数设置",
    "0xB7": "耗材数据传输",
    "0xB8": "执行NVRAM数据传输",
    "0xB9": "引擎固件更新",
    "0xBA": "读取CHAR型MP参数值",
    "0xBB": "读取SHORT型MP参数值",
    "0xBC": "读取LONG型MP参数值",
    "0xBD": "读取STRING型MP参数值",
    "0xBE": "",
    "0xBF": "查询部品未安装异常状态信息",
    "0xC0": "",
    "0xC1": "查询耗材即将用尽异常状态信息",
    "0xC2": "查询耗材用尽警告状态信息",
    "0xC3": "",
    "0xC4": "查询校准警告状态信息",
    "0xC5": "查询耗材不匹配异常状态信息",
    "0xC6": "",
    "0xC7": "查询部品未安装异常详细",
    "0xC8": "查询部品不匹配异常详细",
    "0xC9": "查询耗材余量不足警告状态信息",
    "0xCA": "查询耗材剩余量",
    "0xCB": "查询部品使用率",
    "0xCC": "查询耗材寿命信息",
    "0xCD": "查询耗材打印页数信息",
    "0xCE": "查询耗材部品型号/序列号等信息",
    "0xCF": "查询耗材打印页按长度计数的作业数",
    "0xD0": "",
    "0xD1": "",
    "0xD2": "",
    "0xD3": "查询温湿度警告状态信息",
    "0xD4": "查询卡纸异常码",
    "0xD5": "查询不可修复异常码",
    "0xD6": "",
    "0xD7": "查询打印取消结果",
    "0xD8": "查询打印错误发生页面ID",
    "0xD9": "查询卡纸发生页面ID",
    "0xDA": "查询尺寸不一致异常发生页面ID",
    "0xDB": "引擎NVRAM上传",
    "0xDC": "引擎LOG上传",
    "0xDD": "色彩校正后数据提取",
    "0xDE": "查询引擎色彩校正模式",
    "0xDF": "引擎数据导出",
    "0xE0": "引擎耗材数据上传",
    "0xE1": "查询校正类型实施结果",
    "0xE2": "查询耗材特殊检测类型ID",
    "0xE3": "查询耗材特殊检测实施结果",
    "0xE4": "查询耗材变更状态信息",
    "0xE6": "查询碳粉消耗检测实施结果"
}

def detect_encoding(file_path):
    """检测文件的编码格式"""
    with open(file_path, 'rb') as f:
        result = chardet.detect(f.read())
    return result['encoding']

def load_name_csv_to_json(rv_path):
    # 默认数据
    default_data = {
        "指令": ["0x01"],
        "解释": ["查询引擎基本状态"]
    }

    # 如果文件不存在，则创建
    if not os.path.exists(rv_path):
        df = pd.DataFrame(default_data)
        df.to_csv(rv_path, index=False)  # 使用 csv 文件格式

    # 读取文件前检测文件编码
    encoding = detect_encoding(rv_path)

    # 读取文件
    try:
        df = pd.read_csv(rv_path, encoding=encoding)
    except UnicodeDecodeError:
        print("尝试使用 'latin1' 编码...")
        df = pd.read_csv(rv_path, encoding='latin1')

    # 转换为字典
    data_dict = df.set_index('指令')['解释'].to_dict()

    return data_dict

def load_name_csv_to_json_backup(rv_path):
    # 默认数据
    default_data = {
        "指令": ["0x01"],
        "解释": ["查询引擎基本状态"]
    }

    # 如果文件不存在，则创建
    if not os.path.exists(rv_path):
        df = pd.DataFrame(default_data)
        df.to_csv(rv_path, index=False)  # 使用 csv 文件格式

    # 读取文件
    df = pd.read_csv(rv_path)

    # 转换为字典
    data_dict = df.set_index('指令')['解释'].to_dict()

    return data_dict

def load_config_csv_to_json(rv_path):
    # 默认数据
    default_data = {
        'matchkey': ['51', '52', '65', '82'],
        'changematchkey': ['01', '30', '31', '32']
    }

    # 如果文件不存在，则创建
    if not os.path.exists(rv_path):
        df = pd.DataFrame(default_data)
        df.to_csv(rv_path, index=False)  # 使用 csv 文件格式

    # 读取文件
    df = pd.read_csv(rv_path)

    # 处理缺失值
    df['matchkey'] = '0x' + df['matchkey'].fillna('').astype(str)
    df['changematchkey'] = '0x' + df['changematchkey'].fillna('').astype(str)

    # 格式化数字，确保单个数字前面带0
    df['matchkey'] = df['matchkey'].apply(lambda x: x if len(x[2:]) > 1 else x[:2] + '0' + x[2:])
    df['changematchkey'] = df['changematchkey'].apply(lambda x: x if len(x[2:]) > 1 else x[:2] + '0' + x[2:])

    # 转换为字典
    data_dict = {
        "matchkey": df['matchkey'].tolist(),
        "changematchkey": df['changematchkey'].tolist()
    }

    return data_dict


def json_to_name_csv(rv_dict, output_path):
    # 将 JSON 转换为 DataFrame
    df = pd.DataFrame(list(rv_dict.items()), columns=['指令', '解释'])

    # 将 DataFrame 保存为 CSV 文件
    df.to_csv(output_path, index=False)

def load_config_xlsx_to_json(rv_path):
    # 默认数据
    default_data = {
        'matchkey': ['51', '52', '65', '82'],
        'changematchkey': ['01', '30', '31', '32']
    }

    # 如果文件不存在，则创建
    if not os.path.exists(rv_path):
        df = pd.DataFrame(default_data)
        df.to_excel(rv_path, index=False, engine='openpyxl')  # 指定使用openpyxl作为引擎

    # 读取文件
    df = pd.read_excel(rv_path, engine='openpyxl')  # 指定使用openpyxl作为引擎

    # 处理缺失值
    df['matchkey'] = '0x' + df['matchkey'].fillna('').astype(str)
    df['changematchkey'] = '0x' + df['changematchkey'].fillna('').astype(str)

    # 格式化数字，确保单个数字前面带0
    df['changematchkey'] = df['changematchkey'].apply(lambda x: x if len(x[2:]) > 1 else x[:2] + '0' + x[2:])

    # 转换为字典
    data_dict = {
        "matchkey": df['matchkey'].tolist(),
        "changematchkey": df['changematchkey'].tolist()
    }

    # 输出JSON
    # json_data = json.dumps(data_dict, indent=4)
    # print(json_data)

    return data_dict

def load_name_xlsx_to_json(rv_path):
    # 默认数据
    default_data = {
        "指令": ["0x01"],
        "解释": ["查询引擎基本状态"]
    }

    # 如果文件不存在，则创建
    if not os.path.exists(rv_path):
        df = pd.DataFrame(default_data)
        df.to_excel(rv_path, index=False, engine='openpyxl')

    # 读取文件
    df = pd.read_excel(rv_path, engine='openpyxl')

    # 转换为字典
    data_dict = df.set_index('指令')['解释'].to_dict()

    # 输出JSON
    # json_data = json.dumps(data_dict, indent=4)
    # print(json_data)

    return data_dict
    
def json_to_name_xlsx(rv_dict):
    # 将 JSON 转换为 DataFrame
    df = pd.DataFrame(list(rv_dict.items()), columns=['指令', '解释'])
    

if __name__ == "__main__":
    # 检查并创建/写入config.json文件
    keys_dict = load_config_csv_to_json(config_file_path)

    # 检查并创建/写入name.json文件,读取
    names_dict = load_name_csv_to_json(name_file_path)

    default_command = CommandData(head="", send=[], return_values=[])
    repeat_commands = [default_command.copy() for _ in range(256)]

    if len(sys.argv) > 1:
        process_input(sys.argv[1], keys_dict, names_dict)

# commands = parse_commands(lines, keys_dict)
# write_commands_to_file(commands, "out.csv")
# write_commands_to_csv(commands, "out.csv", names_dict)
