"""
ValkyrieEngine 数据 I/O 与预处理模块 (V2.0.0 稳定护航版)
功能：负责与外部 Excel 文件进行交互。包含从源文件读取并清洗项目编号（输入），
以及将内存中的结构化字典列表序列化并导出为最终的报表（输出）。
新增：缺失值 (NaN) 内存安全过滤机制与文件占用时的自动降级保存策略。
"""

import pandas as pd
import config
import time
import os


def is_valid_erp_code(code):
    """
    数据格式校验模块
    功能：对传入的原始业务编号进行字符串格式化与合法性验证。
    """
    # [V2.0.0 健壮性升级] 内存安全拦截：
    # 上游表格中由于合并单元格或错误录入导致的完全空值，在 Pandas 中会被解析为 float 类型的 NaN。
    # 若直接对其调用 str().strip() 会变成字面量 "nan"，从而引发后续逻辑的难以排查的隐患。
    # 此处利用 pd.isna() 进行前置拦截，静默清洗此类脏数据。
    if pd.isna(code):
        return False, ""

    # 数据清洗：将输入值强制转换为字符串类型，并调用 strip() 方法移除两端的空白字符（包含空格、制表符及换行符）。
    # 这一步能有效修复因人工录入表格时误敲空格导致的格式错误。
    code_str = str(code).strip()

    # 复合规则校验：
    # 1. len(code_str) == 11：业务规则要求总长度必须严格为 11 位。
    # 2. code_str.startswith('D')：业务规则要求首字符必须为大写字母 'D'。
    # 3. code_str[1:].isdigit()：使用切片 [1:] 截取掉首位的 'D' 后，剩余的 10 个字符必须全部为纯数字。
    if len(code_str) == 11 and code_str.startswith('D') and code_str[1:].isdigit():
        # 校验通过：返回状态标识 True，以及经过清洗后的标准字符串
        return True, code_str
    else:
        # 校验失败：拦截非法格式数据
        return False, code_str


def load_and_clean_data():
    """
    数据输入与预处理模块
    功能：从配置文件指定的源 Excel 文件中提取目标列数据，逐条经过校验模块过滤，
    最终返回标准化且格式合法的项目编号列表。
    """
    file_path = config.SOURCE_EXCEL_PATH
    col_name = config.COLUMN_NAME_CODE

    print("[数据输入] 正在读取并清洗源始数据表...")

    try:
        # 核心读取逻辑：
        # pd.read_excel(file_path) 将整个 Excel 表格加载进内存。
        # df[col_name] 针对性地提取名为“项目编号”的单列数据。
        # .tolist() 将 Pandas 的列对象 (Series) 转化为 Python 原生的一维列表类型。
        df = pd.read_excel(file_path)
        raw_codes = df[col_name].tolist()
    except Exception as e:
        # 异常捕获：处理诸如文件不存在、路径错误或文件被其他程序（如 WPS）锁定的情况
        raise Exception(f"[系统异常] 数据读取失败，请检查输入文件路径或文件占用状态。错误详情：{e}")

    valid_codes = []  # 初始化空列表，用于存储校验通过的数据
    invalid_count = 0  # 记录拦截的非法数据数量

    # 遍历源始列表中的每一个编号
    for raw_code in raw_codes:
        # 调用格式校验函数，接收其返回的布尔值状态及处理后的字符串
        is_valid, clean_code = is_valid_erp_code(raw_code)

        if is_valid:
            # 存入合规的编号
            valid_codes.append(clean_code)
        else:
            invalid_count += 1
            print(f"[数据清洗] 拦截非标准格式或空值编号：[{raw_code}]，已从队列中静默剔除。")

    print(
        f"[数据输入] 清洗任务完成。源数据共 {len(raw_codes)} 条，保留有效数据 {len(valid_codes)} 条，剔除无效数据 {invalid_count} 条。")

    # 返回纯净的一维数组供后续核心业务模块遍历
    return valid_codes


def save_data_to_excel(data_list):
    """
    数据输出与序列化模块
    功能：将装载着业务处理结果的字典列表转换为结构化的二维数据表，并持久化保存为 Excel 文件。
    """
    # 前置数据校验：如果传入的数据列表为空，则阻断导出流程，防止生成无意义的空文件
    if not data_list:
        print("[数据输出] 警告：输出队列为空，本次运行未生成任何结果文件。")
        return

    output_file = config.OUTPUT_EXCEL_NAME
    print(f"\n[数据输出] 正在将 {len(data_list)} 条业务记录序列化并导出至：{output_file}")

    try:
        # 核心转换逻辑：
        # pd.DataFrame(data_list) 将包含多个相同结构字典的列表映射为标准的二维数据框 (DataFrame)。
        # 字典的键 (Key) 将被自动解析为 Excel 的表头列名，字典的值 (Value) 对应具体的行数据。
        df = pd.DataFrame(data_list)

        # 数据持久化：
        # to_excel() 方法将数据框写入指定路径。
        # 参数 index=False 的作用是：禁止向最终生成的 Excel 中写入 Pandas 内部自动生成的默认行索引 (0, 1, 2...)。
        df.to_excel(output_file, index=False)

        print(f"[数据输出] 导出任务圆满成功，成果文件已保存至：{output_file}")

    except PermissionError:
        # [V2.0.0 健壮性升级] 防丢失降级保存策略：
        # 极高频场景：业务人员在程序运行期间打开了 output.xlsx 查看历史数据且忘记关闭，
        # 导致底层抛出 PermissionError（拒绝访问）。若不予处理，所有辛苦抓取的内存数据将随程序结束而灰飞烟灭。
        # 此处触发降级策略：提取原文件名，在末尾动态追加时间戳后缀生成备份文件。
        fallback_time = time.strftime("%Y%md_%H%M%S")
        base_name, ext = os.path.splitext(output_file)
        fallback_file = f"{base_name}_备份_{fallback_time}{ext}"

        print(f"[系统警报] 目标文件 {output_file} 正被其他程序锁定，写入权限被拒绝！")
        print(f"[数据抢救] 正在触发防丢失降级策略，将数据转存至安全副本：{fallback_file}")

        try:
            df.to_excel(fallback_file, index=False)
            print(f"[数据输出] 抢救成功！请后续在备份文件中查看结算数据。")
        except Exception as e2:
            print(f"[系统异常] 灾难性错误：备份文件同样无法写入。错误详情：{e2}")

    except Exception as e:
        # 捕获其他非权限类的 I/O 异常（如磁盘空间满、路径非法等）
        print(f"[系统异常] 序列化导出发生未预期错误。错误详情：{e}")


# ---------------- 单独测试与调试入口 ----------------
if __name__ == '__main__':
    print("================ 模块独立调试：数据 I/O ================")

    try:
        # 接收预处理后的数据列表
        result_list = load_and_clean_data()

        # 切片打印前 5 个元素，供开发者核对数据清洗效果
        print("\n[调试输出] 数据预处理结果抽样 (前 5 条有效记录)：")
        print(result_list[:5])
    except Exception as error:
        print(error)