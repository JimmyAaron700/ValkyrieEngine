import pandas as pd
import config


def is_valid_erp_code(code):
    """
    积木块 D：编号安检机
    """
    # 强制转成字符串，并刨掉两头的全部空格、回车等
    code_str = str(code).strip()

    # 开始三大连环安检
    if len(code_str) == 11 and code_str.startswith('D') and code_str[1:].isdigit():
        # 安检通过！返回 True，并且把洗干净的、没有多余空格的编号一起交回去
        return True, code_str
    else:
        # 安检失败，原路打回
        return False, code_str


def load_and_clean_data():
    """
    积木块 E：数据清洗车间 (纯列表极简版)
    """
    file_path = config.SOURCE_EXCEL_PATH
    col_name = config.COLUMN_NAME_CODE

    print(f"\n正在读取并清洗源数据表...")

    try:
        # 直接使用你手册里的绝招1：读取某一列，直接变成一维数组(纯列表)
        df = pd.read_excel(file_path)
        raw_codes = df[col_name].tolist()
    except Exception as e:
        raise Exception(f"读取 Excel 失败！原因：{e}")

    valid_codes = []  # 准备一个空的一维数组，只装洗干净的有效编号
    invalid_count = 0

    for raw_code in raw_codes:
        # 接收安检机返回的两个值：是否合格，以及洗干净的字符串
        is_valid, clean_code = is_valid_erp_code(raw_code)

        if is_valid:
            # 存入干净的编号
            valid_codes.append(clean_code)
        else:
            invalid_count += 1
            print(f"拦截不合规编号：[{raw_code}]，已丢弃。")

    print(f"数据清洗完毕！共读取 {len(raw_codes)} 条，保留有效数据 {len(valid_codes)} 条，剔除 {invalid_count} 条。")

    # 把纯纯的编号列表扔出去
    return valid_codes


def save_data_to_excel(data_list):
    """
    积木块 J：成果打包车间 (导出模块)
    功能：把装满字典的列表，直接转换成 Excel 表格并保存。
    """
    # 如果传进来的列表是空的（比如今天没查任何数据），直接拦住报错
    if not data_list:
        print("⚠️ 警告：没有抓取到任何数据，取消导出 Excel。")
        return

    output_file = config.OUTPUT_EXCEL_NAME
    print(f"\n📦 后勤部：正在将 {len(data_list)} 条记录打包导出到 {output_file} ...")

    try:
        # 绝招：用 Pandas 把“字典列表”瞬间变回“二维表格(DataFrame)”
        # 只要字典里的键（项目编号、项目名称等）是一致的，Pandas 会自动把它们变成 Excel 的表头
        df = pd.DataFrame(data_list)

        # index=False 的意思是：不要把 Pandas 内部自带的 0,1,2,3 行号写进 Excel 里，保持表格干净
        df.to_excel(output_file, index=False)

        print(f"✅ 导出成功！请在项目文件夹下查看 [{output_file}]。")

    except Exception as e:
        print(f"❌ 导出失败！请检查文件是不是已经打开了忘记关，导致程序写不进去。错误信息：{e}")



# ---------------- 输入检测单独测试区 ----------------
# 只有你在当前文件右键 Run 的时候，这里才会执行。
if __name__ == '__main__':
    # 记得先在 config.py 里把 SOURCE_EXCEL_PATH 改成你真实的测试表格路径
    try:
        # 拿一个变量接住清洗好的列表
        result_list = load_and_clean_data()

        # 打印前 5 个看看效果，是不是都是干干净净的字符串
        print("\n🎯 测试结果！提取到的干净列表前 5 个长这样：")
        print(result_list[:5])
    except Exception as error:
        print(error)