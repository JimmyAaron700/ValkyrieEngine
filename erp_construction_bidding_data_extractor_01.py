"""
ValkyrieEngine 工程维度中标查询-数据提取模块 (V2.2.6 完美对齐版)
文件名：erp_construction_bidding_data_extractor_01.py

【版本更新说明】
- V2.2.6: [Excel 列序调整]
          将“总状态”和“工程数”移动至父级信息的末尾，
          确保前几列（编号、名称、金额）与功能1（项目维度）的表格结构保持一致，
          方便后续直接复制粘贴合并。

【架构设计综述】
本模块是 V2.2.0 版本及其后续的核心引擎，负责解决“工程维度”数据的深层抓取难题。
相较于功能1的项目维度，本模块面临三大挑战：
1. 【数据结构复杂】：需要将 1 个父项目和 5 个子工程的数据拍扁在同一行（宽表结构）。
2. 【DOM 陷阱多】：ERP 前端存在 span.val, xformflag 等多种嵌套方式，必须“穿透”提取。
3. 【版本不统一】：历史数据（老版本）和新数据（新版本）字段不一致，需动态识别。

【核心特性】
- 🛡️ 智能熔断：利用 `known_count`（工程数）精准控制搜索次数，绝不浪费一次 HTTP 请求。
- 🕵️ 深度挖掘：内置 `get_deep_text` 穿透器，无视前端嵌套层级。
- 🚑 异常熔断：子工程任何一个报错，总状态立即标记为“需复核”，实现一票否决。
- 💾 实时落地：每处理完一条，立即存入 Excel，确保数据资产零风险。
- 📟 实时监控：终端全字段、高精度透视输出，所见即所得。
"""

import time
import erp_construction_bidding_01  # 导入环境导航模块，用于“回城卷轴”自愈
import data_excel  # 导入数据 I/O 模块，用于“实时存档”

# =========================================================
# 🛠️ 基础工具区：数值清洗与字典初始化
# =========================================================

def parse_money(text):
    """
    [数据清洗] 金额标准化工具
    -------------------------------------------------------
    原理：ERP 系统导出的金额通常是 "1,316,300.00" 这种带千分位逗号的字符串。
    本函数负责将这些“脏数据”洗成干净的 float 类型。
    """
    if not text:
        return 0.0
    try:
        # 1. 移除逗号 ',' 2. 移除首尾空格 strip() 3. 强转 float
        clean_str = str(text).replace(',', '').strip()
        return float(clean_str)
    except:
        return 0.0

def get_mega_record_template(code, known_count):
    """
    [数据结构] 超级宽表模板生成器
    -------------------------------------------------------
    设计哲学：为了保证 Pandas 导出 Excel 时列名绝对对齐，
    我们在循环开始前必须先生成一个“全字段、带默认值”的字典。

    【V2.2.6 调整说明】：
    调整了字典 Key 的插入顺序。
    现在：编号 -> 名称 -> 各类金额 -> 总状态 -> 工程数 -> 子工程...
    目的：与功能1的表头对齐，方便合并。
    """
    record = {
        # --- [父级] 项目维度汇总信息 ---
        "项目编号": code,
        "项目名称": "",      # 将从任一子工程中回填
        "项目工程总造价(元)": 0.0,
        "市政道路修复费": 0.0,
        "小区道路修复费": 0.0,
        "绿化修复费": 0.0,
        "发包金额": 0.0,
        "打捆招标名称": "",
        "项目中标金额": 0.0,

        # [V2.2.6] 移动到父级末尾，方便 Excel 对齐
        "总状态": "初始化",
        "工程数": known_count,
    }

    # --- [子级] 工程维度详情 (儿子 _01 到 _05) ---
    for i in range(1, 6):
        suffix = f"_{i:02d}"  # 生成 _01, _02 ...
        record[f"工程名称{suffix}"] = ""
        record[f"工程造价(元){suffix}"] = 0.0
        record[f"市政道路修复费{suffix}"] = 0.0
        record[f"小区道路修复费{suffix}"] = 0.0
        record[f"绿化修复费{suffix}"] = 0.0
        record[f"发包金额{suffix}"] = 0.0
        record[f"打捆招标名称{suffix}"] = ""
        record[f"中标金额{suffix}"] = 0.0
        record[f"状态{suffix}"] = "初始化"

    return record


# =========================================================
# 🕵️ DOM 深度挖掘工具区 (核心黑科技)
# =========================================================

def get_deep_text(value_td):
    """
    [核心工具] DOM 穿透提取器
    -------------------------------------------------------
    痛点：ERP 系统的前端代码极不规范，数据可能藏在 span.val, xformflag, input 等位置。
    本函数像一个“钻地机”，层层向下嗅探，直到挖出数据。
    """
    try:
        # 策略 A: 优先匹配 .val (工程造价特征)
        val_span = value_td.ele('.val', timeout=0.1)
        if val_span:
            return val_span.text.strip()

        # 策略 B: 匹配 xformflag (通用表单特征)
        xform = value_td.ele('tag:xformflag', timeout=0.1)
        if xform:
            return xform.text.strip()

        # 策略 C: 兜底逻辑，直接取 TD 的直属文本
        return value_td.text.strip()
    except:
        return ""


def extract_field_by_label(tab, label_keywords):
    """
    [核心工具] 模糊标签定位器 (Relative Positioning)
    -------------------------------------------------------
    原理：模拟人类视觉逻辑——“先找到表头(Label)，再找它右边那个格子(Data)”。
    """
    for label in label_keywords:
        try:
            # 1. 定位表头 Label (精确匹配文本)
            label_ele = tab.ele(f'tag:label@@text():{label}', timeout=0.5)

            if label_ele:
                # 2. 相对定位：找父级 TD 的下一个兄弟 TD
                data_td = label_ele.parent().next('tag:td', timeout=0.5)
                if data_td:
                    # 3. 调用穿透器提取数据
                    return get_deep_text(data_td)
        except:
            continue
    return ""


# =========================================================
# 📄 页面动作区：详情提取与版本判断
# =========================================================

def extract_detail_data(detail_tab, suffix):
    """
    [业务逻辑] 单个详情页的数据提取与版本判定
    """
    # 【重要】显式等待 2 秒，确保详情页 DOM 树完全渲染
    detail_tab.wait(2)

    result = {}

    # 1. [版本判定锚点] 寻找“绿化修复费”
    is_new_version = detail_tab.ele('tag:label@@text():绿化修复费', timeout=2)

    if is_new_version:
        status_text = "新版本工程维度发包"
        # 定义新版本字段映射表
        fields_map = {
            "工程名称": ["工程名称"],
            "工程造价(元)": ["工程造价", "工程造价(元)", "工程造价（元）"],
            "市政道路修复费": ["市政道路修复费"],
            "小区道路修复费": ["小区道路修复费"],
            "绿化修复费": ["绿化修复费"],
            "发包金额": ["发包金额"],
            "打捆招标名称": ["打捆招标名称"],
            "中标金额": ["中标金额", "工程中标价", "工程中标价(元)"]
        }
    else:
        status_text = "老版本工程维度发包"
        # 老版本字段较少
        fields_map = {
            "工程名称": ["工程名称"],
            "工程造价(元)": ["工程造价", "工程造价(元)", "工程造价（元）"],
            "打捆招标名称": ["打捆招标名称"],
            "中标金额": ["中标金额", "工程中标价", "工程中标价(元)"]
        }

    result[f"状态{suffix}"] = status_text

    # 2. [批量抓取] 遍历映射表
    for inner_key, label_list in fields_map.items():
        val = extract_field_by_label(detail_tab, label_list)
        result[f"{inner_key}{suffix}"] = val

    # 3. [特别提取] 项目名称 (用于填充父级)
    project_name = extract_field_by_label(detail_tab, ["项目名称"])
    if project_name:
        result["_TEMP_PROJECT_NAME"] = project_name

    return result


def search_and_process_suffix(page, search_tab, code, i, mega_record):
    """
    [流程控制] 单个后缀 (如 _02) 的搜索、点击、提取全流程
    """
    suffix = f"_{i:02d}"
    full_code = f"{code}{suffix}"

    # 1. [UI 清理] 清除输入框里的残留标签
    old_tag = search_tab.ele('text:主题:', timeout=1)
    if old_tag:
        try:
            old_tag.parent().ele('@class=cancel').click()
        except:
            old_tag.next().click()
        search_tab.wait(1)

    # 2. [输入检索]
    search_box = search_tab.ele('@data-lui-placeholder=请输入主题', timeout=5)
    if not search_box:
        search_box = search_tab.ele('@placeholder=请输入主题', timeout=5)

    search_box.clear().input(f'{full_code}\n')
    search_tab.wait(4) # 等待列表刷新

    # 3. [结果判定] 唯一性校验
    target_pattern = f"{full_code}-"
    target_ele = search_tab.ele(f'text:{target_pattern}', timeout=2)

    if not target_ele:
        # [实时监控] 打印未命中状态
        print(f"  -> [{suffix}] 状态: 未发包/项目维度发包")
        mega_record[f"状态{suffix}"] = "未发包/项目维度发包"
        return

    # 4. [进入详情]
    target_ele.click()
    detail_tab = page.latest_tab

    try:
        # 5. [提取数据]
        sub_data = extract_detail_data(detail_tab, suffix)

        # 6. [数据回填]
        for k, v in sub_data.items():
            if k == "_TEMP_PROJECT_NAME":
                if not mega_record["项目名称"]:
                    mega_record["项目名称"] = v
            else:
                mega_record[k] = v

        # 7. [实时全字段监控] (精度控制)
        c_status = mega_record.get(f"状态{suffix}", "N/A")
        c_name = mega_record.get(f"工程名称{suffix}", "")
        c_bundle = mega_record.get(f"打捆招标名称{suffix}", "")

        # 金额类：取出并清洗，以便打印时格式化
        c_bid = parse_money(mega_record.get(f"中标金额{suffix}", 0))
        c_cost = parse_money(mega_record.get(f"工程造价(元){suffix}", 0))
        c_muni = parse_money(mega_record.get(f"市政道路修复费{suffix}", 0))
        c_comm = parse_money(mega_record.get(f"小区道路修复费{suffix}", 0))
        c_green = parse_money(mega_record.get(f"绿化修复费{suffix}", 0))
        c_contract = parse_money(mega_record.get(f"发包金额{suffix}", 0))

        print(f"  -> [{suffix}] 提取成功 | 状态: {c_status}")
        print(f"      工程名称: {c_name}")
        print(f"      标段名称: {c_bundle}")
        print(f"      中标金额: {c_bid:.2f} | 工程造价: {c_cost:.2f}")
        print(f"      市政修复: {c_muni:.2f} | 小区修复: {c_comm:.2f}")
        print(f"      绿化修复: {c_green:.2f} | 发包金额: {c_contract:.2f}")

    except Exception as e:
        print(f"  -> [{suffix}] 数据提取异常: {e}")
        mega_record[f"状态{suffix}"] = "提取异常(需检查)"
    finally:
        detail_tab.close()


# =========================================================
# 🚀 主控循环区
# =========================================================

def run_data_cycle(page, search_tab, enriched_data, output_file):
    """
    [总控制器] 批量数据检索主循环
    """
    total = len(enriched_data)
    all_results = []

    for index, item in enumerate(enriched_data, start=1):
        code = item.get("项目编号")
        known_count = item.get("工程数", 3)

        print(f"\n[任务进度 {index}/{total}] 处理项目: {code} (已知工程数: {known_count})")

        mega_record = get_mega_record_template(code, known_count)

        # --- 内部循环：处理 _01 到 _05 ---
        for i in range(1, 6):
            suffix = f"_{i:02d}"

            # 【逻辑分支 1】智能熔断
            if i > known_count:
                mega_record[f"状态{suffix}"] = "无此工程"
                continue

            # 【逻辑分支 2】搜索与提取
            try:
                search_and_process_suffix(page, search_tab, code, i, mega_record)
            except Exception as e:
                # 【严重异常处理】
                print(f"  -> [{suffix}] 严重错误 (页面卡死): {e}")
                mega_record[f"状态{suffix}"] = "网页卡死失败"

                print("  [自愈程序] 正在执行环境重置...")
                erp_construction_bidding_01.reset_and_back_to_home(page)
                search_tab = erp_construction_bidding_01.setup_search_environment(page)

        # --- 循环结束：执行汇总与总状态判定 ---
        print(f"  [数据汇总] 正在聚合数据并判定总状态...")

        try:
            # 1. 判定总状态 (风控核心：一票否决制)
            has_error = False
            for i in range(1, 6):
                s = mega_record.get(f"状态_{i:02d}", "")
                if "异常" in s or "失败" in s or "卡死" in s:
                    has_error = True
                    break

            if has_error:
                mega_record["总状态"] = "数据提取异常(需复核)"
            else:
                mega_record["总状态"] = mega_record.get("状态_01", "未知")

            # 2. 金额汇总
            sum_mapping = [
                ("项目工程总造价(元)", "工程造价(元)"),
                ("市政道路修复费", "市政道路修复费"),
                ("小区道路修复费", "小区道路修复费"),
                ("绿化修复费", "绿化修复费"),
                ("发包金额", "发包金额"),
                ("项目中标金额", "中标金额")
            ]

            for parent_key, child_prefix in sum_mapping:
                total_val = 0.0
                for i in range(1, 6):
                    child_key = f"{child_prefix}_{i:02d}"
                    val = parse_money(mega_record.get(child_key))
                    total_val += val
                mega_record[parent_key] = total_val

            # 3. 补充信息回填
            for i in range(1, 6):
                val = mega_record.get(f"打捆招标名称_{i:02d}")
                if val:
                    mega_record["打捆招标名称"] = val
                    break

            if not mega_record["项目名称"]:
                mega_record["项目名称"] = "名称提取失败或未发包"

            # [实时全字段监控] (父级)
            print("-" * 50)
            print(f"  [父级汇总] {code} 结算完毕")
            print(f"      总状态  : {mega_record['总状态']}")
            print(f"      工程数量: {known_count}")
            print(f"      项目名称: {mega_record['项目名称']}")
            print(f"      标段名称: {mega_record['打捆招标名称']}")
            print(f"      总中标额: {mega_record['项目中标金额']:.2f}")
            print(f"      总造价  : {mega_record['项目工程总造价(元)']:.2f}")
            print(f"      市政总额: {mega_record['市政道路修复费']:.2f}")
            print(f"      小区总额: {mega_record['小区道路修复费']:.2f}")
            print(f"      绿化总额: {mega_record['绿化修复费']:.2f}")
            print(f"      发包总额: {mega_record['发包金额']:.2f}")
            print("-" * 50)

        except Exception as agg_error:
            print(f"  [汇总异常] 数据聚合计算时发生错误: {agg_error}")
            mega_record["总状态"] = "汇总计算异常"

        all_results.append(mega_record)

        # 【实时存档】
        data_excel.save_data_to_excel(all_results, output_file)

    return all_results