"""
ValkyrieEngine 模块五：项目基础信息查询与数据提取核心
文件名：erp_information_data_extractor.py

【架构说明】
本模块采用“检索-校验-提取-聚合”的线性处理流程，结合多层级异常重试机制。
核心任务是遍历目标项目编号，控制浏览器完成交互，并针对异步加载的页面数据进行稳定抓取。

【核心机制解析】
1. 状态分流逻辑：通过读取页面顶部的“项目数”与“工程数”，动态将任务划分至三种处理分支（Case 1/2/3）。
2. 异步加载校验：通过持续轮询特定的DOM（文档对象模型）元素文本状态，确保在提取前前端数据已完全渲染，避免数据为空或读取脏数据。
3. 梯度容错机制：针对网络延迟或DOM结构卡死，设定最大重试次数。异常发生时，优先尝试局部页面刷新；若局部恢复失效，则触发跨模块的全局环境重置。
4. 数据聚合与后处理：在底层数据提取完成后，执行业务层面的数据清洗（如去除空白项）和去重合并，直接生成可用于最终报表的汇总前置列。
"""

import time
import data_excel
import erp_information

def get_field_mapping():
    """
    DOM元素选择器注册表
    统一管理所需提取字段名称与前端页面ID的映射关系，便于后期页面结构变动时集中维护。
    注：包含辅助校验字段 "#gcbh"（工程编号），用于后续的异步加载状态比对。
    """
    return {
        "工程编号": "#gcbh",
        "项目名称": "#xmmc",
        "项目状态": "#xmztzw",
        "工程名称": "#gcmc",
        "开工日期": "#kgrq",
        "工程状态": "#gcztzw",
        "项目经理": "#xmjl",
        "质监员": "#zjy",
        "施工单位": "#zbdw",
        "监理单位": "#jldw",
        "分包单位": "#fbdw",
        "设计师": "#sjs",
        "项目负责人": "#xmfzr"
    }

def get_information_template(code, status):
    """
    数据结构初始化模块
    创建标准的字典模板，预置“汇总区”与“01-05详细区”的所有键值对。
    作用：确保后续输出到Excel时，无论该项目实际包含几个子工程，生成的列顺序和列数保持绝对一致，避免DataFrame合并时发生错位。
    """
    record = {
        "项目编号": code,
        "工程数": status,
        # 预留用于最终报表展示的聚合字段
        "汇总_项目名称": "",
        "汇总_项目状态": "",
        "汇总_施工单位": "",
        "汇总_分包单位": ""
    }

    # 根据字段注册表，按顺序生成01到05的占位符
    mapping = get_field_mapping()
    for i in range(1, 6):
        suffix = f"_{i:02d}"
        for field in mapping.keys():
            record[f"{field}{suffix}"] = ""

    return record

def wait_for_data_load(tab, suffix_id, exact_suffix=None):
    """
    异步数据渲染校验组件
    用于解决点击列表项后，右侧详情面板数据加载存在延迟的问题。

    参数说明:
        tab: 当前浏览器标签页对象
        suffix_id: 当前处理的工程编号后缀（仅用于日志输出记录进度）
        exact_suffix: 期望出现的单据尾标（如 "_02"）。若传入此参数，则启用严格比对模式。
    """
    print(f"        -> [{suffix_id}] 校验模块启动：轮询监测详情面板数据渲染状态...")
    end_time = time.time() + 15  # 设定全局超时阈值为15秒

    while time.time() < end_time:
        if exact_suffix:
            # 【严格比对模式】适用于多工程切换场景（Case 3）
            # 持续读取工程编号元素，必须确认内容已变更为当前指定的尾标，防止提取到上一次点击的缓存数据
            gcbh_val = tab.ele('#gcbh', timeout=0.1).text.strip()
            if exact_suffix in gcbh_val:
                # 确认数据刷新后，强制等待0.8秒，确保前端的全局透明加载遮罩完全被移除
                tab.wait(0.8)
                print(f"        -> [{suffix_id}] 尾标校验通过: 当前获取编号为 {gcbh_val}")
                return True
        else:
            # 【基础比对模式】适用于单工程加载场景（Case 2）
            # 无需校验尾标，只需确认核心必填字段（项目名称）已不再为空字符串
            name_val = tab.ele('#xmmc', timeout=0.1).text.strip()
            if name_val != "":
                tab.wait(0.8)
                print(f"        -> [{suffix_id}] 数据渲染完成: 首字段获取内容为 {name_val[:10]}...")
                return True

        # 每次轮询间隔0.5秒，降低对CPU和DOM渲染引擎的占用
        time.sleep(0.5)

    # 若超出15秒条件仍未成立，主动抛出异常，中断当前逻辑并交由外层重试机制接管
    raise Exception(f"[{suffix_id}] 页面异步数据请求超时，触发异常阻断逻辑")

def extract_one_engineering(tab, suffix_id):
    """
    单节点DOM解析与数据提取模块
    遍历映射表，通过精确ID选择器快速提取当前页面的文本内容。
    """
    mapping = get_field_mapping()
    data = {}
    for field, selector in mapping.items():
        try:
            val = tab.ele(selector, timeout=2).text.strip()
            data[field] = val
        except:
            # 捕获因字段缺失引发的异常，默认赋空值，保证程序平稳运行
            data[field] = ""
    return data

def run_data_cycle(page, tab, codes_list, output_file):
    """
    批量查询与生命周期管控主循环
    """
    total = len(codes_list)
    all_results = []

    for index, code in enumerate(codes_list, start=1):
        print(f"\n[任务进度 {index}/{total}] 开始分配处理线程: {code}")

        max_try = 3
        current_record = None

        # 异常容错机制：针对单个编号设定最高三次的执行尝试
        for attempt in range(1, max_try + 1):
            try:
                # ---------------------------------------------------------
                # 第一阶段：系统交互与查询触发
                # ---------------------------------------------------------
                input_box = tab.ele('#projectcode', timeout=10)
                if not input_box:
                    raise Exception("无法定位查询输入框，判定页面DOM结构已失效")

                input_box.clear().input(code)
                tab.ele('#btnquery', timeout=5).click()

                # 查询指令发出后，系统会触发全局遮罩阻断用户操作。
                # 此处强制挂起3.5秒，规避因过早执行后续DOM查询导致的交互无效问题。
                print(f"    [系统状态] 尝试次数 {attempt}：查询指令已发送，执行系统响应等待...")
                tab.wait(3.5)

                # ---------------------------------------------------------
                # 第二阶段：数据总览状态嗅探
                # ---------------------------------------------------------
                p_count, e_count = erp_information.get_header_counts(tab)
                print(f"    [数据分析] 识别到当前列表信息：包含项目数 {p_count}，工程数 {e_count}")

                # ---------------------------------------------------------
                # 第三阶段：依据工程结构特征的分类提取逻辑
                # ---------------------------------------------------------
                if p_count == 0:
                    # 分支1：无数据情况，生成查无此项目记录
                    print("    [逻辑流向] 执行 Case 1 流程: 未检索到有效数据。")
                    current_record = get_information_template(code, "查无此项目")
                else:
                    # 设定工程状态标签，无子工程时默认标识为0
                    status_label = e_count if e_count > 0 else 0
                    current_record = get_information_template(code, status_label)

                    if e_count == 0:
                        # 分支2：单工程模式。根据编号直接在列表DOM中定位并点击跳转
                        print(f"    [逻辑流向] 执行 Case 2 流程: 触发编号 {code} 详情页跳转...")
                        list_item = tab.ele(f'text:{code}', timeout=5)
                        if list_item:
                            list_item.click()
                            tab.wait(1.5)  # 交互延迟缓冲

                            # 调用异步校验模块（基础比对模式）
                            wait_for_data_load(tab, "_01", exact_suffix=None)

                            data = extract_one_engineering(tab, "_01")
                            for k, v in data.items(): current_record[f"{k}_01"] = v
                        else:
                            raise Exception(f"Case 2 DOM寻址失败：未能在查询结果中定位到预期编号 {code}")

                    else:
                        # 分支3：多子工程模式。通过构建具有顺序后缀的编号循环定位点击
                        print(f"    [逻辑流向] 执行 Case 3 流程: 侦测到 {e_count} 个子项目，启动顺序提取机制...")
                        for i in range(1, min(e_count, 5) + 1):
                            suffix = f"_{i:02d}"
                            target_code = f"{code}{suffix}"

                            target_ele = tab.ele(f'text:{target_code}', timeout=5)
                            if target_ele:
                                target_ele.click()
                                tab.wait(1.5)

                                # 调用异步校验模块（严格比对模式：传入当前后缀，确保数据源变更）
                                wait_for_data_load(tab, suffix, exact_suffix=suffix)

                                data = extract_one_engineering(tab, suffix)
                                for k, v in data.items(): current_record[f"{k}{suffix}"] = v
                            else:
                                raise Exception(f"Case 3 DOM寻址失败：列表内缺失单据 {target_code}")

                    # 对于未达到最高列数（5个）的剩余字段，进行占位符填充处理
                    start_fill = max(e_count, 1 if e_count == 0 else e_count) + 1
                    for j in range(start_fill, 6):
                        current_record[f"项目名称_{j:02d}"] = "无此工程"

                # ---------------------------------------------------------
                # 第四阶段：状态重置与正常跳出
                # ---------------------------------------------------------
                tab.ele('#btnclear').click()
                tab.wait(1)
                break  # 当前编号所有流程执行无误，主动终止重试循环

            except Exception as e:
                # ---------------------------------------------------------
                # 异常接管与自愈处理模块
                # ---------------------------------------------------------
                print(f"    [异常捕获] 流程中断: {e}")
                if attempt < max_try:
                    print("    [容错机制] 尝试执行局部视图重载 (DOM Refresh)...")
                    try:
                        tab.refresh()
                        tab.wait(3)
                        # 视图重载后执行二次校验，确认核心DOM元素是否恢复
                        if not tab.ele('#projectcode', timeout=2):
                            raise Exception("刷新后核心DOM结构依然残缺")
                    except Exception as inner_e:
                        print(f"    [深度容错] 局部恢复失败 ({inner_e})，触发全局环境重建机制！")
                        # 当系统级假死导致局部刷新无效时，调用外部模块执行跨页签的彻底清理与重新导航
                        erp_information.reset_and_back_to_home(page)
                        tab = erp_information.setup_search_environment(page)
                else:
                    print("    [中断判定] 已达到最大重试上限，记录异常状态并跳过该任务。")
                    current_record = get_information_template(code, "运行异常跳过")

        # ---------------------------------------------------------
        # 第五阶段：数据清洗与聚合（后处理）
        # ---------------------------------------------------------
        if current_record and current_record.get("工程数") not in ["查无此项目", "运行异常跳过"]:
            # 直接提取首个工程的核心属性映射至汇总级
            current_record["汇总_项目名称"] = current_record.get("项目名称_01", "")
            current_record["汇总_项目状态"] = current_record.get("项目状态_01", "")

            sg_list = []
            fb_list = []
            # 定义过滤字典，用于在聚合时剔除无意义的占位文本和空值
            invalid_texts = ["", "无此工程", "查无此项目", "长时间转圈/执行异常"]

            # 遍历所有可能存在的子工程后缀，执行数据归集与去重逻辑
            for i in range(1, 6):
                sg_val = current_record.get(f"施工单位_{i:02d}", "").strip()
                fb_val = current_record.get(f"分包单位_{i:02d}", "").strip()

                if sg_val not in invalid_texts and sg_val not in sg_list:
                    sg_list.append(sg_val)
                if fb_val not in invalid_texts and fb_val not in fb_list:
                    fb_list.append(fb_val)

            # 将归集后的数组元素通过固定分隔符拼接为最终可展示的字符串格式
            current_record["汇总_施工单位"] = " / ".join(sg_list)
            current_record["汇总_分包单位"] = " / ".join(fb_list)

            # 在控制台格式化输出最终聚合结果，提供实时状态监控
            print("-" * 50)
            print(f"  [聚合监控] {code} 后处理执行完毕")
            print(f"      子工程总量: {current_record.get('工程数')}")
            print(f"      项目主名称: {current_record.get('汇总_项目名称')}")
            print(f"      项目主状态: {current_record.get('汇总_项目状态')}")
            print(f"      涉及施工方: {current_record.get('汇总_施工单位')}")
            print(f"      涉及分包方: {current_record.get('汇总_分包单位')}")
            print("-" * 50)

        # ---------------------------------------------------------
        # 第六阶段：数据持久化
        # ---------------------------------------------------------
        # 每个原子任务执行完毕后，立即进行I/O操作写入硬盘，防止后续流程崩溃导致数据丢失
        all_results.append(current_record)
        data_excel.save_data_to_excel(all_results, output_file)

    return all_results