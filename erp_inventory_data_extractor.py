"""
ValkyrieEngine 盘点流程查询-数据提取模块 (V2.1.0 报表结构优化版)
功能：在列表页通过“DOM 嗅探”技术，精准检测项目维度下各个工程后缀的流转状态，
      并顺便从已结束的单据标题中提取项目名称。
特性：
  1. 极速模式：无需进入详情页，利用显式等待在列表页直接扫描。
  2. 乱序兼容：严格独立检测每一个后缀，完美应对 ERP 业务中经常出现的乱序流转场景。
  3. 动态边界防御（预留接口）：加入 known_count 参数，精确区分合法待办的“未盘点”与超出范围的“无此工程”。
  4. 智能名称捕获：不依赖特定后缀，只要列表里有记录就能抓取名称。
  5. [V2.1.0 新增] 报表结构扩展：
     - 新增“工程数”列，展示该项目实际包含的工程数量。
     - 扩展状态列至 _05 (max_columns=5)，覆盖更多业务场景，且不影响检索效率。
"""

import time
import erp_inventory  # 导入盘点的专属页面初始化模块，用于调用异常自愈重置功能
import data_excel  # 引入数据 I/O 模块，用于实现实时自动存档


def get_inventory_record(code, known_count=3, max_columns=5):
    """
    标准数据结构生成器 (动态边界+名称+工程数增强版)
    功能：为每个项目维度编号初始化统一的字典映射模板。保证后续利用 Pandas 导出数据时，
         DataFrame 的列名严格对齐，绝对不发生表格错位。

    参数解析：
      - known_count: 该项目实际拥有的工程数量。目前版本默认传 3，未来通过新模块传入准确值。
      - max_columns: [V2.1.0 修改] 默认值扩展为 5。无论实际有几个工程，Excel 表格固定输出 5 列状态。
                     (强制对齐 _01 到 _05 列不塌陷)。
    """
    # 建立基础字典，存入项目维度的根编号
    # 【V2.1.0 新增】初始化 '项目名称' 字段。默认值为 "未找到已结束单据"。
    # 【V2.1.0 新增】初始化 '工程数' 字段，排在第三列。
    record = {
        "项目编号": code,
        "项目名称": "未找到已结束单据",
        "工程数": known_count  # 直接记录基础库查到的真实数量，方便人工核对
    }

    # 动态初始化状态：利用 known_count 精准区分“合法未做”和“压根没有”
    for i in range(1, max_columns + 1):
        # 格式化字符串，生成 "_01工程状态" ... "_05工程状态" 等标准 Key
        suffix_key = f"_{i:02d}工程状态"

        if i <= known_count:
            # 【情况 A】在已知数量范围内：默认它是“未盘点”。
            # （后续通过页面的雷达嗅探，如果发现了对应的元素，再把它改写为“结束”）
            record[suffix_key] = "未盘点"
        else:
            # 【情况 B】超出已知数量范围：连查都不用查，直接判死刑“无此工程”。
            # 这样输出的报表会极其清晰，彻底消灭模棱两可的数据。
            # 即使我们将 max_columns 加到了 5，对于只有 2 个工程的项目，
            # _03, _04, _05 会直接在这里被填上“无此工程”，根本不会触发 HTTP 请求，效率零损失！
            record[suffix_key] = "无此工程"

    return record


def extract_project_name(full_text):
    """
    [V2.2.0 新增] 项目名称提取器
    逻辑：根据用户指令，项目名称藏在第一个短横杠和第二个短横杠之间。
    示例：Dxxxxxxxxxx_01-xx路xx弄...-直埋-项目...
    切割后：['Dxxxxxxxxxx_01', 'xx路xx弄...', '直埋', ...]
    提取索引为 1 的部分。
    """
    try:
        # 简单的字符串切片，高效且直接
        parts = full_text.split('-')
        # 防御性判断：确保切出来的段数够多，防止脏数据导致越界报错
        # 用户指示：如果有人瞎搞乱填导致格式不对，我们也认了，至少保证程序不崩。
        if len(parts) >= 3:
            return parts[1].strip()
        else:
            # 如果格式确实太烂（比如没有两个短横杠），就返回原文本或错误提示
            return "名称格式异常(无法解析)"
    except:
        return "名称解析错误"


def search_and_process_single(page, search_tab, code, known_count=3):
    """
    单一项目检索与定向嗅探模块
    功能：处理目标查询页面的残留状态，输入目标编号发起检索，并对已知范围内的工程节点进行独立无序扫描。
    """
    try:
        # ==========================================
        # 阶段 1：页面状态重置 (残留标签清理)
        # ==========================================
        # 逻辑：系统执行搜索后，输入框会被隐藏，取而代之的是“主题: Dxxx”的展示标签。
        # 必须定位并清除该标签，才能恢复搜索输入框进行下一次查询。
        old_tag = search_tab.ele('text:主题:', timeout=2)

        if old_tag:
            print("[检索准备] 检测到历史查询残留标签，正在执行清除操作...")
            try:
                # 优先逻辑：通过 parent() 向上一层寻找包裹标签的父容器，再寻找包含 cancel 样式的关闭按钮
                old_tag.parent().ele('@class=cancel').click()
            except:
                # 备用逻辑：若父容器结构变化，直接尝试点击标签文本旁边的下一个相邻节点
                old_tag.next().click()

            # 【护航级维稳】标签被清除后，页面会触发局部重绘显示输入框，需强制挂起等待 2 秒
            search_tab.wait(2)

        # ==========================================
        # 阶段 2：输入框定位与检索触发
        # ==========================================
        # 优先使用 Landray 框架底层静态属性 data-lui-placeholder 定位，避免受页面状态影响
        search_box = search_tab.ele('@data-lui-placeholder=请输入主题', timeout=10)

        # 若底层属性失效，降级使用常规的 placeholder 属性进行定位
        if not search_box:
            search_box = search_tab.ele('@placeholder=请输入主题', timeout=10)

        # 清除输入框内可能存在的数据，填入新编号，并追加换行符 \n 模拟物理回车操作，触发强力搜索
        search_box.clear().input(f'{code}\n')
        print(f"[数据检索] 检索指令已发送，当前处理编号：[{code}]，等待服务器响应...")

        # 【V2.1.0 修复：放慢节奏】将原本的 4 秒增加到 6 秒。
        # 给 ERP 系统的后台数据库留出足够的检索与渲染时间，防止我们过快介入导致“竞态条件”错过数据。
        search_tab.wait(6)

        # ==========================================
        # 阶段 3：工程维度独立无序嗅探 (核心狙击逻辑 + 动态边界 + 名称抓取)
        # ==========================================
        # 调用初始化函数，把 known_count 传进去
        # 【V2.1.0】这里 max_columns 默认已经是 5 了，record 里会预置好 5 个状态坑位
        record = get_inventory_record(code, known_count=known_count)

        print(f"[业务判定] 正在对 [{code}] 展开工程维度雷达扫描 (系统已知该项目工程数: {known_count})...")

        # 标记位：用于防止重复抓取名称。
        is_name_extracted = False

        # 【核心性能优化】：只扫描已知范围内的后缀！
        # 这里的循环保证了我们不会错过任何一个存在的后缀。
        # 【重要】即使 max_columns 是 5，如果 known_count 是 2，循环只跑 2 次。绝不浪费时间！
        for i in range(1, known_count + 1):
            # 格式化生成当前需要扫描的后缀，如 "_01"
            suffix = f"_{i:02d}"
            # 组合成精确的特征靶标，必须附带短横杠（如 "D1221420034_01-"），这是防止误判其他相似编号的定海神针！
            target_text = f"{code}{suffix}-"

            # 使用隐式等待短暂探测（timeout=1 足矣，因为我们前面 wait(6) 已经确保列表渲染完毕了）
            # 【V2.2.0 修改】我们需要获取这个元素对象，不仅仅是判断它存在，还要读它的文本
            target_element = search_tab.ele(f'text:{target_text}', timeout=1)

            if target_element:
                # 1. 更新状态：哪怕只有这一个存在，状态也能被正确记录
                record[f"{suffix}工程状态"] = "结束"
                print(f"  --> [命中] 发现靶标 [{target_text}]，状态更新为：结束")

                # 2. [V2.2.0 新增] 顺手牵羊抓取项目名称
                if not is_name_extracted:
                    full_text = target_element.text # 获取整行主题文本
                    extracted_name = extract_project_name(full_text)
                    record["项目名称"] = extracted_name
                    is_name_extracted = True # 标记已完成
                    print(f"      [信息捕获] 已从单据 [{suffix}] 中提取名称：{extracted_name}")
            else:
                # 没找到就脱靶，由于我们在 get_inventory_record 里已经把它设为了"未盘点"，这里其实什么都不用做
                print(f"  --> [脱靶] 未见靶标 [{target_text}]，维持合法状态：未盘点")

        # 扫描完毕，直接返回这张映射得清清楚楚的字典
        return record

    except Exception as e:
        # 捕获检索及 DOM 交互过程中引发的系统级异常（如断网、页面彻底卡死无响应）
        print(f"[系统警报] 处理编号 [{code}] 时发生页面崩溃或响应超时，错误信息：{e}")
        # 向上级总控模块抛出自定义异常，请求重置干预
        raise Exception("SearchTimeout")


def run_data_cycle(page, search_tab, codes_data, output_file):
    """
    批量数据检索主控循环模块
    功能：遍历待处理的数据，调用单次查询逻辑。包含应对页面级卡顿的自愈重启策略与实时存档。

    参数解析：
      - codes_data: 接收包含字典的列表 (如 [{'项目编号': 'D123', '工程数': 2}])。
    """
    total = len(codes_data)
    all_results = []

    for index, item in enumerate(codes_data, start=1):

        # ==========================================
        # 兼容层：智能解析传入的数据类型
        # ==========================================
        if isinstance(item, dict):
            # 未来的理想状态：item 是一个字典，直接提取编号和准确的工程数
            code = item.get("项目编号")
            # 【重点】这里获取到的真实工程数，会一路传给 get_inventory_record，覆盖掉默认的 3
            known_count = item.get("工程数", 3)
        else:
            # 现在的状态：item 只是一个字符串编号，我们就保守起见，默认查到 3
            code = item
            known_count = 3

        print(f"\n[任务进度 {index}/{total}] 开始分配任务，当前执行编号: {code} (计划精确嗅探 {known_count} 个工程)")

        # 设定单次任务的最大容错重试次数
        max_retries = 2

        for attempt in range(1, max_retries + 1):
            try:
                # 把解析出的 code 和 known_count 透传给底层核心处理函数
                record = search_and_process_single(page, search_tab, code, known_count=known_count)
                all_results.append(record)

                print(f"[任务完成] [{code}] 处理完毕，已压入内存栈。")
                break  # 当前编号处理成功，跳出重试循环，执行下一个编号

            except Exception as e:
                # 接收来自下层抛出的超时异常，触发自愈干预
                print(f"[自愈干预] 第 {attempt} 次处理失败，正在启动浏览器环境重置程序...")

                if attempt < max_retries:
                    # 联动专属盘点模块的重置功能，清理多余标签并刷新恢复干净的业务页面
                    erp_inventory.reset_and_back_to_home(page)
                    search_tab = erp_inventory.setup_search_environment(page)
                else:
                    # 超过最大重试次数，判定该数据异常或网络中断严重
                    print(f"[业务放弃] 编号 [{code}] 导致程序反复超时，已跳过该节点。")

                    # 【核心容错机制】：生成一个包含已知信息的报错字典，保证总体进度不受单一数据影响，且列名依旧对齐！
                    # 这里调用的 get_inventory_record 也会生成带有“工程数”字段的记录，保持队形整齐
                    error_record = get_inventory_record(code, known_count=known_count)

                    # 报错时，名称也得占位
                    error_record["项目名称"] = "抓取失败(网页卡死)"

                    # 强行遍历，把带有"工程状态"的键的值全部覆盖为报错提示
                    for k in error_record:
                        if "工程状态" in k:
                            error_record[k] = "网页连续卡死失败"
                    all_results.append(error_record)

        # ======================================================================
        # 【极致护航级实时自动存档机制】
        # 逻辑：每当一个编号经过 1~2 次重试处理完毕，立即序列化并覆盖写入硬盘。
        # 作用：确保即便程序在下一秒崩溃，之前的所有劳动成果都已安全落盘。
        # ======================================================================
        print(f"[自动存档] 正在执行进度同步，当前已安全保存 {len(all_results)} 条业务记录...")
        data_excel.save_data_to_excel(all_results, output_file)

    return all_results