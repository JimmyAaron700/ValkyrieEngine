import time
import erp_construction_bidding  # 导入页面初始化模块，用于调用其内置的页面重置功能
import data_excel  # [V2.0.0 新增] 引入数据 I/O 模块，用于实现实时自动存档机制


def get_empty_record(code, status):
    """
    标准数据结构生成器
    功能：为每个项目编号初始化统一的字典映射模板。
    无论业务执行成功、匹配失败还是网络异常，均强制返回包含所有字段的字典。
    这保证了后续利用 Pandas 导出数据时，DataFrame 的列名严格对齐，防止表格错位。
    """
    return {
        "项目编号": code,
        "项目名称": 0,
        "项目工程总造价(元)": 0,
        "市政道路修复费": 0,
        "小区道路修复费": 0,
        "绿化修复费": 0,
        "发包金额": 0,
        "打捆招标名称": 0,
        "项目中标金额": 0,
        "状态": status
    }


def extract_detail_data(detail_tab, code):
    """
    详情页数据提取模块
    功能：在详情页中，基于表头文本的相对偏移量，精准提取对应的资金数据。
    内置了针对不规范前端代码的容错与兼容处理。
    """
    print(f"[数据提取] 正在解析目标编号 [{code}] 的明细数据...")

    # 初始化默认的返回模板，并标记初始状态为完成
    record = get_empty_record(code, "完成")

    # 显式等待 2 秒：确保新打开的详情页面 DOM 树和内部 AJAX 数据完全渲染完毕
    detail_tab.wait(2)

    # 定义需要按序提取的字段名列表
    fields_to_extract = [
        "项目名称", "项目工程总造价(元)", "市政道路修复费",
        "小区道路修复费", "绿化修复费", "发包金额",
        "打捆招标名称", "项目中标金额"
    ]

    for field in fields_to_extract:
        try:
            # 【定位策略 1：左侧表头严格匹配】
            # 语法解析：tag:td@@class=td_normal_title@@text():{field}
            # 逻辑：限定节点必须是 <td> 标签，且 class 属性必须为 td_normal_title，同时内部文本包含目标字段。
            # 作用：这种强约束可以防止误抓页面上其他包含该字段文本的无关大模块。
            label_td = detail_tab.ele(f'tag:td@@class=td_normal_title@@text():{field}', timeout=3)

            # [V2.0.0 健壮性增强] 防御性编程：确保表头真实存在后再进行相对偏移
            if not label_td:
                print(f"[数据提取] 警告：页面中未找到表头 [{field}]")
                record[field] = "抓取缺失"
                record["状态"] = "部分字段异常"
                continue

            # 【定位策略 2：右侧数据宽泛匹配 (相对节点偏移)】
            # 逻辑：获取上一步定位到的 label_td 节点的下一个兄弟节点 (next sibling)，限定其为 <td> 标签。
            # 作用：因为 ERP 系统部分数据单元格缺失规范的 class 属性，所以放弃精确匹配 class，
            # 只要它是紧随表头之后的单元格，即认定为包含目标金额的数据节点。
            value_td = label_td.next('tag:td', timeout=2)

            if value_td:
                # 提取数据节点内的纯文本内容
                raw_text = value_td.text

                # 【数据清洗逻辑】
                # 将系统底层自带的换行符 (\n) 与制表符 (\t) 替换为空字符串，最后用 strip() 移除首尾残留的空白字符。
                clean_text = raw_text.replace('\n', '').replace('\t', '').strip()
                record[field] = clean_text
            else:
                record[field] = "数据节点缺失"
                record["状态"] = "部分字段异常"

        except Exception as e:
            # 容错处理：若该页面缺失某个字段（如无绿化费），或页面结构发生改变，记录异常并不中断程序
            print(f"[数据提取] 字段解析异常: {field}，底层错误: {e}")
            record[field] = "抓取异常"
            record["状态"] = "部分字段异常"

    return record


def search_and_process_single(page, search_tab, code):
    """
    单一项目检索与状态判定模块
    功能：处理目标查询页面的残留状态，输入目标编号发起检索，并根据检索结果的数量
    进行条件分支处理（未发包/工程维度发包、待确认、精确命中）。
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

            # 标签被清除后，页面会触发局部重绘显示输入框，需强制挂起等待 2 秒
            search_tab.wait(2)

        # ==========================================
        # 阶段 2：输入框定位与检索触发
        # ==========================================
        # 优先使用 Landray 框架底层静态属性 data-lui-placeholder 定位，避免受页面状态影响
        search_box = search_tab.ele('@data-lui-placeholder=请输入主题', timeout=10)

        # 若底层属性失效，降级使用常规的 placeholder 属性进行定位
        if not search_box:
            search_box = search_tab.ele('@placeholder=请输入主题', timeout=10)

        # 清除输入框内可能存在的数据，填入新编号，并追加换行符 \n 模拟物理回车操作，触发搜索
        search_box.clear().input(f'{code}\n')

        print(f"[数据检索] 检索指令已发送，当前处理编号：[{code}]，等待服务器响应...")
        search_tab.wait(4)

        # ==========================================
        # 阶段 3：检索结果校验与分流
        # ==========================================
        # 利用项目编号附带的短横杠（如 D123-）作为业务唯一标识，获取所有匹配的记录节点。
        results = search_tab.eles(f'text:{code}-')
        result_count = len(results)

        if result_count == 0:
            print("[业务判定] 数据库未返回匹配记录，状态标记：未发包/工程维度发包。")
            return get_empty_record(code, "未发包/工程维度发包")

        elif result_count > 1:
            print(f"[业务判定] 检测到 {result_count} 条同名记录，为避免数据混淆，状态标记：待确认。")
            return get_empty_record(code, "待确认")

        else:
            print("[业务判定] 精确命中单一业务记录，准备深入抓取明细...")
            # 触发唯一记录的点击事件，打开详情页
            results[0].click()

            # 给底层系统留出响应打开新标签页的微小时间差
            page.wait(1)

            # 获取最新弹出的详情页句柄
            detail_tab = page.latest_tab

            # [V2.0.0 健壮性增强] 引入 try...finally 确保即使提取报错也能强制销毁标签页，防止内存溢出
            try:
                final_record = extract_detail_data(detail_tab, code)
                return final_record
            finally:
                print(f"[资源回收] 正在关闭编号 [{code}] 的详情层对象。")
                detail_tab.close()

    except Exception as e:
        # 捕获检索及 DOM 交互过程中引发的系统级异常（如断网、页面彻底卡死无响应）
        print(f"[系统警报] 处理编号 [{code}] 时发生页面崩溃或响应超时，错误信息：{e}")
        # 向上级总控模块抛出自定义异常，请求介入处理
        raise Exception("SearchTimeout")


def run_data_cycle(page, search_tab, codes_list, output_file):
    """
    批量数据检索主控循环模块
    功能：遍历待处理的编号列表，调用单次查询逻辑。包含应对页面级卡顿的自愈重启策略。
    """
    total = len(codes_list)
    all_results = []

    # 利用 enumerate 生成带序号的迭代，提供任务进度监控
    for index, code in enumerate(codes_list, start=1):
        print(f"\n[任务进度 {index}/{total}] 开始分配任务，当前执行编号: {code}")

        # 设定单次任务的最大容错重试次数
        max_retries = 2

        for attempt in range(1, max_retries + 1):
            try:
                # 尝试执行单一查询处理链
                record = search_and_process_single(page, search_tab, code)
                all_results.append(record)

                print(f"[任务完成] 成功构建数据映射: {record}")
                break  # 当前编号处理成功，跳出重试循环，执行下一个编号

            except Exception as e:
                # 接收来自下层 search_and_process_single 抛出的异常
                print(f"[自愈干预] 第 {attempt} 次处理失败，正在启动浏览器环境重置程序...")

                if attempt < max_retries:
                    # 调用页面初始化模块的重置功能，清理多余标签并刷新首页
                    erp_construction_bidding.reset_and_back_to_home(page)
                    # 重新执行从首页导航至查询页面、重置筛选条件的初始化操作
                    search_tab = erp_construction_bidding.setup_search_environment(page)
                    print(f"[自愈干预] 浏览器状态已重置，准备对编号 [{code}] 重新发起请求...")
                else:
                    # 超过最大重试次数，判定该数据异常或网络中断严重
                    print(f"[业务放弃] 编号 [{code}] 导致程序反复超时，已跳过该节点。")
                    # 保留业务日志，记录错误状态，确保总体进度不受单一数据影响
                    all_results.append(get_empty_record(code, "网页连续卡死失败"))

        # ======================================================================
        # 【V2.0.0 新增：极致护航级实时自动存档机制】
        # 垂直水平对齐说明：本行代码与上面的 for attempt 循环对齐。
        # 逻辑：每当一个编号经过 1~2 次重试处理完毕（无论结果是成功还是失败标记），
        # 立即将当前内存中已搜集的 all_results 序列化并覆盖写入硬盘。
        # 作用：确保即便程序在下一秒崩溃，之前的所有劳动成果都已安全落盘，绝不白跑。
        # ======================================================================
        print(f"[自动存档] 正在执行结算进度同步，当前已安全保存 {len(all_results)} 条业务记录...")
        data_excel.save_data_to_excel(all_results, output_file)

    return all_results


# ---------------- 单独测试与调试入口 ----------------
if __name__ == '__main__':
    from DrissionPage import ChromiumPage

    print("================ 核心业务模块独立调试 ================")
    print("[调试提醒] 请确认当前浏览器已登录 ERP，并处于已设置好初步筛选条件的查询页面。")
    input("按回车键开始注入测试序列...")

    # 接管当前处于激活状态的浏览器进程
    page = ChromiumPage()
    current_search_tab = page

    # 定义包含各类边界情况的测试用例列表
    test_codes = [
        "D1251420211",  # 预期输出：精确抓取数据
        "D9999999999",  # 预期输出：未发包/工程维度发包
        "D1251420013"   # 预期输出：另一条正常数据
    ]

    # 执行批处理流程
    final_data = run_data_cycle(page, current_search_tab, test_codes)

    print("\n[调试结束] 最终合并的结构化数据列表如下：")
    for data in final_data:
        print(data)