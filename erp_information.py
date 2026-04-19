"""
ValkyrieEngine 模块五：环境导航与状态嗅探组件
文件名：erp_information.py

【组件职责】
本模块负责维持 RPA 流程的“环境一致性”，包括业务页面的定向导航、浏览器上下文的初始化验证
以及关键业务指标（项目数、工程数）的预解析。
"""

import time
import re

def reset_and_back_to_home(page):
    """
    浏览器上下文恢复组件 (Browser Context Recovery)

    【设计巧思】
    在分布式或长时间运行的自动化任务中，浏览器可能会因为内存溢出或页面逻辑死锁产生冗余页签。
    本函数通过遍历 Tab 句柄池，强制销毁除主索引页以外的所有业务逻辑页，并触发全局刷新，
    确保后续的重试逻辑能够在一个确定的、纯净的初始状态下启动，避免逻辑干扰。
    """
    print("[系统维护] 正在执行环境重置，同步浏览器上下文状态...")
    all_tabs = page.tab_ids
    if len(all_tabs) > 1:
        # 保留索引句柄，销毁所有业务衍生句柄
        page.close_tabs(all_tabs[1:])
    page.refresh()
    page.wait(1)
    print("[系统维护] 浏览器环境已重置至基准线。")

def setup_search_environment(page):
    """
    业务工作台导航处理器 (Workflow Navigation Handler)

    【技术要点】
    1. 动态导航：从系统首页通过路由点击进入“项目流程工作台”。
    2. 句柄移交：实时获取最新产生的标签页对象，实现从导航态到查询态的平滑过渡。
    3. 状态检查机制：在移交控制权前，通过探测 ID 为 #projectcode 的核心输入组件，
       验证 DOM 树是否已经完成了关键业务元素的挂载，确保下游提取模块不会因环境未就绪而崩溃。
    """
    max_try = 3
    for attempt in range(1, max_try + 1):
        print(f"\n[导航任务 - 尝试 {attempt}/{max_try}] 正在请求进入：项目流程工作台...")
        try:
            # 触发业务路由点击
            page.ele('text:项目流程工作台', timeout=15).click()

            # 获取最新的标签页句柄
            tab = page.latest_tab

            # 执行 DOM 就绪校验
            tab.ele('#projectcode', timeout=15)
            print("[导航任务] 环境校验通过，工作台句柄已就绪。")
            return tab
        except Exception as e:
            print(f"[导航任务] 导航流程中断或 DOM 挂载超时: {e}")
            if attempt < max_try:
                reset_and_back_to_home(page)
            else:
                # 达到最大重试上限后，向调用栈抛出致命异常，触发主程序的终结逻辑
                raise Exception("Fatal Error: 无法建立有效的业务查询上下文。")

def get_header_counts(tab):
    """
    业务元数据解析器 (Metadata Sniffing Module)

    【技术要点】
    1. 数据非标解析：系统顶部的汇总信息通常以“指标名(数值)”的非结构化字符串形式存在。
    2. 正则表达式驱动：利用 re 库精准捕获括号内的整数值，实现从 UI 文本到程序逻辑变量的类型转换。
    3. 逻辑解耦：将“查到了多少个项目”与“如何查询”分离，为 Case 1/2/3 的逻辑分流提供判别依据。
    """
    try:
        # 强制同步等待，确保 AJAX 回调已完成 UI 数据更新
        tab.wait(1)

        # 严格映射业务 ID：#aatest(项目数) 与 #bbtest(工程数)
        p_text = tab.ele('#aatest', timeout=10).text
        e_text = tab.ele('#bbtest', timeout=10).text

        # 提取数值元数据
        p_count = int(re.search(r'\((\d+)\)', p_text).group(1)) if p_text else 0
        e_count = int(re.search(r'\((\d+)\)', e_text).group(1)) if e_text else 0

        return p_count, e_count
    except Exception as e:
        print(f"    [警告] 元数据解析失败，可能由于 AJAX 响应结构异常: {e}")
        return 0, 0