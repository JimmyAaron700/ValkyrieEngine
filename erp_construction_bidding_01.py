"""
ValkyrieEngine 页面初始化与环境导航模块 (功能2专属：工程维度_01)
文件名：erp_construction_bidding_01.py
功能：控制浏览器导航至“施工委托（招标）”页面，并完成筛选条件的初始化。
区分：本模块专用于【工程维度】查询，与功能1的【项目维度】入口不同。
"""

def reset_and_back_to_home(page):
    """
    页面状态重置模块 (复用逻辑)
    功能：在系统发生卡顿或异常时被调用，负责清理浏览器中产生的所有无效标签页，
    并将唯一保留的首页刷新，使系统恢复至初始的纯净状态。
    """
    print("[系统维护] 正在清理无响应的页面，重置浏览器环境...")

    # 获取当前浏览器进程中所有已打开标签页的内部 ID
    all_tabs = page.tab_ids

    # 状态判定：如果列表长度大于 1，说明除了最开始的首页，还有新弹出的业务页面
    if len(all_tabs) > 1:
        # 页面清理：保留首页 (索引0)，关闭后续所有衍生标签页
        page.close_tabs(all_tabs[1:])

    # 状态初始化：刷新首页，清除缓存造成的卡顿
    page.refresh()

    # 强制等待 1 秒，确保网页后端的 JavaScript 脚本初始化完成
    page.wait(1)
    print("[系统维护] 首页状态已重置，准备重新发起业务请求。")


def apply_search_conditions(tab):
    """
    查询条件初始化模块 (功能2定制版)
    功能：在“施工委托（招标）”页面加载后，挂载筛选条件。
    """
    print("[参数配置] 正在等待页面数据加载稳定...")

    # [核心防错机制] 强制等待 2 秒。
    # 原理：ERP 系统的 AJAX 异步刷新特性，刚进页面时元素可能不稳定。
    # 这一步是为了防止 DOM 树还未构建完成就急着操作。
    tab.wait(2)

    print("[参数配置] 正在设置业务筛选条件...")

    # 【定位与操作 1】：通过元素属性锁定单选按钮
    # 寻找 html 属性 title 值等于“结束”的节点并点击
    tab.ele('@title=结束', timeout=15).click()
    print("[参数配置] 已勾选 '结束' 状态...")

    # 再次强制等待 1 秒，确保点击引发的局部 AJAX 重绘完成
    tab.wait(1)

    # 【定位与操作 2】：清除默认的时间限制
    # 步骤 A：锁定“创建时间”的外部容器区块 (data-criterion-key=docCreateTime)
    time_box = tab.ele('@data-criterion-key=docCreateTime', timeout=15)

    # 步骤 B：在容器内部寻找取消按钮 (class=cancel) 并点击
    time_box.ele('@class=cancel', timeout=15).click()
    print("[参数配置] 已清除默认的 '创建时间' 限制...")

    # ==========================================
    # [显式等待锚点]：使用 SR 前缀作为加载完成的信号
    # ==========================================
    print("[参数配置] 正在等待系统拉取全量历史数据 (工程维度)...")
    try:
        # 【逻辑复用】和功能1一样，列表里的单据编号通常包含 "SR"
        # 给系统 30 秒宽容度，嗅探到 SR 说明列表渲染完毕
        tab.ele('text:SR', timeout=30)
        print("[参数配置] 成功嗅探到 'SR' 编号，列表数据渲染完毕！准备移交控制权...")

        # 渲染完再给浏览器 1 秒钟的喘息时间，平复 DOM 树
        tab.wait(1)
    except Exception as e:
        # 致命超时处理
        print(f"[系统警报] 致命超时：30秒内未检测到 'SR' 数据，网络严重阻塞！")
        raise Exception("ListRenderTimeout: 列表渲染严重超时。")


def setup_search_environment(page):
    """
    查询环境导航总控模块 (功能2入口)
    功能：导航至“施工委托（招标）”并初始化。
    """
    max_try = 3  # 定义最大容错重试次数

    for attempt in range(1, max_try + 1):
        print(f"\n[环境导航 - 第 {attempt}/{max_try} 次尝试] 正在进入【施工委托（招标）】界面...")

        try:
            # 1. 点击一级菜单 '流程查询'
            page.ele('text:流程查询', timeout=15).click()
            print("[环境导航] 触发菜单栏 '流程查询'")

            # 移交控制权至新页面
            tab = page.latest_tab

            # 2. 点击二级菜单 (注意：这里是功能2的核心差异点！)
            # 功能1点的是“施工委托招标(项目)”，这里我们要点“施工委托（招标）”
            tab.ele('text:施工委托（招标）', timeout=15).click()
            print("[环境导航] 成功进入【施工委托（招标）】业务列表...")

            # 3. 执行筛选条件挂载
            apply_search_conditions(tab)

            print("[环境导航] 业务查询环境初始化完毕，检索条件已挂载。")

            # 返回初始化完毕的标签页对象
            return tab

        except Exception as e:
            print(f"[系统警报] 第 {attempt} 次加载未响应或渲染超时，错误详情：{e}")

            if attempt < max_try:
                # 失败则重置回首页
                reset_and_back_to_home(page)
            else:
                print("[致命错误] 页面响应超时，已放弃重试。")
                raise Exception("Fatal Error：无法突破导航拦截，已达到最大重试次数。")