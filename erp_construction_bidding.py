"""
ValkyrieEngine 页面初始化与环境导航模块 (功能1专属)
功能：控制浏览器导航至“施工委托招标”页面，并完成筛选条件的初始化。
"""

def reset_and_back_to_home(page):
    """
    页面状态重置模块
    功能：在系统发生卡顿或异常时被调用，负责清理浏览器中产生的所有无效标签页，
    并将唯一保留的首页刷新，使系统恢复至初始的纯净状态。
    """
    print("[系统维护] 正在清理无响应的页面，重置浏览器环境...")

    # 获取当前浏览器进程中所有已打开标签页的内部 ID，返回一个列表（如 ['id_1', 'id_2']）
    all_tabs = page.tab_ids

    # 状态判定：如果列表长度大于 1，说明除了最开始的首页，还有新弹出的业务页面
    if len(all_tabs) > 1:
        # 页面清理：使用 Python 列表切片 [1:]，跳过索引为 0 的首页，将后续所有衍生的标签页批量关闭
        page.close_tabs(all_tabs[1:])

    # 状态初始化：此时浏览器只剩下首页，调用 refresh() 重新加载当前页面的 DOM 树，清除缓存造成的卡顿
    page.refresh()

    # 强制等待 1 秒，确保网页后端的 JavaScript 脚本初始化完成
    page.wait(1)
    print("[系统维护] 首页状态已重置，准备重新发起业务请求。")


def apply_search_conditions(tab):
    """
    查询条件初始化模块
    功能：在目标查询页面加载后，通过精准的 DOM 定位技术，修改默认的筛选状态。
    重点包含：勾选特定单据状态、解除默认的时间范围限制。
    """
    print("[参数配置] 正在等待页面数据加载稳定...")

    # [核心防错机制] 强制等待 2 秒。
    # 原理：ERP 系统的部分表格采用 AJAX 异步局部刷新。刚进入页面时元素可能正在重绘，
    # 强行等待可避免因抓取到“即将被替换的旧元素”而引发的 ElementLostError（元素失效异常）。
    tab.wait(2)

    print("[参数配置] 正在设置业务筛选条件...")

    # 【定位与操作 1】：通过元素属性锁定单选按钮
    # 语法解析：'@title=结束' 表示在整个网页结构中，寻找 html 属性 title 值等于“结束”的节点。
    # 操作逻辑：设定 15 秒的隐式等待，一旦该节点在 DOM 中出现，立刻触发底层 JavaScript 的 click() 点击事件。
    tab.ele('@title=结束', timeout=15).click()
    print("[参数配置] 已勾选 '结束' 状态...")

    # [V2.0.0 回滚补充] 再次强制等待 1 秒，确保上一步点击引发的局部 AJAX 重绘完成
    tab.wait(1)

    # 【定位与操作 2】：通过层级关系精准锁定特定按钮（防止误点）
    # 步骤 A (寻找父节点)：通过底层框架自定义属性 'data-criterion-key' 锁定“创建时间”的外部容器区块。
    # 这种 data- 属性通常是系统底层的业务标识，比 class 或 id 更稳定。
    time_box = tab.ele('@data-criterion-key=docCreateTime', timeout=15)

    # 步骤 B (相对定位与点击)：在上一步找到的 time_box 容器内部，继续向下寻找包含 'class=cancel' 属性的元素（即取消按钮）。
    # 这种“先找大盒子，再找小按钮”的链式定位法，彻底杜绝了点到页面其他位置同名 cancel 按钮的风险。
    time_box.ele('@class=cancel', timeout=15).click()
    print("[参数配置] 已清除默认的 '创建时间' 限制...")

    # ==========================================
    # [V2.1.0 优化] 智能显式等待：防范首页大数据量加载造成的竞态条件
    # ==========================================
    print("[参数配置] 正在等待系统拉取全量历史施工委托数据，这可能需要较长时间...")
    try:
        # 【核心狙击逻辑】小郁提出的神级锚点：寻找申请单编号的通用前缀 "SR"
        # 给系统 30 秒的极限宽容度，只要能嗅探到列表里有任意一个 SR，就说明转圈结束，列表渲染完毕！
        tab.ele('text:SR', timeout=30)
        print("[参数配置] 成功嗅探到 'SR' 编号，列表数据渲染完毕！准备移交控制权...")

        # 为了极度安全，渲染完再给浏览器 1 秒钟的喘息时间，平复 DOM 树
        tab.wait(1)
    except Exception as e:
        # 【V2.1.1 紧急修复】绝不能强制放行！果断抛出致命异常！
        # 这个异常会被外层的 setup_search_environment 捕获，从而完美触发 reset_and_back_to_home 机制。
        print(f"[系统警报] 致命超时：30秒内未检测到 'SR' 数据，网络严重阻塞！")
        raise Exception("ListRenderTimeout: 列表渲染严重超时，拒绝执行后续脏数据抓取。")


def setup_search_environment(page):
    """
    查询环境导航总控模块
    功能：控制浏览器从首页逐步导航至具体的业务查询列表，并调用 apply_search_conditions 完成条件设置。
    内置多轮重试的自愈机制，以应对企业内网偶尔的延迟与阻断。
    """
    max_try = 3  # 定义最大容错重试次数

    for attempt in range(1, max_try + 1):
        print(f"\n[环境导航 - 第 {attempt}/{max_try} 次尝试] 正在进入施工委托查询界面...")

        try:
            # 【定位与操作 3】：通过文本模糊匹配进行导航
            # 语法解析：'text:流程查询' 表示在页面上寻找可见文本包含“流程查询”的元素并执行点击。
            page.ele('text:流程查询', timeout=15).click()
            print("[环境导航] 触发菜单栏 '流程查询'")

            # 上一步的点击会触发浏览器打开一个新的标签页，
            # 必须通过 page.latest_tab 将代码的控制权（句柄）移交至最新弹出的页面对象上。
            tab = page.latest_tab

            # 在新获取控制权的页面中，继续通过文本特征查找目标业务入口并点击
            tab.ele('text:施工委托招标(项目)', timeout=15).click()
            print("[环境导航] 成功进入施工委托业务列表...")

            # 页面导航完成，调用独立模块执行具体的筛选框勾选逻辑
            apply_search_conditions(tab)

            print("[环境导航] 业务查询环境初始化完毕，检索条件已挂载。")

            # 返回初始化完毕的标签页对象，供后续的数据提取模块进行循环搜索
            return tab

        except Exception as e:
            # 异常捕获与日志输出：记录引发超时的具体节点错误信息
            # 【完美联动】：apply_search_conditions 抛出的 ListRenderTimeout 会直接被这里接住！
            print(f"[系统警报] 第 {attempt} 次加载未响应或渲染超时，错误详情：{e}")

            # 自愈逻辑分流：如果未达到最大重试上限，则调用重置模块清理页面；否则向主程序抛出致命异常。
            if attempt < max_try:
                reset_and_back_to_home(page)
            else:
                print("[致命错误] 页面响应超时，已放弃重试。")
                raise Exception("Fatal Error：无法突破导航拦截，已达到最大重试次数，自动化引擎强制停机。")