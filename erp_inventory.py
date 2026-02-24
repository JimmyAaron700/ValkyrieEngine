"""
ValkyrieEngine 盘点流程查询业务模块 (V2.1.0版)
功能：负责处理【项目材料竣工数量盘点】业务线的前期页面导航与环境初始化。
"""


def reset_and_back_to_home(page):
    """
    页面状态重置模块
    功能：在系统发生卡顿或异常时被调用，负责清理浏览器中产生的所有无效标签页，
    并将唯一保留的首页刷新，使系统恢复至初始的纯净状态。
    """
    print("[系统维护] 正在清理无响应的页面，重置浏览器环境...")

    # 获取当前浏览器进程中所有已打开标签页的内部 ID，返回一个列表
    all_tabs = page.tab_ids

    # 状态判定：如果列表长度大于 1，说明除了最开始的首页，还有新弹出的业务页面
    if len(all_tabs) > 1:
        # 页面清理：跳过索引为 0 的首页，将后续所有衍生的标签页批量关闭
        page.close_tabs(all_tabs[1:])

    # 状态初始化：此时浏览器只剩下首页，重新加载 DOM 树
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
    print("[参数配置] 正在等待盘点页面数据加载稳定...")

    # [核心防错机制] 强制等待 2 秒，确保 AJAX 异步数据渲染完毕，护航级稳定性保障！
    tab.wait(2)

    print("[参数配置] 正在设置业务筛选条件...")

    # 【定位与操作 1】：通过元素属性锁定单选按钮
    tab.ele('@title=结束', timeout=15).click()
    print("[参数配置] 已勾选 '结束' 状态...")

    # 强制等待 1 秒，确保上一步点击引发的局部 AJAX 重绘完成
    tab.wait(1)

    # 【定位与操作 2】：通过层级关系精准锁定特定按钮（防止误点）
    # 锁定“创建时间”的外部容器区块
    time_box = tab.ele('@data-criterion-key=docCreateTime', timeout=15)

    # 在容器内部点击取消按钮
    time_box.ele('@class=cancel', timeout=15).click()
    print("[参数配置] 已清除默认的 '创建时间' 限制...")

    # ==========================================
    # [V2.1.0 修复] 智能显式等待：拦截首页大数据量加载造成的转圈卡顿
    # ==========================================
    print("[参数配置] 正在等待系统拉取全量历史盘点数据，这可能需要较长时间...")

    try:
        # 给系统 30 秒的极限宽容度，只要能嗅探到列表里有任意一个 _01-，就说明转圈结束了！
        tab.ele('text:_01-', timeout=30)
        print("[参数配置] 成功嗅探到列表数据渲染完毕！准备移交控制权...")

        # 为了极度安全，渲染完再给浏览器 1 秒钟的喘息时间，平复 DOM 树
        tab.wait(1)
    except Exception as e:
        # 【V2.1.1 紧急修复】绝不能强制放行！果断抛出致命异常！
        # 这个异常会被外层的 setup_search_environment 捕获，从而完美触发 reset_and_back_to_home 机制。
        print(f"[系统警报] 致命超时：30秒内未检测到 '_01-' 数据，网络严重阻塞！")
        raise Exception("ListRenderTimeout: 列表渲染严重超时，拒绝执行后续脏数据抓取。")

def setup_search_environment(page):
    """
    查询环境导航总控模块
    功能：控制浏览器从首页逐步导航至具体的盘点业务查询列表，并挂载初始筛选条件。
    """
    max_try = 3  # 定义最大容错重试次数

    for attempt in range(1, max_try + 1):
        print(f"\n[环境导航 - 第 {attempt}/{max_try} 次尝试] 正在进入盘点流程查询界面...")

        try:
            # 触发菜单栏 '流程查询'
            page.ele('text:流程查询', timeout=15).click()
            print("[环境导航] 触发菜单栏 '流程查询'")

            # 将代码的控制权（句柄）移交至最新弹出的页面对象上
            tab = page.latest_tab

            # 【核心修改点】：在新的页面中，寻找盘点业务专属的入口并点击
            tab.ele('text:项目材料竣工数量盘点', timeout=15).click()
            print("[环境导航] 成功进入【项目材料竣工数量盘点】业务列表...")

            # 页面导航完成，调用独立模块执行具体的筛选框勾选逻辑
            apply_search_conditions(tab)

            print("[环境导航] 盘点业务查询环境初始化完毕，检索条件已挂载。")

            # 返回初始化完毕的标签页对象，供后续的数据提取模块使用
            return tab

        except Exception as e:
            print(f"[系统警报] 第 {attempt} 次加载未响应，错误详情：{e}")

            if attempt < max_try:
                reset_and_back_to_home(page)
            else:
                print("[致命错误] 页面响应超时，已放弃重试。")
                raise Exception("Fatal Error：无法突破导航拦截，已达到最大重试次数，自动化引擎强制停机。")