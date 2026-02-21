def reset_and_back_to_home(page):
    """
    功能：一旦发现报错，关掉所有多余的新页面，刷新首页。
    """
    print("正在清理卡死的页面...")

    # 拿到当前所有标签页的 ID 列表
    all_tabs = page.tab_ids

    # 如果除了首页之外还有别的页面，全部关掉
    if len(all_tabs) > 1:
        # [1:] 表示跳过列表里的第一个(首页)，关闭后面的所有
        page.close_tabs(all_tabs[1:])

    # 因为 page 本身就一直绑定在首页上，直接让它刷新即可！
    page.refresh()
    # 刚刷新完首页，给它1秒钟的缓冲时间
    page.wait(1)
    print(">>> 首页已刷新，准备重新发起。")


def apply_search_conditions(tab):
    """
    功能：专门负责在这个页面里精准点击“结束”和“时间叉叉”。
    """
    print(">>> 正在等待页面数据加载稳定...")
    # 【核心防翻车】：强行等待 2 秒！
    # 这是应对 ERP 系统异步加载、局部刷新的最有效手段，让它把该刷新的部分彻底刷新完。
    tab.wait(2)

    print(">>> 正在设置筛选条件...")

    # 页面稳定后，出手点击“结束”
    tab.ele('@title=结束', timeout=15).click()
    print(">>> 已勾选 '结束' 状态...")

    # 同样，稳妥起见，找时间盒子和叉叉
    time_box = tab.ele('@data-criterion-key=docCreateTime', timeout=15)
    time_box.ele('@class=cancel', timeout=15).click()
    print(">>> 已清空 '创建时间' 限制...")


def setup_search_environment(page):
    """
    主控调度中心 (保持和之前一样)
    """
    max_try = 3

    for attempt in range(1, max_try + 1):
        print(f"\n[第 {attempt}/{max_try} 次尝试] 开始进入查询界面...")

        try:
            page.ele('text:流程查询', timeout=15).click()
            print(">>> 已点击 '流程查询'")

            tab = page.latest_tab

            tab.ele('text:施工委托招标(项目)', timeout=15).click()
            print(">>> 已进入施工委托列表...")

            # 调用修改后的积木块 B
            apply_search_conditions(tab)

            print("所有筛选条件加载成功！")
            return tab

        except Exception as e:
            # 打印具体的错误原因，以后排错更直观
            print(f"警报：第 {attempt} 次加载失败，原因：{e}")

            if attempt < max_try:
                # 调用修改后的积木块 A
                reset_and_back_to_home(page)
            else:
                print("无法进入搜索页面：页面超时！")
                raise Exception("Fatal Error：已达到最大重试次数，ValkyrieEngine 强制停机。")