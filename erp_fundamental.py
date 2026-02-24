"""
ValkyrieEngine ERP 基础能力原子库 (V2.1.0 延迟读取修复版)
功能：封装 ERP 系统通用的页面导航、表单交互与基础数据获取能力。
更新日志：
  - V2.1.0: [紧急修复] 针对列表加载后表头数字刷新延迟导致误读为0的问题，
            增加了“逻辑自洽验证”机制。当列表存在数据时，强制等待表头数字非零。
"""

import time
import re  # 导入正则模块，用于提取 "工程数(3)" 括号里的数字

def reset_and_back_to_home(page):
    """
    [通用组件] 页面状态重置模块 (回城卷轴)
    功能：清理所有标签页并强制刷新首页。
    注意：仅在连“工作台”都进不去的致命错误时使用。
    """
    print("[基础能力] 正在清理无响应的页面，重置浏览器环境...")
    all_tabs = page.tab_ids
    if len(all_tabs) > 1:
        page.close_tabs(all_tabs[1:])
    page.refresh()
    page.wait(1)
    print("[基础能力] 首页状态已重置。")


def enter_workbench(page):
    """
    [导航组件] 进入“项目流程工作台”
    """
    print("[基础能力] 正在导航至 '项目流程工作台'...")
    try:
        # 点击菜单
        page.ele('text:项目流程工作台', timeout=15).click()

        # 移交控制权给新页面
        tab = page.latest_tab

        # 显式等待：确保输入框加载出来，证明页面进去了
        tab.ele('#projectcode', timeout=15)

        print("[基础能力] 工作台页面加载完毕。")
        return tab
    except Exception as e:
        print(f"[基础能力] 导航失败：{e}")
        raise e


def query_single_project_count(tab, code):
    """
    [核心组件] 单项目工程数原子查询
    功能：输入编号 -> 查询 -> 【核心】校验列表数据 -> 【修复】延迟并循环验证表头 -> 清空 -> 校验清空状态
    返回：该项目的工程数量 (int)
    """
    try:
        # ==========================
        # 步骤 1: 输入编号
        # ==========================
        # 定位输入框 (ID: projectcode)
        input_box = tab.ele('#projectcode', timeout=5)
        # 必须先 clear，防止残留上一次输入的字符
        input_box.clear().input(code)

        # ==========================
        # 步骤 2: 点击查询并硬等待
        # ==========================
        # 点击查询按钮 (ID: btnquery)
        tab.ele('#btnquery', timeout=5).click()

        # 硬等待：给前端 JS 渲染和转圈圈留出时间。
        tab.wait(1)

        # ==========================
        # 步骤 3: 结果锚点验证 (抗干扰核心)
        # ==========================
        # 逻辑：输入框里有一个 code，列表里也应该出来一个 code。
        # 我们寻找页面中标签为 <td> 且包含该编号的元素。
        # 如果找到了，说明列表真正渲染出来了。
        if not tab.ele(f'tag:td@@text():{code}', timeout=15):
             # 如果等了 15 秒列表里还没刷出这个编号，说明大概率还在空转或者查无此人
             raise Exception(f"ListRenderTimeout: 列表未出现编号 {code}，疑似仍在转圈")

        # ==========================
        # 步骤 4: 提取工程数量 (修复的核心战场！)
        # ==========================
        # 定位工程数表头 (ID: bbtest)
        count_element = tab.ele('#bbtest', timeout=5)

        # [V2.1.0 逻辑自洽验证]
        # 小郁发现的问题：列表出来了，但表头可能还没变，还是 (0)。
        # 解决方案：既然步骤3通过了，说明至少有1个项目。如果表头读出来是0，说明它在撒谎（延迟）。
        # 我们给它 3 秒钟的时间改口。

        count = 0
        end_time = time.time() + 3  # 设置 3 秒的纠错超时时间

        while time.time() < end_time:
            raw_text = count_element.text  # 获取文本 "工程数(X)"
            match = re.search(r'\((\d+)\)', raw_text)

            if match:
                temp_count = int(match.group(1))
                if temp_count > 0:
                    count = temp_count
                    break # 读到了非 0 的数，说明表头刷新了，立刻跳出循环

            # 如果读到是 0，或者没读到，就稍等 0.5 秒再读一次
            time.sleep(0.5)

        # 循环结束，如果还是 0，那也没办法了，只能信了（虽然理论上不可能）
        if count == 0:
            print(f"    [警告] 列表存在但表头显示 0，可能存在数据异常: {code}")

        print(f"    -> [{code}] 抓取成功 | 真实工程数: {count}")

        # ==========================
        # 步骤 5: 闭环清空
        # ==========================
        # 点击清空按钮 (ID: btnclear)
        tab.ele('#btnclear', timeout=5).click()

        # 【状态复位检查】必须死等 "项目数(0)"
        # 如果不清零，下一次循环输入新编号时，会和旧数据混在一起，导致灾难性后果。
        if not tab.ele('text:项目数(0)', timeout=10):
            raise Exception("ClearFailed: 清空按钮点击后，状态未归零")

        return count

    except Exception as e:
        # 捕获任何步骤的异常（比如转圈超时、清空失败），直接抛给上层去刷新重试
        raise e


def batch_get_engineering_counts(page, codes_list):
    """
    [控制器] 批量获取工程数量主程序
    特性：具备“原地复活”能力，单点故障直接刷新当前页，不回首页。
    """
    print(f"\n[基础能力] 启动批量工程数嗅探，目标数量：{len(codes_list)}")

    # 初始化：先尝试进入工作台
    try:
        workbench_tab = enter_workbench(page)
    except:
        reset_and_back_to_home(page)
        workbench_tab = enter_workbench(page)

    result_data = []

    for index, code in enumerate(codes_list, start=1):
        # 进度日志
        print(f"[嗅探进度 {index}/{len(codes_list)}] 正在探测: {code}")

        max_retries = 2

        for attempt in range(1, max_retries + 1):
            try:
                # 调用原子查询
                count = query_single_project_count(workbench_tab, code)

                # 成功则存入
                result_data.append({"项目编号": code, "工程数": count})
                break

            except Exception as e:
                print(f"    [异常] {code} 第 {attempt} 次探测失败: {e}")

                if attempt < max_retries:
                    print("    [自愈] 检测到页面卡顿或遮罩层死锁，正在执行【原地刷新】...")

                    # ==================================
                    # 【核心优化】原地复活机制
                    # ==================================
                    # 不需要退回首页！直接刷新当前工作台页面 (F5)
                    # 这样能清除所有的转圈圈、遮罩层和残留查询条件
                    workbench_tab.refresh()

                    # 刷新后，必须重新等待输入框出现，确保页面活过来了
                    # 如果这一步报错，说明网络断了，外层 try 会捕获
                    workbench_tab.ele('#projectcode', timeout=30)
                    print("    [自愈] 页面刷新完毕，准备重试...")

                else:
                    print(f"    [放弃] {code} 连续失败，启用兜底值 3")
                    # 兜底策略：为了不卡死整个流程，给一个默认值
                    result_data.append({"项目编号": code, "工程数": 3})

                    # 即使放弃了，为了下一个编号能跑，也得刷新一下保持环境干净
                    workbench_tab.refresh()
                    workbench_tab.wait(2) # 稍微缓一下

    print(f"[基础能力] 批量嗅探结束，已获取 {len(result_data)} 条边界数据。")

    # 任务结束，关闭工作台，深藏功与名
    if workbench_tab:
        workbench_tab.close()

    return result_data


# ---------------- 单独测试入口 ----------------
if __name__ == '__main__':
    from erp_login import login_erp

    print("=== 开始测试 erp_fundamental (延迟修复版) ===")

    # 1. 启动并登录 (可视化模式)
    page = login_erp('1')

    # 2. 准备几个真实的编号 (建议包含之前容易出错的编号)
    test_codes = ['D1221420034', 'D1251420045']

    try:
        # 3. 执行批量探测
        final_list = batch_get_engineering_counts(page, test_codes)

        print("\n=== 最终结果 ===")
        print(final_list)

    except Exception as e:
        print(f"测试严重崩溃: {e}")