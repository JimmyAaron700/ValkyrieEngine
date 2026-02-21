import config
import data_excel
import erp_login
import erp_construction_bidding
import erp_data_extractor


def start_valkyrie_engine():
    """
    ValkyrieEngine 自动化处理主程序
    功能：负责统筹调度各子模块，按既定业务流程完成数据预处理、系统登录、网页检索及结果导出。
    """
    print("=" * 50)
    print("              ValkyrieEngine 自动化处理程序启动")
    print("=" * 50)

    try:
        # 步骤 1：数据预处理
        # 调用 data_excel 模块，从配置文件指定的 Excel 中读取 ERP 项目编号。
        # 此步骤包含格式校验与清洗机制，剔除不符合规范的脏数据。
        print("\n[系统执行 1/5] 开始数据预处理...")
        target_codes = data_excel.load_and_clean_data()

        # 数据校验拦截：若有效编号列表为空，则无后续执行必要，直接终止程序。
        if not target_codes:
            print("[提示] 源表格中未发现有效的 ERP 编号，程序终止运行。")
            return

        # 步骤 2：系统鉴权与登录
        # 调用 erp_login 模块初始化浏览器实例（ChromiumPage）。
        # 程序将自动填充账号密码，并挂起等待用户手动完成验证码验证。
        print("\n[系统执行 2/5] 启动浏览器并执行系统登录...")
        page = erp_login.login_erp()

        # 步骤 3：初始化查询环境
        # 调用 erp_construction_bidding 模块，控制浏览器导航至“施工委托招标”目标页面。
        # 自动执行前置动作：勾选“结束”状态，清除默认的创建时间限制，准备就绪。
        print("\n[系统执行 3/5] 正在进入目标查询界面并设置筛选条件...")
        search_tab = erp_construction_bidding.setup_search_environment(page)

        # 步骤 4：核心数据提取循环
        # 调用 erp_data_extractor 模块，将清洗后的 target_codes 列表传入。
        # 模块内部将循环执行：输入编号 -> 检索 -> 判定结果唯一性 -> 进入详情页抓取数据 -> 关闭详情页。
        # 附带页面异常自愈机制，防止单次查询卡顿导致程序崩溃。
        print("\n[系统执行 4/5] 开启自动化搜索与数据提取流程...")
        final_results = erp_data_extractor.run_data_cycle(page, search_tab, target_codes)

        # 步骤 5：成果导出与保存
        # 调用 data_excel 模块，将抓取产生的字典列表（包含成功数据及异常状态标记）
        # 转换为结构化的 DataFrame，并保存至配置文件指定的输出 Excel 路径。
        print("\n[系统执行 5/5] 数据提取完毕，正在导出结果文件...")
        data_excel.save_data_to_excel(final_results)

        print("\n" + "=" * 50)
        print("              ValkyrieEngine 自动化处理任务执行完毕")
        print("=" * 50)

    except Exception as e:
        # 全局异常捕获机制
        # 拦截所有下层模块未能自行处理且向上抛出的严重异常（如彻底断网、达到最大重试次数上限等）。
        # 确保程序安全停机，并向控制台输出详细的错误追踪信息，便于后续排查。
        print("\n" + "!" * 50)
        print(f"          [系统异常] 程序因以下错误终止运行：\n    {e}")
        print("!" * 50)


# 程序启动入口
if __name__ == '__main__':
    start_valkyrie_engine()