import config
import data_excel
import erp_login
import erp_construction_bidding
import erp_construction_bidding_data_extractor

def get_run_mode():
    """
    运行模式配置模块
    功能：在引擎启动初期拦截控制流，初始化底层驱动的运行策略。
    """
    print("\n[系统配置] 请选择 ValkyrieEngine 运行时策略：")
    print("  1. 经典可视化模式 (前台渲染，适用于运行状态监控)")
    print("  2. 静默模式 (后台执行，进程隔离，防止影响正常操作)")

    while True:
        mode = input("[系统配置] 请输入策略编号 (1 或 2): ").strip()
        if mode in ['1', '2']:
            return mode
        print("[输入异常] 校验失败，仅支持输入字符 1 或 2，请重新输入。")


def feature_1_project_bidding():
    """
    功能模块 1：中标金额查询（项目维度）业务主程序
    功能：负责统筹调度各子模块，按既定业务流程完成数据预处理、系统登录、网页检索及结果导出。
    """
    print("\n" + "-" * 50)
    print("        开始执行 [模块 1：中标金额查询（项目维度）]")
    print("-" * 50)

    # 【架构规范】提前声明浏览器句柄
    # 防止在步骤1读取Excel报错时，finally 块调用未定义的 page 变量导致二次崩溃
    page = None

    try:
        # [V2.0.0 新增] 获取运行模式指令
        run_mode = get_run_mode()

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
        # 将运行模式参数下发至登录模块，由其决定是否进行会话迁移
        page = erp_login.login_erp(run_mode)

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
        final_results = erp_construction_bidding_data_extractor.run_data_cycle(page, search_tab, target_codes)

        # 步骤 5：成果导出与保存
        # 调用 data_excel 模块，将抓取产生的字典列表（包含成功数据及异常状态标记）
        # 转换为结构化的 DataFrame，并保存至配置文件指定的输出 Excel 路径。
        print("\n[系统执行 5/5] 数据提取完毕，正在导出结果文件...")
        data_excel.save_data_to_excel(final_results)

        print("\n" + "=" * 50)
        print("            模块 1 自动化处理任务执行完毕")
        print("=" * 50)

    except Exception as e:
        # 全局异常捕获机制
        # 拦截所有下层模块未能自行处理且向上抛出的严重异常（如彻底断网、达到最大重试次数上限等）。
        # 确保程序安全停机，并向控制台输出详细的错误追踪信息，便于后续排查。
        print("\n" + "!" * 50)
        print(f"          [系统异常] 程序因以下错误终止运行：\n    {e}")
        print("!" * 50)

    finally:
        # 【生命周期终结与资源回收】
        # 无论程序是正常执行到最后，还是中途被 return 阻断，抑或是由于严重报错进入 except 块，
        # 只要控制流即将离开该函数，系统就会强制进入 finally 块执行这里的资源清理。
        if page is not None:
            print("\n[系统维护] 正在执行浏览器生命周期终结与资源回收...")
            page.quit()
            print("[系统维护] 底层浏览器进程已安全彻底销毁，内存已释放。")


def main_engine_hub():
    """
    ValkyrieEngine 主控路由中枢
    功能：作为程序统一入口，提供多业务模块的导航菜单，支持平滑扩展新功能。
    """
    print("=" * 50)
    print("           ValkyrieEngine 自动化处理程序启动")
    print("=" * 50)

    while True:
        print("\n[主控中枢] 欢迎使用 ValkyrieEngine，请选择需要执行的业务功能：")
        print("  1. 中标金额查询（项目维度） [已上线]")
        print("  2. 中标金额查询（工程维度） [待开发]")
        print("  3. 盘点情况查询 [待开发]")
        print("  4. 结算情况查询 [待开发]")
        print("  5. 项目基础信息查询（ERP状态、总包、分包、项目经理） [待开发]")
        print("  6. 项目类别查询 [待开发]")
        print("  7. 退出系统")

        choice = input("\n[主控中枢] 请输入功能编号 (1-7): ").strip()

        if choice == '1':
            # 路由跳转：分配至功能 1 对应的业务线
            feature_1_project_bidding()
            break  # 业务执行完毕后结束程序。如需持续运行，可删除此 break 返回主菜单

        elif choice in ['2', '3', '4', '5', '6']:
            # 占位符：为后续开发预留的扩展接口
            print(f"\n[系统提示] 功能模块 {choice} 暂未上线，正在规划开发中，敬请期待...")

        elif choice == '7':
            print("\n[系统提示] 正在安全退出系统。")
            break

        else:
            print("\n[输入异常] 无法识别的指令，请输入 1-7 之间的有效数字。")


# 程序启动入口
if __name__ == '__main__':
    main_engine_hub()