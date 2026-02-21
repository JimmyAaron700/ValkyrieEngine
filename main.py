import config
import data_excel
import erp_login
import erp_construction_bidding
import erp_data_extractor


def start_valkyrie_engine():
    """
    ValkyrieEngine 主引擎启动程序
    """
    print("=" * 50)
    print("    🚀 ValkyrieEngine (女武神引擎) 核心启动 🚀")
    print("=" * 50)

    try:
        # 第一步：后勤部先动，读取并清洗 Excel 里的 ERP 编号
        print("\n>>> [系统指令 1] 启动数据预处理...")
        target_codes = data_excel.load_and_clean_data()

        # 如果表里一个有效编号都没有，直接停机，没必要去登录网页了
        if not target_codes:
            print("⚠️ 源表格中没有发现有效的 ERP 编号，引擎自动中止。")
            return

        # 第二步：突破大门，手动登录验证码
        print("\n>>> [系统指令 2] 请求接管浏览器与系统登录...")
        page = erp_login.login_erp()

        # 第三步：布置查询战场，设置“结束”和“时间”等条件
        print("\n>>> [系统指令 3] 正在进入施工委托招标界面并设置条件...")
        search_tab = erp_construction_bidding.setup_search_environment(page)

        # 第四步：核心数据收割大循环
        print("\n>>> [系统指令 4] 开启全自动搜索与数据抓取序列...")
        # 把刚才洗干净的 target_codes 喂给循环收割机
        final_results = erp_data_extractor.run_data_cycle(page, search_tab, target_codes)

        # 第五步：收尾结算，成果导出
        print("\n>>> [系统指令 5] 任务结束，开始打包导出成果...")
        data_excel.save_data_to_excel(final_results)

        print("\n" + "=" * 50)
        print("    🎉 ValkyrieEngine 全部任务完美执行完毕！ 🎉")
        print("=" * 50)

    except Exception as e:
        # 如果有任何子模块扔出了“致命红牌 (raise Exception)”，主控台会在这里稳稳接住，并安全停机
        print("\n" + "!" * 50)
        print(f"    🚨 引擎紧急停机！发生致命错误：\n    {e}")
        print("!" * 50)


# 启动开关
if __name__ == '__main__':
    start_valkyrie_engine()