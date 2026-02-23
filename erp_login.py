from DrissionPage import ChromiumPage, ChromiumOptions
import config


def login_erp(run_mode):
    """
    ERP 系统鉴权与会话管理模块 (V2.0.0)
    功能：拉起独立的图形化浏览器完成人工鉴权，并实现全维度的状态迁移。
    新增了视口防塌陷、内存防溢出、证书穿透等底层稳定性补丁，确保无头模式下的高可靠性。
    """
    print("[身份鉴权] 正在分配独立的图形化资源，初始化前置浏览器实例...")

    # [V2.0.0 升级] 强制弹出一个全新的独立浏览器窗口
    # 使用 auto_port() 避开默认的 9222 端口，防止被系统后台现有的 Chrome 进程静默劫持
    gui_options = ChromiumOptions()
    gui_options.auto_port()
    # 【新增】强制图形化界面启动时最大化，保持与 1.0 版本一致的视觉体验
    gui_options.set_argument('--start-maximized')
    page = ChromiumPage(gui_options)

    while True:
        print("[身份鉴权] 建立网络连接，请求登录网关...")
        # 访问配置文件中设定的 ERP 地址。如果当前已在该页面，将执行刷新操作清理缓存状态。
        page.get(config.ERP_URL)

        # 定位账户输入框与密码输入框。
        # 使用 clear() 清空可能存在的历史缓存数据，确保凭证输入准确无误。
        if page.ele('@name=j_username', timeout=2):
            page.ele('@name=j_username').clear().input(config.USERNAME)
            page.ele('@name=j_password').clear().input(config.PASSWORD)

        print("[身份鉴权] 账户凭证已自动填充。")
        print("[身份鉴权] 暂停执行：请手动输入验证码，并点击登录进入 ERP 首页。")

        # 挂起主线程，等待用户的终端输入确认。以确保页面确实已跳转至首页，而不是停留在错误提示页。
        user_choice = input("[系统鉴权] 请确认是否已成功登录并看到首页？(输入 y 确认继续，输入 n 重新刷新页面): ")

        # 将用户的输入统一转换为小写进行判定，增强大小写容错性
        if user_choice.lower() == 'y':
            print("[系统鉴权] 登录状态确认完毕，放行后续业务流程。")
            break  # 跳出校验循环，解除程序阻塞状态

        elif user_choice.lower() == 'n':
            print("[系统鉴权] 收到重试指令，正在重新加载登录页面...")
            # 流程未触发 break，循环将重置回起点，重新请求页面并再次输入凭证

        else:
            print("[系统鉴权] 无法识别的输入指令，为确保流程安全，默认执行重试操作...")

    # [V2.0.0 新增] 会话状态的全维度跨进程迁移逻辑
    if run_mode == '2':
        print("[进程调度] 检测到静默策略，启动全维度会话迁移序列...")

        # 1. 记录重定向后的真实业务主页 URL (URL Context Capture)
        real_home_url = page.url

        # 2. 提取全域 Cookie 字典
        session_cookies = page.cookies()

        # 3. 提取 LocalStorage 与 SessionStorage (Web Storage Serialization)
        # 针对现代框架，将前端存储数据通过 JS 引擎序列化为 JSON 字符串
        local_storage_data = page.run_js("return JSON.stringify(window.localStorage);")
        session_storage_data = page.run_js("return JSON.stringify(window.sessionStorage);")

        # 释放前置图形化进程占用的系统资源
        page.quit()
        print("[进程调度] 前置图形化主进程已销毁。正在注入稳定性参数...")

        # 实例化配置对象并注入稳定性补丁
        headless_options = ChromiumOptions()
        headless_options.auto_port()
        headless_options.headless(True)

        # --- 稳定性核心参数注入 ---
        # A. 固化桌面级视口，防止由于无头默认小窗口导致的响应式菜单折叠 (视口坍塌陷阱防御)
        headless_options.set_argument('--window-size=1920,1080')
        headless_options.set_argument(
            '--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')

        # B. 内存与进程隔离优化：突破沙盒共享内存限制，降低 OOM 崩溃率
        headless_options.set_argument('--disable-dev-shm-usage')
        headless_options.set_argument('--no-sandbox')

        # C. 性能优化与网络穿透：忽略自签名证书错误，屏蔽自动化控制特征
        headless_options.set_argument('--blink-settings=imagesEnabled=false')
        headless_options.set_argument('--disable-gpu')
        headless_options.set_argument('--ignore-certificate-errors')
        headless_options.set_argument('--disable-blink-features=AutomationControlled')
        # -----------------------------

        # 在后台重构一个全新的无头浏览器实例
        headless_page = ChromiumPage(headless_options)

        # 步骤 A：向真实的业务主页发起首次请求，建立合法的域名上下文
        print("[进程调度] 正在建立无头实例的域名上下文...")
        headless_page.get(real_home_url)

        # 步骤 B：域名上下文建立后，反序列化并注入凭证与前端存储数据
        print("[进程调度] 正在执行跨进程 Web 状态注入...")
        headless_page.set.cookies(session_cookies)

        inject_storage_js = """
            let ls = JSON.parse(arguments[0] || '{}');
            for (let k in ls) { window.localStorage.setItem(k, ls[k]); }

            let ss = JSON.parse(arguments[1] || '{}');
            for (let k in ss) { window.sessionStorage.setItem(k, ss[k]); }
        """
        headless_page.run_js(inject_storage_js, local_storage_data, session_storage_data)

        # 步骤 C：执行最终刷新，激活会话状态与完整 DOM 渲染
        print(f"[进程调度] 正在激活业务系统会话：{real_home_url}")
        headless_page.get(real_home_url)
        headless_page.wait.load_start()

        print("[进程调度] 身份凭证已深度激活，底层控制器控制权已交接至无头实例。")
        return headless_page

    else:
        # 经典策略分支：直接向上层调用栈返回原始的图形化浏览器句柄
        print("[进程调度] 经典策略已启用，维持当前可视状态，流程控制权释放。")
        return page


# 模块独立测试入口：仅当直接运行本文件时执行，被主控程序导入时将被忽略。
if __name__ == '__main__':
    login_erp('1')