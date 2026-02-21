from DrissionPage import ChromiumPage
import config  # 导入外部配置参数


def login_erp():
    """
    ERP 系统鉴权登录模块
    功能：初始化浏览器环境，自动填充登录凭证，并挂起等待用户完成人工验证（如验证码）。
    在用户确认登录成功后，返回浏览器实例供后续业务模块调用。
    """

    # 初始化 Chromium 浏览器实例，接管浏览器控制权
    page = ChromiumPage()

    # 开启状态校验循环，确保系统成功登录后才放行后续流程
    while True:
        print("[系统鉴权] 正在导航至 ERP 登录页面...")
        # 访问配置文件中设定的 ERP 地址。如果当前已在该页面，将执行刷新操作清理缓存状态。
        page.get(config.ERP_URL)

        # 定位账户输入框与密码输入框。
        # 使用 clear() 清空可能存在的历史缓存数据，确保凭证输入准确无误。
        page.ele('@name=j_username').clear().input(config.USERNAME)
        page.ele('@name=j_password').clear().input(config.PASSWORD)

        print("[系统鉴权] 账户凭证已自动填充。")
        print("[系统鉴权] 暂停执行：请手动输入验证码，并点击登录进入 ERP 首页。")

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

    # 将浏览器实例返回给主控程序，以便其他业务模块接力操作当前已登录的页面
    return page


# 模块独立测试入口：仅当直接运行本文件时执行，被主控程序导入时将被忽略。
if __name__ == '__main__':
    login_erp()