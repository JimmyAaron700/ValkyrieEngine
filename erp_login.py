from DrissionPage import ChromiumPage
import config  # 把刚才写的全局变量拉过来用


def login_erp():
    """负责登录ERP，返回登录成功后的浏览器控制权"""

    # 启动浏览器 (相当于开辟一个进程)
    page = ChromiumPage()

    # 开启 while(1) 无限循环防错机制
    while True:
        print("🚀 正在前往 ERP 登录页...")
        # .get() 相当于浏览器地址栏输入回车，如果已经在该页面，相当于刷新
        page.get(config.ERP_URL)

        # 核心：用 @name=xxx 来精准锁定元素并输入
        # .clear() 是为了防止刷新后框里有残留残留，先清空再输入，更稳
        page.ele('@name=j_username').clear().input(config.USERNAME)
        page.ele('@name=j_password').clear().input(config.PASSWORD)

        print(">>> 账号密码已自动填好！")
        print(">>> 请手动输入验证码，并点击登录进入首页。")

        # Python 里的 input 就等于 C 语言的 scanf，程序会在这里卡住，等你敲回车
        user_choice = input("登录成功并看到首页了吗？(输入 y 确认，输入 n 重新来过): ")

        # 判断你输入的指令 (.lower() 是为了防止你大写输入了 Y 或 N)
        if user_choice.lower() == 'y':
            print("✅ 登录模块验证通过，成功突破大门！")
            break  # 关键点：打破死循环，放行！

        elif user_choice.lower() == 'n':
            print("🔄 收到重试指令，准备刷新页面重填...")
            # 这里不用写代码，只要没有 break，while 循环就会回到开头重新执行 page.get()

        else:
            print("❓ 没看懂你的输入，为了安全起见，默认重试...")

    # 循环结束后，把浏览器的“遥控器”交还给主程序
    return page


# 这是一个调试开关：只有你单独运行这个文件时，下面两行才会执行。
# 如果是被 main.py 调用，这部分不会运行。
if __name__ == '__main__':
    login_erp()