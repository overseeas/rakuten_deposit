from robocorp.tasks import task
from robocorp import windows
from robocorp import vault

desktop = windows.desktop()
@task
def minimal_task():
    # login_tokko()
    tokko_main = windows.find_window('subname:"特攻店長 - メインメニュー"')
    tokko_main.find("id:pbOrderList").click()
    
    tokko_order = windows.find_window('subname:"特攻店長 - 受注一覧"')
    tokko_order.inspect()


def login_tokko():
    desktop.windows_run("C:\Program Files (x86)\TokkoTencho\特攻店長\特攻店長.exe")
    login = vault.get_secret("tokko")
    tokko_login = windows.find_window('subname:"特攻店長 - ログイン"')
    tokko_login.find("id:txtUserID").set_value(login["id"])
    tokko_login.find("id:txtPassword").set_value(login["password"], validator = None)
    tokko_login.find("id:btnLogin").click()
