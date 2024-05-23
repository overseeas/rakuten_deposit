from robocorp.tasks import task
from robocorp import windows
from robocorp import vault

desktop = windows.desktop()
@task
def minimal_task():
    #login_tokko()
    #tokko_main = windows.find_window('subname:"特攻店長 - メインメニュー"')
    #tokko_main.find("id:pbOrderList").click()
    set_search_config()



def login_tokko():
    desktop.windows_run("C:\Program Files (x86)\TokkoTencho\特攻店長\特攻店長.exe")
    login = vault.get_secret("tokko")
    tokko_login = windows.find_window('subname:"特攻店長 - ログイン"')
    tokko_login.find("id:txtUserID").set_value(login["id"])
    tokko_login.find("id:txtPassword").set_value(login["password"], validator = None)
    tokko_login.find("id:btnLogin").click()

def set_search_config():
    tokko_order = windows.find_window('subname:"特攻店長 - 受注一覧"')
    #initialize status
    #tokko_order.select("", locator="id:cmbStatus")
    #tokko_order.find("name:詳細検索▼").click()
    detailed_options = tokko_order.find_many("id:cmbSearchType", search_strategy="all", search_depth=6)
    detailed_options = detailed_options[:-1]

    detail_option(detailed_options)
    

    #name:詳細検索▼ id:btnChangeSearchDetail
    """
    for i in combobox:
        print(i)
    """

def detail_option(options):
    options[0].select("・処理状況（テキスト）")