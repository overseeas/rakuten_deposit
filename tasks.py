import sys
sys.path.append("C:/Users/LINK_RPA_001/Desktop/robocorp/modules")
print(sys.path)
from robocorp.tasks import task
from robocorp import windows
from robocorp import vault
import Tokko_Tencho.Main
import Tokko_Tencho.Order
import datetime

desktop = windows.desktop()
@task
def minimal_task():

    today = datetime.datetime.now()
    
    """login = vault.get_secret("tokko")
    print("<")
    Tokko_Tencho.Main.open()
    print(">")
    Tokko_Tencho.Main.login(login["id"], login["password"])
    Tokko_Tencho.Main.select_menu("受注一覧")
    Tokko_Tencho.Order.initialize()
    Tokko_Tencho.Order.option_text(0, "・処理状況（テキスト）", "入金待ち")
    Tokko_Tencho.Order.option_list(1, "楽天ステータス", "発送待ち")
    if Tokko_Tencho.Order.search() > 0:
        #Send mail
        Tokko_Tencho.Order.list_all_click()
        Tokko_Tencho.Order.send_mails("入金完了メール")

        #発注分類：受注発注
        Tokko_Tencho.Order.option_text(2,"発注分類","受注発注")
        if Tokko_Tencho.Order.search() > 0:
            Tokko_Tencho.Order.list_all_click()"""
    order_window = windows.find_window('subname:"特攻店長 - 受注一覧"')
    order_window.find("id:btnEditPlural").click()
    subwindow = order_window.find_child_window("id:FrmOrderBatchUpdater")
    subwindow.find("id:cbEditStatus").click()
    subwindow.find("id:cmbStatus").select("入金済み")
    subwindow.find("id:cbActionLog").click()
    subwindow.find("id:txtActionLog").set_value(datetime.date.strftime(datetime.datetime.now(), "%m/%d　発送待ち確認"))