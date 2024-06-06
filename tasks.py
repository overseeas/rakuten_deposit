import sys
sys.path.append("../modules/Tokko_Tencho")
print(sys.path)
from robocorp.tasks import task
from robocorp import windows
from robocorp import vault
import Tokko_Tencho
import datetime

desktop = windows.desktop()
@task
def minimal_task():
    login_info = vault.get_secret("tokko")

    menu = Tokko_Tencho.Main()
    menu.open()
    menu.login(login_info["id"], login_info["password"])
    menu.select_menu("受注一覧")

    order = Tokko_Tencho.Order()
    order.initialize()
    order.option_text(0, "・処理状況（テキスト）", "入金待ち")
    order.option_list(1, "楽天ステータス", "発送待ち")
    if order.search() > 0:
        orders_list = order.get_values_from_list("ID")
        orders = ", ".join(map(str, orders_list))
        #Send mail
        order.list_all_click()
        order.send_mails("入金完了メール")

        #発注分類：受注発注
        status_change(orders, order, "受注発注")
        #発注分類：直送
        status_change(orders, order, "直送")
        #発注分類：在庫品
        status_change_direct(orders, order)

    menu.close()



def status_change(orders, order, option):
    order.initialize()
    order.option_text(0, "・処理状況（テキスト）", "入金待ち")
    order.option_list(1, "楽天ステータス", "発送待ち")
    order.input_order_id(orders)
    order.option_text(0,"・発注分類", option)
    if order.search() > 0:
        order.list_all_click()
        subwindow = order.open_bulk_change()
        subwindow.find("id:cbEditStatus").click()
        subwindow.find("id:cmbStatus").select("入金済み")
        confirming_change(order, subwindow)

def status_change_direct(orders, order):
    order.initialize()
    order.option_text(0, "・処理状況（テキスト）", "入金待ち")
    order.option_list(1, "楽天ステータス", "発送待ち")
    order.input_order_id(orders)
    if order.search() > 0:
        order.list_all_click()
        subwindow = order.open_bulk_change()
        subwindow.find("id:cbEditStatus").click()
        subwindow.find("id:cmbStatus").select("出荷日確定")
        subwindow.find("id:cbEditShipmentDueDate").click()
        confirming_change(order, subwindow)

def confirming_change(order, window):
    window.find("id:cbActionLog").click()
    window.find("id:txtActionLog").set_value(datetime.date.strftime(datetime.datetime.now(), "%m/%d　発送待ち確認"))
    window.find("id:btnEditPlural").click()
    order.wait()
    order.find("control:ButtonControl class:Button path:1|1|1").click()
    order.wait()
    order.find("control:ButtonControl class:Button path:1|1|1").click()
    order.wait()
