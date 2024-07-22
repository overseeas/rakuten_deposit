import sys
sys.path.append("../modules/Tokko_Tencho")
from robocorp.tasks import task
from robocorp import windows
from robocorp import vault
import Tokko_Tencho
import datetime
from robocorp import log
from RPA.Email.ImapSmtp import ImapSmtp
from RPA.Outlook.Application import Application
import traceback
import holidays

desktop = windows.desktop()
@task
def minimal_task():

    #check if today is working date
    if datetime.date.strftime(datetime.datetime.now(), "%y-%m-%d") in holidays.JP() or datetime.datetime.today().weekday() in [5,6]:
        return False
    
    login_info = vault.get_secret("Tokko")
    bank = False
    convinience = False    

    #銀行
    menu = Tokko_Tencho.Main()
    menu.open()
    try:
        menu.login(login_info["id"], login_info["password"])
        menu.select_menu("受注一覧")

        order = Tokko_Tencho.Order()
        order.initialize()
        order.option_text(0, "・処理状況（テキスト）", "入金待ち")
        order.option_list(1, "楽天ステータス", "発送待ち")
        order.input_order_number("r")
        if order.search() > 0:
            orders_list = order.get_values_from_list("ID")
            orders = ", ".join(map(str, orders_list))

            #発注分類：受注発注
            status_change(orders, order, "受注発注", "入金待ち")
            #発注分類：直送
            status_change(orders, order, "直送", "入金待ち")
            #発注分類：在庫品
            status_change_direct(orders, order, "入金待ち")
        bank = True
    except:
        traceback.print_tb()
    menu.close()
    

    #コンビニ
    menu = Tokko_Tencho.Main()
    menu.open()
    try:
        menu.login(login_info["id"], login_info["password"])
        menu.select_menu("受注一覧")

        order = Tokko_Tencho.Order()
        order.initialize()
        order.option_text(0, "・処理状況（テキスト）", "収納依頼")
        order.option_list(1, "楽天ステータス", "発送待ち")
        if order.search() > 0:
            orders_list = order.get_values_from_list("ID")
            orders = ", ".join(map(str, orders_list))
            #Send mail
            order.list_all_click()
            order.send_mails("入金完了メール")

            #発注分類：受注発注
            status_change(orders, order, "受注発注", "収納依頼")
            #発注分類：直送
            status_change(orders, order, "直送", "収納依頼")
            #発注分類：在庫品
            status_change_direct(orders, order, "収納依頼")
        convinience = True
    except:
        traceback.print_tb
    menu.close()

    if convinience and bank:
        mailto("chumon@maido-diy.jp", "★完了報告★楽天入金処理完了報告", "楽天の銀行、コンビニ払いの入金処理しました。")
    else:
        mailto("chumon@maido-diy.jp", "★エラー報告★楽天入金処理完了報告", "楽天の銀行、コンビニ払いの入金処理に失敗しました。")

def mailto(to, subject, body):
    secrets = vault.get_secret("Mail")
    app = Application()
    app.open_application()
    app.send_email(
        recipients=to,
        cc_recipients=["yang@proszet.com", "funaki@proszet.com"],
        subject=subject,
        body=body
    )

def status_change(orders, order, option, status):
    order.initialize()
    order.option_text(0, "・処理状況（テキスト）", status)
    order.option_list(1, "楽天ステータス", "発送待ち")
    order.input_order_id(orders)
    order.option_text(2,"・発注分類", option)
    if order.search() > 0:
        order.list_all_click()
        subwindow = order.open_bulk_change()
        subwindow.find("id:cbEditStatus").click()
        subwindow.find("id:cmbStatus").select("入金済み")
        confirming_change(order, subwindow)

def status_change_direct(orders, order, status):
    order.initialize()
    order.option_text(0, "・処理状況（テキスト）", status)
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
    assert window.find("control:ButtonControl class:Button path:1|1")
    window.find("control:ButtonControl class:Button path:1|1").click()
    order.wait()
    assert window.find("control:ButtonControl class:Button path:1|1")
    window.find("control:ButtonControl class:Button path:1|1").click()
    order.wait()
