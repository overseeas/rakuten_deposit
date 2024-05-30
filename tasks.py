import sys
sys.path.append("C:/Users/LINK_RPA_001/Desktop/robocorp/modules")
print(sys.path)
from robocorp.tasks import task
from robocorp import windows
from robocorp import vault
import Tokko_Tencho.main
import Tokko_Tencho.order

desktop = windows.desktop()
@task
def minimal_task():
    """
    login = vault.get_secret("tokko")
    Tokko_Tencho.main.login_tokko(login["id"], login["password"])
    Tokko_Tencho.main.select_menu("受注一覧")
    Tokko_Tencho.order.initialize()
    Tokko_Tencho.order.option_text(0, "・処理状況（テキスト）", "入金待ち")
    Tokko_Tencho.order.option_list(1, "楽天ステータス", "発送待ち")
    """
    order_nums = Tokko_Tencho.order.search()