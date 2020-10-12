from KGlobal import Toolbox
from os.path import join, dirname

import sys


if getattr(sys, 'frozen', False):
    app_path = sys.executable
    ico_dir = sys._MEIPASS
else:
    app_path = __file__
    ico_dir = dirname(__file__)

tool = Toolbox(app_path, logging_folder="01_Event_Logs", logging_base_name="Mail_Distro", max_pool_size=100)
attach_dir = join(tool.local_config_dir, '03_Attachments')
icon_path = join(ico_dir, 'Mail_Distro.ico')
icon_path2 = join(ico_dir, 'Mail_Distro_Settings.ico')

email_engine = tool.default_exchange_conn()
sql_engine = tool.default_sql_conn()
local_config = tool.local_config
