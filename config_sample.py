import windowmanager
import constants as C

windowmanager.ignore_list = [
    {"class_name": "Windows.UI.Core.CoreWindow"},
    {"class_name": "TaskManagerWindow"},
    {"class_name": "Microsoft-Windows-SnipperToolbar"},
    {"class_name": "Qt5QWindowIcon", "title": "GtransWeb"},
    {"class_name": "screenClass"}
    ]
windowmanager.workspace_n = 9
windowmanager.network_interface = "Dell Wireless 1820A 802.11ac"

append_list = [
        {"mod": C.NOCONVERT | C.WIN, "key": "a", "function": windowmanager.focus_up,
            "param": [+1]}
        ]
