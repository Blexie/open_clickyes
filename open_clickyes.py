import pywinauto
import time
from infi.systray import SysTrayIcon
running = None
if running:
    icon="clickyes_running.ico"
else:
    icon="clickyes.ico"


def start_clickyes(systray):
    global running
    if running:
        print("Stopping Clickyes...")
        running = False
        pass
    else:
        print("Clickyes Running!")
        running = True
        clickyes()

    

def clean_exit(systray):
    global running
    running = False
    raise SystemExit

def clickyes():
    while running:
        time.sleep(1)
        try:
            app = pywinauto.Application().connect(title='Microsoft Outlook', class_name="#32770")
            dlg = app['Microsoft Outlook']
            dlg.Allow.click()
        except (pywinauto.findbestmatch.MatchError, pywinauto.findwindows.ElementNotFoundError, pywinauto.base_wrapper.ElementNotEnabled):
            pass




menu_options = (("Start", None, start_clickyes),)
systray = SysTrayIcon(icon, "Open ClickYes v0.1", menu_options, on_quit=clean_exit)
systray.start()
