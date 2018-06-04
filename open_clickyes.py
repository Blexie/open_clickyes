import pywinauto
import time
from infi.systray import SysTrayIcon


class SysTray:
    icon = "clickyes.ico"
    def start_clickyes(systray):
        if ClickYes.running:
            print("Stopping ClickYes.")
            ClickYes.running = False
            SysTray.update_systray_stopped()
        else:
            print("Starting ClickYes!")
            ClickYes.running = True
            SysTray.update_systray_running()
            ClickYes.toggle()

    def clean_exit(systray):
        raise SystemExit

    def update_systray_running():
        SysTray.systray.update(icon = "clickyes_running.ico")

    def update_systray_stopped():
        SysTray.systray.update(icon = "clickyes.ico")

    menu_options = (("Toggle (double-click)", None, start_clickyes),)
    systray = SysTrayIcon(icon, "Open ClickYes v0.1", menu_options, on_quit=clean_exit)


class ClickYes:
    running = ""
    def toggle():
        while ClickYes.running:
            time.sleep(1)
            try:
                app = pywinauto.Application().connect(title='Microsoft Outlook', class_name="#32770")
                dlg = app['Microsoft Outlook']
                dlg.Allow.click()
            except (pywinauto.findbestmatch.MatchError, pywinauto.findwindows.ElementNotFoundError, pywinauto.base_wrapper.ElementNotEnabled):
                pass

SysTray.systray.start()
