import pywinauto
import time
from infi.systray import SysTrayIcon
from ctypes import windll, Structure, c_long, byref


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
            try:
                app = pywinauto.Application().connect(title='Microsoft Outlook', class_name="#32770")['Microsoft Outlook']
                app.Allow.Wait('ready', retry_interval=0.1)
                app.Allow.Click()
                #Gross Hack: See https://stackoverflow.com/questions/50682448/pywinauto-click-fails-until-mouse-is-physically-clicked
                #Might be able to catch mouse clicks and stop it interfering with wherever (0, 0) happens to be?
                pos = queryMousePosition()
                pywinauto.mouse.release()
                pywinauto.mouse.move(coords=pos)
                #/Hack
                pass


            except (pywinauto.findbestmatch.MatchError, pywinauto.findwindows.ElementNotFoundError):
                time.sleep(0.1)
                pass

class POINT(Structure):
    _fields_ = [("x", c_long), ("y", c_long)]

def queryMousePosition():
    pt = POINT()
    windll.user32.GetCursorPos(byref(pt))
    return  pt.x, pt.y

SysTray.systray.start()
