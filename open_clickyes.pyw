import time
from ctypes import windll, Structure, c_long, byref
from tkinter import *
import pywinauto
from infi.systray import SysTrayIcon

class SysTray:
    icon="clickyes.ico"
    def start_clickyes(systray):
        if ClickYes.running:
            #print("Stopping ClickYes.")
            ClickYes.running = False
            SysTray.update_systray_stopped()
        else:
            #print("Starting ClickYes!")
            ClickYes.running = True
            SysTray.update_systray_running()
            ClickYes.toggle()

    def clean_exit(systray):
        raise SystemExit

    def update_systray_running():
        SysTray.systray.update(icon="clickyes_running.ico")

    def update_systray_stopped():
        SysTray.systray.update(icon="clickyes.ico")

    menu_options = (("Toggle (double-click)", None, start_clickyes),)
    systray = SysTrayIcon(icon, "Open ClickYes v0.1", menu_options, on_quit=clean_exit)


class ClickYes:
    running = ""
    def toggle():
        while ClickYes.running:
            try:
                app = pywinauto.Application().connect(title='Microsoft Outlook', class_name="#32770")#['Microsoft Outlook']
                dlg = app['Microsoft Outlook']
                dlg.Allow.Wait('ready', retry_interval=0.1)
                dlg.Allow.Click()
                #Gross Hack: See https://stackoverflow.com/questions/50682448/pywinauto-click-fails-until-mouse-is-physically-clicked
                #Might be able to catch mouse clicks and stop it interfering with wherever (0, 0) happens to be?
                pos = queryMousePosition()
                pywinauto.mouse.click()
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
#Second part of hack to fix mouse not responding...#
class SplashScreen(Frame):
    def __init__(self, master=None, width=1, height=1, useFactor=True, takefocus = False):
        Frame.__init__(self, master)
        self.pack(side=TOP, fill=BOTH, expand=YES)

        # get screen width and height
        w = 3
        h = 3
        # calculate position x, y
        x = 0
        y = 0
        self.master.geometry('%dx%d+%d+%d' % (w, h, x, y))
        
        self.master.overrideredirect(True)
        self.lift()

SysTray.systray.start()

if __name__ == '__main__':
    root = Tk()

    sp = SplashScreen(root)
    sp.config(bg="#3366ff")

    m = Label(sp, text="Dummy splash for Open_ClickYes")
    m.pack(side=TOP, expand=YES)
    m.config(bg="#3366ff", justify=CENTER, font=("calibri", 29))
    root.call('wm', 'attributes', '.', '-topmost', '1')
    root.mainloop()
#end
