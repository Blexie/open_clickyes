import pywinauto

def autoaccept():
    app = pywinauto.Application().connect(class_name="#32770")
    dlg = app['Microsoft Outlook']
    dlg.Allow.click()

autoaccept()
