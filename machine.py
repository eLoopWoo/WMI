import wmi
import _winreg
import win32api
import win32con


x = wmi.WMI()
for process in x.Win32_Process():
    print process.name

