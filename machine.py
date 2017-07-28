import wmi
import _winreg
import win32api
import win32con


x = wmi.WMI()
for process in x.Win32_Process():
    print process.name

for disk in x.Win32_LogicalDisk(DriveType=3):
    print disk

for interface in x.Win32_NetworkAdapterConfiguration(IPEnabled=1):
    print interface

for s in x.Win32_StartupCommand():
    print s

for share in x.Win32_Share():
    print share

print x