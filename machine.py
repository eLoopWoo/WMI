import wmi
import _winreg
import win32api
import win32con


class machine(object):
    def __init__(self, ip=None):
        if ip:
            self.wmi = wmi.WMI(ip)
        else:
            self.wmi = wmi.WMI()

    def list_all_running_processes(self):
        running_processes = {}
        for process in self.wmi.Win32_Process():
            running_processes[process.ProcessId] = process.Name
        return running_processes

    def percentage_free_space_each_disk(self):
        disks = {}
        for disk in self.wmi.Win32_LogicalDisk(DriveType=3):
            disks[disk.Caption] = "%0.2f%% free" % (100.0 * long(disk.FreeSpace) / long(disk.Size))
        return disks

    def show_network_ip_mac_addresses(self):
        network = {}
        for interface in self.wmi.Win32_NetworkAdapterConfiguration(IPEnabled=1):
            # print interface.Description, interface.MACAddress
            flag = 1
            for ip_address in interface.IPAddress:
                if flag > 0:
                    temp = ip_address
                    flag *= -1
                else:
                    network[temp] = ip_address
                    flag *= -1
        return network

    def show_running_on_startup_paths(self):
        startup_paths = []
        for s in self.wmi.Win32_StartupCommand():
            startup_paths.append("[%s] %s <%s>" % (s.Location, s.Caption, s.Command))
        return startup_paths

    def list_shared_drives(self):
        shared_drives = {}
        for share in self.wmi.Win32_Share():
            shared_drives[share.Name] = share.Path
        return shared_drives

    def list_print_jobs(self):
        print_jobs = {}
        for printer in self.wmi.Win32_Printer():
            print printer.Caption
            for job in self.wmi.Win32_PrintJob(DriverName=printer.DriverName):
                print_jobs[printer.Caption] = job.Document
        return print_jobs

    def list_disk_partitions(self):
        hard_drives = {}
        for physical_disk in self.wmi.Win32_DiskDrive():
            hard_drives[physical_disk.Caption] = {}
            for partition in physical_disk.associators("Win32_DiskDriveToDiskPartition"):
                hard_drives[physical_disk.Caption][partition.Caption] = {}
                for logical_disk in partition.associators("Win32_LogicalDiskToPartition"):
                    hard_drives[physical_disk.Caption][partition.Caption][logical_disk.Caption] = {}
        return hard_drives

    def find_drive_types(self):
        DRIVE_TYPES = {
            0: "Unknown",
            1: "No Root Directory",
            2: "Removable Disk",
            3: "Local Disk",
            4: "Network Drive",
            5: "Compact Disc",
            6: "RAM Disk"
        }
        drives = {}
        for drive in self.wmi.Win32_LogicalDisk():
            drives[drive.Caption] = DRIVE_TYPES[drive.DriveType]
        return drives

    def run_process_minimised(self, proc):
        SW_SHOWMINIMIZED = 1
        startup = self.wmi.Win32_ProcessStartup.new(ShowWindow=SW_SHOWMINIMIZED)
        pid, result = self.wmi.Win32_Process.Create(
            CommandLine=proc,
            ProcessStartupInformation=startup
        )
        return pid

    def find_current_wallpaper(self):
        full_username = win32api.GetUserNameEx(win32con.NameSamCompatible)
        desktops = []
        for desktop in self.wmi.Win32_Desktop(Name=full_username):
            desktops.append((desktop.Wallpaper or "[No Wallpaper]",
                             desktop.WallpaperStretched, desktop.WallpaperTiled))
        return desktops

    def list_registry_keys(self):
        r = wmi.Registry()
        result, names = r.EnumKey(
            hDefKey=_winreg.HKEY_LOCAL_MACHINE,
            sSubKeyName="Software"
        )
        keys = []
        for key in names:
            keys.append(key)
        return keys