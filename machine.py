import datetime
import os

import wmi
import _winreg
import win32api
import win32con


class Machine(object):
    def __init__(self, ip=None):
        if ip:
            self.wmi = wmi.WMI(ip)
        else:
            self.wmi = wmi.WMI()

    def get_operating_system(self):
        operating_systems = list()
        for os in self.wmi.Win32_OperatingSystem():
            operating_systems.append(os.Caption)
        return operating_systems

    def get_processes(self):
        processes = {}
        for process in self.wmi.Win32_Process():
            processes[process.ProcessId] = process.Name
        return processes

    def get_percentage_free_space_each_disk(self):
        disks = {}
        for disk in self.wmi.Win32_LogicalDisk(DriveType=3):
            disks[disk.Caption] = "%0.2f%% free" % (100.0 * long(disk.FreeSpace) / long(disk.Size))
        return disks

    def get_network_ip_mac_addresses(self):
        network = {}
        for interface in self.wmi.Win32_NetworkAdapterConfiguration(IPEnabled=1):
            interface_info = dict()
            for index, data in enumerate(interface.IPAddress):
                if index == 0:
                    interface_info['ipv4_address'] = data
                if index == 1:
                    interface_info['ip_mac_address'] = data
                if index == 2:
                    interface_info['ipv6_address'] = data
                else:
                    if 'undefined_data' not in interface_info:
                        interface_info['undefined_data'] = []
                    interface_info['undefined_data'].append(data)
            if interface_info:
                name = interface_info.values()[0]
                network[name] = interface_info
        return network

    def get_running_on_startup_paths(self):
        startup_paths = []
        for s in self.wmi.Win32_StartupCommand():
            startup_paths.append("[%s] %s <%s>" % (s.Location, s.Caption, s.Command))
        return startup_paths

    def get_shared_drives(self):
        shared_drives = {}
        for share in self.wmi.Win32_Share():
            shared_drives[share.Name] = share.Path
        return shared_drives

    def get_printer_info(self):
        printer_info = {}
        for printer in self.wmi.Win32_Printer():
            for job in self.wmi.Win32_PrintJob(DriverName=printer.DriverName):
                printer_info[printer.Caption] = job.Document
        return printer_info

    def get_disk_partitions(self):
        hard_drives = {}
        for physical_disk in self.wmi.Win32_DiskDrive():
            hard_drives[physical_disk.Caption] = {}
            for partition in physical_disk.associators("Win32_DiskDriveToDiskPartition"):
                hard_drives[physical_disk.Caption][partition.Caption] = {}
                for logical_disk in partition.associators("Win32_LogicalDiskToPartition"):
                    hard_drives[physical_disk.Caption][partition.Caption][logical_disk.Caption] = {}
        return hard_drives

    def get_drives_type(self):
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

    def get_current_wallpaper(self):
        desktops = []
        for desktop in self.wmi.Win32_Desktop():
            desktops.append((desktop.Wallpaper or "[No Wallpaper]",
                             desktop.WallpaperStretched, desktop.WallpaperTiled))
        return desktops

    @staticmethod
    def get_registry_keys():
        r = wmi.Registry()
        hklm = 0x80000002
        result, names = r.EnumKey(
            hDefKey=hklm,
            sSubKeyName="SOFTWARE"
        )
        keys = []
        for key in names:
            keys.append(key)
        return keys

    @staticmethod
    def set_registry_key(path):
        r = wmi.Registry()
        hklm = 0x80000002
        result, = r.CreateKey(
            hDefKey=hklm,
            sSubKeyName=r"{}".format(path)
        )
        return result

    @staticmethod
    def set_registry_value(path, value_name, value_content):
        r = wmi.Registry()
        hklm = 0x80000002
        result, = r.SetStringValue(
            hDefKey=hklm,
            sSubKeyName=r"{}".format(path),
            sValueName="{}".format(value_name),
            sValue="{}".format(value_content)
        )
        return result

    def set_schedule_job(self, command, minutes):
        one_minutes_time = datetime.datetime.now() + datetime.timedelta(minutes=minutes)
        job_id, result = self.wmi.Win32_ScheduledJob.Create(
            Command=r"{}".format(command),
            StartTime=wmi.from_time(one_minutes_time)
        )
        return job_id

    @staticmethod
    def get_schedule_jobs():
        schedule_info = dict()
        title = None
        for line in os.popen("schtasks.exe"):
            if line == '\n':
                title = None
                continue
            if not title:
                title = line
                schedule_info[title] = list()
            else:
                schedule_info[title].append(line)
        return schedule_info

    def get_info(self):
        machine_info = dict()
        machine_info['operating_system_info'] = self.get_operating_system()
        machine_info['network_info'] = self.get_network_ip_mac_addresses()
        machine_info['diks_info'] = self.get_disk_partitions()
        machine_info['proc_info'] = self.get_processes()
        machine_info['shared_info'] = self.get_shared_drives()
        machine_info['space_info'] = self.get_percentage_free_space_each_disk()
        machine_info['stratup_info'] = self.get_running_on_startup_paths()
        machine_info['printer_info'] = self.get_printer_info()
        machine_info['drive_info'] = self.get_drives_type()
        machine_info['wallpaper_info'] = self.get_current_wallpaper()
        machine_info['registry_info'] = self.get_registry_keys()
        machine_info['schedule_jobs_info'] = self.get_schedule_jobs()
        return machine_info
