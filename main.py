from machine import Machine
from pprint import pprint


def main():
    import wmi
    c = wmi.WMI()
    for os in c.Win32_OperatingSystem():
        print os.Caption

    me = Machine()
    # pprint(me.list_registry_keys())
    pprint(me.get_info())


if __name__ == '__main__':
    main()
