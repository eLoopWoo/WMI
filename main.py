from machine import machine
from pprint import pprint

def main():
    me = machine()
    neighbors_info = me.show_network_ip_mac_addresses()
    diks_info = me.list_disk_partitions()
    pprint(neighbors_info)
    print "-"*80
    pprint(diks_info)

if __name__ == '__main__':
    main()