from machine import Machine
from pprint import pprint


def main():
    me = Machine()
    pprint(me.get_info())


if __name__ == '__main__':
    main()
