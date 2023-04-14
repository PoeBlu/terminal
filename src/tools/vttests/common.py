################################################################################
#                                                                              #
# Copyright (c) Microsoft Corporation.
# Licensed under the MIT license.
#                                                                              #
################################################################################

import sys, os
import time # time.sleep is in seconds

if os.name == 'nt':
    import ctypes
    hOut = ctypes.windll.kernel32.GetStdHandle(-11)
    out_modes = ctypes.c_uint32()
    ENABLE_VT_PROCESSING = ctypes.c_uint32(0x0004)
    ctypes.windll.kernel32.GetConsoleMode(hOut, ctypes.byref(out_modes))
    out_modes = ctypes.c_uint32(out_modes.value | 0x0004)
    ctypes.windll.kernel32.SetConsoleMode(hOut, out_modes)
    setcp_result = ctypes.windll.kernel32.SetConsoleOutputCP(65001)
    if not setcp_result:
        gle = ctypes.windll.kernel32.GetLastError()
        print(f'SetConsoleOutputCP failed with error {gle}')
        exit(0)
    setcp_result = ctypes.windll.kernel32.SetConsoleCP(65001)
    if not setcp_result:
        gle = ctypes.windll.kernel32.GetLastError()
        print(f'SetConsoleCP failed with error {gle}')
        exit(0)
    import codecs
    codecs.register(lambda name: codecs.lookup('utf-8') if name == 'cp65001' else None)
    sys.stdout = codecs.getwriter('utf8')(sys.stdout)
    sys.stderr = codecs.getwriter('utf8')(sys.stderr)

def write(s):
    sys.stdout.write(s)

def esc(seq):
    write(f'\x1b{seq}')

def csi(seq):
    sys.stdout.write(f'\x1b[{seq}')

def osc(seq):
    sys.stdout.write(f'\x1b]{seq}\x07')

def cup(r=0, c=0):
    csi('H') if (r==0 and c==0) else csi(f'{r};{c}H')

def cupxy(x=0, y=0):
    cup(y+1, x+1)

def margins(top=0, bottom=0):
    csi(f'{top};{bottom}r')

def clear_all():
    cupxy(0,0)
    csi('2J')

def sgr(code=0):
    csi(f'{code}m')

def sgr_n(seq=[]):
    csi(f"{';'.join(str(code) for code in seq)}m")

def tbc():
    """
    Clear all tabstops from the terminal.
    Tab will take the cursor to the last column, then the first column.
    """
    csi('3g')

def ht():
    write('\t')

def cbt():
    csi('Z')

def hts(column=-1):
    if column > 0:
        csi(f';{column}H')
    esc('H')

def alt_buffer():
    csi('?1049h')

def main_buffer():
    csi('?1049l')

def flush(timeout=0):
    sys.stdout.flush()
    time.sleep(timeout)

def set_color(index, r, g, b):
    osc('4;{};rgb:{:02X}/{:02X}/{:02X}'.format(index, r, g, b))

if __name__ == '__main__':
    clear_all()
    print('This is the VT Test template.')
