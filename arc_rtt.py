#! /usr/bin/env python2.6
import sys, os
import signal, re, xlrd, subprocess, xlsxwriter

def signal_handler(sig, frame):
    print('Exiting gracefully Ctrl-C detected...')
    sys.exit()

def read_from_book():
    book = xlrd.open_workbook('EBA_Arc_status_report.xlsx')
    arc = book.sheet_by_name('EBA_Arc_status_report')
    return arc

def open_xls_to_write():
    book = xlsxwriter.Workbook('RA-EBA-Results.xlsx')
    min = book.add_worksheet('Min')
    max = book.add_worksheet('Max')
    avrg = book.add_worksheet('Avrg')
    return book, min, max, avrg

def close_xls_book(book):
    book.close()

def write_values(worksheet, value, row, column):
    worksheet.write(row, column, value)

def device_lookup(device):
    result = subprocess.Popen(
    ['grep -i ' + device + ' /etc/hosts'],
    stdout=subprocess.PIPE,
    shell=True)

    return result.stdout.read().split('#')[0].split()

def rtt_values(result):
    values = re.findall(r'\d*\.+\d*', result.stdout.read().strip())
    return values

def ping(arc_devices):
    device = arc_devices[int(round(len(arc_devices)/2))].strip()
    eba_ip = device_lookup(device)
    raH_ip = device_lookup(arc_devices[0].strip())
    raT_ip = device_lookup(arc_devices[len(arc_devices)-1].strip())

    result = subprocess.Popen(
        ['rcomauto ' + device + ' "ping ' + raH_ip[0] + ' source ' + eba_ip[0] + ' rapid timeout 1" | grep round-trip'],
        stdout=subprocess.PIPE,
        shell=True)
    raH_values = rtt_values(result)

    result = subprocess.Popen(
        ['rcomauto ' + device + ' "ping ' + raT_ip[0] + ' source ' + eba_ip[0] + ' rapid timeout 1" | grep round-trip'],
        stdout=subprocess.PIPE,
        shell=True)
    raT_values = rtt_values(result)

    return device, raH_values, raT_values

def main():
    arc = read_from_book()
    book, min, max, avrg = open_xls_to_write()
    for i in range(arc.nrows):
        arc_devices = arc.cell_value(i,1).split('>')
        device, raH_values, raT_values = ping(arc_devices)
        print device, raH_values, raT_values

        write_values(min, device, i, 0)
        write_values(avrg, device, i, 0)
        write_values(max, device, i, 0)

        write_values(min, arc_devices[0], i, 3)
        write_values(avrg, arc_devices[0], i, 3)
        write_values(max, arc_devices[0], i, 3)

        write_values(min, arc_devices[len(arc_devices)-1], i, 4)
        write_values(avrg, arc_devices[len(arc_devices)-1], i, 4)
        write_values(max, arc_devices[len(arc_devices)-1], i, 4)

        write_values(min, arc.cell_value(i,1), i, 5)
        write_values(avrg, arc.cell_value(i,1), i, 5)
        write_values(max, arc.cell_value(i,1), i, 5)

        write_values(min, raH_values[0], i, 1)
        write_values(avrg, raH_values[1], i, 1)
        write_values(max, raH_values[2], i, 1)
        write_values(min, raT_values[0], i, 2)
        write_values(avrg, raT_values[1], i, 2)
        write_values(max, raT_values[2], i, 2)


if __name__ == '__main__':
    signal.signal(signal.SIGINT, signal_handler)  # catch ctrl-c and call handler to terminate the script
    main()
