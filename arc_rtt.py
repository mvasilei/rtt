#! /usr/bin/env python2.6
import sys, os
import signal, re, xlrd, subprocess, xlsxwriter, csv

def signal_handler(sig, frame):
    print('Exiting gracefully Ctrl-C detected...')
    sys.exit()

def open_xls_to_write():
    book = xlsxwriter.Workbook('RA-EBA-Results.xlsx')
    min = book.add_worksheet('Min')
    max = book.add_worksheet('Max')
    avrg = book.add_worksheet('Avrg')
    return book, min, max, avrg

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
    count = 0
    book, min, max, avrg = open_xls_to_write()
    with open('EBA_Arc_status_report.csv') as csvfile:
        spamreader = csv.reader(csvfile, delimiter=':')
        for row in spamreader:
            count += 1
            if len(row) > 1 and 'UNCLOSED' not in row[1] and 'Count' not in row[0]:
                arc_devices = row[1].split('|')[0].split('>')

                device, raH_values, raT_values = ping(arc_devices)
                print device, raH_values, raT_values

                write_values(min, device, count, 0)
                write_values(avrg, device, count, 0)
                write_values(max, device, count, 0)

                write_values(min, arc_devices[0], count, 3)
                write_values(avrg, arc_devices[0], count, 3)
                write_values(max, arc_devices[0], count, 3)

                write_values(min, arc_devices[len(arc_devices)-1], count, 4)
                write_values(avrg, arc_devices[len(arc_devices)-1], count, 4)
                write_values(max, arc_devices[len(arc_devices)-1], count, 4)

                write_values(min, row[1], count, 5)
                write_values(avrg, row[1], count, 5)
                write_values(max, row[1], count, 5)

                write_values(min, raH_values[0], count, 1)
                write_values(avrg, raH_values[1], count, 1)
                write_values(max, raH_values[2], count, 1)
                write_values(min, raT_values[0], count, 2)
                write_values(avrg, raT_values[1], count, 2)
                write_values(max, raT_values[2], count, 2)

    book.close()

if __name__ == '__main__':
    signal.signal(signal.SIGINT, signal_handler)  # catch ctrl-c and call handler to terminate the script
    main()
