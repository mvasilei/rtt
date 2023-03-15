#! /usr/bin/env python2.6
import sys
import signal, re, xlrd, subprocess, xlsxwriter

def signal_handler(sig, frame):
    print('Exiting gracefully Ctrl-C detected...')
    sys.exit()

def read_from_book():
    book = xlrd.open_workbook('RA_Topology_lookup.xlsx')
    sheet = book.sheet_by_name('RA Topology')
    return sheet

def open_xls_to_write():
    book = xlsxwriter.Workbook('RA_Results.xlsx')
    min = book.add_worksheet('Min')
    max = book.add_worksheet('Max')
    avrg = book.add_worksheet('Avrg')

    return book, min, max, avrg

def rtt_values(result):
    values = re.findall(r'\d+\.+\d*', result.stdout.read().strip())
    print result.stdout.read().strip()
    return values

def ping(ra_name, px_name):
    result = subprocess.Popen(
        ["grep -i " + px_name + " /etc/hosts | awk {'print $1'}"],
        stdout=subprocess.PIPE,
        shell=True)
    px_ip = result.stdout.read().strip()

    result = subprocess.Popen(
        ["grep -i " + ra_name + " /etc/hosts | awk {'print $1'}"],
        stdout=subprocess.PIPE,
        shell=True)
    ra_ip = result.stdout.read().strip()

    result = subprocess.Popen(
        ['rcomauto ' + ra_name + ' "ping ' + px_ip + ' source ' + ra_ip + ' rapid timeout 1" | grep round-trip'],
        stdout=subprocess.PIPE,
        shell=True)
    ra_rtt = rtt_values(result)
    return ra_rtt

def write_values(worksheet, value, row, column):
    worksheet.write(row, column, value)

def main():
    xls_row = 0
    sheet = read_from_book()
    book, min, max, avrg = open_xls_to_write()
    for i in range(1,sheet.nrows):
        if i%2:
            print 'Collecting values for ' + sheet.cell_value(i, 5) + '>' + sheet.cell_value(i, 1)
            write_values(min, sheet.cell_value(i, 5) + '>' + sheet.cell_value(i, 1), xls_row, 0)
            write_values(avrg, sheet.cell_value(i, 5) + '>' + sheet.cell_value(i, 1), xls_row, 0)
            write_values(max, sheet.cell_value(i, 5) + '>' + sheet.cell_value(i, 1), xls_row, 0)

            ra_rtt = ping(sheet.cell_value(i, 5), sheet.cell_value(i, 1))
            write_values(min, ra_rtt[0], xls_row, 1)
            write_values(avrg, ra_rtt[1], xls_row, 1)
            write_values(max, ra_rtt[2], xls_row, 1)
            xls_row += 1

            print 'Collecting values for ' + sheet.cell_value(i, 5) + '>' + sheet.cell_value(i+1, 1)
            write_values(min, sheet.cell_value(i, 5) + '>' + sheet.cell_value(i+1, 1), xls_row, 0)
            write_values(avrg, sheet.cell_value(i, 5) + '>' + sheet.cell_value(i+1, 1), xls_row, 0)
            write_values(max, sheet.cell_value(i, 5) + '>' + sheet.cell_value(i+1, 1), xls_row, 0)

            ra_rtt = ping(sheet.cell_value(i, 5), sheet.cell_value(i, 1))
            write_values(min, ra_rtt[0], xls_row, 1)
            write_values(avrg, ra_rtt[1], xls_row, 1)
            write_values(max, ra_rtt[2], xls_row, 1)
            xls_row += 1
        else:
            if sheet.cell_value(i, 5) != sheet.cell_value(i - 1, 5):
                print 'Collecting values for ' + sheet.cell_value(i, 5) + '>' + sheet.cell_value(i, 1)
                write_values(min, sheet.cell_value(i, 5) + '>' + sheet.cell_value(i, 1), xls_row, 0)
                write_values(avrg, sheet.cell_value(i, 5) + '>' + sheet.cell_value(i, 1), xls_row, 0)
                write_values(max, sheet.cell_value(i, 5) + '>' + sheet.cell_value(i, 1), xls_row, 0)

                ra_rtt = ping(sheet.cell_value(i, 5), sheet.cell_value(i, 1))
                write_values(min, ra_rtt[0], xls_row, 1)
                write_values(avrg, ra_rtt[1], xls_row, 1)
                write_values(max, ra_rtt[2], xls_row, 1)
                xls_row += 1

                print 'Collecting values for ' + sheet.cell_value(i, 5) + '>' + sheet.cell_value(i-1, 1)
                write_values(min, sheet.cell_value(i, 5) + '>' + sheet.cell_value(i-1, 1), xls_row, 0)
                write_values(avrg, sheet.cell_value(i, 5) + '>' + sheet.cell_value(i-1, 1), xls_row, 0)
                write_values(max, sheet.cell_value(i, 5) + '>' + sheet.cell_value(i-1, 1), xls_row, 0)

                ra_rtt = ping(sheet.cell_value(i, 5), sheet.cell_value(i, 1))
                write_values(min, ra_rtt[0], xls_row, 1)
                write_values(avrg, ra_rtt[1], xls_row, 1)
                write_values(max, ra_rtt[2], xls_row, 1)
                xls_row += 1

    book.close()

if __name__ == '__main__':
    signal.signal(signal.SIGINT, signal_handler)  # catch ctrl-c and call handler to terminate the script
    main()
