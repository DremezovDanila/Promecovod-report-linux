from pyModbusTCP.client import ModbusClient
from pyModbusTCP import utils
#from fpdf import FPDF
import fpdf
import time
import datetime
import os
import subprocess
from typing import Any
from threading import Thread
from calendar import monthrange
import math


#  Create new class for PDF document based on original class from FPDF lib.
class Pdf1(fpdf.FPDF):

    #  Method creating report in PDF.
    #  table_data comprises list of data divided into columns and rows without '/n' in the end of rows.
    def create_report(self, report_data: list[list[str]], report_params: list):
        if type(report_data) != list or len(report_data) <= 1:
            #  data variable includes table head and table contents.
            report_data = [['Data', 'is', 'wrong'], [0, 0, 0]]
        date_and_time = datetime.datetime.now()
        #  Initialise list of file head.
        report_head = report_data.pop(0)
        #  Initialise list of file content.
        report_content = report_data
        #  Formatting pdf document.
        left_margin, top_margin, right_margin = 5, 5, -5
        margins = (left_margin, top_margin, right_margin)
        row_height = 6.2
        self.set_margins(left_margin, top_margin, right_margin)
        cur_x, cur_y = left_margin, top_margin
        self.set_xy(cur_x, cur_y)
        #  Create the first page.
        self.add_page()
        #  Download the font supporting Unicode and set it. *It has to be added before used.
        self.add_font(family="NotoSans", style="", fname=r"NotoSans-Regular.ttf", uni=True)
        #  Filling in general information.
        #  First line with current date on the left and company name on the right.
        self.set_font(family="NotoSans", size=14)
        self.cell(w=(self.w - (margins[0] + abs(margins[2]))) / 2, h=self.font_size,
                  txt="{:%Y.%m.%d}".format(date_and_time), ln=0, align="L")
        self.cell(w=(self.w - (margins[0] + abs(margins[2]))) / 2, h=self.font_size, txt='ООО "Промэковод"', ln=0,
                  align="R")
        cur_x = left_margin
        cur_y += 10
        self.set_xy(cur_x, cur_y)
        #  Site name.
        self.set_font("NotoSans", size=20)
        self.cell(self.w - 20, self.font_size * 1.5, txt=report_params[1], ln=0, align="C")
        cur_x = left_margin
        cur_y += 10
        self.set_xy(cur_x, cur_y)
        #  Table title.
        self.set_font("NotoSans", size=14)
        self.cell(self.w - 20, self.font_size * 2, txt=f"Отчет по суточным расходам воды и электроэнергии", ln=0,
                  align="C")
        cur_x = left_margin
        cur_y += 10
        self.set_xy(cur_x, cur_y)
        #  Drawing a table.
        #  Find widths of columns in table using specified font.
        self.set_font("NotoSans", size=14)
        #  Minimum column width is 20, else the meaning of a number of lines is rough.
        #  Adding some width to meanings of table data.
        for i in range(len(report_content)):
            for j in range(len(report_head)):
                if j == 0:
                    report_content[i][j] = report_content[i][j].rjust(3, " ")
                elif j == 1:
                    report_content[i][j] = report_content[i][j].rjust(12, " ")
                else:
                    #  Check column width coefficient for adequacy.
                    if int(site_params[7]) < 0:
                        column_width_coef = 0
                    elif int(site_params[7]) > 30:
                        column_width_coef = 30
                    else:
                        column_width_coef = int(site_params[7])
                    report_content[i][j] = report_content[i][j].rjust(6 + column_width_coef, " ")
        #  col_width equals to the width of the largest meaning in a respective column of the table.
        report_elem_width = list()
        for i in range(len(report_content)):
            report_elem_width.append([])
            for j in range(len(report_head)):
                report_elem_width[i].append(self.get_string_width(report_content[i][j]))
        #  Find the longest lines in columns and assign these values to head titles length.
        #  Reverse massive of contents length (columns into rows).
        report_elem_width_rev = [[0 for j in range(len(report_content))] for i in range(len(report_head))]
        for i in range(len(report_content)):
            for j in range(len(report_head)):
                report_elem_width_rev[j][i] = report_elem_width[i][j]
        #  Finding widths of column name cells.
        report_head_width = list()
        for i in range(len(report_head)):
            #  Indents from edges of the cell equal 2.
            report_head_width.append(max(report_elem_width_rev[i]) + 4.0)
        #  Find how many lines head titles take.
        report_head_lines_num = list()
        report_head_len = list()
        for i in range(len(report_head)):
            report_head_len.append(self.get_string_width(report_head[i]) + 2.0)
            #  How many lines a string takes.
            if 0 < report_head_width[i] < report_head_len[i]:
                report_head_lines_num.append(math.ceil(report_head_len[i] / report_head_width[i]))
            else:
                report_head_lines_num.append(1)
        #  Draw a table head.
        #  Calculate how many Line Feed characters are used in every head cell relating to the highest head cell.
        for i in range(len(report_head)):
            self.multi_cell(w=report_head_width[i],
                            h=row_height,
                            txt="\n" + report_head[i] + "\n" * 2 + "\n" * (
                                        max(report_head_lines_num) - report_head_lines_num[i]),
                            border=1,
                            align='C')
            cur_x += report_head_width[i]
            self.set_xy(cur_x, cur_y)
        cur_x = left_margin
        #  Make top indent according to number of lines in a head cell.
        cur_y += row_height * (max(report_head_lines_num) + 2)
        self.set_xy(cur_x, cur_y)
        for i in range(len(report_content)):
            for j in range(len(report_head)):
                self.multi_cell(w=report_head_width[j], h=row_height, txt=report_content[i][j], border=1, align='C')
                cur_x += report_head_width[j]
                self.set_xy(cur_x, cur_y)
            cur_x = left_margin
            cur_y += row_height
            self.set_xy(cur_x, cur_y)


    #  Override footer method so that it numerates pages.
    def footer(self):
        self.set_y(-15)
        self.set_x(10)
        self.set_font(family="NotoSans", size=12)
        self.cell(w=self.w - 20, h=10, txt='Страница %s из ' % self.page_no() + '{nb}', align='C')


#  Function for checking config file and adding site parameters to a list of site parameters.
def get_sites_params(config_path: str = "config"):
    #  Check config text file and get the list of row sites data.
    if os.path.isfile(config_path):
        config_file_exist = True
        print("test 1.2")
        with open(config_path, "r", encoding="UTF-8") as config_file:
            print("test 1.3")
            config_list = config_file.readlines()
            config_list_len = len(config_list)
            sites_row_params = list()
            for num in range(config_list_len):
                if config_list[num].split(';')[0] == "site name":
                    print("test 1.4")
                    if num + 1 <= config_list_len:
                        print("test 1.5")
                        site_row_params_list = config_list[num + 1].split(';')
                        site_row_params_list.pop(-1)
                        if len(site_row_params_list) == 8:
                            print("test 1.6")
                            sites_row_params.append(site_row_params_list)
    else:
        config_file_exist = False

    #  Check correctness of row sites data and create a list of sites data.
    if config_file_exist and len(sites_row_params) > 0:
        config_params_off = False
    else:
        config_params_off = True

    sites_params = list()
    if not config_params_off:
        for site_row_params in sites_row_params:
            print("test 1.7")
            site_params_off = False
            site_params = list()
            try:
                site_params.append(site_row_params[0])  #  Site name.
                site_params.append(site_row_params[1])  #  Site name in Russian.
                site_params.append(site_row_params[2])  #  Report text file directory.
                site_params.append(site_row_params[3])  #  IP address and TCP port.
                site_ip_str = site_row_params[3].split(':')[0]
                int(site_row_params[3].split(':')[1])
                site_ip_row_list = site_ip_str.split('.')
                if len(site_ip_row_list) == 4:
                    for el in site_ip_row_list:
                        if 0 < int(el) < 255:
                            pass
                        else:
                            site_params_off = True
                            break
                else:
                    site_params_off = True
                site_params.append(int(site_row_params[4]))  #  Parameter number.
                site_params.append(site_row_params[5])  #  PDF report directory.
                site_params.append(site_row_params[6])  #  Columns names.
                if len(site_row_params[6].split(":")) != site_params[4]:
                    site_params_off = True
                site_params.append(int(site_row_params[7]))  #  Column width coefficient (int, 0 - 30).
                print("test 1.8")
            except:
                print("test 1.8.1")
                site_params_off = True
            finally:
                print("test 1.9")
                if not site_params_off:
                    sites_params.append(site_params)

    return config_params_off, sites_params


#  Global variables.
finish_main_process: bool = False
restore_data_start: bool = False
restore_data_site_name: str = ""
restore_data_report_date: str = ""
print_pdf_start: bool = True
print_pdf_site_name: str = ""
print_pdf_report_date: str = ""
command: str = ""

#  Cyclic function implementing console for operating the program.
def operate_program():
    global finish_main_process
    global restore_data_start
    global restore_data_site_name
    global restore_data_report_date
    global print_pdf_start
    global print_pdf_site_name
    global print_pdf_report_date
    global command

    while not finish_main_process:
        print("Enter your command:\n")
        command = input()
        #  Finish main process.
        if command == "finish" or command == "Finish" or command == "FINISH":
            finish_main_process = True
            command = ""
        #  Restore data.
        elif command == "restore" or command == "Restore" or command == "RESTORE":
            restore_data_site_name = str(input("Name of object (ex.:vzu_borodinsky): "))
            restore_data_report_date = str(input("Year and month (ex.:2023_05): "))
            restore_data_start = True
            command = ""
        #  Print data.
        elif command == "print" or command == "Print" or command == "PRINT":
            print_pdf_site_name = str(input("Name of object (ex.:vzu_borodinsky): "))
            print_pdf_report_date = str(input("Year and month (ex.:2023_05): "))
            print_pdf_start = True
            command = ""


#  Get current date and time.
cur_datetime = datetime.datetime.now()
cur_date = cur_datetime.date()
cur_time = cur_datetime.time()
cur_time_hour = cur_time.hour
cur_time_min = cur_time.minute
cur_time_sec = cur_time.second
cur_date_day = datetime.datetime.today().day
cur_date_month = datetime.datetime.today().month
cur_date_year = datetime.datetime.today().year
report_cur_datetime = datetime.datetime.now()
#  Log parameters.
logs_num_line = 0
logs_separator = " " * 4

#  Set new directory with fonts for FPDF library.
fpdf.set_global("SYSTEM_TTFONTS", os.path.join(os.path.dirname(__file__), f"{os.pardir}/etc/fonts"))

#  Start another thread for processing console commands.
th = Thread(target=operate_program, args=())
th.start()

if not os.path.isdir("../logs"):
    os.makedirs("../logs")

if not os.path.isdir("../reports"):
    os.makedirs("../reports")

with open("../logs/logs", "w", encoding="UTF-8") as log:
    log.write("logs start...\n")

print("test 1.0")
#  Main loop.
if __name__ == "__main__":
    while not finish_main_process:
        # Get current date and time.
        cur_datetime = datetime.datetime.now()
        cur_date = cur_datetime.date()
        cur_time = cur_datetime.time()
        cur_time_hour = cur_time.hour
        cur_time_min = cur_time.minute
        cur_time_sec = cur_time.second
        cur_date_day = datetime.datetime.today().day
        cur_date_month = datetime.datetime.today().month
        cur_date_year = datetime.datetime.today().year

        #  Start polls once a day at specified time (next day after report get formed in PLC). PLC draws report at
        #  23:59:59 and client polls start several minutes later.
        #if cur_time_hour == 0 and cur_time_min == 5 and cur_time_sec == 0:
        if cur_time_sec == 0:
            print("test 1.1")
            #  Get date of previous day (report).
            if cur_date_day == 1 and cur_date_month == 1:
                report_last_year = cur_date_year - 1
                report_last_month = 1
                report_last_day = 31
            elif cur_date_day == 1 and cur_date_month != 1:
                report_last_year = cur_date_year
                report_last_month = cur_date_month - 1
                #  Last day of previous month.
                report_last_day = monthrange(int(report_last_year), int(report_last_month))[1]
            else:
                report_last_year = cur_date_year
                report_last_month = cur_date_month
                report_last_day = cur_date_day - 1
            #  Last day datetime (when last report got formed).
            report_last_datetime = report_cur_datetime
            #  New day datetime (when next report get formed).
            report_cur_datetime = datetime.datetime.now()

            #  To interact with a network drive we have to add it to the local directory by mapping it to a letter (x:).
            #  Check whether drive x: has already created. If not, then create network drive x:.
            if not os.path.isdir("x:"):
                #  The command sets utf-8 (65001 code page) for "cmd.exe".
                command_1 = "chcp 65001"
                cmd_command_1 = subprocess.run(command_1, shell=True, capture_output=True)
                with open("../logs/logs", "a", encoding="UTF-8") as logs:
                    logs.write(f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                               f"Change encoding: {cmd_command_1.stdout}\n")
                #  The command maps network share to drive letter x: on the local system by SMB with authorization.
                command_2 = r'net use x: "\\192.168.1.2\ics" ddd2232 /user:dremezov /persistent:yes'
                try:
                    cmd_command_2 = subprocess.run(command_2, shell=True, capture_output=True)
                    with open("../logs/logs", "a", encoding="UTF-8") as logs:
                        logs.write(f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                                   f"Create network drive X: {cmd_command_2.stdout}\n")
                #  Process exception while system command performing in subprocess.
                except subprocess.CalledProcessError:
                    with open("../logs/logs", "a", encoding="UTF-8") as logs:
                        logs.write(f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                                   f"Error of creating drive X:.\n")
            else:
                with open("../logs/logs", "a", encoding="UTF-8") as logs:
                    logs.write(f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                               f"Drive X exists.\n")

            config_params_off, sites_params = get_sites_params()

            #  Check config text file and get the list of row sites data.
            #if os.path.isfile(r"config"):
            #    print("test 1.2")
            #    with open(r"config", "r", encoding="UTF-8") as config_file:
            #        print("test 1.3")
            #        config_list = config_file.readlines()
            #        config_list_len = len(config_list)
            #        sites_row_params = list()
            #        for num in range(config_list_len):
            #            if config_list[num].split(';')[0] == "site name":
            #                print("test 1.4")
            #                if num + 1 <= config_list_len:
            #                    print("test 1.5")
            #                    site_row_params_list = config_list[num + 1].split(';')
            #                    site_row_params_list.pop(-1)
            #                    if len(site_row_params_list) == 8:
            #                        print("test 1.6")
            #                        sites_row_params.append(site_row_params_list)
            #else:
            #    sites_row_params = list()

            #  Check correctness of row sites data and create a list of sites data.
            #sites_params = list()
            #for site_row_params in sites_row_params:
            #    print("test 1.7")
            #    site_row_params_off = False
            #    site_params = list()
            #    try:
            #        site_params.append(site_row_params[0])  #  Site name.
            #        site_params.append(site_row_params[1])  #  Site name in Russian.
            #        site_params.append(site_row_params[2])  #  Report text file directory.
            #        site_params.append(site_row_params[3])  #  IP address and TCP port.
            #        site_ip_str = site_row_params[3].split(':')[0]
            #        site_tcp_port_str = site_row_params[3].split(':')[1]
            #        site_ip_row_list = site_ip_str.split('.')
            #        if len(site_ip_row_list) == 4:
            #            site_ip_list = list()
            #            for el in site_ip_row_list:
            #                site_ip_list.append(int(el))
            #            site_ip = site_ip_str
            #        else:
            #            site_row_params_off = True
            #        site_tcp_port = int(site_tcp_port_str)
            #        site_params.append(int(site_row_params[4]))  #  Parameter number.
            #        site_params.append(site_row_params[5])  #  PDF report directory.
            #        site_params.append(site_row_params[6])  #  Columns names.
            #        if len(site_row_params[6].split(":")) != site_params[4]:
            #            site_row_params_off = True
            #        site_params.append(int(site_row_params[7]))  #  Column width coefficient (int, 0 - 30).
            #        print("test 1.8")
            #    except:
            #        print("test 1.8.1")
            #        site_row_params_off = True
            #    finally:
            #        print("test 1.9")
            #        if not site_row_params_off:
            #            sites_params.append(site_params)

            #  ModbusTCP polling.
            for site_params in sites_params:
                print("test 1.10")
                #  Get number of days in restoring report month.
                site_report_days_in_month = monthrange(int(report_last_year), int(report_last_month))[1]
                #  Create report list of a site.
                site_report_list2 = [['' for column in range(len(site_params[6].split(":")) + 2)]
                                     for row in range(site_report_days_in_month + 1)]
                #  Get column names from configuration file.
                site_report_col_names = ["№", "Дата"]
                site_report_col_names = site_report_col_names + site_params[6].split(":")
                #  Fill the first line in a site report.
                for col_num in range(len(site_report_col_names)):
                    site_report_list2[0][col_num] = site_report_col_names[col_num]
                #  Fill the first two columns in a site report.
                for day in range(site_report_days_in_month):
                    site_report_list2[day + 1][0] = f"{day + 1}"
                    site_report_list2[day + 1][1] = f"{report_last_year}.{report_last_month:02}.{(day + 1):02}"
                #  Get a directory for report text file of a site.
                site_report_dir = f"{site_params[2]}/{site_params[0]}_{report_last_year}_{report_last_month}"
                #  Initialise MBTCP connection.
                site_mbtcp_client = ModbusClient(host=site_params[3].split(":")[0],
                                                 port=int(site_params[3].split(":")[1]),
                                                 unit_id=1,
                                                 timeout=8.0,
                                                 auto_open=True,
                                                 auto_close=True)

                #  Poll parameters for object.
                poll_list2 = []
                poll_list2_real = []
                for param_num in range(site_params[4]):
                    print("test 1.11")
                    #  Write year, month, parameter number.
                    poll_w_request = \
                        site_mbtcp_client.write_multiple_registers(0,
                                                                  [int(report_last_year),
                                                                             int(report_last_month),
                                                                             param_num + 1])
                    #  Read parameter values for month.
                    poll_r_request = site_mbtcp_client.read_input_registers(0, site_report_days_in_month * 2)
                    #  Draw a list of registers.
                    poll_list2.append(poll_r_request)
                #  If polls succeeded then format data.
                if site_mbtcp_client.last_error == 0:
                    print("test 1.11.1")
                    for poll_list in poll_list2:
                        #  Conversion list of 2int16 to list of long32.
                        poll_r_request_long_32 = utils.word_list_to_long(poll_list, big_endian=True)
                        #  Conversion of list of long32 to list of real.
                        poll_r_request_list_real = []
                        for elem in poll_r_request_long_32:
                            poll_r_request_list_real.append(utils.decode_ieee(int(elem)))
                        #  Add every list of reals to list2 of reals.
                        poll_list2_real.append(poll_r_request_list_real)
                    #  Convert rows to columns in list2 of real values.
                    poll_list2_real_rows = len(poll_list2_real)
                    if poll_list2_real_rows > 0:
                        poll_list2_real_columns = len(poll_list2_real[0])
                    else:
                        poll_list2_real_columns = 0
                    poll_list2_real_converted = [[0.0 for col in range(poll_list2_real_rows)]
                                                 for row in range(poll_list2_real_columns)]
                    for row_num in range(poll_list2_real_rows):
                        for col_num in range(poll_list2_real_columns):
                            poll_list2_real_converted[col_num][row_num] = poll_list2_real[row_num][col_num]
                    #  Fill a site report by polled data.
                    for day in range(site_report_days_in_month):
                        for param_num in range(site_params[4]):
                            site_report_list2[day + 1][param_num + 2] = f"{format(poll_list2_real_converted[day][param_num], '.1f')}"
                else:
                    print("test 1.11.2")
                    print(site_mbtcp_client.last_error, "ModbusTCP error code")
                    with open("../logs/logs", "a", encoding="UTF-8") as logs:
                        logs.write(f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                                   f"{site_params[0]} site ModbusTCP poll error.\n")
                #  Write data into report text file.
                with open(f"{site_params[2]}/{site_params[0]}_{report_last_year}_{report_last_month}", "w",
                          encoding="UTF-8") as site_report_file:
                    print("test 1.12")
                    for row in site_report_list2:
                        for elem in row:
                            site_report_file.write(f"{elem};")
                        site_report_file.write("\n")

            #  PDF document printing.
            for site_params in sites_params:
                print("test 1.13")
                site_report_dir = f"{site_params[2]}/{site_params[0]}_{report_last_year}_{report_last_month}"
                if os.path.isfile(site_report_dir):
                    with open(site_report_dir, "r", encoding="UTF-8") as site_report_file:
                        site_report_data_list = site_report_file.readlines()
                        site_report_data_list2 = list()
                        for row in site_report_data_list:
                            site_report_data_list2.append(row.split(";")[:-1])
                    #  Declare instance of Pdf1.
                    report_pdf = Pdf1()
                    #  Set {nb} alias for PDF document footer.
                    report_pdf.alias_nb_pages()
                    report_pdf.create_report(site_report_data_list2, site_params)
                    if os.path.isdir(f"../{site_params[5]}/{site_params[1]}"):
                        report_pdf.output(f"../{site_params[5]}/{site_params[1]}/{site_params[1]}.pdf")
                    else:
                        os.makedirs(f"../{site_params[5]}/{site_params[1]}")
                        report_pdf.output(f"../{site_params[5]}/{site_params[1]}/{site_params[1]}.pdf")
            print("test 1.14")
            time.sleep(1.0)


            #  Function for printing PDF report from text file.
            if print_pdf_start:
                print_pdf_start = False
                print_report_input_data_ok = False
                #  Check input console data for relevance.
                try:
                    print_pdf_report_date_list = print_pdf_report_date.split("_")
                    print_pdf_report_date_year = print_pdf_report_date_list[0]
                    print_pdf_report_date_month = print_pdf_report_date_list[1]
                    print_pdf_report_date_year_int = int(print_pdf_report_date_year)
                    print_pdf_report_date_month_int = int(print_pdf_report_date_month)
                except:
                    with open("../logs/logs", "a", encoding="UTF-8") as logs:
                        logs.write(f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                                   f"Incorrect input data for printing PDF report.\n")
                else:
                    if len(print_pdf_report_date_list) == 2 and \
                            len(print_pdf_report_date_year) == 4 and \
                            (len(print_pdf_report_date_month) == 2 or len(print_pdf_report_date_month) == 1) and \
                            2022 < print_pdf_report_date_year_int < 2040 and \
                            (1 <= print_pdf_report_date_month_int <= 12):
                        print_report_input_data_ok = True

                    else:
                        with open("../logs/logs", "a", encoding="UTF-8") as logs:
                            logs.write(f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                                       f"Incorrect date or site name for printing PDF report.\n")

                if print_report_input_data_ok:
                    #  Get path to text file.
                    print_report_data_path = os.path.join(os.path.dirname(__file__), f"{os.pardir}/reports/"
                                                                                         f"{print_pdf_site_name}_{print_pdf_report_date_year_int}_{print_pdf_report_date_month_int}")
                    if os.path.isfile(print_report_data_path):
                        with open(print_report_data_path, "r", encoding="UTF-8") as print_report_file:
                            print_report_data_list = print_report_file.readlines()
                            print_report_data_list2 = list()
                            for row in print_report_data_list:
                                print_report_data_list2.append(row.split(";")[:-1])

                        config_params_off, sites_params = get_sites_params()
                        print_report_site_params = list()
                        if not config_params_off:
                            for site_params in sites_params:
                                if site_params[0] == print_pdf_site_name:
                                    print_report_site_params = site_params
                                    break

                        if print_report_site_params != list():
                            #  Declare instance of Pdf1.
                            print_pdf_file = Pdf1()
                            #  Set {nb} alias for PDF document footer.
                            print_pdf_file.alias_nb_pages()
                            print_pdf_file.create_report(print_report_data_list2, print_report_site_params)
                            if os.path.isdir(f"../{print_report_site_params[5]}/{print_report_site_params[1]}"):
                                print_pdf_file.output(f"../{print_report_site_params[5]}/{print_report_site_params[1]}/{print_report_site_params[1]}.pdf")
                            else:
                                os.makedirs(f"../{print_report_site_params[5]}/{print_report_site_params[1]}")
                                print_pdf_file.output(f"../{print_report_site_params[5]}/{print_report_site_params[1]}/{print_report_site_params[1]}.pdf")


























