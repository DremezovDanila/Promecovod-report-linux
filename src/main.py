from fpdf import FPDF
import math


with open('file', 'r') as f:
    file = f.readlines()
if len(file) <= 1:
#    file = [['Data', 'is', 'wrong'], [0, 0, 0]]
    file = ['Data;is;wrong;\n', '0;0;0;\n']

for i in range(len(file)):
    file[i] = file[i].split(";")
    file[i].pop(-1)
# Initialise list of file head
report_head = file.pop(0)
# Initialise list of file content
report_content = file

pdf1 = FPDF()
# Format pdf document
left_margin, top_margin, right_margin = 5, 5, 5
row_height = 5
font_size = 8
pdf1.set_margins(5, 5, 5)
cur_x, cur_y = left_margin, top_margin
# Create the first page
pdf1.add_page()
pdf1.set_font(family="Arial", size=font_size)
# Minimum column width is 20, else the meaning of a number of lines is rough
# col_width equals to a width of the longest meaning in a respective column of the table
report_elem_width = list()
for i in range(len(report_content)):
    report_elem_width.append([])
    for j in range(len(report_head)):
        report_elem_width[i].append(pdf1.get_string_width(report_content[i][j]) + 2)
# Find the longest lines in columns and assign these values to head titles length
# Reverse massive of contents length (columns into rows)
report_elem_width_rev = [[0 for j in range(len(report_content))] for i in range(len(report_head))]
for i in range(len(report_content)):
    for j in range(len(report_head)):
        report_elem_width_rev[j][i] = report_elem_width[i][j]
report_head_width = list()
for i in range(len(report_head)):
    # Indents from edges of the cell equal 2
    report_head_width.append(max(report_elem_width_rev[i]) + 2)
# Find how many lines head titles take
report_head_lines_num = list()
report_head_len = list()
for i in range(len(report_head)):
    report_head_len.append(pdf1.get_string_width(report_head[i]))
    # How many lines a string takes
    if report_head_len[i] > 0 and report_head_len[i] > report_head_width[i] - 2 and report_head_width[i] > 2:
        report_head_lines_num.append(math.ceil(report_head_len[i] / (report_head_width[i] - 2)) + 2)
    else:
        report_head_lines_num.append(3)
# Draw a table head
# Calculate how many Line Feed characters are used in every head cell relating to the highest head cell
for i in range(len(report_head)):
    pdf1.multi_cell(w=report_head_width[i],
                    h=row_height,
                    txt="\n" + report_head[i] + "\n" * (max(report_head_lines_num) - report_head_lines_num[i] + 2),
                    border=1,
                    align='C')
    cur_x += report_head_width[i]
    pdf1.set_xy(cur_x, cur_y)
cur_x = left_margin
# Make top indent according to number of lines in a head cell
cur_y = top_margin + row_height * max(report_head_lines_num)
pdf1.set_xy(cur_x, cur_y)
for i in range(len(report_content)):
    for j in range(len(report_head)):
        pdf1.multi_cell(w=report_head_width[j], h=row_height, txt=report_content[i][j], border=1, align='C')
        cur_x += report_head_width[j]
        pdf1.set_xy(cur_x, cur_y)
    cur_x = left_margin
    cur_y += row_height
    pdf1.set_xy(cur_x, cur_y)
pdf1.output("new_pdf.pdf")
