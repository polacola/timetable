from openpyxl import Workbook, load_workbook  # 对排序不友好 尾大不掉
from openpyxl.styles import Font
from openpyxl.styles import PatternFill
from openpyxl.styles import Alignment
from openpyxl.styles import Border, Side
from openpyxl.worksheet.properties import PageSetupProperties
# from openpyxl.worksheet.pagebreak import Break

import os
import time
import pandas as pd  # 对样式不友好 弃用
import win32com.client as win32  # pip install pywin32  #用于格式转化
from win32com.client import DispatchEx
import fitz  # pdf转图片  #PyMuPDF

import base64
import hashlib
import json
import requests

import xlwings as xw

path = os.getcwd()  # r"F:\桌面\新建文件夹"  # r"{}".format(input())   #路径
f_n = os.listdir(path)

align = Alignment(horizontal='center', vertical='center', wrap_text=True)  # 剧中选项

font_1 = Font(  # ”讲课“样式
    name="华光准圆_CNKI",
    color="006100",  # 颜色
    size=11,  # 设定文字大小
    bold=False,  # 设定为粗体
    italic=False  # 设定为斜体
)
font_2 = Font(  # ”实验“样式
    name="华光准圆_CNKI",
    color="9C5700",  # 颜色
    size=11,  # 设定文字大小
    bold=False,  # 设定为粗体
    italic=False  # 设定为斜体
)
font_3 = Font(  # "在线"样式
    name="华光准圆_CNKI",
    color="757672",  # "545454"  # 颜色
    size=11,  # 设定文字大小
    bold=False,  # 设定为粗体
    italic=False  # 设定为斜体
)
font_4 = Font(  # ”讨论“样式 以及 “其他”样式
    name="华光准圆_CNKI",
    color="9C0006",  # 颜色
    size=11,  # 设定文字大小
    bold=False,  # 设定为粗体
    italic=False  # 设定为斜体
)
font_5 = Font(  # 标题样式
    name="华光准圆_CNKI",
    color="000000",  # 颜色
    size=11,  # 设定文字大小
    bold=False,  # 设定为粗体
    italic=False  # 设定为斜体
)

font_title = Font(  # 表头样式
    name="华光准圆_CNKI",
    color="000000",  # 使用预置的颜色常量
    size=20,  # 设定文字大小
    bold=True,  # 设定为粗体
    italic=False  # 设定为斜体
)

border_NOR = Border(left=Side(border_style='thin', color='000000'),
                    # values=('dashDot','dashDotDot', 'dashed','dotted','double','hair', 'medium', 'mediumDashDot',
                    # 'mediumDashDotDot','mediumDashed', 'slantDashDot', 'thick', 'thin')
                    right=Side(border_style='thin', color='000000'),
                    top=Side(border_style='thin', color='000000'),
                    bottom=Side(border_style='thin', color='000000'))


def set_font(read_col, cell, ws_new_in_set, row_in_set):
    if read_col == "Y" and cell.value == "讲课":
        col = 0
        for each_cell_wb in ws_new_in_set["{}".format(row_in_set)]:
            ws_new_in_set["{}{}".format("ABCDEFGHI"[col], row_in_set)].font = font_1
            ws_new_in_set["{}{}".format("ABCDEFGHI"[col], row_in_set)].fill = PatternFill("solid", "C6EFCE")
            col += 1
    elif read_col == "Y" and cell.value == "实验":
        col = 0
        for each_cell_wb in ws_new_in_set["{}".format(row_in_set)]:
            ws_new_in_set["{}{}".format("ABCDEFGHI"[col], row_in_set)].font = font_2
            ws_new_in_set["{}{}".format("ABCDEFGHI"[col], row_in_set)].fill = PatternFill("solid", "FFEB9C")
            col += 1
    elif read_col == "Y" and cell.value == "在线":
        col = 0
        for each_cell_wb in ws_new_in_set["{}".format(row_in_set)]:
            ws_new_in_set["{}{}".format("ABCDEFGHI"[col], row_in_set)].font = font_3
            ws_new_in_set["{}{}".format("ABCDEFGHI"[col], row_in_set)].fill = PatternFill("solid", "e9ebe3")
            col += 1
    elif read_col == "Y" and cell.value == "讨论":
        col = 0
        for each_cell_wb in ws_new_in_set["{}".format(row_in_set)]:
            ws_new_in_set["{}{}".format("ABCDEFGHI"[col], row_in_set)].font = font_4
            ws_new_in_set["{}{}".format("ABCDEFGHI"[col], row_in_set)].fill = PatternFill("solid", "FFC7CE")
            col += 1
    elif read_col == "Y" and cell.value == "授课性质":
        col = 0
        for each_cell_wb in ws_new_in_set["{}".format(row_in_set)]:
            ws_new_in_set["{}{}".format("ABCDEFGHI"[col], row_in_set)].font = font_5
            ws_new_in_set["{}{}".format("ABCDEFGHI"[col], row_in_set)].fill = PatternFill("none")
            col += 1
    elif read_col == "Y":
        col = 0
        for each_cell_wb in ws_new_in_set["{}".format(row_in_set)]:
            ws_new_in_set["{}{}".format("ABCDEFGHI"[col], row_in_set)].font = font_4
            ws_new_in_set["{}{}".format("ABCDEFGHI"[col], row_in_set)].fill = PatternFill("solid", "FFC7CE")
            col += 1
    ws_new_in_set.row_dimensions[row_in_set].height = 40  # 可自定义行高 删除此行可自动换行 课程内容过长时适用（如病理学） 但十分丑陋 建议手动调整
    ws_new_in_set['{}{}'.format(ws_creat_col, row_in_set)].alignment = align
    ws_new_in_set['{}{}'.format(ws_creat_col, row_in_set)].border = border_NOR


def set_col_width(ws_new_in_set, flag_a1=False):
    ws_new_in_set.column_dimensions['A'].width = 7.0  # 列宽 col_width
    if flag_a1:
        ws_new_in_set.column_dimensions['A'].width = 20.0  # 列宽 col_width
    ws_new_in_set.column_dimensions['B'].width = 7.0
    ws_new_in_set.column_dimensions['C'].width = 12.0
    ws_new_in_set.column_dimensions['D'].width = 7.0
    ws_new_in_set.column_dimensions['E'].width = 7.0
    ws_new_in_set.column_dimensions['F'].width = 40.0
    ws_new_in_set.column_dimensions['G'].width = 10.0
    ws_new_in_set.column_dimensions['H'].width = 10.0
    ws_new_in_set.column_dimensions['I'].width = 10.0


def new_filename(set_):
    new_f_name = ""
    if len(set_) == 1:
        for course_name in set_:
            new_f_name = course_name
        return new_f_name
    for course_name in set_:
        new_f_name = new_f_name + "、" + course_name
    return new_f_name[1:]


# def new_excel(): #exec()无法在函数内部修改局部变量 所以无法封装 我知道非常丑陋 但已经没有精力改了 #又不是不能用(•ิ_•ิ)
#     exec("wb_new{} = Workbook() ".format(a)) # 创建一个工作簿对象 exec(" ".format(a))
#     # exec("global wb_new{}".format(a))
#     exec("wb_new{}.pageSetUpPr = PageSetupProperties(fitToPage=True, autoPageBreaks=False)".format(a))
#     exec("ws_created{} = wb_new{}.create_sheet('教学进程', 0)".format(a,a))# 在索引为0的位置创建一个sheet页
#     # exec("global ws_created{}".format(a))
#     exec("ws_created{}.sheet_properties.tabColor = 'ff72BA'".format(a))  # 设置一个颜色（16位）
#     exec("ws_created{}['A1'].value = '{}'".format(a,os.path.splitext(file)[0]))  # 写入表头(之后更新为课程名称)
#     exec("ws_created{}['A1'].alignment = align".format(a))
#     # ws["A1"].border = border_NOR  #默认标题无边框
#     exec("ws_created{}['A1'].font = font_title".format(a))
#     exec("ws_created{}.row_dimensions[1].height = 30".format(a))  # 标题行高
#
#     exec("ws_created{}.merge_cells('A1:I1')".format(a))
#     exec("ws_created{}.merge_cells('A2:I2')".format(a))
#     exec("ws_created{}['A2'].value = '                ver1.0 by CDH{}'".format(a,time.strftime("%Y", time.localtime())))
#     exec("ws_created{}['A2'].alignment = align".format(a))
#     exec("ws_created{}['A2'].font = Font(name='华光准圆_CNKI', size=8, bold=False, italic=True)".format(a))
def excel_to_pdf(excel_path_, pdf_path_):
    xlApp = DispatchEx("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = 0
    books = xlApp.Workbooks.Open(excel_path_, False)
    books.ExportAsFixedFormat(0, pdf_path_)
    books.Close(False)
    xlApp.Quit()


def pdf_to_imgs(pdf_path, imgs_path, imgs_name):
    pdfDoc = fitz.open(pdf_path)
    for pg in range(pdfDoc.page_count):
        page = pdfDoc[pg]
        rotate = int(0)
        # 每个尺寸的缩放系数为1.3，这将为我们生成分辨率提高2.6的图像。
        # 此处若是不做设置，默认图片大小为：792X612, dpi=96
        zoom_x = 4  # (1.33333333-->1056x816)   (2-->1584x1224)
        zoom_y = 4
        mat = fitz.Matrix(zoom_x, zoom_y).prerotate(rotate)  # 注意大小写 不同版本大小写不同
        pix = page.get_pixmap(matrix=mat, alpha=False)

        if not os.path.exists(imgs_path):  # 判断存放图片的文件夹是否存在
            os.makedirs(imgs_path)  # 若图片文件夹不存在就创建
        if pg == 0:
            pix.save(imgs_path + '/' + f'{imgs_name}.png')
        else:
            pix.save(imgs_path + '/' + f'{imgs_name}_{pg}.png')
    # pdfDoc = fitz.open("pdf_path")
    # page = pdfDoc.loadPage(0)  # PDF页数
    # pix = page.getPixmap()
    # pix.writePNG(imgs_path)  # 保存 此路径格式不同


if not os.path.exists(r"{}\output".format(path)):  # 生成output文件夹
    os.makedirs(r"{}\output".format(path))
if not os.path.exists(r"{}\output_pdf".format(path)):
    os.makedirs(r"{}\output_pdf".format(path))
if not os.path.exists(r"{}\output_date".format(path)):
    os.makedirs(r"{}\output_date".format(path))  # images是否生成可选，路径在函数内生成


# def sort_excel(excel_name):
#     # 读取上一步保存的Excel文件
#     df = pd.read_excel(excel_name, sheet_name="教学进程",header=2)
#     df_value = df.sort_values(by=["日期","节次"], ascending=True)
#     # 保存文件
#     writer = pd.ExcelWriter(excel_name)
#     df_value.to_excel(writer, sheet_name='教学进程', index=False)
#     writer.save()


def sort_excel(excel_name):
    app = xw.App(visible=False, add_book=False)
    wb = app.books.open(excel_name)
    sht = wb.sheets['教学进程']

    atable = sht.range('A4:I{}'.format(row_all_with_data)).value  # 先拿值出来处理
    df = pd.DataFrame(atable, columns=['课程名称', '周次', '日期', '星期', "节次", "授课内容", "授课地点", "授课教师",
                                       "授课性质"])
    sort_df = df.sort_values(by=["日期", "节次"], ascending=True)
    # 使用这个函数进行排序
    # ascending=False 是降序排序

    sht.range('A4:I{}'.format(row_all_with_data)).value = sort_df.values.tolist()
    wb.save(excel_name)
    app.quit()


date_today = time.strftime("%Y-%m-%d", time.localtime())
with open(r'{}\output\备注信息_{}.txt'.format(path, date_today), "w", encoding="utf-8") as f:
    f.write(
        f'================以下是备注信息，请仔细核对{time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())}================\n')

set_course_name_all = set()
wb_new_all = Workbook()  # 创建一个工作簿对象
wb_new_all.pageSetUpPr = PageSetupProperties(fitToPage=True, autoPageBreaks=False)

ws_created_all = wb_new_all.create_sheet('教学进程', 0)  # 在索引为0的位置创建一个sheet页
ws_created_all.sheet_properties.tabColor = 'ff72BA'  # 设置一个颜色（16位）

ws_created_all["A1"].value = "TEMP"  # 写入表头
ws_created_all["A1"].alignment = align
# ws["A1"].border = border_NOR  #默认标题无边框
ws_created_all["A1"].font = font_title
ws_created_all.row_dimensions[1].height = 100  # 总表标题行高

ws_created_all.merge_cells("A1:I1")
ws_created_all.merge_cells("A2:I2")
ws_created_all["A2"].value = "                  ver1.0 by CDH{}".format(time.strftime("%Y", time.localtime()))
ws_created_all["A2"].alignment = align
ws_created_all["A2"].font = Font(name="华光准圆_CNKI", size=8, bold=False, italic=True)
row_all = 3  # 从第三行开始
temp = 3
row_number_last_read = 0
count_title_all = 0
# count_print_imgs=0
flag = False
# ============================================================
for file in f_n:
    if os.path.isfile(r'{}\{}'.format(path, file)) and (os.path.splitext(file)[1].lower() == ".xls" or
                                                        os.path.splitext(file)[
                                                            1].lower() == ".xlsx"):  # openpyxl 不支持xls
        print("分析源文件：" + file)
        if os.path.splitext(file)[1] == ".xls":
            filename = r'{}\{}'.format(path, file)
            Excelapp = win32.gencache.EnsureDispatch(
                'Excel.Application')
            workbook = Excelapp.Workbooks.Open(filename)
            workbook.SaveAs(filename.replace('xls', 'xlsx'), FileFormat=51)
            workbook.Close()
            Excelapp.Application.Quit()
            os.remove(filename)  # 删除源文件

        wb_read = load_workbook(r'{}\{}.xlsx'.format(path, os.path.splitext(file)[0]))  # 读取源文件
        for each_sheet in wb_read.sheetnames:
            set_course_name = set()
            # new_excel()  #本应如此优雅

            wb_new = Workbook()  # 创建一个工作簿对象
            wb_new.pageSetUpPr = PageSetupProperties(fitToPage=True, autoPageBreaks=False)

            ws_created = wb_new.create_sheet('教学进程'.format(os.path.splitext(file)[0]), 0)  # 在索引为0的位置创建一个sheet页
            ws_created.sheet_properties.tabColor = 'ff72BA'  # 设置一个颜色（16位）

            ws_created["A1"].value = "{}".format(os.path.splitext(file)[0])  # 写入表头
            ws_created["A1"].alignment = align
            # ws["A1"].border = border_NOR  #默认标题无边框
            ws_created["A1"].font = font_title
            ws_created.row_dimensions[1].height = 30  # 标题行高

            ws_created.merge_cells("A1:I1")
            ws_created.merge_cells("A2:I2")
            ws_created["A2"].value = "                  ver1.0 by CDH{}".format(time.strftime("%Y", time.localtime()))
            ws_created["A2"].alignment = align
            ws_created["A2"].font = Font(name="华光准圆_CNKI", size=8, bold=False, italic=True)

            ws_read = wb_read[each_sheet]
            count_P = 0  # 用于课程内容栏目有额外内容的情况，计数取到P列的次数，第二次分析内容
            for ws_read_col in "ABDFGPUWYRSP":
                colA_To_P = ws_read['{}'.format(ws_read_col)]  # 取一整列
                row = 3  # 前几行是原来的表头
                if count_P == 1 and ws_read_col == "P":
                    row_all += row_number_last_read
                    last_temp = temp  # 用于第二遍读到P  此时的row_all继续增加 用last_temp 回退到上一个值
                    temp = row_all
                    # print(str(row_all)+"更新---------------------")
                else:
                    row_all = temp
                if (ws_read_col != "R" and ws_read_col != "S" and ws_read_col != "P") or (
                        count_P == 0 and ws_read_col == "P"):
                    if (count_P == 0 and ws_read_col == "P"): count_P += 1
                    ws_creat_col = "ABCDEFGHI"["ABDFGPUWY".index(ws_read_col)]
                    for each_cell in colA_To_P[2:-2]:  # 先不看课程备注信息；最后有两行无用信息

                        # 写入新文件
                        ws_created['{}{}'.format(ws_creat_col, row)].value = str(each_cell.value).replace("\n",
                                                                                                          "")  # 写入新文件的数据

                        if colA_To_P[2:-2].index(each_cell) == 0 and count_title_all < 9:
                            ws_created_all['{}{}'.format(ws_creat_col, row_all)].value = str(each_cell.value).replace(
                                "\n", "")  # 写入新文件的数据
                            count_title_all += 1
                            row_number_last_read = len(colA_To_P[2:-2])
                        elif colA_To_P[2:-2].index(each_cell) == 0 and count_title_all == 9:
                            flag = True  # 标题行在总表只用写一遍，所以row_all要往回一行
                            row_number_last_read = len(colA_To_P[2:-2]) - 1
                        elif ws_read_col == "A":  # 总表需要显示课程名称
                            ws_created_all['{}{}'.format(ws_creat_col, row_all)].value = str(
                                ws_read["H{}".format(row)].value).replace(
                                "\n", "")
                        elif ws_read_col == "G":  # 更改“节次”为文本，便于排序
                            ws_created_all['{}{}'.format(ws_creat_col, row_all)].value = str(each_cell.value).replace(
                                "\n", "")
                            ws_created_all['{}{}'.format(ws_creat_col, row_all)].number_format = '@'

                        else:
                            ws_created_all['{}{}'.format(ws_creat_col, row_all)].value = str(each_cell.value).replace(
                                "\n", "")
                        set_font(ws_read_col, each_cell, ws_created, row)  # 设置当前行样式
                        set_font(ws_read_col, each_cell, ws_created_all, row_all)
                        if flag == True:
                            row_all -= 1  # 标题行在总表只用写一遍，所以row_all要往回一行
                            flag = False
                        row += 1
                        row_all += 1


                elif (ws_read_col == "R" or ws_read_col == "S") or (ws_read_col == "P"):

                    count_print = 0
                    row += 1  # col初始为4 因为有标题

                    if ws_read_col == "P":
                        row_all = last_temp
                    for each_cell in colA_To_P[3:-2]:
                        if (each_cell.value != "" and ws_read_col != "P") or (
                                ("在线" in each_cell.value) or ("钉钉" in each_cell.value) or (
                                "直播" in each_cell.value) or ("更大" in each_cell.value) or (
                                        "线上" in each_cell.value)):
                            if count_print == 0: print(
                                "————————————————————————————————————————————————\n注意！此课程有备注信息，请仔细核对。备注中出现在线、钉钉等字样会设置为浅色填充样式（课程性质不一定为在线）");count_print += 1
                            notes = "{0:{6}^4}老师的{1}({2}{3})出现备注信息：日期{5}  {4} ".format(
                                ws_created["H{}".format(row)].value, "【" + str(ws_read["H{}".format(row)].value) + "】",
                                ws_read_col, row, each_cell.value, ws_read["D{}".format(row)].value, chr(12288))
                            print(notes)
                            with open(r'{}\output\备注信息_{}.txt'.format(path, date_today), "a",
                                      encoding="utf-8") as f:
                                f.write(notes + "\n")

                        if ("在线" in each_cell.value) or ("钉钉" in each_cell.value) or (
                                "直播" in each_cell.value) or ("更大" in each_cell.value) or (
                                "线上" in each_cell.value):
                            col = 0
                            for each_cell_wb in ws_created["{}".format(row)]:
                                ws_created["{}{}".format("ABCDEFGHI"[col], row)].font = font_3  # "在线样式"
                                ws_created["{}{}".format("ABCDEFGHI"[col], row)].fill = PatternFill("solid",
                                                                                                    "e9ebe3")  # PatternFill("none")B8CCE4

                                # print(row_all)
                                ws_created_all["{}{}".format("ABCDEFGHI"[col], row_all)].font = font_3  # "在线样式"
                                ws_created_all["{}{}".format("ABCDEFGHI"[col], row_all)].fill = PatternFill("solid",
                                                                                                            "e9ebe3")  # PatternFill("none")B8CCE4
                                col += 1

                        if ws_read_col == "P":
                            set_course_name.add(str(ws_read["H{}".format(row)].value).replace(" ", ""))
                            set_course_name_all.add(str(ws_read["H{}".format(row)].value).replace(" ", ""))
                        row += 1
                        row_all += 1
                    row_all = temp

            # 信息读取完毕 开始设置新文件列宽和页面布局 （行高已经在之前设置）
            ws_created["A1"].value = "{}".format(new_filename(set_course_name))  # 写入表头
            set_col_width(ws_created, False)

            ws_created.print_options.horizontalCentered = True  # 页面布局
            ws_created.print_options.verticalCentered = False

            ws_created.sheet_properties.pageSetUpPr.fitToPage = True  # 调整为一页 ws.page_setup.fitToPage
            ws_created.page_setup.fitToHeight = True

            if len(wb_read.sheetnames) == 1:
                wb_new.save(r'{}\output\{}.xlsx'.format(path, new_filename(set_course_name)))  # 将创建的工作簿保存

                excel_path = r'{}\output\{}.xlsx'.format(path, new_filename(set_course_name))
                pdf_path = r'{}\output_pdf\{}.pdf'.format(path, new_filename(set_course_name))
                img_name = new_filename(set_course_name)
            else:
                wb_new.save(r'{}\output\{}_{}.xlsx'.format(path, new_filename(set_course_name), each_sheet))

                excel_path = r'{}\output\{}_{}.xlsx'.format(path, new_filename(set_course_name), each_sheet)
                pdf_path = r'{}\output_pdf\{}_{}.pdf'.format(path, new_filename(set_course_name), each_sheet)
                img_name = "{}_{}".format(new_filename(set_course_name), each_sheet)
            wb_new.close()  # 最后关闭文件

            # 转pdf
            excel_to_pdf(excel_path, pdf_path)
            # 转图片
            img_path = r'{}\output_images'.format(path)
            pdf_to_imgs(pdf_path, img_path, img_name)

            print(r'写入：output\{}.xlsx'.format(img_name))  # 名称都是new_filename(set_course_name)
            print(r'写入：output_pdf\{}.pdf'.format(img_name))
            print(r'写入：output_images\{}.png'.format(img_name))
            print("====================写入完成！====================")
            with open(r'{}\output\备注信息_{}.txt'.format(path, date_today), "a", encoding="utf-8") as f:
                f.write(r'--------以上是【{}】的备注信息--------'.format(new_filename(set_course_name)) + "\n\n")
        wb_read.close()  # 关闭

        old_source_file_name = r'{}\{}.xlsx'.format(path, os.path.splitext(file)[0])
        new_source_file_name = r'{}\{}(origin).xlsx'.format(path, new_filename(set_course_name),
                                                            os.path.splitext(file)[0])
        os.renames(old_source_file_name, new_source_file_name + "0")
        os.renames(new_source_file_name + "0", new_source_file_name)

ws_created_all["A1"].value = "{}".format(new_filename(set_course_name_all))  # 写入表头
ws_created_all["A3"].value = "课程名称"
set_col_width(ws_created_all, True)

ws_created_all.print_options.horizontalCentered = True  # 页面布局
ws_created_all.print_options.verticalCentered = False

ws_created_all.sheet_properties.pageSetUpPr.fitToPage = True  # 调整为一页 ws.page_setup.fitToPage
ws_created_all.page_setup.fitToHeight = False

row_all_with_data = len(ws_created_all["I"])  # 总表行数
wb_new_all.save(r'{}\output\{}门课程.xlsx'.format(path, len(set_course_name_all)))  # 将创建的工作簿保存

wb_new_all.close()  # 最后关闭文件
sort_excel(r'{}\output\{}门课程.xlsx'.format(path, len(set_course_name_all)))
print(r'写入：output\{}门课程.xlsx'.format(len(set_course_name_all)))

old_txt_file_name = r'{}\output\备注信息_{}.txt'.format(path, date_today)  # 重命名备注txt
new_txt_file_name = r'{}\output\{}门课程_备注信息_{}.txt'.format(path, len(set_course_name_all), date_today)
print(r'写入：output\{}门课程_备注信息'.format(len(set_course_name_all)))
try:
    os.renames(old_txt_file_name, new_txt_file_name)

except:
    print("备注信息txt更名失败！（已存在同名文件）\n已存入{}.txt文件".format(date_today))

# 转总表pdf可选项
# excel_path_all = r'{}\output\{}门课程.xlsx'.format(path, len(set_course_name_all))
# pdf_path_all= r'{}\output_pdf\{}门课程.pdf'.format(path, len(set_course_name_all))
# excel_to_pdf(excel_path_all,pdf_path_all)
# print(r'写入：output_pdf\{}门课程.pdf'.format(len(set_course_name_all)))
# wb_read = load_workbook(r'{}\{}.xlsx'.format(path, os.path.splitext(file)[0]))
# wb_read(r'{}\output\{}.xlsx'.format(path, new_filename(set_course_name)))
# wb_read.close()
# --------------------
# ------
#到是的f

print("====================写入完成！====================")

# ver 1.0 2023.9.2
# Copyright (c) 2023 CDH
# molu2003@foxmail.com
