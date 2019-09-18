import sys
import xlrd
import xlwt

def styleFullColour(co,ro,colour):
    styleOK = xlwt.easyxf()
    pattern = xlwt.Pattern()#一个实例化的样式类
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN # 固定的样式
    pattern.pattern_fore_colour = colour#背景颜色
    styleOK.pattern = pattern
    ws.write(co, ro, "",style=styleOK)


if __name__ == "__main__":
    print("statr...!")
    wb = xlwt.Workbook()
    ws = wb.add_sheet('A Test Sheet') # 增加sheet

    print(xlwt.Style.colour_map['yellow'])  # 背景颜色
    # styleOK = xlwt.easyxf()
    # pattern = xlwt.Pattern()#一个实例化的样式类
    # pattern.pattern = xlwt.Pattern.SOLID_PATTERN # 固定的样式
    # pattern.pattern_fore_colour = xlwt.Style.colour_map['red']#背景颜色
    # styleOK.pattern = pattern

    for i in range(10):
        styleFullColour(0,i,4)

    wb.save('example.xls')   # 保存xls
