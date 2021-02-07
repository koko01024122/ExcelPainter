import os
from tkinter import *
import tkinter.filedialog
import xlsxwriter
def RGB_to_Hex(rgb):
    RGB = rgb.split(',')            # 将RGB格式划分开来
    color = '#'
    for i in RGB:
        num = int(i)
        # 将R、G、B分别转化为16进制拼接转换并大写  hex() 函数用于将10进制整数转换成16进制，以字符串形式表示
        color += str(hex(num))[-2:].replace('x', '0').upper()

    return color

if __name__ == '__main__':
    from PIL import Image
    workbook = xlsxwriter.Workbook('C:\\Users\\koko0\\Desktop\\python.xlsx')
    worksheet = workbook.add_worksheet('Python')
    root = tkinter.Tk()  # 创建一个Tkinter.Tk()实例
    root.withdraw()  # 将Tkinter.Tk()实例隐藏
    default_dir = "C:\\Users\\koko0\\Desktop\\"
    file_path = tkinter.filedialog.askopenfilename(title=u'选择文件', initialdir=(os.path.expanduser(default_dir)))
    im = Image.open(file_path)
    rgb_im = im.convert('RGB')


    for i in range(rgb_im.height):
        for j in range(rgb_im.width):
            r, g, b = rgb_im.getpixel((j,i))
            cell_format = workbook.add_format()
            # cell_format.set_h
            cell_format.set_bg_color(RGB_to_Hex("{},{},{}".format(r,g,b)))
            worksheet.write(i,j,'',cell_format)
    worksheet.set_column('A:ZZ', 2)
    workbook.close()

