# python-excel
图片
视频
可转
 0、安装依赖；
 1、将文件命名为“1.jpg”保存在项目根目录下，其他后缀名名称自行更改代码，也可复制下面代码覆盖全部也可行（临时写了可能有错误）；
 2、输出文件保存在“outputs”文件下，默认为“1.xlsx”;
 

 
from optparse import Option
import cv2
import os
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter


def toHex(rgb):
    color = ""
    color += str(hex(rgb[0])).replace('x', '0')[-2:]
    color += str(hex(rgb[1])).replace('x', '0')[-2:]
    color += str(hex(rgb[2])).replace('x', '0')[-2:]
    return color


def video2xlsx(cap, size=(0, 0), jump=0, outputPixelSize=(1, 1)):
    tabelNum = 1
    # while True:
    print('正在处理第' + str(tabelNum) + '帧...')
        # delay = jump
        # while delay != 0:
        #     ret, frame = cap.read()
        #     delay -= 1
    frame = cap
        # if size != (0, 0):
    
    w, h= frame.shape[:2]
    size=(int(h*0.5),int(w*0.5))
    frame = cv2.resize(frame, size)
    w, h= frame.shape[:2]
    wb = Workbook()
    ws = wb.active
    for i in range(h):
        ws.column_dimensions[get_column_letter(i+1)].width = outputPixelSize[0]
    for i in range(w):
        ws.row_dimensions[i+1].height = outputPixelSize[1]
    print("开始采集")
    print("h:",h,"w:",w)
    for i in range(h):
        for j in range(w):
            fill = PatternFill(fill_type="solid", fgColor=toHex(frame[j][i]))
            ws.cell(j + 1, i + 1).fill = fill
    print("结束采集")
    print("正在保存")
    try:
        wb.save('./outputs/' + str(tabelNum) + '.xlsx')
    except:
        print('保存失败，请检查文件是否已被打开')
    print("保存完成")
    if cv2.waitKey(40) & 0xFF == ord('q'):
        print('完成！')
        os.system('pause')
        # break
    tabelNum += 1

 
 
 
 def pathE(path1):
    while path1:
        if os.path.exists(path1):
            print("路径:"+path1)
            path1 = ""
        else:
            print('未找到文件，请重新输入：重启吧，我还没写，哈哈！！')
            path1 = ""
 
 option = input('请输入数字“1”或“2”选择图片文件路径输入方式（回车默认1.jpg)：\n1 输入图片的绝对路径；2 保存在项目根目录再输入图片名称及后缀名\n')
if option == "1":
    test = input('请输入图片的绝对路径(不能出现中文):')
    pathE(test)
    cap = cv2.imread(test)
elif option == "2" :
    test = input('请保存在项目根目录再输入图片名称及后缀名（不能出现中文）：')
    capName = os.path.abspath(".")
    test = capName+"\\"+test
    pathE(test)
    cap = cv2.imread(test)
elif option == "":
    test = "1.jpg"
    pathE(test)
    cap = cv2.imread(test)

sizeW = input('请输入缩小后宽度（默认不缩小）：')
sizeH = input('请输入缩小后高度（默认不缩小）：')
if sizeW == '' or sizeH == '':
    sizeW = '0'
    sizeH = '0'
jump = input('请设置每跳过多少帧采集一次（默认为0）：')
if jump == '':
    jump = '0'
outputkr = input('请设置输出单元格宽（默认为1）：')
if outputkr == '':
    outputkr = '1'
outputgc = input('请设置输出单元格高（默认为1）：')
if outputgc == '':
    outputgc = '1'

video2xlsx(cap, (int(sizeW), int(sizeH)), int(jump), (int(outputkr), int(outputgc)))
