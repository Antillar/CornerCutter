from PIL import Image, ImageOps, ImageEnhance
import cv2
import os
from pathlib import Path
from pytesseract import pytesseract 
import xlwt
import re
import subprocess


def pad_number(_match):
    number = int(_match.group(1))
    return format(number, "03d")


def get_offsets(_width, _height):
    k_x1 = 0
    k_y1 = 0
    k_x2 = 1
    k_y2 = 1
    a_x1 = 0
    a_y1 = 0
    a_x2 = 1
    a_y2 = 1
    o_x1 = 0
    o_y1 = 0
    o_x2 = 1
    o_y2 = 1

    # FullHD monitors 16:9 (1920x1080) and 16:10 (1920x1200)
    if 0.95 < _width / 1920 < 1.05 and (0.95 < _height / 1080 < 1.05 or 0.95 < _height / 1200 < 1.05):
        k_x1 = 860
        k_y1 = 696
        k_x2 = 905
        k_y2 = 724
        a_x1 = 860
        a_y1 = 747
        a_x2 = 905
        a_y2 = 776
        o_x1 = 800
        o_y1 = 799
        o_x2 = 905
        o_y2 = 828

    # UltraWide FullHD monitors 21x9 (2560x1080)
    if 0.95 < _width / 2560 < 1.05 and 0.95 < _height / 1080 < 1.05:
        k_x1 = 1180
        k_y1 = 696
        k_x2 = 1225
        k_y2 = 724
        a_x1 = 1180
        a_y1 = 747
        a_x2 = 1225
        a_y2 = 776
        o_x1 = 1130
        o_y1 = 799
        o_x2 = 1225
        o_y2 = 828

    # 2K monitors 16:9 (2560x1440) and 16:10 (2560x1600)
    if 0.95 < _width / 2560 < 1.05 and (0.95 < _height / 1440 < 1.05 or 0.95 < _height / 1600 < 1.05):
        k_x1 = 1136
        k_y1 = 928
        k_x2 = 1207
        k_y2 = 965
        a_x1 = 1136
        a_y1 = 997
        a_x2 = 1207
        a_y2 = 1035
        o_x1 = 1090
        o_y1 = 1066
        o_x2 = 1207
        o_y2 = 1104

    # UltraWide 2K monitors 21:9 (3440x1440)
    if 0.95 < _width / 3440 < 1.05 and 0.95 < _height / 1440 < 1.05:
        k_x1 = 1580
        k_y1 = 928
        k_x2 = 1647
        k_y2 = 965
        a_x1 = 1580
        a_y1 = 997
        a_x2 = 1647
        a_y2 = 1035
        o_x1 = 1530
        o_y1 = 1066
        o_x2 = 1647
        o_y2 = 1104

    # 4K monitor 16x9 (3840x2160)
    if 0.95 < _width / 3840 < 1.05 and 0.95 < _height / 2160 < 1.05:
        k_x1 = 1720
        k_y1 = 1392
        k_x2 = 1810
        k_y2 = 1448
        a_x1 = 1720
        a_y1 = 1494
        a_x2 = 1810
        a_y2 = 1552
        o_x1 = 1600
        o_y1 = 1598
        o_x2 = 1810
        o_y2 = 1656

    # Tel_Rivka's setup: 2K 16:9 monitor bottom left, on the right 2x FullHD monitors on top of each other. (4480x2160)
    if 0.95 < _width / 4480 < 1.05 and 0.95 < _height / 2160 < 1.05:
        k_x1 = 1136
        k_y1 = 1648
        k_x2 = 1207
        k_y2 = 1686
        a_x1 = 1136
        a_y1 = 1717
        a_x2 = 1207
        a_y2 = 1755
        o_x1 = 1090
        o_y1 = 1786
        o_x2 = 1207
        o_y2 = 1824
    return k_x1, k_y1, k_x2, k_y2, a_x1, a_y1, a_x2, a_y2, o_x1, o_y1, o_x2, o_y2


rootdir = os.path.dirname(os.path.realpath(__file__))
rootdir = rootdir + '\\raw'

# If you installed tesseract to somewhere else, replace the path to executable here
path_to_tesseract = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
pytesseract.tesseract_cmd = path_to_tesseract

# Uses an external converter to transform .jxr file to .png file, by default - root directory of CornerCutter
path_to_jxr_converter = r"jxr_to_png.exe"

if not os.path.exists('cut'):
    os.makedirs('cut')

raw_results = []

for subdir, dirs, files in os.walk(rootdir):
    for file in files:
        fileName = Path(str(os.path.join(subdir, file)))
        fileNameNoExt, fileExt = os.path.splitext(fileName)
        newFileName = fileNameNoExt + '.png'
        if fileExt == '.jxr':
            subprocess.run([path_to_jxr_converter, fileName, newFileName])
            os.remove(fileName)

for subdir, dirs, files in os.walk(rootdir):
    for file in files:
        player = os.path.basename(subdir)
        fileName = Path(str(os.path.join(subdir, file))).stem
        fileName = re.sub(r"(?:^|(?<=[^0-9]))([0-9]{1,3})(?=$|[^0-9])", pad_number, fileName)
        path_K = os.path.join('cut\\' + player, fileName + '_K.png')
        path_A = os.path.join('cut\\' + player, fileName + '_A.png')
        path_O = os.path.join('cut\\' + player, fileName + '_O.png')
        rawImg = Image.open(os.path.join(subdir, file)).convert('RGB')
        img_width, img_height = rawImg.size
        K_x1, K_y1, K_x2, K_y2, A_x1, A_y1, A_x2, A_y2, O_x1, O_y1, O_x2, O_y2 = get_offsets(img_width, img_height)
        rawImg = ImageOps.invert(rawImg)

        # Cuts out all the "corners" except the necessary info
        img_K = rawImg.crop((K_x1, K_y1, K_x2, K_y2))
        img_A = rawImg.crop((A_x1, A_y1, A_x2, A_y2))
        img_O = rawImg.crop((O_x1, O_y1, O_x2, O_y2))
        if not os.path.exists('cut\\' + player):
            os.makedirs('cut\\' + player)

        # Lower the number to reduce the effect of horizontal strikethrough lines but also lower the quality of numbers
        img_K = ImageEnhance.Contrast(img_K).enhance(2)
        img_A = ImageEnhance.Contrast(img_A).enhance(2)
        img_O = ImageEnhance.Contrast(img_O).enhance(2)
        
        img_K.save(path_K)
        img_A.save(path_A)
        img_O.save(path_O)
        
        # denoise algorithm through OpenCV2 - not ideal, but I cannot come up with anything better,
        # perhaps AI upscaling and more sharpening would yield better results
        denoised = cv2.fastNlMeansDenoising(cv2.imread(path_K,  cv2.IMREAD_GRAYSCALE), None, 3, 7, 21)
        out_binary = cv2.threshold(denoised, 0, 255, cv2.THRESH_OTSU)[1]
        cv2.imwrite(path_K, out_binary)
        denoised = cv2.fastNlMeansDenoising(cv2.imread(path_A,  cv2.IMREAD_GRAYSCALE), None, 3, 7, 21)
        out_binary = cv2.threshold(denoised, 0, 255, cv2.THRESH_OTSU)[1]
        cv2.imwrite(path_A, out_binary)
        denoised = cv2.fastNlMeansDenoising(cv2.imread(path_O,  cv2.IMREAD_GRAYSCALE), None, 3, 7, 21)
        out_binary = cv2.threshold(denoised, 0, 255, cv2.THRESH_OTSU)[1]
        cv2.imwrite(path_O, out_binary)
        
        # Has huge problems with recognition of number 1 due to dogshit font EA used (1 looks like |), otherwise works fine
        text_K = pytesseract.image_to_string(Image.open(path_K), lang="eng", config='--psm 7 --oem 3 -c tessedit_char_whitelist=0123456789')
        text_A = pytesseract.image_to_string(Image.open(path_A), lang="eng", config='--psm 7 --oem 3 -c tessedit_char_whitelist=0123456789')
        text_O = pytesseract.image_to_string(Image.open(path_O), lang="eng", config='--psm 7 --oem 3 -c tessedit_char_whitelist=0123456789')
        if text_K[:-1].isnumeric():
            int_K = int(text_K[:-1])
        else:
            int_K = -1
        if text_A[:-1].isnumeric():
            int_A = int(text_A[:-1])
        else:
            int_A = -1
        if text_O[:-1].replace(',', '').replace('.', '').isnumeric():
            int_O = int(text_O[:-1].replace(',', '').replace('.', ''))
        else:
            int_O = -1
        
        raw_results += [[player, fileName, int_A, int_K, '', int_O]]

# Sorting the raw results array by player and filename
raw_results = sorted(sorted(raw_results, key=lambda x: x[1]), key=lambda x: x[0])

wb = xlwt.Workbook()
xlwt.add_palette_colour('errors', 0x21)
wb.set_colour_RGB(0x21, 255, 0, 0)
xlwt.add_palette_colour('warnings', 0x22)
wb.set_colour_RGB(0x22, 255, 255, 0)
xlwt.add_palette_colour('totals', 0x23)
wb.set_colour_RGB(0x23, 200, 200, 200)
style_error = xlwt.easyxf('pattern: pattern solid, fore_colour errors')
style_warning = xlwt.easyxf('pattern: pattern solid, fore_colour warnings')
style_total = xlwt.easyxf('pattern: pattern solid, fore_colour totals')
ws = wb.add_sheet('Raw results')
header = ['Player', 'Filename', 'Assists', 'Kills', 'K+A', 'Objective score']
for column, heading in enumerate(header):
    ws.write(0, column, heading, style_total)

prev_player = ''
num_blanks = 0
row_start_player = 0
row_end_player = 0
row = 0
for row, results_array in enumerate(raw_results):
    if prev_player != results_array[0]:
        if prev_player != '':
            row_end_player = row + num_blanks + 1
            ws.write(row + num_blanks + 1, 0, prev_player, style_total)
            ws.write(row + num_blanks + 1, 1, 'TOTALS', style_total)
            ws.write(row + num_blanks + 1, 2, xlwt.Formula('SUM(C' + str(row_start_player) + ':C' + str(row_end_player) + ')'), style_total)
            ws.write(row + num_blanks + 1, 3, xlwt.Formula('SUM(D' + str(row_start_player) + ':D' + str(row_end_player) + ')'), style_total)
            ws.write(row + num_blanks + 1, 4, xlwt.Formula('SUM(C' + str(row + num_blanks + 2) + ':D' + str(row + num_blanks + 2) + ')'), style_total)
            ws.write(row + num_blanks + 1, 5, xlwt.Formula('SUM(F' + str(row_start_player) + ':F' + str(row_end_player) + ')'), style_total)
            num_blanks = num_blanks + 1
        prev_player = results_array[0]
        row_start_player = num_blanks + row + 2
    for column, value in enumerate(results_array):
        if value == -1:
            ws.write(num_blanks + row + 1, column, value, style_error)
        else: 
            if ((column == 2 or column == 3) and value >= 100) or (column == 5 and (value >= 10000 or value < 100)):
                ws.write(num_blanks + row + 1, column, value, style_warning)
            else:
                ws.write(num_blanks + row + 1, column, value)
            
ws.write(row + num_blanks + 2, 0, prev_player, style_total)
ws.write(row + num_blanks + 2, 1, 'TOTALS', style_total)
ws.write(row + num_blanks + 2, 2, xlwt.Formula('SUM(C' + str(row_end_player + 2) + ':C' + str(row + num_blanks + 2) + ')'), style_total)
ws.write(row + num_blanks + 2, 3, xlwt.Formula('SUM(D' + str(row_end_player + 2) + ':D' + str(row + num_blanks + 2) + ')'), style_total)
ws.write(row + num_blanks + 2, 4, xlwt.Formula('SUM(C' + str(row + num_blanks + 3) + ':D' + str(row + num_blanks + 3) + ')'), style_total)
ws.write(row + num_blanks + 2, 5, xlwt.Formula('SUM(F' + str(row_end_player + 2) + ':F' + str(row + num_blanks + 2) + ')'), style_total)

wb.save('RawData.xls')
