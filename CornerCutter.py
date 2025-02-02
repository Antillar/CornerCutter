from PIL import Image, ImageOps, ImageEnhance
import cv2
import os
from pathlib import Path
from pytesseract import pytesseract 
import xlwt

rootdir = os.path.dirname(os.path.realpath(__file__))
rootdir = rootdir + '\\raw'

#If you installed tesseract to somewhere else, replace the path to executable here
path_to_tesseract = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
pytesseract.tesseract_cmd = path_to_tesseract 

if not os.path.exists('cut'):
    os.makedirs('cut')


raw_results = []

for subdir, dirs, files in os.walk(rootdir):
    for file in files:
        player = os.path.basename(subdir)
        fileName = Path(os.path.join(subdir, file)).stem
        path_K = os.path.join('cut\\' + player, fileName + '_K.png')
        path_A = os.path.join('cut\\' + player, fileName + '_A.png')
        path_O = os.path.join('cut\\' + player, fileName + '_O.png')
        rawImg = Image.open(os.path.join(subdir, file)).convert('RGB')
        img_width, img_height = rawImg.size
        rawImg = ImageOps.invert(rawImg)
        
        #Full HD monitor (1920x1080) 16:9
        if( img_height > 1080 * 0.9 and img_height < 1080 * 1.1 
            and img_width > 1920 * 0.9 and img_width < 1920 * 1.1):
            K_x1 = 860
            K_x2 = 905
            K_y1 = 696
            K_y2 = 724
            
            A_x1 = 860
            A_x2 = 905
            A_y1 = 747
            A_y2 = 776
            
            O_x1 = 800
            O_x2 = 905
            O_y1 = 799
            O_y2 = 828
            
        #2K monitor (2560x1440) 16:9
        if( img_height > 1440 * 0.9 and img_height < 1440 * 1.1 
            and img_width > 2560 * 0.9 and img_width < 2560 * 1.1):
            K_x1 = 1136
            K_x2 = 1207
            K_y1 = 928
            K_y2 = 965
            
            A_x1 = 1136
            A_x2 = 1207
            A_y1 = 997
            A_y2 = 1035
            
            O_x1 = 1090
            O_x2 = 1207
            O_y1 = 1066
            O_y2 = 1104
        
        #Cuts out all the "corners" except the necessary info
        img_K = rawImg.crop((K_x1, K_y1, K_x2, K_y2))
        img_A = rawImg.crop((A_x1, A_y1, A_x2, A_y2))
        img_O = rawImg.crop((O_x1, O_y1, O_x2, O_y2))
        if not os.path.exists('cut\\' + player):
            os.makedirs('cut\\' + player)
        
        img_K = ImageEnhance.Contrast(img_K).enhance(2)
        img_A = ImageEnhance.Contrast(img_A).enhance(2)
        img_O = ImageEnhance.Contrast(img_O).enhance(2)
        
        img_K.save(path_K)
        img_A.save(path_A)
        img_O.save(path_O)
        
        #denoise algorythm through OpenCV2 - not ideal, but I cannot come up with anything better, perhaps AI upscaling and more sharpening would yield better results       
        denoised = cv2.fastNlMeansDenoising(cv2.imread(path_K,  cv2.IMREAD_GRAYSCALE), None, 3, 7, 21)
        out_binary=cv2.threshold(denoised, 0, 255, cv2.THRESH_OTSU )[1] 
        cv2.imwrite(path_K,out_binary)
        denoised = cv2.fastNlMeansDenoising(cv2.imread(path_A,  cv2.IMREAD_GRAYSCALE), None, 3, 7, 21)
        out_binary=cv2.threshold(denoised, 0, 255, cv2.THRESH_OTSU )[1] 
        cv2.imwrite(path_A,out_binary)
        denoised = cv2.fastNlMeansDenoising(cv2.imread(path_O,  cv2.IMREAD_GRAYSCALE), None, 3, 7, 21)
        out_binary=cv2.threshold(denoised, 0, 255, cv2.THRESH_OTSU )[1] 
        cv2.imwrite(path_O,out_binary)
        
        #Has huge problems with recognition of number 1 due to dogshit font EA used (1 looks like |), otherwise works fine
        text_K = pytesseract.image_to_string(Image.open(path_K), lang="eng", config='--psm 7 --oem 3 -c tessedit_char_whitelist=0123456789')
        text_A = pytesseract.image_to_string(Image.open(path_A), lang="eng", config='--psm 7 --oem 3 -c tessedit_char_whitelist=0123456789')
        text_O = pytesseract.image_to_string(Image.open(path_O), lang="eng", config='--psm 7 --oem 3 -c tessedit_char_whitelist=0123456789')
        if(text_K[:-1].isnumeric()):
            int_K = int(text_K[:-1])
        else:
            int_K = -1
        if(text_A[:-1].isnumeric()):
            int_A = int(text_A[:-1])
        else:
            int_A = -1
        if(text_O[:-1].replace(',', '').replace('.', '').isnumeric()):
            int_O = int(text_O[:-1].replace(',', '').replace('.', ''))
        else:
            int_O = -1
        
        raw_results += [[player, fileName, int_K, int_A, int_O]]

wb = xlwt.Workbook()
ws = wb.add_sheet('Raw results')
header = ['Player', 'Filename', 'Kills', 'Assists', 'Objective score']
for column, heading in enumerate(header):
    ws.write(0, column, heading)

for row, results_array in enumerate(raw_results):
    for column, value in enumerate(results_array):
        ws.write(row + 1, column, value)

wb.save('RawData.xls')
