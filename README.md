# CornerCutter
Program to transform Star Wars Battlefront 2 screenshots of personal tab of match results to Excel sheet.

## About the program
The core feature of the program is to use OCR to convert image into text. However, due to different languages the players have the game in, the text in the image can be vastly different, thus will provide poor quality result if put through OCR pipeline without any assumptions. To overcome the problem, firstly only the pieces with numbers are cut out from the image, which are near the center (hense the name corner cutter) and OCR is run on those.
There are problems with OCR, main one being - EA uses crappy font (which is even different for the language) and also adds those orange glare around everything. So the images need to be cleaned up before running the OCR.
Also, some barebones support was implemented for processing several players' data without it interfering with each other. Could be expanded in the future.
Currently the program only supports FullHD 16:9 AND 2K 16:9 screenshots, support for other resolutions can be added easily though once the examples of such screenshots are provided.

## How to run the program
1. The program requires Python v3, can be downloaded from here: [Python Website](https://www.python.org/downloads/windows/).
2. The OCR algorythm used is Tesseract OCR - the most standalone available, not the best, but it makes do, easiest to get up running though. Can be downloaded here: [SourceForge page](https://sourceforge.net/projects/tesseract-ocr.mirror/).
3. The OCR will try to install in Program Files X86 by default, if another location is chosen, the path to it is to be updated in the .py file.
4. Once both Python and Tesseract are installed, obtain the python modules necessary to run the program with the execution of the following command in CMD (once per module):
```
py -m pip install MODULENAME
```
These modules are required: **Pillow, xlwt, pytesseract, opencv-python**.

The program can be run via CMD with the following command:
```
py CornerCutter.py
```

## How to prepare data:
1. Create folder named "raw" inside the folder with CornerCutter.py file.
2. Inside "raw" folder create one folder for each player. If the player folder contains no files, it is ignored further.
3. Inside each player folder place the screenshots received from the player.
4. Run the CornerCutter.py.
5. The results should appear in RawData.xls
6. Open RawData.xls, check for any -1's - this means OCR failed to process a certain image. Go to "cut" folder, then to that player folder, find a prepared cut (_K for kills, _A for assists, _O for objective score) and manually imput it in the database. Unfortunately, OCR is not perfect. Better results are only available with pretrained model.
7. Check results for outliers, like 271 kills is certainly not true, needs to be manually double-checked. Another drawback of base trained OCR.
8. Once all -1's are filled with data and all outliers corrected, the data can be further used in formulas.
