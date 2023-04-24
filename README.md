# csv2excel

This is a Python script to simply generate an excel file from a csv file. The contents of csv file will be copied to the excel file. The file name of the excel file inherits the one of the csv file.

## Usage on Windows

Suppose Python system installed. The script uses [openpyxl](https://pypi.org/project/openpyxl/). So
```
py -m pip install openpyxl
```
if you have not installed it. For example,
```
py csv2excel.py sample.csv
```
generates `sample.xlsx` at the same directory. Similarly,
```
py csv2excel.py sample.csv sample2.csv
```
generates `sample.xlsx` and `sample2.xlsx` at the same directory. Any file which doesn't end with `.csv` will be skipped.

## Making script standalone executable on Windows

I recommend to use [Nuitka](https://github.com/Nuitka/Nuitka) alternative to [Pyinstaller](https://pypi.org/project/pyinstaller/).
```
py -m pip install nuitka
```
You may be asked Yes or No several times. Then ask Yes. These are for mingw64 with c language complier on Windows. Also
```
py -m pip install imageio
```
will prepare [imageio](https://pypi.org/project/imageio/) to use an png icon file.

Now you can make an exe file (everything included) of `csv2excel.py` as follows.
```
py -m nuitka --disable-console --onefile --windows-icon-from-ico=icon.png csv2excel.py
```
After a few minutes, you will find `csv2excel.exe` in the directory `dist`.

This enables us to use our command without Python system. You can use the command line like
```
csv2excel.exe sample.csv sample2.csv
```
to generate two Excel files. Note that your anti-virus software might prevent [Nuitka](https://github.com/Nuitka/Nuitka) from making an exe file. You can avoid this by turning off the software temporarily at your own risk.

## Drag and drop

On Windows desktop, you can generate Excel files by selecting two csv files with mouse, and dragging and dropping them to the icon of `csv2excel.exe`. This is another way of the previous command line. You will see `sample.xlsx` and `sample2.xlsx` in the same directory as the csv files.