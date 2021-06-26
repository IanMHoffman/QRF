# QRF
 QRF (Quick Reference File)
 Handy things that I seem to forget. Will probably need to organize this in a better way later.


# Python

## File Handling

Picking out components of a file path
```python
from pathlib import Path

path
PosixPath('/home/user/python/test.md')

# .name: the file name without any directory
path.name
'test.md'

# .stem: the file name without the suffix
path.stem
'test'

# .suffix: the file extension
path.suffix
'.md'

# .parent: the directory containing the file, or the parent directory if path is a directory
path.parent
PosixPath('/home/user/python')

path.parent.parent
PosixPath('/home/user')

# .anchor: the part of the path before the directories
path.anchor
'/'
```

Handling Multiple desired file types
```python
img_formats = ['.bmp', '.jpg', '.jpeg', '.png', '.tif', '.tiff', '.dng']
vid_formats = ['.mov', '.avi', '.mp4', '.mpg', '.mpeg', '.m4v', '.wmv', '.mkv']

p = str(Path(path))  # os-agnostic
p = os.path.abspath(p)  # absolute path
files = []
all_files = []

all_files = sorted(Path(p).rglob('*.*'))
for x in all_files:
    if re.match('Snagged_TO BE CROPPED', x.parts[len(x.parts) - 3], re.IGNORECASE):
        files.append(x)

images = [x for x in files if os.path.splitext(x)[-1].lower() in img_formats]
videos = [x for x in files if os.path.splitext(x)[-1].lower() in vid_formats]
ni, nv = len(images), len(videos)
```

Find all the folders/directories with TQDM progress bar.
```python
myList = []
for file in tqdm.tqdm(os.listdir(startPath),"Locating Folders"):
    if os.path.isdir(os.path.join(startPath, file)):
        myList.append(os.path.join(startPath, file))
```

## import logging
```python
import logging

logging.basicConfig(filename='my-log.log', filemode='w', level=logging.DEBUG)

# Message Types
logging.debug("debug message")
logging.info("info message")
logging.warning("warning message")
logging.error("error message")
logging.critical("critical message")
```

## import xlrd

Read in all excel tabs to one list of dictionaries
```python
def readExcelFile(excelFile):
    # Open Excel Workbook
    book = xlrd.open_workbook(excelFile)
    sheets = book.sheets()
    disc_list = []
    # Read Data From Sheets One at a Time
    for sheet in tqdm.tqdm(sheets,'Processing excel sheet Tabs'):
        
        worksheet = book.sheet_by_name(sheet.name)
        header = worksheet.row_values(0) # Select first row as header
        values = [worksheet.row_values(i) for i in range(1, sheet.nrows)] # Skip first row of headers

        for value in values:
            for item in value:
                item = str(item)
            # zip the headings and row values together into a dict
            disc_list.append(dict(zip(header, value)))

    return(disc_list)
```

Read in one excel tab by name to a list of dictionaries
```python
def readExcelFile(excelFile, sheet):
    # Open Excel Workbook
    book = xlrd.open_workbook(excelFile)
    disc_list = []
    worksheet = book.sheet_by_name(sheet)
    header = worksheet.row_values(0) # Select 2nd row as header
    values = [worksheet.row_values(i) for i in range(1, sheet.nrows)] # Skip first row of headers

    for value in values:
        for item in value:
            item = str(item)
        # zip the headings and row values together into a dict
        disc_list.append(dict(zip(header, value))) 

    return(disc_list)
```
## import pyodbc

Read the tables in a access database.
```python
def printDatabaseTables(db_path):
    cnxn = pyodbc.connect('Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+str(db_path)+';')
    cursor = cnxn.cursor()
    for row in cursor.tables():
        print(row.table_name)
```

Open a table and zip headers to content to make a list of dictionaries.
```python
def readDatabase(db_path, table):
    db_list = []
    cnxn = pyodbc.connect('Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+str(db_path)+';')
    cursor = cnxn.cursor()
    sql = 'select * from ' + str(table)
    cursor.execute(sql)
    dbData = cursor.fetchall()
    desc = cursor.description
    colNames = [col[0] for col in desc]
    for row in dbData:
        db_list.append(dict(zip(colNames, row)))
    return db_list
```

# VBA

## Protect All Sheets
```
Sub Protect()
      'Loop through all sheets in the workbook
      For i = 1 To Sheets.Count
         Sheets(i).Protect
      Next i
End Sub
```

## Unprotect All Sheets
```
Sub UnProtect()
      'Loop through all sheets in the workbook
      For i = 1 To Sheets.Count
         Sheets(i).UnProtect
      Next i
End Sub
```