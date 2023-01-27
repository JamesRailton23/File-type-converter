import jpype
import asposecells

# starts java VM
jpype.startJVM()
from asposecells.api import Workbook

# input file path
print('insert file path without " " ')
filePath = input()
# Output file path
print('insert export File path without " " and with file name and file type')
fileExport = input()

# Creates new file based on output file name
Workbook = Workbook(filePath)
Workbook.save(fileExport)

# program is done completed message is shown
print("file exported, Press enter to Exit")
input()

# python program shutsdown
jpype.shutdownJVM()
