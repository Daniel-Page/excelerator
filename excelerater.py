import time
start = time.time()
import os
os.system('cls') # Clear terminal
import glob
from win32com import client
from pdfCropMargins import crop

dirName = os.path.dirname(__file__) # Directory that this script is in
print("Excel files found in " + dirName + ":")

fileNames = [] # Excel file names found
i = 1 # File counter
# Find all Excel in the same directory as this script
for file in glob.glob("*.xlsx"):
    fileNames.append(file)
    print(str(i) + ": " + file) # Display the Excel files found
    i += 1

# Check whether the output folder exists
outputExists = os.path.isdir("output")

# Create a new directory if output does not exist 
if not outputExists:
  os.mkdir("output")

excel = client.Dispatch("Excel.Application") # Open Microsoft Excel

print("\nFiles exported to " + os.path.join(dirName + '\\output') + ":")
i = 1 # File export counter
for file in fileNames:
    sheets = excel.Workbooks.Open(os.path.join(dirName, file)) # Read Excel File
    work_sheets = sheets.Worksheets[0] # Only first sheet exported
    work_sheets.ExportAsFixedFormat(0, os.path.join(dirName + '\\output\\', file.strip(".xlsx") + '.pdf'))
    sheets.Close(True)
    print(str(i) + ": " + file.strip(".xlsx") + '.pdf')
    i += 1

print("\nPDFs cropped in " + os.path.join(dirName + '\\output'))
i = 1 # File export counter
for file in fileNames:
    os.chdir(dirName + '\\output\\')
    # [-p4 left bottom right top] - PCT: percentage
    # [-a4 left bottom right top] - BP: big point
    crop(["-a4", "0.3","0.3","-0.2","0","-p4", "0","0","0","0", file.strip(".xlsx") + '.pdf'], quiet=True) # Quiet argument removes printed errors
    print(str(i) + ": " + file.strip(".xlsx") + '_cropped.pdf')
    i += 1

print("\nPDFs deleted from " + os.path.join(dirName + '\\output'))
i = 1 # File delete counter
for file in fileNames:
  os.remove(os.path.join(dirName + '\\output\\', file.strip(".xlsx") + '.pdf'))
  print(str(i) + ": " + file.strip(".xlsx") + '.pdf')
  i += 1

end = time.time()
print("\nFinished in " + str(round(end - start,2)) + "s\n")