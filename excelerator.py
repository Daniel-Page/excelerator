import time
start = time.time()
import os
os.system('cls') # Clear terminal
import glob
from win32com import client
from pdfCropMargins import crop

dirName = os.path.dirname(__file__) # Directory that this script is in

fileNames = [] # Excel file names found
i = 1 # File counter
# Find all Excel in the same directory as this script
for file in glob.glob("*.xlsx"):
    if i == 1:
      print("Excel files found in " + dirName + ":")
    fileNames.append(file)
    print(str(i) + ": " + file) # Display the Excel files found
    i += 1

if len(fileNames) > 0:

  # Check whether the output folder exists
  outputExists = os.path.isdir("output")

  # Create a new directory if output does not exist 
  if not outputExists:
    os.mkdir("output")

  excel = client.Dispatch("Excel.Application") # Open Microsoft Excel

  print("\nPDFs exported to " + os.path.join(dirName + '\\output') + ":")
  
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
      # -c p: calculates bounding box with pdftoppm
      # -t NUM: threshold
      # -a4 BP BP BP BP: absolute offset
      # -p4 PCT PCT PCT PCT: percentage retain 
      # Order is "left","bottom","right","top"
      # -mo: modify original
      # -x DPI, -y DPI
      # quiet: stop errors from being printed
      crop(["-c","p","-t","10","-a4", "0","0","0","0.05","-p4", "0","0","0","0","-mo","-x","500","-y","500", file.strip(".xlsx") + '.pdf'], quiet=True)
      os.remove(os.path.join(dirName + '\\output\\', file.strip(".xlsx") + '_uncropped.pdf'))
      print(str(i) + ": " + file.strip(".xlsx") + '.pdf')
      i += 1

  end = time.time()
  print("\nFinished in " + str(round(end - start,2)) + "s\n")

else:
  print("No Excel files found in " + dirName + "\n")