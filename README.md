# Excelerator  

A script that automatically exports and crops Excel tables as individual PDF files. This is useful if you want to add Excel tables to a LaTeX document (with the table widths, fonts and font sizes consistent).     

## Compatibility and Requirements

Works with:  
- Python v3.10.5 (64-bit)
- Windows 10 Pro
- Excel for Microsoft 365 v2206 (64-bit)   

To install the required Python libraries:
```
pip install -r requirements.txt
```

## Running  

To run the script:
```
python .\excelerator.py
```

All .xlsx files in the same directory as the script will be processed and exported to an output folder.

<!-- Future improvements:
- Handle PDF file already open
- Check if excel actually installed
- Check it is Windows is running the script
- Implement for Mac
- Turn into executable
- Automatic Excel table width manipulation
- Automatically adjust font and font size
- Provide example of making width and font consistent with .tex file
- Cycle through different sheets in Excel files
- Status updates changes for single/multiple files e.g. PDF and PDFs -->
