# Excelerator  

A script that automatically exports and crops Excel tables as individual PDF files. This is useful if you want to add Excel tables to a LaTeX document (assuming the widths and text sizes are consistent).     

## Compatibility and Requirements

Works with:  
- Python 3.10.5 (64-bit)
- Windows 10 Pro
- Excel for Microsoft 365 Version 2206 (64-bit)   

To install the required Python libraries:  
```
pip install -r requirements.txt
```

## Running  

To the run the script:
```
python .\excelerater.py
```

All .xlsx files in the same directory as the script will be processed and exported to an output folder.

<!-- To do:
- Handle pdf already open
- Check consistency for different table sizes
- Check if excel actually installed
- Check it is Windows running script
- Turn into executable
- Automatic Excel table width manipulation
- Provide example of making width and font consistent with .tex file
- Cycle through different sheets in Excel files
- Status updates changes for single/multiple files e.g. PDF and PDFs -->
