# Excelerator  

A script that automatically exports and crops Excel tables as individual pdf files.  

This is useful if you want to add Excel tables to a LaTeX document (assuming the widths and text sizes are consistent).  

However, it only works with Excel installed on Windows and very slight undercropping can occur around the edges of the tables (this can be tuned by changing the absolute offsets).  

Last tested with Python 3.10.5 (64-bit).  

To install the required Python libraries:  
```
pip install -r requirements.txt
```

To the run the script:
```
python .\excelerater.py
```

All .xlsx files in the same directory as the script will be processed and exported to an output folder.

<!-- To do:
Handle pdf already open
Check consistency for different table sizes
Check if excel actually installed
Check it is Windows running script
Turn into executable
Excel table width manipulation
Cycle through different sheets in Excel files
PDF and PDFs print -->
