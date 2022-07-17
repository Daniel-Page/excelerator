# Excelerator  

A script that automatically exports and crops Excel tables as individual pdf files.  

This is useful if you want to add Excel tables to a LaTeX document (assuming the width is consistent).  

However, it only works with Excel installed on Windows and very slight undercropping can occur around the edges of the tables (this can be tuned by changing the absolute offsets).  

To install the required Pyhton libraries:  
```
pip install -r requirements.txt
```

To the run the script:
```
python .\excelerater.py
```

<!-- To do:
Handle no excel documents
Handle pdf already open
Check consistency for different table sizes
Check if excel actually installed
Check it is Windows running script
Turn into executable
Excel table width manipulation
Cycle through different sheets in Excel files
What if folder is empty
PDF and PDFs print
Check requiements contain all dependencies -->