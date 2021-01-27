# FileNameFinder.py

## About:
Detects filenames ending in .cpp and .h in a excel spreadsheet. Outputs all of the found filenames in a excel spreadsheet. Tested on Windows 10.
It is by no means perfect and can be adjusted for your specific means.

## Set up:
1. Make sure Python (tested on version 3.7) and pip (tested on v19.03) are installed. 
2. Once Python is installed. Install the library openpyxl:
```pip install openpyxl```
3. Copy the script into a local folder. It will run better when not on a network drive.

## How to:
1. Open up the directory of the local copy of FileNameFinder.py in a terminal or powershell.
2. Run the command:
```python FileNameFinder.py```
3. You will be prompted to enter the input file location and name. Write the full path. Ex: C:\Workspace\Scripts\Filename.xlsx
4. You will next be prompted next to enter the file location and name you want the output to be. Write the full path. Ex: C:\Workspace\Scripts\output.xlsx
5. The program will display how many filenames it found and the list of filenames in the terminal. The output file will be placed in the location you specified.

## Common Errors
1. Make sure the output filename doesn't already exist