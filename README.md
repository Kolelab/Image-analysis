# Image-analysis

Scripts are written in Microsoft Visual Basic and were confirmed to run on Windows compatible computers in Excel 2013 or newer. This code provided here is only the source code and needs to be imported into Microsoft Visual Basic 

Image splitter: The splitter is implemented to work on .da files as saved by Neuroplex (RedShirt Imaging).
Every other frame of the .da file is split and as a final result two files are saved (odd 1-3-5 and so on and even 2-4-6...)

∆F/F calculator: Raw fluorescence values are imported from .txt files. Text files have to be formated in columns in which each column are the data for one ROI. The ROI with the lowest F values is assumed to be the background ROI and is subtracted from all given ROIs. The average fluorescence of each column is then determined and used to calculate ∆F/F. The output is saved in a new tab in the Excel file. 
