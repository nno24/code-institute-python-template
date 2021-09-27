## .xls to .map
This program will convert X,Y coordinates in meters from a .xls format to a .map format and add a text label corresponding the "ID" to the map object.
The output will be a .txt file that can further be  used inside a .map file.
The program will use the first and last X,Y coordinates to the corresponding "ID" in meters and add a text label
equal to the "ID".

The user must specify input document in .xls format, id column, X coordinate column, and Y coordinate column inside the python program.

## Future enhancements
-Add support for specifying workbook,outputfile,id,x,and y columns outside the python code.
-Add support for higher resolution, using more than first and last coordinate.
-Add support for inputing the converted data directly to a .map file, usually this is an existing file that will be modified.

![alt text]("input_vs_output.JPG" "Title")