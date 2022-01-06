# ReadDB

2021 year

My customer had a third party program. It allowed copying only one cell at a time.
The customer asked me to develop a program that copies all the table and writes to a file.

Firstly, set coordinates of cells in one raw + one coordinate for a scrolling button.
Thus, the whole table is get read.

Here I used keyboard hooks to set coordinates, start and stop reading.

F10 - Set coords of cells
F11 - Start reading
F12 - Stop reading, Write to file