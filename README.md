#studentresult

resultanalysis.py: Download the students result from website http://www.tekerala.org/.

Usage: Type at the command prompt

python copytxt.py -i <inputfile> -o <outputfile> 

inputfile     =  name of the input file. Must be a excel file  (with extension .xlsx)  containing register number in first column and date of birth in second coulmn. Register no and DOB must be in the first excel sheet with name Sheet1 
outputfile    =  name of the output file. Must be a .xlsx file.

eg:

python resultanalysis.py -i regno.xlsx -o result.xlsx



Bugs and issues can be repoted to author at williamdoyleaf@gmail.com
