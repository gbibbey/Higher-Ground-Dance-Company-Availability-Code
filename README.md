# Higher-Ground-Dance-Company-Availability-Code
highlights the availability for dancers in order to choose which dances they can be in
Follow closely for the code to work properly:

1. Download your XXX XXX Dance Preferences (Responses) sheet from Google Sheets

2. Make sure there are NO DUPLICATE FORMS from people that may have submitted multiple times 
   (can alphabetize to check, but make sure you put it back to submission order before sorting 
   begins to keep it fair) and make sure each person's name is spelled the same way for each preference.

3. Make sure all the times in the Sunday-Thursday columns are on the left side of the column and not the right
   (or else google will think it is a date instead of a time)

4. "Save As" this spreadsheet as an Excel 97-2003 Workbook and change the name to FullResponseSheet 

5. Add it to this "Avilability Code" folder

6. Look at the dancesAndTimes.xls and HAND TYPE all the new dance names/days/times so it is EXACTLY like the 
   old one (or duplicate the file, change it, and then add it to the folder)

7. open availability_code.py 
	a. you will need to have python downloaded (I prefer installing the VSCode extension)
	b. run pip install pandas in a terminal
	c. run pip install xlrd in terminal
	d. run pip install openpyxl in terminal
	e. i think that is everything you need to install, but the terminal will tell you when you try to run it if	
	   there is something else it needs

8. In the FullResponseSheet.xls sheet, count from 0 to find the column number of the "Sunday" column

9. In the availability_code.py line 12, change the range(x, y) to have x be the column number of the 
   "Any other comments or concerns" column starting from 0 and y to be the column number of the "Sunday" 
   column starting from 0

10. In line 44, change the index = to be the column number of the "Sunday" column and then fix the index 
	numbers following to be increased by 1

11. Run the availability_code script and it should add a highlightedAvailability.xlsx to the Avilability Code 
    folder.
	In this sheet, the dancers names are in the first row and all of the dances they CAN be in are underneath 
	their name with the green highlights being ones they signed up for.

Happy Sorting! Text me if you have questions 708-200-6793!
