AIR3550 Instructions

SETUP: 
- There are no required passwords to run this code other than the passwords
of the users. Based on the dataset we used, all premade accounts have the password: Omniguy144$
- Our database is an excel file. The path for it is specified
in the database_connect() function of the functions.cs file within
public Excel.Workbook database_connect(). This method is around line 235.
STEPS:
	- Open the solution for Air3550
	- Go to functions.cs within the solution
	- Under Excel.Workbook database_connect(), change:
	- "Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\twild\Desktop\FinalAir3550Database.xlsx");"
	  to the appropriate path for the database.
	- Build the solution
	- You are now ready to run the application!

Pre-populated account User IDs and passwords:
	- Customer		- 278101	- Omniguy144$	
	- Load Engineer		- 293874	- Omniguy144$
	- Marketing Manager	- 923487	- Omniguy144$
	- Flight Manager	- 998324	- Omniguy144$
	- Accountant		- 384782	- Omniguy144$

BUGS/PROBLEMS:
- During a final double-check of the requirements, we realized some problems
with our flight searching page. We implemented the date to search for the flight
as a range instead of a single flight. We also forgot to allow the user to 
specify the time that they were hoping to travel and return on. We realized
these issues too late into development (the final two days before the final
submission) and did not feel we had enough time to rework our function.

- In the event that a bug that we missed causes the code to crash, there is a
specific process that must be implemented to restabilize the database. First,
open your task manager. Then, scroll down to background processes. It would
probably be best to sort the running programs by name. Then, find all 
instances of Microsoft Excel running in the background processes and end those
tasks. These are any open excel workbooks that were not closed prior to the
code crashing.

-There seems to be some minor issues when calculating flight time for creating a new flight

- Within flight manifest, the manifest function correctly adjusts Flight.Taken from FALSE to TRUE
However, it does not update User data. USer data will remain the same as if flight never took off

- A similar issue to the above occurs when deleting a flight. Deleting a flight properly updates
the flights sheet, but not the user sheet.