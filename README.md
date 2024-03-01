# Higher-Ground-Dance-Company-Show-Order-
Formating the Spreadsheets
==============================================================================================
1. Create a new google sheet called "SortedDances"

Row 1 should be the names of all the dances.

Row 2 should be the style of that dance

Row 3-... should the names of all dancers in the dance (be sure to take out the names of those who dropped)

Then go through and make the show opening dance the first column, the show closing dance the last column,
and be sure the two dances bordering intermission are in the correct middle columns for the amount of dances
that semester.

There should be NO formatting in this spreadsheet (i.e. no bold names, no highlighted names)

2. Download this spreadsheet as an Excel file.

3. Download a blank excel sheet called "ShowOrder"

Running the Code
==============================================================================================
1. Download HGDC.m

2. Open MATLAB and add HGDC.m, SortedDances.xlsx, and ShowOrder.xlsx to your MATLAB Drive

3. Open HGDC.m and Run it

4. Fill out the text boxes shown in the pop up. 
	For the Excel Names include (.xlsx) at the end of the file.

	Examples:

	Enter Input Excel Name: (SortedDances.xlsx)

	Enter Output Excel Name: (ShowOrder.xlsx)

	Enter Column Number for First Song: (1)

	Enter Column Number for Last Song: (44)

	Enter Column Number for Song Before Intermission: (22)

	Enter Column Number for Song After Intermission: (23)

	Enter Desired Number of Different Ways to Sort: (10)

	You can ask for how ever many ways you want but if you try to run it again after that amount of options,
	be sure to change the input excel file or else it will take longer to run. (i.e. save one of the previous
	output options with the correct name and then use it as the input file)

	A general rule of thumb when picking the show order is to make sure the tap dances are separated (i.e. an
	equal amount in both acts and spread apart)

6. Once all the boxes are full, click OK and it will start running. The command window should display "Warning: Added specified spreadsheet." for every desired number of ways sorted you specified earlier.

7. Once the program is done running, download the ShowOrder.xlsx to view all the output options.
