# Higher-Ground-Dance-Company-Show-Order-
Excel File Format
==============================================================================================
To make sure that the code runs correctly the excel file needs to follow an exact format:
	1. Each Column is a single dance
	2. Row 1 is the title for the dance
	3. Row 2 is the type of dance i.e. (Hip Hop, Jazz, ...)
	4. Row 3 to ... is the dancers in the dance

Also to make sure that it is read correctly make sure that the file name of the excel does not contain any spaces.
Format of the excel should follow the file "ShowOrderSping22.xlsx"




Key Notes when running the executable
==============================================================================================
1. Run the exe file "HGDC_Dance_Sorter.exe"

2. Fill out the text boxes shown in the pop up. 
	For the Excel Names include (.xlsx) at the end of the file.
	
	Examples:
	Enter Input Excel Name: (ShowOrder.xlsx)
	Enter Output Excel Name: (SorderDances.xlsx)
	Enter Column Number for First Song: (1)
	Enter Column Number for Last Song: (39)
	Enter Column Number for Song Before Intermission: (9)
	Enter Column Number for Song After Intermission: (5)
	Enter Desired Number of Different Ways to Sort: (7)

3. Once everything is fill correctly run the program in which it will generate the output file.
   The program goes through many iterations of dances so it will take a while to create different dance orders.
   Do not open the generated excel file until a pop up saying the operation is complete.
