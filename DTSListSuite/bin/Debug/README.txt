#################################################################################
#
#   README
#   Written by Michael Wallace
#   Last Edit 03/18/2014
#	Current Version v3.0
#
#################################################################################

HOW TO USE
__________

To make a fresh DTS List:
	- Ensure a valid DTSLIST.txt file is in the program dir/input in the following format:
			CS ADM 6257
			CS ADM 6322
			CS LAB 9751
			etc...
	- Run the exe and choose (1) at the menu.
	- The program will aggregate data from the DTS database and compile it in to a
		spreadsheet named spreadsheetoutput.csv in dir/input.

To find DTS changes
	- In the DTS Master Sheet in Google Drive, export the spreadsheet as a .csv.
	- Rename this file MASTERDTSLIST.csv and place it in the program dir/input.
	- Run the exe and choose (2) at the menu.
	- The program will compare fields of the spreadsheet only for DTSs which have
		been updated on or before the Last Checked date.
	- A list of Rejected/Reclassed/Changed DTSs will be output in the program dir/output
		named DTSChanges.txt.
	- If you'd like a Google Drive formatted list of updated DTSs, choose (Y) when
		prompted and a spreadsheet named spreadsheetchanges.csv will be output to
		the program dir/output.

To upload a spreadsheet to Google Drive:
	- In the Google Drive DTS Master Sheet, choose File -> Import.
	- Choose the correct .csv file (output/spreadsheetoutput.csv or output/spreadsheetchanges.csv).
	- Ensure the 'Separator character' is set to 'Tab'
	- Choose 'Replace data starting at selected cell'
	- Upon pressing 'Import,' the spreadsheet will be imported and the corresponding
		cells will be replaced.

To generate URLs from the DTS#:
	- In the Google Drive DTS Master Sheet, choose Tools -> Script Manager.
	- Highlight 'makeHyperlink' and click 'Run'
	- Each cell in the DTS# column will turn in to clickable hyperlinks.

VERSION HISTORY
_______________

v3.0
-New GUI!
-Deprecated unused text
-Errors now print to console, while everything else prints to screen

v2.2
-Various columns moved around to fit new spreadsheet format
-Status now checks for 6.07 first, then 6.1. If neither, nil
-Added support for parsing Keywords
-Added support for parsing Due Dates

v2.1
-Added support for Change numbers, Shipping rings, LSS, CCD, and MU2
-Shifted columns to match new spreadsheet
-Bug fix where Submitted DTSs were getting their programs cleared

v2.0
-Refactorization of all methods
-Removed 'View Updates DTSs'
-Removed 'Completed Releases'
-Performance increase by 250% in web aggregation
-Code length decreased by 47.5%
-Methods are now less specific, and can take more polymorphic data

v1.7
-Added support for DTSs to be shipped
-Added support for parsing change numbers for 6.07
-Changed Production Rings method to be for only DTSs completed at 6.07

v1.6
-Edited loadcsv method to deal with new .csv export rules
-Logic change to match new urls in .csv export rules
-Bug fix where quoting would multiply on itself
-Bug fix where spreadsheetchanges.csv was posting multiples of the same DTS
-Bug fix where DTS numbers were showing up as URLs
-Bug fix where invalid URLs in compare method would crash out

v1.5
-Added Production Ring method!
-Added DTS Standard method!
-Added the ability to open this README from the main menu
-Added the ability to output a Rej/Rec list in Method 1
-Added comments where necessary
-Added graceful exits in niche crash scenarios
-Edited some display methods for clarity
-Bug fix where files wouldn't close until program exit
-Bug fix where 6.07 releases weren't filing
-Bug fix where PP numbers were interpreted at strings, not floats
-Bug fix where "No Changes" wouldn't be output to DTSChanges.txt in some cases
-Bug fix where larger PP numbers would sometimes error out
-Bug fix where method two would fail to return to the main menu
-Bug fix where unformatted input would result in a crash

v1.4
-Refactored code into separate files
-Various bug fixes involving refactoring
-Bug fix where errUrl was displaying a nil value
-Bug fix where YN prompt only accepted some values
-Bug fix where program icon wasn't displaying correctly
-Clean up of distributable directory

v1.3
-Added functionality for checking Rejected/Reclassed DTSs
-Added functionality to output a Blank Value .csv listing only changed DTSs
-Added/clarified comments at busy sections
-Bug fix where program would crash when missing DTSLIST.txt
-Bug fix where exe would crash while exiting
-Bug fix where unordered programs would count as changes
-Bug fix where missing MASTERDTSLIST.csv wouldn't exit gracefully
-Bug fix where blankValList was returning old information
-Resorting of some code

v1.2
-Changed messages for DTSChanges.txt
-Added separate instructions for each method
-Bug fix where comparing Assc DTSs would sometimes fail
-Bug fix where quotation marks and apostrophes would fail out parsing
-Bug fix where the compare parser was skipping the second field

v1.1
-Bug fix where quotation marks were being read as "&quot;" in compare method
-Bug fix where Assc DTSs were not checking for nil when comparing
-Resorting of some code

v1.0
-Added compare method!
-Added functionality to print changes to .txt file
-Bug fix where compare to nil was failing
-Bug fix where dates weren't formatting correctly
-Added support for importing MASTERDTSLIST.csv from Google

v0.4
-Added menu support
-Refactored code for readability
-Added support for Date Checked

v0.3
-Added support for checking intranet status
-Added support for gracefully exiting upon inability to load DTS page
-Bug fix for Associated DTS Requests filing wrong when there are multiple
-Bug fix for 5.x releases showing up in csv
-Bug fix for red Priority formatting deleting Priority number
-Bug fix where only the first comma in the desc would be deleted
-Formatting fix where .'s would show up at "xXx" in the program listing
-Changed csv delimiter from comma to \t in order to allow to Google formatting
-Formatting fix to have all DTS numbers be hyperlinks. This is achieved by a script ran
	outside of this program.

v0.2
-Separated screens from menus/process/end
-Finished parsing all fields, need to fix formatting
-Introduced ability to do a final screening on strings to remove useless punctuation
-Fixed compatibility with Google Docs
-Bug fix where missing a trailing CR in DTSLIST.txt would crash out

v0.1
-Updated welcome text
-Introduced ability to read a DTS list(DTSLIST.txt) in the format of:
	PRODUCT APPL DTS#
	PRODUCT APPL DTS#
#NOTE: DTSLIST.txt **MUST** have a trailing CR or list will error.
-Introduced ability to parse fields, see below for list of items available
-Introduced ability to write gathered lists to .csv file


PERSONAL NOTES
______________

Parse Fields
0 C/R				NN		/0/
1 Product			Y		[0]
2 Appl				Y		[1]
3 DTS#				Y		[2]
4 Desc				Y		[3]
5 Status			Y		[4]
6 Pri				Y		[5]
7 Ass'd DTS			Y		[6]
8 Rel				Y		[7]
9 Rel PP			Y		[8]
10 Programs 		Y		[9]
11 Date Checked		Y		[10]
12 Change Num		Y		[11]
13 Due Date 		Y		[12]
14 To Be Shipped	Y		[13]
15 Shipping TEST	Y		[14]
16 Shipping LIVE	Y		[15]
17 LSS				Y		[16]
18 CCD				Y		[17]
19 MU2				Y		[18]
-- Comments 		NN		RETIRED
-- In Prod			NN		RETIRED
-- Rejected			NN		RETIRED
-- Reclassed		NN		RETIRED