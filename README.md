NLMUSD Propery Loan Agreement
=============================

Norwalk La Mirada USD Property Loan form creator that uses an Excel file 
from Follett Destiny export as source to populate forms for teachers' Property Loan Agreement.

#### Run in Ubuntu Linux
One of the libraries used will not compile/install easily on Mac,
so use Ubuntu to run this. Easy if you have Parallels on your Mac
or an Ubuntu box lying around. 

#### No Unicode Support
If there are non-ascii chars in the Excel file, the program barfs.

#### Expects one exact file name
This program expects the exact file PatronCircReportJob203634.xls
in a folder called 'Excel_files'. Edit the source to point to your Excel file.

##### Backgrounds and template Word file included
I've also included the Word and pdf templates, as well as the background that needs to be
in the root folder of the project.

##### NLMUSD Form Works just how I need it
This project works exactly as needed, but only for the exact placement of the form used by Norwalk La Mirada Unified School District. It would be easy to figure out how to adapt it to another school that uses Follett Destiny, though, because the database info and columns output from Follett would probably have the same sentinel info, "Asset", in column 9.
I have not tested this, though, since custom reports may add columns before this, thereby throwing all my column indexes off.
Just use a background for your form, and figure out all the numbers for the placement of text. 

I think reportlab uses points, which are 72 to an inch. I just iteratively stepped the text into the right positions, though. It took time.

#### TODO: Find a really good way to sort the items checked out.
As it is, there is some logic to find computers, document cameras and cameras, projectors, and iPads, and try to place them on the lines on the form on the front page. The logic is very basic, and does a simple swap if a match is found. This means that if it finds one match for computer, it puts it in slot[0] for display, but if it finds another in, say, slot[9], it simply swaps them, moving the first-found computer way down in slot[9].

#### TODO: Remove all the padding.
I thought the logic would work better as a "quick & dirty" sort if I padded the front spots with empty data, and moved found items into slot[0] through slot[5] if found. This leaves a bunch of empty-data elements in the list, and I simply pass over them if I'm outputting to the second page (the back of the sheet). 
This is partly necessitated by the poor original form, which I duplicated in Word and of which I am using a static image. 
It does mean that the output almost certainly requires a printer that can duplex print. Since it's easier to find a duplex printer than to get the district to use a better form, that's the route I chose to take.


#### TODO: Count the items for each person beforehand, so a second sheet is not required.
There are zero or more items printed on the back of each sheet (if duplexed as intended). The text "(continued on back of page)" is printed whether or not there are items to be printed on the back. This is an offshoot of all the empty-item padding used.

#### TODO: pass the Excel file in from the command line.

#### TODO: See if program really needs lxml module.
I was going to parse the html file using this module, but Follett Destiny uses opaque urls and probably some session info.
There's probably no iterative way to step through pages based on name or patron ID, 'curl' or 'wget' the needed pages, and process them. Since Follett Destiny has reports that can be output to Excel, not a deal-breaker. Just use an .xls file.

