
*******************************************************************************************************
 ObjectToExcel									
 Params:										
	rootObj 				= 	Query or array of queries					
	paramStruct (Optional)	= 	Structure with all needed parameters for the cfc.  Currently 
					these are the supported options: 
				 	localeDates [true/false]: Use LSisDate() or isDate()
				  	autoFilter [true/false]: Filter the headers or not 
				  	worksheetNames [list]:  List of worksheet names, if there are 						missing worksheet names a default will be used. 
					groupField[string]: Name of the field you want to group the 
						query using
					groupFieldDisplay[string]: The field to use as a display 
						on the  
				  
*******************************************************************************************************


Version 1.0
- The CFC is designed to consume a query or an array of query objects.  Pass either in and 
  the CFC will handle the rest.  
- The second parameter takes a simple list of worksheet names.  If you do not pass in this
  variable the CFC will default the worksheet names.  The same is true if your list does 
  not equal the number of queries available
- The third optional parameter is a boolean that sets the auto filter of the worksheet 
  headers.  This addition is courtesy of James McCullough.
- DateTime items can be sent in as a date only or as a date and a time.  There is a style 
  descriptor in getStyles() that will set the proper type
- Boolean values accept 'true/false' and 'yes/no'.  

Version 1.1
	Thanks to Richard Davies:
	- Binary data is now skipped.
	- Query columns are in the same order as the database query.  
	- LSisDate support
	- 24-hour date values
- There is a new structure for parameters in anticipation of the addition of the other xml 
  options in the specification document.
- There are only three new values currently, but the worksheet names have been moved into 
  the structure.    

Version 1.2
	Thanks to James McCullough for the newest addition.
	-Ability to create separate sheets based on a grouped query.

All feedback is appreciated.  I am going to be adding an optional struct parameter for 
functions and other features.   