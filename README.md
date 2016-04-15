# excel-lib
XQuery library for generating excel files using MarkLogic.

### install

    mlpm install excel-lib
	mlpm deploy -u <user> -p <password> -H <host> -P <port>


## Guide

### Ingest the template

As a first step the users are required to ingest the template into MarkLogic. The library decomposes the template into ooxml files. Once the files are decomposed they are then organized into collection. Collections are way or organizing documents in MarkLogic and this allows users to isolate the template files from other documents in MarkLogic.

The library allows users to analyze the template and its various parts. For every worksheet in the excel document there is a corresponding sheet*.xml. This document will be the primary document we will be working on for generating the report. E.g. If the users ingest a template which has 5 worksheets, once ingested, we will have access to sheet1.xml thru sheet5.xml. we would work on each of these 5 documents to pull in the data for the worksheet.

#### example

Once you've deployed and installed the extension create an excel template file, a simple example can be found in the sample folder.

We can then run the following in query console:

	xquery version "1.0-ml";
	
	import module namespace excel-lib= "http://marklogic.com/solutions/htx/excel-lib" 
		at "/ext/mlpm_modules/excel-lib/excel-lib.xqy";
		
	let $REPORT_FILE := "C:/sample_template.xlsx"

	let $COLLECTIONS := ("excel_files")
		  
	return excel-lib:load-report-template-file($REPORT_FILE, $COLLECTIONS);

note the following:

-  The `$REPORT_FILE` variable that refers to the location of the template file on your system
-  The `$COLLECTIONS` variable that will be the collections assigned to the template file as well as the decomposed ooxml files ingested into marklogic

### Generate Data

This is the step where we leverage MarkLogic connectors / api to generate data for the excel worksheet which will be compiled into the excel report.
We could use either XQuery or leverage ODBC/SQL connectors to run SQL command and generate results from the SQL in a format that is required for pushing data into the excel worksheet.


#### example

We can generate data using a range of ways but the data must obey the following criteria:

- Be an xml document inside marklogic containing `<detail>` elements
- Each excel cell must correspond to a detail element ie.`<detail cell="A2">Cell_Content</detail>`
- The `cell="X#"` attribute must be a legal cell reference


To insert correctly formatted data into the database we could run the following code:

	xquery version "1.0-ml";

	xdmp:document-insert("data.xml",
	   <data>
	    <detail cell="A2">Test Name1</detail>
	    <detail cell="B2">25</detail>
	    <detail cell="A3">Test name2</detail>
	    <detail cell="B3">40</detail>
	  </data> 
	)

After ingesting the template and generating the data to be used in the spreadsheet we can then update the sheets using:

	xquery version "1.0-ml";

	import module namespace excel-lib= "http://marklogic.com/solutions/htx/excel-lib" 
		at "/ext/mlpm_modules/excel-lib/excel-lib.xqy";
	
	let $REPORT_FILE := "C:/sample_template.xlsx"
	let $COLLECTIONS := ("excel_files")
	
	let $DATA_FILE := "data.xml"
	let $SHEET_URI := "xl/worksheets/sheet1.xml"
	let $SHEET_SHORT_NAME := "sheet1.xml"
	  
	return excel-lib:process-results-for-sheet($DATA_FILE, $SHEET_URI, $REPORT_FILE, $SHEET_SHORT_NAME);

note the following:

-  The `$REPORT_FILE` variable that refers to the template file we previously ingested
-  The `$SHEET_URI` and `$SHEET_SHORT_NAME` variables that refer to the sheet we want to use within the template file
-  And `$DATA_FILE` that refers to the data formatted for use within the database

### Package and Dispatch

The library runs the code to generate data in a recursive manner for all the sheets in the workbook.
This code is pushed on to MarkLogic task server as a scheduled task to further automate the process. 
The task server run the code at specified time, generate the data file, Once the data file is generate the library picks up the data file and compiles an excel document. This document can then be sent to users via email using MarkLogic SMTP server support and the xdmp:email api call. 

#### example

After performing the previous steps we can generate the final file using:

	xquery version "1.0-ml";
	
	import module namespace excel-lib= "http://marklogic.com/solutions/htx/excel-lib" 
		at "/ext/mlpm_modules/excel-lib/excel-lib.xqy";
	
	let $REPORT_FILE := "C:/sample_template.xlsx"
	
	let $COLLECTIONS := ("excel_files", "reportResults")
	
	let $FINAL_FILE := "result.xlsx"
	
	return excel-lib:generate-report($REPORT_FILE, $FINAL_FILE, $COLLECTIONS);

note the following:

-  The `$REPORT_FILE` variable that refers to the template file we previously ingested
-  The `$COLLECTIONS` variable that will be the collections assigned to our final file
-  And `$FINAL_FILE` which will be the uri of our generated excel file which will be inserted into the database for further use.

