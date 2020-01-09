# Query-Listener

=========
#### Windows form application that automates the processing of BIR FORM 2316 

### Frequently Asked Questions

#### For Query Listener v4.0

##### General Information

The Query Listener is a program designed to automate the process of the BIR form 2316. This is done by aggregating the data from disparate excel files into a single query capable Access .accdb file. 

Notice: As of 2019, this program is marked OBSOLETE, and requires an update to modern standards.

###### Features include: 

+	Searching entries by ID
+	CRUD operations (Create, Read, Update, Delete)
+	Automated importing of excel files into access
+	Injecting formula data into a given Excel file
+	Joining similar excel data with access data where ID’s match
+	Injecting Started and Ended periods where ID’s match
+	Lastly, exporting the data into the 2316 Excel form and automatically generating a PDF file of it.


If you wish to review the documentation of the program, you need only press the [HELP] button docked on the north-east section of the program.



###Frequently Asked Questions/Error clarifications:


+	“What file types does the program interact with?”

Excel: .xlsx (2007), Access: .accdb (2007). It is required that the target machine has these versions installed in order for the program to function correctly.

+	“NullpointerException?”

This can mean that the target file has missing or null data. Enter the correct data into the file that matches the database structure. This can also mean that the file you are trying to open is either invalid or missing. Try restarting the program and make sure any process of Excel is not opened through your task manager.

+	“What is the proper database format?”

[ID],[FirstName],[LastName],[GrossIncome],[LessTNT],[TaxableIncomeCE],[TaxableIncomePE],[GrossTaxableIncome],[LessTE],[LessPPH],[NetTax],[TaxDue],[HeldTaxCE],[HeldTaxPE],[TotalTax],[TIN]
Datatypes: N=Number/Double, V=Varchar
 VVVNNNNNNNNNNNNV
Tax Identification Number minimum digits = 12
[ID],[Started],[Ended]
VVV
Failure to meet the database structure requirements may result in unexpected behavior and null fields.

+	“Excel file is in use/Read-only mode?” Followed by a crash.

This happens if the excel file is already open in another program. Close the program and make sure it is not running as a process in task manager.
 

+	OleDb/ACE/Jet/Access is not registered in the local machine?

Install the two drivers mentioned in the download page. It may be possible that you only require one of these.
+	Overflow

This error is triggered when you try to enter a number that is too large for Access to hold, generally it’s limit is 64K.

+	“It would create a duplicate in the primary key, index, or unique field”

The data you are trying to import already exists in the Access database. Either remove the duplicate entry or change the ID. The ID is the unique reference. When clicking the plus button and this happens, consider the structure of your excel file. 

+	“Where can I find the Access database file or the program itself?”

The file is located in your user folder, example: C:\Users\[UserName]\AppData\Local\Apps\2.0\[WeirdLetterFormation]
The folder AppData may be hidden, so enable the preview of hidden files.
The program was designed such that there should be no need to manipulate the data of the Access database directly. Do perform all CRUD operations within the context of the program to avoid any undesired bugs or malfunctions that may not have been discovered during the testing phase.

+	“Href:(0Fx33b211)?” or any similar error series

Error is caused by incompatibility with the excel file. Ensure the data structure matches before inserting the data.

+	“The INSERT INTO statement contains the following unknown field name ‘FIELD NAME’”

This happens when you are trying to manipulate an Excel file that doesn’t match the criteria of the insert or append query. If the field mentions the field name, take note of this and adjust your fields accordingly, you may have misspelled the header somewhere or gotten the order confused.

+	“You must enter a value in the ABCD.ID field”

ABCD is the temporary access database table containing the period data: Started and Ended. This usually happens when you manipulate the fields in a way that uses more space than the required column cells (3) then deleted it. It is unknown why this issue occurs, it would state this error even though the ID field contains matching data. It has been tested that even when the table data is copied to another table, the error will persist such that even if you cleared the entire sheet and rewrite the table the correct way, the error will continue to persist. To solve this, collect the values of the table data and migrate it to a completely new work sheet.
It is highly encouraged that you test your data set through the program before preparing the complete document as the error is internal and may be impossible to fix, the only solution would be to work around it.

+	Incorrectly formatted PDF data, missing numbers, data in scientific notation, unnecessarily completed dates, ETC.

The generated PDF document is based completely on the selected Excel file, ensure that the fields of the 2316 document are tailored to your needs by changing the font sizes, number formats, date to text format. Also ensure that you are using the MODIFIED Form 2316 as the fields of the original 2316 document are invalid for injecting text into.




 
### Contact Information:
Contact August Bryan N. Florese to report bugs or errors, or to update the functionality of the program.
  GitHub repository: https://github.com/Aroueterra/Query-Listener
  Email address: Aroueterra@gmail.com

#### Required Materials
+	MODIFIED 2316 Excel File
+	Query Listener Software System
#### Required Installations
+	AccessDatabaseEngine  2007 Connectivity
+	AccessDatabaseEngine_x86 or AccessDatabaseEngine_X64
+	In the case of misaligned or enlarged font sizes, you may not have the dominant font, Century Gothic, installed
+	Any program that can be used to read PDF files may be installed


Developed by August Bryan N. Florese

