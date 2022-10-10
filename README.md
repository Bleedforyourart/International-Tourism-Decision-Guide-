International Tourism Decision Guide 
Data Preparation and Analysis Using VBA and Python   

Overview:   
Our decision support system consulting company, SIE Consultants Corporation, is building a decision guide for 
companies in the International Tourism industry. We will use information collected by the World Tourism 
Organization (UNWTO), a specialized agency of the United Nations. In addition, we will provide a searchable 
database of UN Countries by Regions with Gross Domestic Product (GDP) data. Our system will provide 
analysis tools for these companies that will help them discover trends and niches in international tourism that 
they should explore.   

Requirements:   
There are two parts to this system, an offline data processor that loads and processes a data file and places 
descriptive data into an Excel Workbook, and a viewer and reporting application that uses this data as well as MS 
Access database data.  

The Python program will begin by displaying the Tkinter file dialog to allow the user to select the file to be 
processed as in the image below. Notice the dialog box has the title ‘Please select the International Tourist 
Arrivals Worldwide file’ and the file types to show in the dialog has been set to ‘Excel Files (*.xlsx) so that only 
Excel files with the .xlsx file extension are shown. 

The Python preparation program will read in this data and process it by processing the data, copying description 
data for the Data Source to the Overview worksheet, copying tourism data to the Data worksheet, and 
calculating and writing the descriptive data to the Statistics worksheet. For each region, read and calculate the 
Mean, Minimum, and Maximum values. 
Create a new workbook named international_tourism_data.xlsx, and create 3 new worksheets in it: Overview, 
Data, and Statistics. 

The Overview tab will contain the processed Data Source description information. The data will be written into 
relevant fields with static labels in A1:A9 as shown above. Column widths must be set at 30 for Column A and 
100 for Column B. All row heights must be set at 20. 
 
The Data tab will contain the processed data from the incoming Data tab. The data will be copied into the range 
A1:F17 as shown above. Column widths must be set at 30 and row heights must be set at 20. 
 
The Statistics tab will contain the processed data from the incoming Data tab, calculating the Mean, Min, and 
Max for each region. The data will be copied into the range B2:F4, with static headers in the Columns (B1:BF) 
and Rows (A1:A4) in Bold as shown above. Column widths must be set at 30 and row heights must be set at 20. 

There will be only one worksheet in the VBA App. Name it the SIE Decision Guide Application. Insert a button on 
that worksheet and set the text to the same thing. This is how your VBA application will be launched. There will 
be no data written in this workbook, only the program. All data that the application uses will come from your 
processed Excel file (international_tourism_data.xlsx) or from the MS Access database (Regions.accdb). 
Insert a module into the VBA program and call it MainModule. Inside that module create a subroutine named 
Sub Main(). Begin the VBA program with a button assigned to the Sub Main(). 

Main Menu User Form 
Sub Main() will call the Main Menu user form that displays four buttons: Chart Builder, Query Builder, About Our 
Datasources, and Cancel. Cancel will simply exit the application. 
 
By default, when the user form is initialized, the Data Source is set to Annual Report. Chart Types are relevant to 
the Data Source as follows: 
Data Source: Annual Report has the Chart Types (as shown above): 
-Stacked Area (xlAreaStacked) 
-Stacked Column (xlColumnStacked) 
-Stacked Bar (xlBarStacked) 
Data Source: Statistical Data has the Chart Types (as shown above): 
-Clustered Column (xlColumnClustered) 
-Stacked Column (xlColumnStacked) 
-Clustered Bar (xlBarClustered) 

The Create Chart button will create a chart based on the user’s selections. The chart will be saved as an image 
file and displayed on the Chart Display user form (examples shown below) using that image. These charts will be 
built using the data from the processed Excel file (international_tourism_data.xlsx). Annual Report data will 
come from the Data worksheet and Statistical Data will come from the Statistics worksheet.  
The Cancel button will unload this user form and return control to the Main Menu. 

Query Builder User Form 
The Query Builder Button on the Main Menu will display the Query Builder user form. The Query Builder user 
form has a listbox that lists all countries from the Country table. The contents of the Country Listbox, when 
selected, drive the Results Listbox (shown below). The contents of the Results Listbox, when selected, drive the 
Country Detail Textboxes to the right (as shown below). All data on this page is read-only, with no updates by 
the end user. The Cancel button will unload this user form and return control to the Main Menu. 

Country ListBox: Fill the contents of this listbox based on a query to the Country table in the Regions 
database. 
Results Listbox: Fill the contents of this listbox based on a query to the Country, Region, SubRegion, and 
UN_GDP_ByCountry tables that match the current selection in the Country Listbox. 
Country Detail Textboxes: Fill the contents of these textboxes based currently selected year and GDP 
data in the Results Listbox. 

All of the data for this user form will come from the Regions.accdb MS Access file (shown below). We will not be 
altering this file or the tables, just using them to read the data. Our file contains four tables: Country, Region, 
SubRegion, and UN_GDP_ByCountry. These tables contain data downloaded from the UN. 

The Region table has a relationship with the Country table using the RegionCode field (RegionCode is the 
Primary Key in Region, and a Foreign Key in Country) Likewise, the SubRegion table has a relationship with the 
Country table using the SubregionCode field (SubregionCode is the Primary Key in SubRegion, and aForeign Key 
in Country). Finally, the relationship for the Country and UN_GDP_ByCountry table is built on the two non-key 
fields Country.CountryName and UN_GDP_ByCountry.Country. 

 

 
