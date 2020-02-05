# **Instructions**

<br>

# **Table of Contents**
1. [Summary](#Summary) .......................................................................................................................................................................... 2
2. [Instructions](#Instructions) ...................................................................................................................................................................... 2
3. [Instructions - Inserting KMLs into XML](#Instructions-Inserting) ................................................................................................................ 2
4. [Formatting Guidelines](#Formatting) ................................................................................................................................................. 2
5. [Visual Representation of How to Use Application](#Visual-Rep) ........................................................................................... 3

 <br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>
 <br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>

## **Summary:**<a name="Summary"></a>

The CSV2KML2XML application allows for the generation of many KML files from a single CSV. 
These KMLs can then be inserted into an XML file to meet the formatting regulations of the federal government.
___
## **Instructions:**<a name="Instructions"></a>
1. Make sure the CSV follows the formatting guidelines as outlined below [(see Formatting Guidelines)](#csv-format)
2. Click on the CSV to KML button and select a CSV file from any available directory. This file will be used to generate the KML files. 
3. The generated KML files will be saved to the 'KML Files' folder that will be generated if one does not already exist.
4. Once the application has finished generating the KML files, the user will be able to view all the generated files. **KML files with data validation errors will not be generated.**
___
## **Instructions - Inserting KMLs into XML:**<a name="Instructions-Inserting"></a>
1. This program will take all of the KML files found in the 'KML Files' folder and insert them into the XML file.
2. To insert the KML files into the XML file, click on the KML to XML button, then select the XML file to modify.
3. Once the application has completed its task, a popup will display to inform the user that the XML file has been modified.
___
## **Formatting Guidelines (Technical):**<a name="Formatting"></a>

1. The spreadsheet file must be a CSV (Comma Separated Values), XLS and XLSX will not work.
2. The CSV must be formatted as follows: <a name="csv-format"></a>

    | CaseNumber | Name | Description | Latitude | Longitude | Group | Validation |
    |------------|------|-------------|----------|-----------|-------|------------|

3. Formula for CSV coordinate validation:<br>**=IF(AND(41.416723<$D2,$D2<56.85012, -95.15699<$E2, $E2<-71.30798),"Pass","Fail")**
4. The CSV **MUST** be sorted by Case Number.
5. Please ensure the XML being used follows the formatting of the federal government.
   
   <span style="color:red;font-weight: bold;">IMPORTANT</span>: Make sure that the XML hierarchy is as follows:
   ```xml
	 <ProjectSubmissions>
		<ProjectSubmission>
			<General_PS_Generals>
				<submissionPTProjectIDCreate> Case Number </submissionPTProjectIDCreate>
			</General_PS_Generals>
			...
			<LocationFile></LocationFile>
    ```
<br>

___
## **Visual Representation of How to Use Application:** <a name="Visual-Rep"></a>

<br>


<div style="text-align:center"><img src="C:/Users/MohiuddinSo/projects/CSV2KML2XML/screenshots/start.png" alt="Start of the Application"/></div><br>
This is what shows when the application is opened <br>

What happens when you click each button:
1. CSV to KML button: <br>
<div style="text-align:center"><img src="C:/Users/MohiuddinSo/projects/CSV2KML2XML/screenshots/start-selectcsvtokml.png" alt="CSV to KML button clicked"/></div><br>
when you click the CSV to KML button, a popup will appear where you can select the **CSV (Comma Seperated Values)** file:<br>
<div style="text-align:center"><img src="C:/Users/MohiuddinSo/projects/CSV2KML2XML/screenshots/opencsv.png" alt="Select CSV File"/></div><br>
You can only select CSV files as XLS and XLSX files are not supported.<br><br>
Once the CSV file is selected, click open and this popup will show up when all KML files are successfully generated:<br>
<div style="text-align:center"><img src="C:/Users/MohiuddinSo/projects/CSV2KML2XML/screenshots/kmlgeneratepopup.png" alt="KML Successfully Generated Popup"/></div><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>
<br>
The application list will appear like this:<br><br>
<div style="text-align:center"><img src="C:/Users/MohiuddinSo/projects/CSV2KML2XML/screenshots/kmlgeneratelistview.png" alt="KML File Names in List"/></div><br>
The KML files that have successfully been generated will show in the list with the name of the file. The KML files that failed validation will show in the list as "Data validation failed: '...last 4 values of file name'".
You can double click the file name in the list and it will open the kml file in any supported application (ie. Notepad, Notepad++, etc..).<br><br><br><br><br>

2. KML to XML button: <br><br>
<div style="text-align:center"><img src="C:/Users/MohiuddinSo/projects/CSV2KML2XML/screenshots/start-selectkmltoxml.png" alt="KML to XML button clicked"/></div><br><br><br><br><br><br><br><br><br><br><br><br><br><br>
when you click the KML to XML button, a popup will appear where you can select the **XML** file:<br>
<div style="text-align:center"><img src="C:/Users/MohiuddinSo/projects/CSV2KML2XML/screenshots/openxml.png" alt="Select CSV File"/></div><br>
Once the XML file is selected, click open and this popup will show up when the KML files have successfully been added to the XML and will also show exceptions (if any) of which KML files were not added to the XML:<br>
<div style="text-align:center"><img src="C:/Users/MohiuddinSo/projects/CSV2KML2XML/screenshots/xmlgeneratepopup.png" alt="KML Successfully Added to XML Popup"/></div><br><br>

3. More Info button: <br>
The More Info button will open this pdf.