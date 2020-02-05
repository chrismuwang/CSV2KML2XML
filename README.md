# CSV2KML2XML

## Authors: Sohaib & Christina

This program can convert csv to kml given that the csv file is formatted as: 

| CaseNumber | Name | Description | Latitude | Longitude | Group | Validation |
|------------|------|-------------|----------|-----------|-------|------------|

The Excel validation for the coordinates must be complete before using it in the program.

This program can also base64 encode kml files and insert the base64 encoding and kml information into the xml with 6 tags modified:
```xml
<LocationFile>
    <KLMdescription></KLMdescription>
    <KLMSubject></KLMSubject>
    <KLMuploadtxt></KLMuploadtxt>
    <file_name_fldLocation></file_name_fldLocation>
    <mimetypeKLMupload></mimetypeKLMupload>
    <isdocumentKLMupload></isdocumentKLMupload>
</LocationFile>
```
