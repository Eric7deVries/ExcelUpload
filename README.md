Practice project to experiment with CRUD funtionality and Excel.

This API has 3 endpoints
UploadExcel
DownloadExcel
DownloadSampleExcel

#UploadExcel:
Upload an Excel file to modify the objects we store.
Here an explaination how you can do CRUD functionaly on these objects.

Create: Add a new record to the Excel and upload. Duplicates will be ignored. 
Read: See DownloadExcel. 
Update: Change the record but keep the mandetory ID the same. This will update the record.
Delete: Send in the ID, but keep all fields empty will delete the record on our side.

Want to extend the object with an extra field?
Add the field and make sure to also do this with existing items.
Note: ID is mandatory! (must be type of long)
For an example layout, call DownloadSampleExcel.

#DownloadExcel:
Downloads the current objects we store in Excel format.

#DownloadSampleExcel:
Downloads sample objects in excel.
These objects are based on a Person object with the following attributes:
ID
Age
Name
DateOfBirth
