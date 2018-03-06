# POI-Demo-with-JAVA-using-JOGET-workflow-engine

POI is Apache's JAVA API for Microsoft Documents. 
In this webservice, I am doing:
1. Get request parameter for HTTPREQUEST i.e. reference no from calling URL.
2. Get data for corresponding reference no from Database
3. Save this data in treemap so that we have can have key value structure.
4. Create a MS Excel File
5. Create an excel sheet in this file
6. Create rows and then for each row create cells.
7. Fill cells with the data from treemap.
8. Format cells by setting borders, background colors, background pattern and fixed width.
9. Save this excel file in Byte Output Stream.
10. Send this stream in HTTPSERVLET response.

Output PDF of generated Excel file:
![alt text](https://drive.google.com/open?id=1P6CBr5fw2SWzK97MutbJThGAHjt_cV0j)
