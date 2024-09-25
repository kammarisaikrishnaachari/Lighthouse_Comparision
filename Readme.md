Lighthouse is an open-source, automated tool for improving the performance, quality, and correctness of your web apps.



User must provide required information/ file path on the Config.properties file to compare two lighthouse json files 
1. log file path[the path which user wants stores the log file]
2. Results file path [the path which user wants to store the results file]
Ex - 
logFilePath=D:\New folder\application.log
excelFilePath=D:\New folder\Results.xlsx

3.Previous report path - user should provide the valid previous json file path
4. Current report path - user should provide the valid current json file path

ex - 
previousReportPath= C:/Users/AcariK/Downloads/linkdev.kantar.com-20240522T191856.json
currentReportPath= C:/Users/AcariK/Downloads/beta.kantarmarketplace.com-20240920T105718.json

Once the user double click the executable file the execution will start, it will fecth the Application.log path and results file path and previous path and current file path.

code will executed and then it gives output as Results excel and Application log file.

Appliaction log will contain all step wise logs like what measures/audits are comparing and validating with info tags, if user unable to get results errors will logged into appliaction.log file

Results excel - this excel will contain two sheets 

1. Metrics sheet will have the below metrics 
 - Performance 
 - accessibility
 - Best Practices 
 - SEO

2. Metrics sub-metrics will have the Performance sub-metrics 
 - first-contentful-paint
 - largest-contentful-paint
 - Speed Index
 - Total blocking time
 - cumulative-layout-shift