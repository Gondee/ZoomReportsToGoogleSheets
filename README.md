# ZoomReportsToGoogleSheets
Easy to generate Zoom Dashboard.

Automatically push zoom attendence logs to a Google sheet. of your choosing. Good for role call. 

This app combines a couple things. 

1. It uses Google Sheets as a dashboard for your attendence logs. This allows you to filter
2. It uses mongoDB cloud for the DB, free edition is fine. 
3. A paid zoom account is needed to access the parts of the API required. 


When you compile this you will notice that I had to add some things to the zoomus python api. Most of these
additions are in the pipeline in merge requests, so i will not detail here. 