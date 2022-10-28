# Google Itinerary Printout

This is a script that will pull events from Google Calendar and Print them out to the computers **default** printer.

## This script requires the following

- Creating a project with google cloud
- Enable the project to have access to Google Calendar
- Create a application account for the project
- Get certificate for application and store somewhere
- Sharing your calendar with the application account
- Creating the scheduled task on your computer/server

## Creating Google Cloud Project

1. Navigate to https://developers.google.com/workspace/guides/create-project and create a new project
![New Google Cloud Project](./Instruction%20Assets/Create%20Project%20Link.png)
![Set Project Name](./Instruction%20Assets/Project%20Name.png)
2. Enable Project to have access to the Google Calendar API
![Admin Side bar](./Instruction%20Assets/API%20%26%20Services.png)
![API Services](./Instruction%20Assets/Enable%20API%20.png)
![API Library](./Instruction%20Assets/Google%20Calendar.png)
![Calendar App](./Instruction%20Assets//Google%20Calendar%202.png)
![Hit Enable](./Instruction%20Assets//Google%20Calendar%203.png)
3. Create credentials that the script will use
![Calendar Enabled](./Instruction%20Assets/Create%20Creds.png)
![Credentials Sidebar](./Instruction%20Assets/Select%20App%20Account.png)
![New Service Account](./Instruction%20Assets/Select%20App%20Account%202.png)
![Assign role](./Instruction%20Assets/Select%20App%20Account%203.png)
![](./Instruction%20Assets/Cred%20Page%201.png)
4. Create a key for this user. It will be used by the script to authenticate with google. Keep the key and the password **safe**!
![Keys Tab](./Instruction%20Assets/key%201.png)
![Download Key](./Instruction%20Assets/p12%20cert.png)
5. Share you calendar with the Service Account you created
![Share Calendar](./Instruction%20Assets/share-calendar.png)
![Shared with service account](./Instruction%20Assets/share-calendar%202.png)

## Create A Scheduled Task with the script

1. Download the directory and unzip in desired directory.
![Download zip from Github and Extract](./Instruction%20Assets/Download%20Zip.png)
![File Explorer](./Instruction%20Assets/Extract%20to%20dr.png)
2. Update the "run.ps1" file with your correct arguments 
    - -certPath > the path to the certificate you downloaded from google
    - -certPath > the password that was given to you for your certificate to work. For security it's better to store this value in an environment variable like "GoogleCertPasswood". To read about environment variables go here. [How to change environment variables on Windows 10](https://medium.com/thedevproject/how-to-change-environment-variables-on-windows-10-56a9c8b26b5b)
![Update the run.ps1 file](./Instruction%20Assets/Update%20run.ps1.png)
3. Create a new task in Task Scheduler for windows
![Open Task Scheduler and create new task](./Instruction%20Assets/Open%20Task%20Scheduler%20and%20Create%20Task.png)
![Set name for task ](./Instruction%20Assets/Create%20task%201.png)
![Create a trigger](./Instruction%20Assets/Create%20Task%202.png)
![Set the trigger to run at desired time and settings](./Instruction%20Assets/Create%20Task%203.png)
![Go to action and create new](./Instruction%20Assets/Create%20Task%204.png)
 
- "arguments" = -file \<pathToThe run.ps1 script\> and the start in
- "start in" = the path to the DIRECTORY the run.ps1 is in
![Add powershell and arguments ](./Instruction%20Assets/Create%20Task%205.png)
![Set conditions](./Instruction%20Assets/Create%20Task%206.png)
![Save by clicking ok](./Instruction%20Assets/Create%20Task%207.png)
4. The task is created and you can hit "Run task" to test the script at anytime.
![Task created](./Instruction%20Assets/Task%20Created.png)