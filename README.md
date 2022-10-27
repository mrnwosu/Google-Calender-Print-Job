# Google Itinerary Printout

This is a script that will pull events from Google Calendar and Print them out

### This script requires the following

- Creating a project with google cloud
- Enable the project to have access to Google Calendar
- Create a application account for the project
- Get certificate for application and store somewhere
- Sharing your calendar with the application account
- Creating the scheduled task on your computer/server

### Creating Google Cloud Project

1. Navigate to https://developers.google.com/workspace/guides/create-project and create a new project
![alt text](./Instruction%20Assets/Create%20Project%20Link.png)
![alt text](./Instruction%20Assets/Project%20Name.png)
2. Enable Project to have access to the Google Calendar API
![alt text](./Instruction%20Assets/API%20%26%20Services.png)
![alt text](./Instruction%20Assets/Enable%20API%20.png)
![alt text](./Instruction%20Assets/Google%20Calendar.png)
![alt text](./Instruction%20Assets//Google%20Calendar%202.png)
![alt text](./Instruction%20Assets//Google%20Calendar%203.png)
3. Create credentials that the script will use
![alt text](./Instruction%20Assets/Create%20Creds.png)
![alt text](./Instruction%20Assets/Select%20App%20Account.png)
![alt text](./Instruction%20Assets/Select%20App%20Account%202.png)
![alt text](./Instruction%20Assets/Select%20App%20Account%203.png)
![alt text](./Instruction%20Assets/Cred%20Page%201.png)
4. Create a key for this user. It will be used by the script to authenticate with google. Keep the key and the password **safe**!
![alt text](./Instruction%20Assets/key%201.png)
![alt text](./Instruction%20Assets/p12%20cert.png)
5. Share you calendar with the Service Account you created
![alt text](./Instruction%20Assets/share-calendar.png)
![alt text](./Instruction%20Assets/share-calendar%202.png)