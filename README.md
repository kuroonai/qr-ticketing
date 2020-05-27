# qr-tickecting

Generates QR codes as tickets from a guestlist and related information and delivers them to the guests email. This tool requires configuring json file with API in gsheets to provide access to the spreadsheet, which is a well known procedure. Here are the brief steps,

'''
steps to get the json file from your google spreadsheet:
    
    1. install gspread (pip install gspread)
    2. For setting up credentials,  go to Google developers console
    https://console.developers.google.com/cloud-resource-manager
    3. Create or select an existing project
    4. select the project and go to navigation menu on the top left corner
    5. On that menu, select API & services and then credentials
    6. Select Create credentials and choose Service account key
    7. In the Create service account key page, select new service account and
        give a new name to service accound and in role dropdown list select project -> owner
    8. Now download the json file and then put it in the currect working folder
    9. In addition to these steps you might need to enable google drive and gsheets to this project
        by going to, API & services and add them via "Enable API and services button"
    10. Also disable secure apps only option for the email you are about to use.
    11. Finally give the email id under "client email" in the json file, access 
        to the spread sheet you are working on.
'''
