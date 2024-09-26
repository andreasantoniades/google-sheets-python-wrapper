# google_sheets.py

The script's intention is to provide a user with a shortcut to start working with Google sheets using the command line.
It can also be imported and used in other python programs
It however relies on the user performing a few actions to set up a Google Cloud development project of their own.
The reason is to avoid security complexities that may arise from sharing projects, secrets and keys.

The below instructions assume the user is not members of an Google Workspace so they don't have the option to create an internal project.
This makes the instructions apply to everyone with a Google account but Google Workspace users have more options.

To use this script you must perform the below steps:

- Clone the repository locally and optionally make the script available in your PATH 
- Install its dependencies by running
   - `pip3 install -r requirements.txt`
- (you can potentially install the few dependencies manually and any recent versions are unlikely to not work)
- Create a developer profile and a Project in the Google Cloud Console https://console.cloud.google.com/
- Enable the Google Sheets API for the Project
- Proceed with one (or both) of the following authentication setups

By using OAuth you allow the api to interact with Google Sheets as if it were you.
This means the script, for better or worse, automatically gets access to the spreadsheets you have access to.
To use OAuth:

- Set up the OAUth consent screen for the Project
- Add yourself (and any other of your accounts you want to use) as a test user for the project
- Create OAuth client ID credentials
- Download the client using the json format and save it locally, e.g. in
   - `client_secret.json`
- On your local system (needs a web browser) initiate the OAuth flow by running
   - `google_sheets.py get_oauth_token client_secret.json`
- After you go through the consent screen, the script will output your token in json format. Copy this into a file e.g.
   - `/some/convenient/path/oauth_token_file`
- DO NOT SHARE your token (as it can be used to access google spreadsheets as if it were you)
- DO NOT SHARE your client secret (as it can be used to display a consent screen to a user tricking them it's your project)
- set the env variable for the OAuth token (example path/filename, it can be whatever you want)
   - `export GOOGLE_SHEETS_TOKEN=/some/convenient/path/oauth_token_file`

By using a service account you have more fine-tuned control on which spreadsheets the generated token can access.
However you need to share each individual spreadsheet with the service account's email address.
To use a service account:

- Create a new service account for the project
- Create a key for the account and download it e.g as
  - `/some/convenient/path/service_account_token_file`
- set the env variable for the service acount token (example path/filename, it can be whatever you want)
   - `export GOOGLE_SHEETS_TOKEN=/some/convenient/path/service_account_token_file`
 
