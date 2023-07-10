# psms-to-gsheet
Uses the same basic code of pypsd to actively generate and maintain a google spreadsheet

# Usage
First Fork or clone the repo to github.    
The google.oauth credentials, psd email and password as well as sheetlink are feed to the script using environmental variables. The variables have been stored in github environment secrets to maintain privacy and prevent account abuse.      

You can get your own google service account through the following steps:
1. Head to Google Developers Console and create a new project (or select the one you already have).
2. In the box labeled “Search for APIs and Services”, search for “Google Drive API” and enable it.
3. In the box labeled “Search for APIs and Services”, search for “Google Sheets API” and enable it.
4. Go to “APIs & Services > Credentials” and choose “Create credentials > Service account key”.
5. Fill out the form, you can skip the optional parts.
6. Click “Create” and “Done”.
7. Press “Manage service accounts” above Service Accounts.
8. Press on ⋮ near recently created service account and select “Manage keys” and then click on “ADD KEY > Create new key”.
9. Select JSON key type and press “Create”.

You will automatically download a JSON file with credentials. It may look like this:

```json
{
    "type": "service_account",
    "project_id": "api-project-XXX",
    "private_key_id": "2cd … ba4",
    "private_key": "-----BEGIN PRIVATE KEY-----\nNrDyLw … jINQh/9\n-----END PRIVATE KEY-----\n",
    "client_email": "473000000000-yoursisdifferent@developer.gserviceaccount.com",
    "client_id": "473 … hd.apps.googleusercontent.com",
    ...
}
```
Let's create a spreadsheet to write to now.
1. Go to google sheet create a new spreadsheet.
2. Make sure the workbook has at least one sheet in it, the program is hard coded to write to the first sheet that exists in a workbook.
3. Share the sheet with editor access to the 'client-email' you got in the json file.
4. Note down the file link and edit the end to remove any suffix. i.e.
   if the url is https://docs.google.com/spreadsheets/d/1MF.........Fgqd3DgYt_8s/edit#gid=882806501       
   change it to https://docs.google.com/spreadsheets/d/1MF.........Fgqd3DgYt_8s
5. We will need this link later to update in the Environment secrets.

Now we set up the environment secrets for GITHUB workflows as per the blueprint given here.
```
          psdemail: ${{ secrets.PSDEMAIL }}
          psdpass: ${{ secrets.PSDPASS }}
          project_id: ${{ secrets.PROJECT_ID }}
          private_key_id: ${{ secrets.PRIVATE_KEY_ID }}
          private_key: ${{ secrets.PRIVATE_KEY }}
          client_email: ${{ secrets.CLIENT_EMAIL }}
          client_id: ${{ secrets.CLIENT_ID }}
          client_x509_cert_url: ${{ secrets.CLIENT_X509_CERT_URL }}
          sheetlink: ${{ secrets.SHEETLINK }}
```
The left side are the environmental variables that are used by the pypsd-gs.py script.      
The right side are the Github Environment secrets we neet to setup.    
Follow these steps:    
1. Go to settings tab in the repo.
2. Open environments section from the left sidebar.
3. Set up a new environment named 'Autopsd'.
4. Open Configuration page for Environment Secrets by cliking on the environment name('Autopsd').
5. Add all the required enviroment secrets by copying them from json file downloaded earlier, also add your psd password and email and the link for the sheet you created until it looks something like this-
   ![secrets](./images/secrets.png)

You can now run the workflow using github actions tab manually using 'Manaul-exec' action present in the'test.yml' workflow or it will anyways run every hour using the 'Auto-exec' action present in the 'main.yml' workflow.      
If there is a dependency error you can run 'generate environment' workflow to fix it otherwise there is no need to run it.
You may have to enable write access to workflow actions to execute 'generate environment' workflow
1. Go to settings tab in the repo.
2. Open actions section from the left sidebar and select 'general'.
3. Scroll down and enable write access in workflow permissions.

