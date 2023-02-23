//Importing packages and dependencies
const express = require('express');
require('dotenv').config();
const fs = require('fs').promises;
const path = require('path');
const process = require('process');
const {authenticate} = require('@google-cloud/local-auth');
const {google} = require('googleapis');
const SCOPES = ['https://www.googleapis.com/auth/drive'];
const TOKEN_PATH = path.join(process.cwd(), 'token.json');
const CREDENTIALS_PATH = path.join(process.cwd(), 'credentials.json');
const router = express.Router();

//Get route to generate deck
router.get('/generateDeck/:name/presentation/:id', async (req,res)=>{

    console.log("Executing Generate Deck...")

    //Assigning the name of the file
    var name = req.params.name;
    var id = req.params.id;

    try{
    //Google Auth flow  
        async function loadSavedCredentialsIfExist() {
            try {
              const content = await fs.readFile(TOKEN_PATH);
              const credentials = JSON.parse(content);
              return google.auth.fromJSON(credentials);
            } catch (err) {
              return null;
            }
        }
          
        async function saveCredentials(client) {
            const content = await fs.readFile(CREDENTIALS_PATH);
            const keys = JSON.parse(content);
            const key = keys.installed || keys.web;
            const payload = JSON.stringify({
              type: 'authorized_user',
              client_id: key.client_id,
              client_secret: key.client_secret,
              refresh_token: client.credentials.refresh_token,
            });
            await fs.writeFile(TOKEN_PATH, payload);
        }
          
        async function authorize() {
            let client = await loadSavedCredentialsIfExist();
            if (client) {
              return client;
            }
            client = await authenticate({
              scopes: SCOPES,
              keyfilePath: CREDENTIALS_PATH,
            });
            if (client.credentials) {
              await saveCredentials(client);
            }
            return client;
        }
      
        //Function to copy the Master Deck File #1 by Id
        async function copyFileId(name) {
          const authClient = await authorize()
          const drive = google.drive({version: 'v3', auth: authClient});
          const res = await drive.files.copy({
            fileId: id,
            requestBody: {
              name: `DeckAutomation - ${name}`
            }
          });
          const fileId = res.data.id
          if (!fileId) {
            console.log('Error with file generation');
            return;
          }

          const file = await drive.files.get({
            fileId: fileId,
            fields: 'parents',
          });

          // Move the file to the new folder
          const previousParents = file.data.parents
              .map(function(parent) {
                return parent.id;
              })
              .join(',');
          const files = await drive.files.update({
            fileId: fileId,
            addParents: '1FJwplaGQ4SIhRuFbP1s0ubHSRuiZJZHv',
            removeParents: previousParents,
            fields: 'id, parents',
          });
          console.log(`New deck generated with id: ${fileId}`)
          return fileId
        }
        
    //Executing function and send the response with the file link
        copyFileId(name)
        .then(fileLink=>{
            console.log("Generate Deck executed successfully...")
            return res.status(200).json(fileLink);

        })
        .catch(console.error);

    }catch(err){
        console.log(`Error executing Generate Deck: ${err}`);
        return res.status(err.code).send(err.message);
    }
});

module.exports = router;