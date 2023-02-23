//Importing packages and dependencies
const express = require('express');
require('dotenv').config();
const fs = require('fs').promises;
const path = require('path');
const process = require('process');
const {authenticate} = require('@google-cloud/local-auth');
const {google} = require('googleapis');
const SCOPES = ['https://www.googleapis.com/auth/drive','https://www.googleapis.com/auth/presentations'];
const TOKEN_PATH = path.join(process.cwd(), 'token.json');
const CREDENTIALS_PATH = path.join(process.cwd(), 'credentials.json');
const router = express.Router();

//Get route to replace text
router.post('/generateDeck/replaceTextSolvemate', async (req,res)=>{

    console.log("Executing Replace Text...")

    //Assigning the name of the file
    let {companyName,objectId,id} = req.body;
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
      
    //Function to replace the text
        async function replaceText(objectId,companyName,id) {
            const authClient = await authorize()
            const slides = google.slides({version: 'v1', auth: authClient});
            const res = await slides.presentations.batchUpdate({
                presentationId: id,
                requestBody: {
                    requests: [
                        {
                          replaceAllText: {
                          replaceText: companyName,
                          containsText: {
                          matchCase: true,
                          text: "INSERT COMPANY NAME"
                            },
                          pageObjectIds: [
                            objectId
                            ]
                          }
                        }
                      ]
                }
            });
            const result = res.status
            if (result.length === 0) {
                return;
            }  
            return result          
        }
    //Executing function and send the response with the response code
    replaceText(objectId,companyName,id)
        .then(result=>{
            console.log("Replace Text executed successfully...")
            return res.status(200).json({success:result});

        })
        .catch(console.error);

    }catch(err){
        console.log(`Error executing Replace Text: ${err}`);
        return res.status(err.code).send(err.message);
    }
});

module.exports = router;
