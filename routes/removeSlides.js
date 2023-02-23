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

//Get route to generate deck
router.post('/generateDeck', async (req,res)=>{

    console.log("Executing Remove Slides...")

    //Assigning the name of the file
    let {slidesToRemove,id} = req.body;

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
      
    //Function to remove slides from generated deck File by Id
        async function removeSlideId(id,slidesToRemove) {
            const authClient = await authorize()
            const slides = google.slides({version: 'v1', auth: authClient});

            if (!slidesToRemove){
                console.log(`No slides to remove , please try it again`);
            }

            var cantSlides = slidesToRemove.length;

            if (cantSlides>20){
                console.log("Spliting array in two parts for processing")
                var index = slidesToRemove.length;
                var firstPart = slidesToRemove.slice(0, index/2);
                var secondPart = slidesToRemove.slice(index/2,index);
                    
                console.log(`Processing first part: ${firstPart.length} to remove`)

                firstPart.forEach(async (slide) => {
                    const res = await slides.presentations.batchUpdate({
                        presentationId: id,
                        requestBody: {
                            requests: [
                            {
                                deleteObject: {
                                objectId: slide
                                }
                            }
                            ]
                        }
                        });
                })

                console.log(`First part processed: ${firstPart.length} slides removed`)

                setTimeout(()=>{
                    console.log(`Processing second part: ${secondPart.length} to remove`)
                    secondPart.forEach(async (slide) => {
                    const res = await slides.presentations.batchUpdate({
                        presentationId: id,
                        requestBody: {
                            requests: [
                            {
                                deleteObject: {
                                objectId: slide
                                }
                            }
                            ]
                        }
                        });
                    })
                    console.log(`Second part processed: ${secondPart.length} slides removed`);
                },6000)      
            }else{
                console.log('Is not neccessary split the array, can be processed in one time')
                slidesToRemove.forEach(async (slide) => {
                const res = await slides.presentations.batchUpdate({
                    presentationId: id,
                    requestBody: {
                        requests: [
                        {
                            deleteObject: {
                            objectId: slide
                            }
                        }
                        ]
                        }
                    });
                })
            }

            console.log(`Total Slides Removed : ${slidesToRemove.length}`);

            const fileLink = 'https://docs.google.com/presentation/d/' + id + '/edit'

            console.log(`Document url: ${fileLink}`)
            return fileLink 
    }

    //Executing function and send the response with the file link
    removeSlideId(id,slidesToRemove)
        .then(fileLink=>{
            console.log("Remove Slides executed successfully...")
            return res.status(200).json(fileLink);

        })
        .catch(console.error);

    }catch(err){
        console.log(`Error executing Remove Slides: ${err}`);
        return res.status(err.code).send(err.message);
    }
});

module.exports = router;