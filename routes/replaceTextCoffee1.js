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
router.post('/generateDeck/replaceTextCoffee1', async (req,res)=>{

    console.log("Executing Replace Text...")

    //Assigning the name of the file
    let {numberEmployees,objectId1,id} = req.body;

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
        async function replaceText(numberEmployees,objectId1,id) {

            const authClient = await authorize()
            const slides = google.slides({version: 'v1', auth: authClient});
            let resultados = []

                console.log(`Text need to be replace on slide: ${objectId1}`)

                const res1 = await slides.presentations.batchUpdate({
                presentationId: id,
                requestBody: {
                    requests: [
                        {
                          replaceAllText: {
                          replaceText: numberEmployees.toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{C4}}"
                            },
                          pageObjectIds: [
                            objectId1
                            ]
                          }
                        }
                      ]
                }
                })

                const res2 = await slides.presentations.batchUpdate({
                    presentationId: id,
                    requestBody: {
                        requests: [
                            {
                              replaceAllText: {
                              replaceText: Math.ceil(numberEmployees*12*2.8*20.8333).toString(),
                              containsText: {
                              matchCase: true,
                              text: "{{C6}}"
                                },
                              pageObjectIds: [
                                objectId1
                                ]
                              }
                            }
                          ]
                    }
                })

                const res3 = await slides.presentations.batchUpdate({
                    presentationId: id,
                    requestBody: {
                        requests: [
                            {
                              replaceAllText: {
                              replaceText: (12*0.03*2.8*(250/12)*numberEmployees).toString(),
                              containsText: {
                              matchCase: true,
                              text: "{{C11}}"
                                },
                              pageObjectIds: [
                                objectId1
                                ]
                              }
                            }
                          ]
                    }
                })

                const res4 = await slides.presentations.batchUpdate({
                    presentationId: id,
                    requestBody: {
                        requests: [
                            {
                              replaceAllText: {
                              replaceText: Math.ceil(4*12*numberEmployees*((2.8*20.8333*10)/1000)).toString(),
                              containsText: {
                              matchCase: true,
                              text: "{{C17}}"
                                },
                              pageObjectIds: [
                                objectId1
                                ]
                              }
                            }
                          ]
                    }
                })

                const res5 = await slides.presentations.batchUpdate({
                    presentationId: id,
                    requestBody: {
                        requests: [
                            {
                              replaceAllText: {
                              replaceText: Math.round(12*0.05*numberEmployees*((10*(250/12)*2.8))/1000).toString(),
                              containsText: {
                              matchCase: true,
                              text: "{{C20}}"
                                },
                              pageObjectIds: [
                                objectId1
                                ]
                              }
                            }
                          ]
                    }
                })

                const res6 = await slides.presentations.batchUpdate({
                    presentationId: id,
                    requestBody: {
                        requests: [
                            {
                              replaceAllText: {
                              replaceText: Math.round(((0.03*2.8*20.8333*numberEmployees*12)/138)*2).toString(),
                              containsText: {
                              matchCase: true,
                              text: "{{C27}}"
                                },
                              pageObjectIds: [
                                objectId1
                                ]
                              }
                            }
                          ]
                    }
                })
                
                console.log(`Text replaced on slide: ${objectId1}`)
                resultados.push(res1.status,res2.status,res3.status,res4.status,res5.status,res6.status)
          
          if (resultados.length === 0) {
              return;
          }  
          return resultados         
        }
    //Executing function and send the response with the response code
    replaceText(numberEmployees,objectId1,id)
        .then(results=>{
            console.log("Replace Text executed successfully...")
            return res.status(200).json({success:`Total fields updated : ${results.length}`});
        })
        .catch(console.error);

    }catch(err){
        console.log(`Error executing Replace Text: ${err}`);
        return res.status(err.code).send(err.message);
    }
});

module.exports = router;