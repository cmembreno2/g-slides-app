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
router.post('/generateDeck/replaceTextCoffee4', async (req,res)=>{

    console.log("Executing Replace Text...")

    //Assigning the name of the file
    let {objectId4,id} = req.body;

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
        async function replaceText(objectId4,id) {

            const authClient = await authorize()
            const slides = google.slides({version: 'v1', auth: authClient});
            let resultados = []

              console.log(`Text need to be replace on slide: ${objectId4}`)

              var t5 = 380*12
              var t6 = 250*12
              var w6 = 40*2.8
              var w7 = 10*47.74
              var x6= Math.round(w6*12)
              var x7= Math.round(w7*12)   
              var x8 = parseFloat(x6) + parseFloat(x7)
              var x11 = Math.round(40*9.87*12)
              var t7= x8
              var t8 = x11
              var t9 =  parseFloat(t5) + parseFloat(t6) + parseFloat(t7) + parseFloat(t8)
              var t12 = 120
              var t13 = 40
              var t16 = Math.round(t13*2.5)
              var t17 = t16*21*12
              var t19 = Math.round((t17*9)/1000)
              var t20 = Math.round(120*((t17*9)/1000))
              var t22 = parseFloat(t9) + parseFloat(t20)
              var t24 = (t22/t17).toFixed(2)
              var w5 = 10

              const res1 = await slides.presentations.batchUpdate({
                  presentationId: id,
                  requestBody: {
                      requests: [
                          {
                            replaceAllText: {
                            replaceText: t5.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{T5}}"
                              },
                            pageObjectIds: [
                              objectId4
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
                          replaceText: t6.toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{T6}}"
                            },
                          pageObjectIds: [
                            objectId4
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
                          replaceText: t7.toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{T7}}"
                            },
                          pageObjectIds: [
                            objectId4
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
                          replaceText: t8.toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{T8}}"
                            },
                          pageObjectIds: [
                            objectId4
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
                          replaceText: Math.round(t9).toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{T9}}"
                            },
                          pageObjectIds: [
                            objectId4
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
                          replaceText: t12.toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{T12}}"
                            },
                          pageObjectIds: [
                            objectId4
                            ]
                          }
                        }
                      ]
                }
              })

              const res7 = await slides.presentations.batchUpdate({
                presentationId: id,
                requestBody: {
                    requests: [
                        {
                          replaceAllText: {
                          replaceText: t13.toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{T13}}"
                            },
                          pageObjectIds: [
                            objectId4
                            ]
                          }
                        }
                      ]
                }
              })

              const res8 = await slides.presentations.batchUpdate({
                presentationId: id,
                requestBody: {
                    requests: [
                        {
                          replaceAllText: {
                          replaceText: t16.toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{T16}}"
                            },
                          pageObjectIds: [
                            objectId4
                            ]
                          }
                        }
                      ]
                }
              })

              const res9 = await slides.presentations.batchUpdate({
                presentationId: id,
                requestBody: {
                    requests: [
                        {
                          replaceAllText: {
                          replaceText: t17.toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{T17}}"
                            },
                          pageObjectIds: [
                            objectId4
                            ]
                          }
                        }
                      ]
                }
              })

              const res10 = await slides.presentations.batchUpdate({
                presentationId: id,
                requestBody: {
                    requests: [
                        {
                          replaceAllText: {
                          replaceText: t19.toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{T19}}"
                            },
                          pageObjectIds: [
                            objectId4
                            ]
                          }
                        }
                      ]
                }
              })

              const res11 = await slides.presentations.batchUpdate({
                presentationId: id,
                requestBody: {
                    requests: [
                        {
                          replaceAllText: {
                          replaceText: t20.toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{T20}}"
                            },
                          pageObjectIds: [
                            objectId4
                            ]
                          }
                        }
                      ]
                }
              })

              const res12 = await slides.presentations.batchUpdate({
                presentationId: id,
                requestBody: {
                    requests: [
                        {
                          replaceAllText: {
                          replaceText: t22.toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{T22}}"
                            },
                          pageObjectIds: [
                            objectId4
                            ]
                          }
                        }
                      ]
                }
              })

              const res13 = await slides.presentations.batchUpdate({
                presentationId: id,
                requestBody: {
                    requests: [
                        {
                          replaceAllText: {
                          replaceText: t24.toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{T24}}"
                            },
                          pageObjectIds: [
                            objectId4
                            ]
                          }
                        }
                      ]
                }
              })

              const res14 = await slides.presentations.batchUpdate({
                presentationId: id,
                requestBody: {
                    requests: [
                        {
                          replaceAllText: {
                          replaceText: w5.toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{W5}}"
                            },
                          pageObjectIds: [
                            objectId4
                            ]
                          }
                        }
                      ]
                }
              })

              const res15 = await slides.presentations.batchUpdate({
                presentationId: id,
                requestBody: {
                    requests: [
                        {
                          replaceAllText: {
                          replaceText: w6.toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{W6}}"
                            },
                          pageObjectIds: [
                            objectId4
                            ]
                          }
                        }
                      ]
                }
              })

              const res16 = await slides.presentations.batchUpdate({
                presentationId: id,
                requestBody: {
                    requests: [
                        {
                          replaceAllText: {
                          replaceText: w7.toFixed(1).toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{W7}}"
                            },
                          pageObjectIds: [
                            objectId4
                            ]
                          }
                        }
                      ]
                }
              })

              const res17 = await slides.presentations.batchUpdate({
                presentationId: id,
                requestBody: {
                    requests: [
                        {
                          replaceAllText: {
                          replaceText: x6.toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{X6}}"
                            },
                          pageObjectIds: [
                            objectId4
                            ]
                          }
                        }
                      ]
                }
              })

              const res18 = await slides.presentations.batchUpdate({
                presentationId: id,
                requestBody: {
                    requests: [
                        {
                          replaceAllText: {
                          replaceText: x7.toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{X7}}"
                            },
                          pageObjectIds: [
                            objectId4
                            ]
                          }
                        }
                      ]
                }
              })

              const res19 = await slides.presentations.batchUpdate({
                presentationId: id,
                requestBody: {
                    requests: [
                        {
                          replaceAllText: {
                          replaceText: x8.toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{X8}}"
                            },
                          pageObjectIds: [
                            objectId4
                            ]
                          }
                        }
                      ]
                }
              })

              const res20 = await slides.presentations.batchUpdate({
                presentationId: id,
                requestBody: {
                    requests: [
                        {
                          replaceAllText: {
                          replaceText: x11.toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{X11}}"
                            },
                          pageObjectIds: [
                            objectId4
                            ]
                          }
                        }
                      ]
                }
              })

              resultados.push(res1.status,res2.status,res3.status,res4.status,res5.status,res6.status,res7.status,res8.status,res9.status,res10.status,res11.status,res12.status,res13.status,res14.status,res15.status,res16.status,res17.status,res18.status,res19.status,res20.status)
          
          if (resultados.length === 0) {
              return;
          }  
          return resultados         
        }
    //Executing function and send the response with the response code
    replaceText(objectId4,id)
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