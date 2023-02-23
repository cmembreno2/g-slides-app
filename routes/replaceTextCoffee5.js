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
router.post('/generateDeck/replaceTextCoffee5', async (req,res)=>{

    console.log("Executing Replace Text...")

    //Assigning the name of the file
    let {objectId5,id} = req.body;

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
        async function replaceText(objectId5,id) {

            const authClient = await authorize()
            const slides = google.slides({version: 'v1', auth: authClient});
            let resultados = []

              console.log(`Text need to be replace on slide: ${objectId5}`)

              var z5 = 720*12
              var z6 = 350*12
              var ad6 = 20*2.8
              var ad7 = 10*47.74
              var ae6= Math.round(ad6*12)
              var ae7= Math.round(ad7*12)
              var ae8 = parseFloat(ae6) + parseFloat(ae7)
              var ae11 = Math.round(20*5.39*12)
              var z7= ae8
              var z8 = ae11
              var z9 =  parseFloat(z5) + parseFloat(z6) + parseFloat(z7) + parseFloat(z8)
              var z12 = 120
              var z13 = 20
              var z16 = Math.round(20*2.5)
              var z17 = 20*2.5*21*12
              var z19 = Math.round((z17*9)/1000)
              var z20 = Math.round(120*((z17*9)/1000))
              var z22 = parseFloat(z5) + parseFloat(z6) + parseFloat(20*2.8*12) + parseFloat(10*47.74*12) + parseFloat(20*5.39*12) + parseFloat(120*(((20*2.5)*21*12*9)/1000))
              var z24 = z22/z17
              var ad5 = 10
              var ae6= Math.round(ad6*12)
              var ae7= Math.round(ad7*12)

              const res1 = await slides.presentations.batchUpdate({
                  presentationId: id,
                  requestBody: {
                      requests: [
                          {
                            replaceAllText: {
                            replaceText: z5.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{AA5}}"
                              },
                            pageObjectIds: [
                              objectId5
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
                            replaceText: z6.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{AA6}}"
                              },
                            pageObjectIds: [
                              objectId5
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
                            replaceText: z7.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{AA7}}"
                              },
                            pageObjectIds: [
                              objectId5
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
                            replaceText: z8.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{AA8}}"
                              },
                            pageObjectIds: [
                              objectId5
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
                            replaceText: z9.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{AA9}}"
                              },
                            pageObjectIds: [
                              objectId5
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
                            replaceText: z12.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{AA12}}"
                              },
                            pageObjectIds: [
                              objectId5
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
                            replaceText: z13.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{AA13}}"
                              },
                            pageObjectIds: [
                              objectId5
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
                            replaceText: z16.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{AA16}}"
                              },
                            pageObjectIds: [
                              objectId5
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
                            replaceText: z17.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{AA17}}"
                              },
                            pageObjectIds: [
                              objectId5
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
                            replaceText: z19.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{AA19}}"
                              },
                            pageObjectIds: [
                              objectId5
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
                            replaceText: z20.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{AA20}}"
                              },
                            pageObjectIds: [
                              objectId5
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
                            replaceText: Math.round(z22).toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{AA22}}"
                              },
                            pageObjectIds: [
                              objectId5
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
                            replaceText: (z24).toFixed(2).toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{AA24}}"
                              },
                            pageObjectIds: [
                              objectId5
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
                            replaceText: ad5.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{AE5}}"
                              },
                            pageObjectIds: [
                              objectId5
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
                            replaceText: ad6.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{AE6}}"
                              },
                            pageObjectIds: [
                              objectId5
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
                            replaceText: ad7.toFixed(1).toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{AE7}}"
                              },
                            pageObjectIds: [
                              objectId5
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
                            replaceText: ae6.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{AF6}}"
                              },
                            pageObjectIds: [
                              objectId5
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
                            replaceText: ae7.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{AF7}}"
                              },
                            pageObjectIds: [
                              objectId5
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
                            replaceText: ae8.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{AF8}}"
                              },
                            pageObjectIds: [
                              objectId5
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
                            replaceText: ae11.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{AF11}}"
                              },
                            pageObjectIds: [
                              objectId5
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
    replaceText(objectId5,id)
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