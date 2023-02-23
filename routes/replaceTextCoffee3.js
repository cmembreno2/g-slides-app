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
router.post('/generateDeck/replaceTextCoffee3', async (req,res)=>{

    console.log("Executing Replace Text...")

    //Assigning the name of the file
    let {objectId3,id} = req.body;

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
        async function replaceText(objectId3,id) {

            const authClient = await authorize()
            const slides = google.slides({version: 'v1', auth: authClient});
            let resultados = []

                console.log(`Text need to be replace on slide: ${objectId3}`)

                var k5 = 830*12
                var k6 = 350*12
                var o6= Math.round(175*10*3)
                var o7= Math.round(225*10)
                var o8 = parseFloat(o6) + parseFloat(o7)
                var o11 = Math.round(30*5.39*12)
                var k7= o8
                var k8 = o11
                var k9 =  parseFloat(k5) + parseFloat(k6) + parseFloat(k7) + parseFloat(k8)
                var k12 = 120
                var k13 = 30
                var k16 = Math.round(k13*2.5)
                var k17 = k16*21*12
                var k19 = Math.round((k17*9)/1000)
                var k20 = Math.round(120*((k17*9)/1000))
                var k22 = parseFloat(k9) + parseFloat(k20)
                var k24 = (k22/k17).toFixed(2)
                var n5 = 10
                var n6 = n5*3
                var n7 = 1*n5

                const res1 = await slides.presentations.batchUpdate({
                    presentationId: id,
                    requestBody: {
                        requests: [
                            {
                              replaceAllText: {
                              replaceText: k5.toString(),
                              containsText: {
                              matchCase: true,
                              text: "{{L5}}"
                                },
                              pageObjectIds: [
                                objectId3
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
                            replaceText: k6.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{L6}}"
                              },
                            pageObjectIds: [
                              objectId3
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
                            replaceText: k7.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{L7}}"
                              },
                            pageObjectIds: [
                              objectId3
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
                            replaceText: k8.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{L8}}"
                              },
                            pageObjectIds: [
                              objectId3
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
                            replaceText: k9.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{L9}}"
                              },
                            pageObjectIds: [
                              objectId3
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
                            replaceText: k12.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{L12}}"
                              },
                            pageObjectIds: [
                              objectId3
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
                            replaceText: k13.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{L13}}"
                              },
                            pageObjectIds: [
                              objectId3
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
                            replaceText: k16.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{L16}}"
                              },
                            pageObjectIds: [
                              objectId3
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
                            replaceText: k17.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{L17}}"
                              },
                            pageObjectIds: [
                              objectId3
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
                            replaceText: k19.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{L19}}"
                              },
                            pageObjectIds: [
                              objectId3
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
                            replaceText: k20.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{L20}}"
                              },
                            pageObjectIds: [
                              objectId3
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
                            replaceText: k22.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{L22}}"
                              },
                            pageObjectIds: [
                              objectId3
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
                            replaceText: k24.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{L24}}"
                              },
                            pageObjectIds: [
                              objectId3
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
                            replaceText: n5.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{O5}}"
                              },
                            pageObjectIds: [
                              objectId3
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
                            replaceText: n6.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{O6}}"
                              },
                            pageObjectIds: [
                              objectId3
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
                            replaceText: n7.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{O7}}"
                              },
                            pageObjectIds: [
                              objectId3
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
                            replaceText: o6.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{P6}}"
                              },
                            pageObjectIds: [
                              objectId3
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
                            replaceText: o7.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{P7}}"
                              },
                            pageObjectIds: [
                              objectId3
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
                            replaceText: o8.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{P8}}"
                              },
                            pageObjectIds: [
                              objectId3
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
                            replaceText: o11.toString(),
                            containsText: {
                            matchCase: true,
                            text: "{{P11}}"
                              },
                            pageObjectIds: [
                              objectId3
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
    replaceText(objectId3,id)
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