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
router.post('/generateDeck/replaceTextCoffee2', async (req,res)=>{

  console.log("Executing Replace Text...")

  //Assigning the name of the file
  let {numberEmployees,priceKilo,objectId2,id} = req.body;

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
      async function replaceText(numberEmployees,priceKilo,objectId2,id) {

          const authClient = await authorize()
          const slides = google.slides({version: 'v1', auth: authClient});
          let resultados = []
          
        
              console.log(`Text need to be replace on slide: ${objectId2}`)

              var H6 = (325*((200*(numberEmployees/20))/62)).toFixed(2)
              var H7 = (650*(1.6*(numberEmployees/20))).toFixed(2)
              var H8 = parseFloat(H6) + parseFloat(H7)
              var C5=((numberEmployees/20)*1996*12).toFixed(2)
              var C6=((numberEmployees/20)*10900).toFixed(2)
              var C7=H8
              var C8=(12*numberEmployees*5.39).toFixed(2)
              var C9= parseFloat(C5)+parseFloat(C6)+parseFloat(C7)+parseFloat(C8)
              var h6 = (325*((200*(numberEmployees/20))/62))
              var h7 = (650*(1.6*(numberEmployees/20)))
              var h8 = parseFloat(h6) + parseFloat(h7)
              var c5=((numberEmployees/20)*1996*12)
              var c6=((numberEmployees/20)*10900)
              var c7=h8
              var c8=(12*numberEmployees*5.39)
              var c9= parseFloat(c5)+parseFloat(c6)+parseFloat(c7)+parseFloat(c8)
              var c20 = ((numberEmployees*2.5*21*12*10/1000)*priceKilo)
              var C22 = parseFloat(c9) + parseFloat(c20)

              const res1 = await slides.presentations.batchUpdate({
                presentationId: id,
                requestBody: {
                    requests: [
                        {
                          replaceAllText: {
                          replaceText: priceKilo.toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{C12}}"
                            },
                          pageObjectIds: [
                            objectId2
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
                          replaceText: numberEmployees.toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{C13}}"
                            },
                          pageObjectIds: [
                            objectId2
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
                          replaceText: ((numberEmployees/20)*1996*12).toFixed(2).toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{C5}}"
                            },
                          pageObjectIds: [
                            objectId2
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
                          replaceText: ((numberEmployees/20)*10900).toFixed(2).toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{C6}}"
                            },
                          pageObjectIds: [
                            objectId2
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
                          replaceText: H8.toFixed(2).toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{C7}}"
                            },
                          pageObjectIds: [
                            objectId2
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
                          replaceText: (12*numberEmployees*5.39).toFixed(2).toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{C8}}"
                            },
                          pageObjectIds: [
                            objectId2
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
                          replaceText: C9.toFixed(2).toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{C9}}"
                            },
                          pageObjectIds: [
                            objectId2
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
                          replaceText: Math.ceil(numberEmployees*2.5).toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{C16}}"
                            },
                          pageObjectIds: [
                            objectId2
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
                          replaceText: (numberEmployees*2.5*21*12).toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{C17}}"
                            },
                          pageObjectIds: [
                            objectId2
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
                          replaceText: Math.round((numberEmployees*2.5*21*12*10)/1000).toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{C19}}"
                            },
                          pageObjectIds: [
                            objectId2
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
                          replaceText: Math.round(((numberEmployees*2.5*21*12*10)/1000)*priceKilo).toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{C20}}"
                            },
                          pageObjectIds: [
                            objectId2
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
                          replaceText: C22.toFixed(2).toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{C22}}"
                            },
                          pageObjectIds: [
                            objectId2
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
                          replaceText: (C22/(numberEmployees*2.5*21*12)).toFixed(2).toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{C24}}"
                            },
                          pageObjectIds: [
                            objectId2
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
                          replaceText: Math.round((numberEmployees/20)).toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{G5}}"
                            },
                          pageObjectIds: [
                            objectId2
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
                          replaceText: (((numberEmployees/20)*200)/62).toFixed(1).toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{G6}}"
                            },
                          pageObjectIds: [
                            objectId2
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
                          replaceText: ((numberEmployees/20)*1.6).toFixed(1).toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{G7}}"
                            },
                          pageObjectIds: [
                            objectId2
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
                          replaceText: (325*((200*(numberEmployees/20))/62)).toFixed(2).toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{H6}}"
                            },
                          pageObjectIds: [
                            objectId2
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
                          replaceText: (650*(1.6*(numberEmployees/20))).toFixed(2).toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{H7}}"
                            },
                          pageObjectIds: [
                            objectId2
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
                          replaceText: H8.toFixed(2).toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{H8}}"
                            },
                          pageObjectIds: [
                            objectId2
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
                          replaceText: (12*numberEmployees*5.39).toFixed(2).toString(),
                          containsText: {
                          matchCase: true,
                          text: "{{H11}}"
                            },
                          pageObjectIds: [
                            objectId2
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
  replaceText(numberEmployees,priceKilo,objectId2,id)
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