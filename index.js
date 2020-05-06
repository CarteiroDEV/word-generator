const express = require('express')
const app = express();
const docx = require('docx');
const dotenv = require('dotenv');
const bodyParser = require('body-parser');

const fs = require('fs');
const readline = require('readline');
const { google } = require('googleapis');


const { Document, Packer, Paragraph, Table, TableCell, TableRow, TextRun  } = docx;

dotenv.config()
app.use(bodyParser.json())

app.listen(process.env.PORT || 3000);

class NovoDocumento{
  create([respostas, perguntas]){
    
    const document = new Document();
    // let arr = [];
    document.addSection({
      properties: {},
      children: [
        new Paragraph({
          children: [
            new TextRun({
              text: perguntas,
              bold: true,
            }),
            new TextRun({
              text: respostas,
              bold: true,
            })
          ]
        })
        // Not Working 
        // perguntas.map(hit =>{
        //   arr.push(this.criaPergunta(hit))
        //   return arr;
        // })
      ]
    });

    return document

  }

  criaPergunta(text) {
    console.log(text)
    return new Paragraph({
        text: String(text)
    });
  }

}

app.get('/googleDrive', (req, res) => {

  const SCOPES = ['https://www.googleapis.com/auth/drive'];
  const TOKEN_PATH = 'token.json';

  // Load client secrets from a local file.
  fs.readFile('credentials.json', (err, content) => {
      if (err) return console.log('Error loading client secret file:', err);
      authorize(JSON.parse(content), uploadFile);
  });

  /**
   * Create an OAuth2 client with the given credentials, and then execute the
   * given callback function.
   * @param {Object} credentials The authorization client credentials.
   * @param {function} callback The callback to call with the authorized client.
   */
  function authorize(credentials, callback) {
      const { client_secret, client_id, redirect_uris } = credentials.web;
      const oAuth2Client = new google.auth.OAuth2(
          client_id, client_secret, redirect_uris[0]);

      // Check if we have previously stored a token.
      fs.readFile(TOKEN_PATH, (err, token) => {
          if (err) return getAccessToken(oAuth2Client, callback);
          oAuth2Client.setCredentials(JSON.parse(token));
          callback(oAuth2Client);//list files and upload file
          //callback(oAuth2Client, '0B79LZPgLDaqESF9HV2V3YzYySkE');//get file

      });
  }

  /**
   * Get and store new token after prompting for user authorization, and then
   * execute the given callback with the authorized OAuth2 client.
   * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
   * @param {getEventsCallback} callback The callback for the authorized client.
   */
  function getAccessToken(oAuth2Client, callback) {
      const authUrl = oAuth2Client.generateAuthUrl({
          access_type: 'offline',
          scope: SCOPES,
      });
      console.log('Authorize this app by visiting this url:', authUrl);
      const rl = readline.createInterface({
          input: process.stdin,
          output: process.stdout,
      });
      rl.question('Enter the code from that page here: ', (code) => {
          rl.close();
          oAuth2Client.getToken(code, (err, token) => {
              if (err) return console.error('Error retrieving access token', err);
              oAuth2Client.setCredentials(token);
              // Store the token to disk for later program executions
              fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
                  if (err) return console.error(err);
                  console.log('Token stored to', TOKEN_PATH);
              });
              callback(oAuth2Client);
          });
      });
  }

  function uploadFile(auth) {
    const drive = google.drive({ version: 'v3', auth });
    var fileMetadata = {
        'name': 'tempoDeExecução.png'
    };
    var media = {
        mimeType: 'image/png',
        body: fs.createReadStream('tempoDeExecução.png')
    };
    drive.files.create({
        resource: fileMetadata,
        media: media,
        fields: 'id'
    }, function (err, res) {
        if (err) {
            // Handle error
            console.log(err);
        } else {
            console.log('File Id: ', res.data.id);
        }
    });
  }

})

// POST SEM BUFFER DE BASE 64 => Retorna apenas Base64
app.post('/wordBase64', async (req, res) => {
  //req.body.testeManeiro
  const documentCreator = new NovoDocumento();
  const doc = documentCreator.create([req.body.respostas, req.body.perguntas]);

  const b64string = await Packer.toBase64String(doc)
  res.end(b64string)

});

// POST COM BUFFER DE BASE 64 => Retorno Arquivo Pronto
app.post('/wordBuffer', async (req, res) => {

  const doc = new Document();

  doc.addSection({
    properties: {},
    children: [
        new Paragraph({
            children: [
                new TextRun(req.body.testeManeiro),
                new TextRun("Hello World"),
                new TextRun({
                    text: "Foo Barrrrrrrr",
                    bold: true,
                }),
                new TextRun({
                    text: "\tGithub is the best",
                    bold: true,
                }),
            ],
        }),
    ],
  });
          
          
  const b64string = await Packer.toBase64String(doc)

  res.setHeader('Content-Disposition', 'attachment; filename=My Document.docx');
  var doc64 = Buffer.from(b64string, "base64");
    res.writeHead(200, {
    'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'Content-Length': doc64.length
  });
  res.end(doc64);

});

// Teste GET levando URL dowload do arquivo
app.get("/wordGET", async (req, res) => {

  console.log("executing....");
  const doc = new Document();

  doc.addSection({
    properties: {},
    children: [
        new Paragraph({
            children: [
                new TextRun("Hello World"),
                new TextRun({
                    text: "Foo Barrrrrrrr",
                    bold: true,
                }),
                new TextRun({
                    text: "\tGithub is the best",
                    bold: true,
                }),
            ],
        }),
    ],
  });

  const b64string = await Packer.toBase64String(doc)
  console.log(b64string)

  res.setHeader('Content-Disposition', 'attachment; filename=My Document.docx');
  var doc64 = Buffer.from(b64string, "base64");
    res.writeHead(200, {
    'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'Content-Length': doc64.length
  });


  res.end(doc64);

  console.log('escreveu base 64');

});