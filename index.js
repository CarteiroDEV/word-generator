const express = require('express')
const app = express();
const docx = require('docx');
const dotenv = require('dotenv');

const bodyParser = require('body-parser');

//const fs = require('fs');
const { Document, Packer, Paragraph, Table, TableCell, TableRow,TextRun  } = docx;
//import * as docx from "docx";
//import * as fs from "fs";
dotenv.config()

app.use(bodyParser.json())

app.get('/', (req, res) => {
  res.send('Hello World!')
});
app.listen(process.env.PORT || 8000);

app.post('/wordTEST', async (req, res) => {
  // res.end(console.log(req.body.parceiro));

  const doc = new Document();

  doc.addSection({
    properties: {},
    children: [
        new Paragraph({
            children: [
                new TextRun(req.body.testeManeiro)
            ],
        }),
    ],
});

  const b64string = await Packer.toBase64String(doc)

  res.setHeader('Content-Disposition', 'attachment; filename=My Document.docx');
  var doc64 = Buffer.from(b64string, "base64");
    res.writeHead(200, {
    // 'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'Content-Type': 'application/msword',
    'Content-Length': doc64.length
  });
  res.end(doc64);

});

app.get("/word", async (req, res) => {

  console.log("executing....");
  const doc = new Document();

  doc.addSection({
    properties: {},
    children: [
        new Paragraph({
            children: [
                new TextRun("Hello World"),
                new TextRun({
                    text: "Foo Bar",
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