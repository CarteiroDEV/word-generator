const express = require('express')
const app = express();
const docx = require('docx');
const dotenv = require('dotenv');
const bodyParser = require('body-parser');
const { Document, Packer, Paragraph, Table, TableCell, TableRow, TextRun  } = docx;

dotenv.config()
app.use(bodyParser.json())

app.listen(process.env.PORT || 8000);

class NovoDocumento{
  create([perguntas]){
    
    const document = new Document();
    let arr = [];
    document.addSection({
      properties: {},
      children: [
        new Paragraph({
            text: "teste2"
        }),
        new Paragraph({
          text: String(perguntas)
      }),
        ...perguntas.map(hit =>{
          arr.push(
            this.criaPergunta(hit)
          )
          return arr;
        })
      ]
    });

    return document
  }

  criaPergunta(text) {
    console.log(String(text))
    return new Paragraph({
        text: String(text)
    });
  }

}

// POST SEM BUFFER DE BASE 64 => Retorna apenas Base64
app.post('/wordBase64', async (req, res) => {
  //req.body.testeManeiro
  const documentCreator = new NovoDocumento();
  const doc = documentCreator.create([req.body.testeManeiro]);

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