const express = require("express");
const bodyParser = require("body-parser");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const app = express();
const port = 5000;

app.use(bodyParser.json());
app.use(cors());

app.post("/generer-document", (req, res) => {
    // Charger le modèle de document et générer le document final
    const content = fs.readFileSync(path.resolve(__dirname, "template.docx"), "binary");
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
    });
    //reformetter la date dans req.body en xx/xx/xxxx
    req.body.date = (req.body.date).split("-").reverse().join("/");
    doc.render(req.body);

    // Générer le fichier
    const buf = doc.getZip().generate({
        type: "nodebuffer",
        compression: "DEFLATE",
    });

    // Envoyer le document en tant que réponse HTTP
    res.setHeader("Content-Disposition", "attachment; filename=attestation_test.docx");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.send(buf);
});

app.listen(port, () => {
    console.log(`Serveur en cours d'exécution sur le port ${port}`);
});
