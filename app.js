const express = require("express");
const app = express();
const port = 3000;

const formidableMiddleware = require("express-formidable");

// PizZip is required because docx/pptx/xlsx files are all zipped files, and
// the PizZip library allows us to load the file in memory
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const fs = require("fs");
const path = require("path");

const toPdf = require("mso-pdf");

const { v4: uuidv4 } = require("uuid");

const myLogger = function (req, res, next) {
  console.log(
    "DEBUG#" +
      req.originalUrl +
      "#" +
      JSON.stringify(req.fields) +
      "#" +
      JSON.stringify(req.files)
  );
  next();
};

//app.use(express.json()); // for parsing application/json
//app.use(express.urlencoded({ extended: true })); // for parsing application/x-www-form-urlencoded
app.use(formidableMiddleware());

app.use(myLogger);

app.post("/composer/", (req, res) => {
  // Load the docx file as binary content
  const template = fs.readFileSync(
    path.resolve(req.files.template.path),
    "binary"
  );
  // Load json data
  const data = JSON.parse(fs.readFileSync(path.resolve(req.files.data.path)));
  // Load convert data switch
  const convertToPdf = String(req.fields.pdf).toLowerCase() == "true";

  // Unzip the content of the template file
  const zip = new PizZip(template);

  // This will parse the template, and will throw an error if the template is
  // invalid, for example, if the template is "{user" (no closing tag)
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });

  console.log("DEBUG" + JSON.stringify(data));

  // Render the document (Replace {first_name} by John, {last_name} by Doe, ...)
  doc.render(data);

  // Get the zip document and generate it as a nodebuffer
  const buf = doc.getZip().generate({
    type: "nodebuffer",
    // compression: DEFLATE adds a compression step.
    // For a 50MB output document, expect 500ms additional CPU time
    compression: "DEFLATE",
  });

  const filename = uuidv4();
  const docx = filename + ".docx";
  const pdf = filename + ".pdf";

  // buf is a nodejs Buffer, you can either write it to a
  // file or res.send it with express for example.
  fs.writeFileSync(path.resolve(__dirname, docx), buf);

  const options = {
    root: path.join(__dirname),
  };

  if (convertToPdf) {
    toPdf.convert(
      path.resolve(__dirname, docx),
      path.resolve(__dirname, pdf),
      function (errors) {
        if (errors) console.log(errors);

        res.set("Content-Type", "application/pdf");

        res.sendFile(pdf, options, function (err) {
          if (err) {
            console.error("Error sending file:", err);
          } else {
            console.log("Sent:", pdf);
          }

          fs.unlinkSync(path.resolve(__dirname, docx));
          fs.unlinkSync(path.resolve(__dirname, pdf));
        });
      }
    );
  } else {
    res.set(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );

    res.sendFile(docx, options, function (err) {
      if (err) {
        console.error("Error sending file:", err);
      } else {
        console.log("Sent:", docx);
      }

      fs.unlinkSync(path.resolve(__dirname, docx));
    });
  }
});

app.listen(port, () => {
  console.log(`Composer server listening on port ${port}`);
});
