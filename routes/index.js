var express = require("express");
var PizZip = require("pizzip");
// const fspdf = require("fs");
// const { PDFDocument } = require("pdf-lib");
var router = express.Router();
const XLSX = require("xlsx");
// const { v4: uuidv4 } = require("uuid");
const path = require("path");
let lockernumber = "";
let renewalDateForHbs = "";

//for deposit data

let deposit_data = {};
let dep_flag = false;
let name
let name2
let name3
let deposit_info
/* GET home page. */
router.get("/", function (req, res) {
  res.render("index", { title: "Locker Master" });
});

router.post("/user", function (req, res) {
  console.log(req.body.lockerNo);
  pdfConvert(req.body.lockerNo, req.body.date, res);
  res.render("index", {
    title: "Locker Master",
    lockernumber: lockernumber,
    date: renewalDateForHbs,name,name2,name3,deposit_info
  });
});
router.post("/masterDataUpload", function (req, res) {
  let uploadfile = req.files.excel;

  uploadfile.mv("./public/excel/exceldata.xlsx", (err, done) => {
    if (!err) {
      console.log("success file upload");
    } else {
      console.log(err);
    }
  });
  res.render("index", { title: "Locker Master" });
});
router.post("/submit", (req, res) => {
  dep_flag=true
  // Access form data submitted by the user
  deposit_data = {
    deposit_no: req.body.deposit_no,
    dep_amount: req.body.dep_amount,
    dep_type: req.body.dep_type,
    depositer_name: req.body.depositer_name,
    depositer_name_2: req.body.depositer_name_2,
    dep_date: req.body.dep_date,
    dep_mature: req.body.dep_mature,
    dep_mature_value: req.body.dep_mature_value,
  };
  res.render("index", { title: "Locker Master" });
  console.log("deposit amount is" + deposit_data.dep_amount);
});
// router.get("/downloadPDF", function (req, res) {
//   pdfDownload(res);
// });
router.get("/downloadDOC", function (req, res) {
  docDownload(res);
});

// function pdfDownload(res) {
//   const filePath = path.join(__dirname, "../public", "output.pdf");

//   // Set the response headers for file download
//   res.setHeader("Content-Type", "application/pdf");
//   res.setHeader("Content-Disposition", "attachment; filename=output.pdf");
//   res.sendFile(filePath, (err) => {
//     if (err) {
//       console.error(err);
//       res.status(500).send("Error sending file");
//     }
//   });
// }
function docDownload(res) {
  const filePath = path.join(__dirname, "../public/word", "generated.docx");

  // Set the response headers for file download
  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  );
  res.setHeader(
    "Content-Disposition",
    "attachment; filename=" + lockernumber + ".docx"
  );
  res.sendFile(filePath, (err) => {
    if (err) {
      console.error(err);
      res.status(500).send("Error sending file");
    }
  });
}

function pdfConvert(selectedLocker, renewalDate, res) {
  // Load the Excel file
  const workbook = XLSX.readFile("public/excel/exceldata.xlsx");

  // Select the first worksheet
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];

  // Convert the worksheet to JSON
  const exceldata = XLSX.utils.sheet_to_json(worksheet);

  // Process the data as per your requirements

  //process word doc
  const fs = require("fs");
  const Docxtemplater = require("docxtemplater");

  // Load the template
  const templateFile = fs.readFileSync("public/word/docdata.docx");
  var zip = new PizZip(templateFile);
  const docxTemplate = new Docxtemplater(zip);

  var lockerIndex = 0;

  for (i = 0; i <= exceldata.length; i++) {
  
    if (exceldata[i].locker_no == selectedLocker) {
      lockerIndex=true
      name=exceldata[i].name
      name2=exceldata[i].name_2
      name3=exceldata[i].name_3
      if(exceldata[i].dep_amount==" "){
        deposit_info="Deposit not present"
      }else{
        deposit_info=exceldata[i].deposit_no
      }
      if (exceldata[i].dep_amount == " " && dep_flag==true) {
        console.log("deposit not present");
        exceldata[i].deposit_no = deposit_data.deposit_no;
        exceldata[i].dep_amount = deposit_data.dep_amount;
        exceldata[i].dep_type = deposit_data.dep_type;
        exceldata[i].depositer_name = deposit_data.depositer_name;
        exceldata[i].depositer_name_2 = deposit_data.depositer_name_2;
        exceldata[i].dep_date = deposit_data.dep_date;
        exceldata[i].dep_mature = deposit_data.dep_mature;
        exceldata[i].dep_mature_value = deposit_data.dep_mature_value;
      } 

      lockerIndex = i;
      lockernumber = exceldata[i].locker_no;
      exceldata[i].date = renewalDate;
      renewalDateForHbs = renewalDate;
      console.log("renewal date is " + exceldata[i].date);
      console.log("lockerindex" + lockerIndex);
      break;
    }
    // if(i==exceldata.length+1){
    //   res.render("apperror", { errormessage: "Enter Locker number correctly" });
    // }
  
}
  // console.log(exceldata[lockerIndex]);

  const docdata = exceldata[lockerIndex];
  console.log(docdata);

  // Set the data in the template
  docxTemplate.setData(docdata);

  // Perform the template rendering
  docxTemplate.render();

  // Get the generated document buffer
  const generatedDoc = docxTemplate.getZip().generate({ type: "nodebuffer" });

  fs.writeFileSync("public/word/generated.docx", generatedDoc);
  //eliminate "undifined"

  //elima undifiened over
  // docxConverter(
  //   "public/word/generated.docx",
  //   "public/output.pdf",
  //   (err, result) => {
  //     if (err) console.log(err);
  //     else console.log(result); // writes to file for us
  //   }
  // );
  deposit_data = {};
  dep_flag=false
}

module.exports = router;
