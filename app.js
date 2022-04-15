const mailer = require('nodemailer');
const readXlsxFile = require('read-excel-file/node');
const fs = require('fs');
const docx2html = require('docx2html');
const dotenv = require('dotenv');
const reader = require('xlsx');
dotenv.config();

const ATTACHMENT_BASE_DIR = './attachment/';
const DOCS_BASE_DIR = "./doc/"
const INPUT_FILE = "./input.xlsx";

const EMAIL = process.env.EMAIL;
const PASSWORD = process.env.PASSWORD;
const EMAIL_SUBJECT = process.env.EMAIL_SUBJECT;
const LIMIT = 2;
const args = process.argv; 

// Requiring the module
  
// Reading our test file
const file = reader.readFile(INPUT_FILE);
let data = []
  
const sheets = file.SheetNames
  
for(let i = 0; i < sheets.length; i++)
{
   const temp = reader.utils.sheet_to_json(
        file.Sheets[file.SheetNames[i]])
        temp.forEach((res) => {
        data.push(res)
   })
} 

const  transporter = mailer.createTransport({
  service: 'gmail',
  port:587,
  auth: {
    user: EMAIL,
    pass: PASSWORD
  }
});

let files = fs.readdirSync(ATTACHMENT_BASE_DIR);
let attachments = files.map(file=> {
    return {
        path: ATTACHMENT_BASE_DIR + file,
    }
}) 

let docs = fs.readdirSync(DOCS_BASE_DIR); 
let input_document =   docs.length > 0 ? DOCS_BASE_DIR + docs[0] : null ;
console.log("Reading File: input.xlsx");
console.log("Input Document: ", input_document);
console.log("Input Attachments: ", files.join(",   "));

if(args.length > 2) {
    var startInd = parseInt(args[2]);
    if(args.length > 3) {
        var lastInd = parseInt(args[3]);
    }
    
     async function sendEmail(html){ 
        if(lastInd == NaN || lastInd == null || lastInd == undefined) {
            lastInd = data.length-1;
        } 
        for(let i = startInd; i <= lastInd; i++) { 
                await new Promise(next=> {
                    let  mailOptions = {
                        from: EMAIL,
                        to: data[i]["Email"],
                        subject: EMAIL_SUBJECT ,
                        html: html,
                        attachments: attachments
                        };
                        transporter.sendMail(mailOptions, function(error, info){
                            if (error) {
                                let error = new Date().toLocaleString() + "   FAILED: "+ "Sno. "+ data[i]['S.No.']+"  "+data[i]['Email']+ "   "+ data[i]['Company Name(s)']+ "  "+ data[i]['CIN'];
                                writeInFile(error);
                                console.log(error);
                                next();
                            } else {
                                let message = new Date().toLocaleString() + "   SUCCESS: "+ "Sno. "+ data[i]['S.No.']+"  "+ data[i]['Email']+ "   "+ data[i]['Company Name(s)']+ "  "+ data[i]['CIN'];
                                writeInFile(message);
                                console.log(message); 
                                next();
                            }
                        });
                }); 
            } 
    }
    if(input_document == null ) {
        console.log("Input Document Not PRESENT!!!");
        process.exit();
    }
    docx2html(input_document).then(function(html){ 
        sendEmail(html.toString());
    })
    
   
} else {
    console.log("Please provide the starting row serial number");
}
function writeInFile(text) { 
    fs.writeFile('log.txt', text+"\n", { flag: 'a+' }, err => {
        if (err) {
            console.error(err)
            return
        } 
    })
}

