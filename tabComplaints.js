const { google } = require('googleapis');
const moment = require('moment-timezone');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const nodemailer = require("nodemailer");

let accountList = [
    {
        "user": "chirag.p@techworks.co.in",
        "pass": "Byzzy1-jucton-gogkeq"
    },
    {
        "user": "Bharti.singh@techworks.co.in",
        "pass": "gymruc-saKpu8-purnoc"
    },
    {
        "user": "hitesh.kumar@techworks.co.in",
        "pass": "4VqvS&RY*ZFnqaU1"
    },
]

const transporter = nodemailer.createTransport({
    host: "smtp.dreamhost.com",
    port: 465,
    secure: true,
    auth: {
        user: accountList[0].user,
        pass: accountList[0].pass
    }
});

// Google Sheet ID
const spreadsheetId = "1MeBsVwqq-ZQlBO3tFK68ZUyDz1Uud1iQlrLl7vvXVb4";
// Get the current date in Asia/Kolkata timezone
const currentDate = moment().tz("Asia/Kolkata").toDate();
const currentFormattedDate = moment(currentDate).format('DD-MM-YYYY');
let Batch1Arr = ['NSAH', 'WNAG', 'SKAR', 'WPUN', 'WMUM', 'NCHA', 'ECAL', 'NJPR'];
let Batch2Arr = ['WBHO', 'NDEL', 'WAHM', 'EPAT', 'NLUC', 'SHYD', 'SBLR'];
let workbookData = {};
let branchListOfArr = [];

function delay(milliseconds) {
    return new Promise(resolve => {
        setTimeout(resolve, milliseconds);
    });
}

const accessGoogleSheet = async () => {
    // Initialize the authentication client
    const auth = new google.auth.GoogleAuth({
        keyFile: "./credentials.json",
        scopes: ["https://www.googleapis.com/auth/spreadsheets"],
    });

    // Get the authenticated client
    const authClientObject = await auth.getClient();

    // Create the Sheets instance
    const sheets = google.sheets({ version: 'v4', auth: authClientObject });

    getAllWorkbookNames(sheets);
};

const getAllWorkbookNames = async (sheets) => {
    // Get workbook names present in the spreadsheet
    const response = await sheets.spreadsheets.get({
        spreadsheetId: spreadsheetId,
    });

    let sheetNames = response.data.sheets.map(sheet => sheet.properties.title);
    let forDeletion = ['Summary', 'Format'];

    sheetNames = sheetNames.filter(item => !forDeletion.includes(item))
    console.log('Sheet Names:', sheetNames);

    getWorkbookWiseData(sheets, sheetNames);
};

const getWorkbookWiseData = async (sheets, sheetNames) => {
    for (let i = 0; i < sheetNames.length; i++) {
        const sheetName = sheetNames[i];
        try {
            // Fetch data for each sheet
            const response = await sheets.spreadsheets.values.get({
                spreadsheetId: spreadsheetId,
                range: sheetName,
            });

            const data = response.data.values || [];
            console.log(`Data for ${sheetName}:`, data.length - 1);

            // Change array of array data to array of objects like API response
            const [headers, ...rows] = data;
            const result = rows.map(row => Object.fromEntries(headers.map((key, index) => [key, row[index]])));

            if (sheetName == "POC_LIST") {
                result.filter(obj => branchListOfArr.push(obj['Branch']));
            }

            // Filter objects that have the key 'Xtravu Id'
            const filteredResult = result.filter(obj => sheetName != "POC_LIST" ? obj['Branch'] !== undefined && obj['Branch']?.length === 4 && branchListOfArr.includes(obj['Branch']) : obj);

            // Add filtered data to the workbook object
            workbookData[sheetName] = filteredResult;

        } catch (error) {
            console.error(`Error fetching data for ${sheetName}:`, error);
        }
    }

    // Convert data to an .xlsx file
    createSeparateExcelFiles(workbookData);
};

const createSeparateExcelFiles = (workbookData) => {
    // Define the folder path for saving the .xlsx files
    const reportsFolderPath = path.join(__dirname, 'reports');

    // Check if 'reports' folder exists, if not create it
    if (!fs.existsSync(reportsFolderPath)) {
        fs.mkdirSync(reportsFolderPath);
    }

    // Iterate over each sheet's data and create a separate .xlsx file
    for (const sheetName in workbookData) {
        // ignore webook POC-LIST
        if (sheetName !== "POC_LIST") {
            const data = workbookData[sheetName];

            // Create a new workbook for each sheet
            const wb = xlsx.utils.book_new();
            const ws = xlsx.utils.json_to_sheet(data);
            let extractBranchName = sheetName?.split('-')[0];
            let fileName = `${extractBranchName}_${currentFormattedDate}.xlsx`;
            xlsx.utils.book_append_sheet(wb, ws, sheetName);

            // Define file path for each sheet
            const filePath = path.join(reportsFolderPath, fileName);

            // Write the workbook to a file
            xlsx.writeFile(wb, filePath);
            console.log(`Data for ${sheetName} has been written to ${filePath}`);

            // adding file name to POC_LIST

            const targetObject = workbookData['POC_LIST'].find(item => item.Branch == extractBranchName);
            if (targetObject) {
                targetObject['fileName'] = fileName;
            }
        }
    }

    sendMail();
};



const sendMail = async () => {
    let mailsForCC = workbookData['POC_LIST'] ? workbookData['POC_LIST'][0]['Emails (For CC Section)'] : null;
    workbookData['POC_LIST'] = workbookData['POC_LIST'].filter(item => item.Branch !== 'CC');

    if (workbookData['POC_LIST']) {
        console.log(`\n`);

        for (let idx = 0; idx < workbookData['POC_LIST'].length; idx++) {
            const i = workbookData['POC_LIST'][idx];
            let branchName = i?.Branch;
            let emailCCSectionFromBaseSheet = i['Emails (For CC Section)']
            let ccEmails = emailCCSectionFromBaseSheet ? `${emailCCSectionFromBaseSheet}, ${mailsForCC}` : mailsForCC;
            let toEmails = i['Emails (For To Section)'];

            // console.log(`TO MAILS :- ${toEmails} CC MAILS :- ${ccEmails}`);
            // console.log(mailsForCC);


            if (Batch1Arr.includes(branchName)) {
                console.log(`================Processing Branch: (${branchName})================`);

                // Ensure that each reportDelivery finishes before moving to the next
                await reportDelivery(i, toEmails, ccEmails);
                console.log(`============Finished processing Branch: (${branchName})============`);
                console.log(`\n`);
                await delay(5000);
            }
        }
    }
}

async function reportDelivery(i, toEmails, ccEmails) {
    let fromEmail = 'reports@techworks.co.in';

    try {
        // send mail with defined transport object
        const info = await transporter.sendMail({
            from: fromEmail,
            // to: 'hitesh.kumar@techworks.co.in',
            to: toEmails,
            cc: ccEmails,
            subject: `ITC TAB COMPLAINT TRACKER ${i?.Branch} ${currentFormattedDate}`, // Subject line
            html: `<h6>Please find the attachment.</h6>
            <p>&nbsp;</p>
            <table style="width: 420px; font-size: 10pt; font-family: Verdana, sans-serif; background: transparent !important;"
                    border="0" cellspacing="0" cellpadding="0">
                    <tbody>
                            <tr>
                                    <td style="width: 160px; font-size: 10pt; font-family: Verdana, sans-serif; vertical-align: top;"
                                            valign="top">
                                            <p style="margin-bottom: 18px; padding: 0px;"><span
                                                            style="font-size: 12pt; font-family: Verdana, sans-serif; color: #183884; font-weight: bold;"><a
                                                                    href="http://contact.techworksworld.com/" target="_"><img
                                                                            style="width: 120px; height: auto; border: 0;"
                                                                            src="https://i.imgur.com/Eie8F53.png" width="120"
                                                                            border="0" /></a></p>
                                            <p
                                                    style="margin-bottom: 0px; padding: 0px; font-family: Verdana, sans-serif; font-size: 9pt; line-height: 12pt;">
                                                    <a style="color: #e25422; text-decoration: none; font-weight: bold;"
                                                            href="http://www.techworksworld.com" target="_"><span
                                                                    style="text-decoration: none; font-size: 9pt; line-height: 12pt; color: #e25422; font-family: Verdana, sans-serif; font-weight: bold;">www.techworksworld.com</span></a>
                                            </p>
                                    </td>
                                    <td style="width: 30px; min-width: 30px; border-right: 1px solid #e25422;">&nbsp;</td>
                                    <td style="width: 30px; min-width: 30px;">&nbsp;</td>
                                    <td style="width: 200px; font-size: 10pt; color: #444444; font-family: Verdana, sans-serif; vertical-align: top;"
                                            valign="top">
                                            <p
                                                    style="font-family: Verdana, sans-serif; padding: 0px; font-size: 9pt; line-height: 14pt; margin-bottom: 14px;">
                                                    <span
                                                            style="font-family: Verdana, sans-serif; font-size: 9pt; line-height: 14pt;"><span
                                                                    style="font-size: 9pt; line-height: 13pt; color: #262626;"><strong>E:
                                                                    </strong></span><a
                                                                    style="font-size: 9pt; color: #262626; text-decoration: none;"
                                                                    href="mailto:support@techworks.co.in"><span
                                                                            style="text-decoration: none; font-size: 9pt; line-height: 14pt; color: #262626; font-family: Verdana, sans-serif;">support@techworks.co.in</span></a><span><br /></span></span><span><span
                                                                    style="font-size: 9pt; color: #262626;"><strong>T:</strong></span><span
                                                                    style="font-size: 9pt; color: #262626;"> (+91) 11 35007205</span><span><br /></span></span><span><span
                                                                    style="font-size: 9pt; color: #262626;"><strong>A:</strong></span><span
                                                                    style="font-size: 9pt; color: #262626;"> O-7,Second Floor Lajpat
                                                                    Nagar-II,<span>,</span></span><span style="color: #262626;">New
                                                                    Delhi-110024,India</span></span></p>
                                            <p style="margin-bottom: 0px; padding: 0px;"><span><a
                                                                    href="https://www.facebook.com/TechworksSolutionsPvtLtd/"
                                                                    rel="noopener"><img
                                                                            style="border: 0; height: 22px; width: 22px;"
                                                                            src="https://www.mail-signatures.com/signature-generator/img/templates/inclusive/fb.png"
                                                                            width="22" border="0" /></a>&nbsp;</span><span><a
                                                                    href="https://www.linkedin.com/company/ds-techworks-solutions-pvt-ltd/"
                                                                    rel="noopener"><img
                                                                            style="border: 0; height: 22px; width: 22px;"
                                                                            src="https://www.mail-signatures.com/signature-generator/img/templates/inclusive/ln.png"
                                                                            width="22" border="0" /></a>&nbsp;</span><span><a
                                                                    href="https://twitter.com/techworks14" rel="noopener"><img
                                                                            style="border: 0; height: 22px; width: 22px;"
                                                                            src="https://www.mail-signatures.com/signature-generator/img/templates/inclusive/tt.png"
                                                                            width="22" border="0" /></a>&nbsp;</span><span><a
                                                                    href="https://www.youtube.com/@TechworksDigitalSolutions"
                                                                    rel="noopener"><img
                                                                            style="border: 0; height: 22px; width: 22px;"
                                                                            src="https://www.mail-signatures.com/signature-generator/img/templates/inclusive/yt.png"
                                                                            width="22" border="0" /></a>&nbsp;</span><span><a
                                                                    href="https://www.instagram.com/techworks140/"
                                                                    rel="noopener"><img
                                                                            style="border: 0; height: 22px; width: 22px;"
                                                                            src="https://www.mail-signatures.com/signature-generator/img/templates/inclusive/it.png"
                                                                            width="22" border="0" /></a></span></p>
                                    </td>
                            </tr>
                            <tr style="width: 420px;">
                                    <td style="padding-top: 14px;" colspan="4"><a href="https://techworksworld.com/" target="_"><img
                                                            style="width: 420px; height: auto; border: 0;"
                                                            src="https://i.imgur.com/QoPxSPy.png" width="420" border="0" /></a></td>
                            </tr>
                            <tr>
                                    <td style="padding-top: 14px; text-align: justify;" colspan="4">
                                            <table style="width: 420px; font-size: 10pt; font-family: Verdana, sans-serif; background: transparent !important;"
                                                    border="0" cellspacing="0" cellpadding="0">
                                                    <tbody>
                                                            <tr>
                                                                    <td
                                                                            style="font-size: 8pt; color: #b2b2b2; line-height: 9pt; text-align: justify;">
                                                                            The content of this email is confidential and intended
                                                                            for the recipient specified in message only. It is
                                                                            strictly forbidden to share any part of this message
                                                                            with any third party,without a written consent of the
                                                                            sender. If you received this message by mistake,please
                                                                            reply to this message and follow with its deletion,so
                                                                            that we can ensure such a mistake does not occur in the
                                                                            future.</td>
                                                            </tr>
                                                    </tbody>
                                            </table>
                                    </td>
                            </tr>
                    </tbody>
            </table>`,
            attachments: [
                {   // file on disk as an attachment
                    filename: i?.fileName,
                    path: `./reports/${i?.fileName}` // stream this file
                }
            ]
        });

        console.log(`Mail Send Succesfull.. (${i?.Branch})`);
        console.log(`Mail Send to ${toEmails}`);
        console.log(`Mail Send cc ${ccEmails}`);

    } catch (error) {
        console.log(error);
        console.log(i);
    }
}

accessGoogleSheet();
