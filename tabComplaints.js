const { google } = require('googleapis');
const moment = require('moment-timezone');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

// Google Sheet ID
const spreadsheetId = "1MeBsVwqq-ZQlBO3tFK68ZUyDz1Uud1iQlrLl7vvXVb4";

// Get the current date in Asia/Kolkata timezone
const currentDate = moment().tz("Asia/Kolkata").toDate();

const currentFormattedDate = moment(currentDate).format('DD-MM-YYYY');

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
    const workbookData = {};

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

            // change array of array data to array of objects like api response
            const [headers, ...rows] = data;
            const result = rows.map(row => Object.fromEntries(headers.map((key, index) => [key, row[index]])));

            // Add data to the workbook object
            workbookData[sheetName] = result;

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
        const data = workbookData[sheetName];

        // Create a new workbook for each sheet
        const wb = xlsx.utils.book_new();
        const ws = xlsx.utils.json_to_sheet(data);
        const fileName = `${sheetName}_${currentFormattedDate}.xlsx`
        xlsx.utils.book_append_sheet(wb, ws, sheetName);

        // Define file path for each sheet
        const filePath = path.join(reportsFolderPath, fileName);

        // Write the workbook to a file
        xlsx.writeFile(wb, filePath);
        console.log(`Data for ${sheetName} has been written to ${filePath}`);
    }
};

accessGoogleSheet();
