

const express = require('express');
const multer = require('multer');
const path = require('path');
const xlsx = require('xlsx');
const fs = require('fs');


const app = express();
const port = 3002;

app.use(
    express.json({
        limit: "200mb",
    })
);

app.use(
    express.urlencoded({
        extended: true,
        limit: "200mb",
    })
);

app.set("views", path.join(__dirname, "src/views"));
app.use(express.static(path.join(__dirname, "public")));
app.set("view engine", "ejs");

app.get("/", (req, res) => {
    res.render('index.ejs', {
        status: 200,
        message: 'Successful render'
    })
});

const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, 'uploads/');
    },
    filename: function (req, file, cb) {
        cb(null, file.fieldname + '-' + Date.now() + path.extname(file.originalname));
    }
});

const upload = multer({ storage: storage });

function getRandomValue(arr) {
    return arr[Math.floor(Math.random() * arr.length)];
}

app.post('/upload-and-generate', upload.single('fileField'), (req, res) => {

    try {


        const { candidates, textField } = req.body;
        const file = req.file;

        if (!file) {
            return res.status(400).send({ error: 'No file uploaded' });
        }

        const workbook = xlsx.readFile(file.path);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
        // console.log('data::::', data);

        const keys = data[0];  // The first row as keys (section names)
        const result = {};

        keys.forEach((key, index) => {
            result[key] = data.slice(1, -2).map(row => row[index]).filter(value => value !== undefined);
        });

        // console.log('result:::::', result);

        // Extract max and weightage values from the last two rows
        const maxRow = data[data.length - 2];
        const weightageRow = data[data.length - 1];

        // console.log('max row::::::', maxRow);
        // console.log('weightage::::::::', weightageRow);

        // Validate max and weightage values
        keys.forEach((key, index) => {
            if (isNaN(maxRow[index]) || maxRow[index] <= 0) {
                return res.status(400).send({ error: `Invalid max value for ${key}` });
            }
            if (isNaN(weightageRow[index]) || weightageRow[index] <= 0) {
                return res.status(400).send({ error: `Invalid weightage value for ${key}` });
            }
        });

        // Prepare headers: first column for CANDIDATE, then section columns, and a SUM column
        const headers = ['CANDIDATE', ...keys, 'SUM'];

        // Initialize the array to hold rows for the first Excel sheet
        const rows = [headers];

        // console.log('rows::::::::::', rows);

        // Array to store random values for the second Excel sheet calculations
        const randomValues = [];

        // Generate rows for each candidate with random values
        for (let i = 1; i <= candidates; i++) {
            const row = [`Candidate${i}`];
            const candidateRandomValues = {};
            let sum = 0;

            keys.forEach((section, index) => {
                const randomValue = getRandomValue(result[section]);
                row.push(randomValue);
                candidateRandomValues[section] = randomValue;
                sum += randomValue;
            });

            row.push(Math.round(sum));  // Adding the SUM value to the row
            randomValues.push(candidateRandomValues);
            rows.push(row);
        }

        // Prepare rows for the second Excel sheet with calculated values
        const calculatedRows = [headers];

        randomValues.forEach((candidateValues, candidateIndex) => {
            const calculatedRow = [`Candidate${candidateIndex + 1}`];
            let sum = 0;

            keys.forEach((section, index) => {
                const randomValue = candidateValues[section];
                const max = maxRow[index];
                const weightage = weightageRow[index];
                const calculatedValue = (randomValue * weightage) / max;

                calculatedRow.push(Math.round(calculatedValue));  // Formatting the calculated value to 2 decimal places
                sum += calculatedValue;
            });

            calculatedRow.push(Math.round(sum));  // Adding the SUM value to the row
            calculatedRows.push(calculatedRow);
        });

        // Create a single Excel workbook with two sheets
        const workbookOutput = xlsx.utils.book_new();
        const firstWorksheet = xlsx.utils.aoa_to_sheet(rows);
        const secondWorksheet = xlsx.utils.aoa_to_sheet(calculatedRows);

        xlsx.utils.book_append_sheet(workbookOutput, firstWorksheet, 'Randomized Values');
        xlsx.utils.book_append_sheet(workbookOutput, secondWorksheet, 'Calculated Values');

        const filePath = path.join(__dirname, `generated_values_${Date.now()}.xlsx`);
        xlsx.writeFile(workbookOutput, filePath);

        // Send the Excel file back to the client
        res.download(filePath, `${textField}.xlsx`, (err) => {
            if (err) {
                console.error('Error sending the Excel file:', err);
            }

            // Clean up the Excel file after sending
            fs.unlink(filePath, (err) => {
                if (err) {
                    console.error('Error deleting the Excel file:', err);
                }
            });

            // Optionally delete the original uploaded file as well
            fs.unlink(file.path, (err) => {
                if (err) {
                    console.error('Error deleting the uploaded file:', err);
                }
            });
        });

    } catch (err) {
        console.log('this is the error', err)
    }
});




app.listen(port, () => {
    console.log(`Server is listening on ${port}`);
});