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
const storage = multer.memoryStorage();  // Store the file in memory instead of disk

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

        const workbook = xlsx.read(file.buffer);  // Read the file buffer instead of a file path
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

        const keys = data[0];  // The first row as keys (section names)
        const result = {};

        keys.forEach((key, index) => {
            result[key] = data.slice(1, -2).map(row => row[index]).filter(value => value !== undefined);
        });

        // Extract max and weightage values from the last two rows
        const maxRow = data[data.length - 2];
        const weightageRow = data[data.length - 1];

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

                calculatedRow.push(Math.round(calculatedValue));  // Rounding the calculated value to nearest integer
                sum += calculatedValue;
            });

            calculatedRow.push(Math.round(sum));  // Adding the SUM value to the row
            calculatedRows.push(calculatedRow);
        });

        // Create a single Excel workbook with two sheets in memory
        const workbookOutput = xlsx.utils.book_new();
        const firstWorksheet = xlsx.utils.aoa_to_sheet(rows);
        const secondWorksheet = xlsx.utils.aoa_to_sheet(calculatedRows);

        xlsx.utils.book_append_sheet(workbookOutput, firstWorksheet, 'Randomized Values');
        xlsx.utils.book_append_sheet(workbookOutput, secondWorksheet, 'Calculated Values');

        // Write the workbook to a buffer
        const excelBuffer = xlsx.write(workbookOutput, { type: 'buffer', bookType: 'xlsx' });

        // Set the response headers to indicate a file download
        res.setHeader('Content-Disposition', `attachment; filename="${textField}.xlsx"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

        // Send the buffer as the response
        res.send(excelBuffer);

    } catch (err) {
        console.log('this is the error', err);
        res.status(500).send({ error: 'An error occurred while processing the file' });
    }
});




app.listen(port, () => {
    console.log(`Server is listening on ${port}`);
});