const xlsx = require('xlsx');

// Load the Excel file
const workbook = xlsx.readFile('C:/Users/owais/Downloads/Test.xlsx');

// Select the first sheet
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Convert the sheet to JSON with the first row as keys
const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

// Transform the data so that the first row becomes keys and the rest are arrays of values
const keys = data[0]; // First row as keys
const result = {};

keys.forEach((key, index) => {
    // Map each column to an array and filter out undefined values
    result[key] = data.slice(1).map(row => row[index]).filter(value => value !== undefined);
  });
  
  // Output the transformed data
  console.log(result);

