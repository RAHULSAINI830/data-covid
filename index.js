const express = require('express');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');
const path = require('path');

const app = express();
app.use(bodyParser.json());

// Function to read the Excel file
const readExcelFile = () => {
  try {
    const filePath = path.join(__dirname, '../data.xlsx'); // Path to file in root directory
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];  // Get the first sheet name
    const sheet = workbook.Sheets[sheetName];

    if (!sheet) {
      throw new Error(`Sheet not found in file: ${filePath}`);
    }

    // Convert the sheet data to JSON format
    const data = xlsx.utils.sheet_to_json(sheet);
    return data;
  } catch (error) {
    console.error('Error in reading Excel file:', error.message);
    throw error;
  }
};

// Get all COVID-19 data
app.get('/api/data', (req, res) => {
  try {
    const data = readExcelFile();
    res.json(data);
  } catch (error) {
    res.status(500).send(`Error reading Excel file: ${error.message}`);
  }
});

// Get data for a specific country
app.get('/api/data/country/:country', (req, res) => {
  try {
    const data = readExcelFile();
    const countryData = data.filter(
      (row) => row.Country.toLowerCase() === req.params.country.toLowerCase()
    );
    
    if (countryData.length === 0) {
      return res.status(404).send('No data found for this country');
    }

    res.json(countryData);
  } catch (error) {
    console.error(error);
    res.status(500).send('Error reading Excel file');
  }
});

// Get data for a specific date
app.get('/api/data/date/:date', (req, res) => {
  try {
    const data = readExcelFile();
    const dateData = data.filter(
      (row) => row.Date === req.params.date
    );

    if (dateData.length === 0) {
      return res.status(404).send('No data found for this date');
    }

    res.json(dateData);
  } catch (error) {
    console.error(error);
    res.status(500).send('Error reading Excel file');
  }
});

// Start the server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
