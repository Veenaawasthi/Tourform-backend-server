const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const fs = require('fs');
const XLSX = require('xlsx');

const app = express();
const PORT = 5000;

app.use(cors());
app.use(bodyParser.json());

app.get('/', (req, res) => {
  res.send('Server is running!');
});

app.post('/submit', async (req, res) => {
  const data = req.body;
  console.log('Received data:', data); // Log incoming data for debugging

  // Validate the data
  if (!data.name || !data.email) {
    return res.status(400).send('Name and email are required');
  }

  try {
    const filePath = 'TourDetails.xlsx';
    let workbook;

    if (fs.existsSync(filePath)) {
      // Add a retry mechanism in case the file is temporarily locked
      const MAX_RETRIES = 5;
      let retryCount = 0;
      let fileOpened = false;

      while (!fileOpened && retryCount < MAX_RETRIES) {
        try {
          workbook = XLSX.readFile(filePath);
          fileOpened = true;
        } catch (err) {
          if (err.code === 'EBUSY' || err.code === 'EPERM') {
            console.log('File is busy, retrying...');
            retryCount++;
            await new Promise(resolve => setTimeout(resolve, 1000)); // wait for 1 second before retrying
          } else {
            throw err; // rethrow if it's not a file busy error
          }
        }
      }

      if (!fileOpened) {
        throw new Error('Could not open file after multiple retries');
      }
    } else {
      workbook = XLSX.utils.book_new();
    }

    // Convert data to worksheet format
    const worksheet = XLSX.utils.json_to_sheet([data], { origin: -1 });
    const sheetName = 'TourDetails';

    if (workbook.Sheets[sheetName]) {
      XLSX.utils.sheet_add_json(workbook.Sheets[sheetName], [data], { skipHeader: true, origin: -1 });
    } else {
      XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    }

    XLSX.writeFile(workbook, filePath);
    res.status(200).send('Form submitted successfully');
  } catch (error) {
    console.error('Error writing to Excel file:', error.message);
    res.status(500).send('Server error');
  }
});

app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});




  
