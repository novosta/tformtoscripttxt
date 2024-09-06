const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

// Set up Express app
const app = express();
const port = 5000;

// Set up Multer for file uploads
const upload = multer({ dest: 'uploads/' });

// Serve the HTML file upload page
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// Handle file uploads and convert Excel to script
app.post('/convert', upload.single('file'), (req, res) => {
  const filePath = req.file.path;

  // Read the uploaded Excel file
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const data = xlsx.utils.sheet_to_json(worksheet);

  // Create output script
  const outputFilePath = path.join(__dirname, 'uploads', 'output_script.txt');
  const writeStream = fs.createWriteStream(outputFilePath);

  data.forEach((row) => {
    const whatSeen = row['What is seen on Camera'] || '';
    const dialogue = row['Dialogue'] || '';
    if (whatSeen && dialogue) {
      writeStream.write(`(${whatSeen}) ${dialogue}\n\n`);
    } else if (dialogue) {
      writeStream.write(`${dialogue}\n\n`);
    }
  });

  writeStream.end();

  // Send the output file back to the user
  writeStream.on('finish', () => {
    res.download(outputFilePath);
  });
});

// Start the server
app.listen(port, () => {
  console.log(`Server running on http://localhost:${port}`);
});
