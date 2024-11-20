const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const cors = require('cors');

const app = express();
const port = 5000;

// Enable CORS for frontend communication
app.use(cors());

// Middleware to parse JSON data
app.use(express.json());

// Serve static files from the views directory
app.use(express.static(path.join(__dirname, 'views')));

// Configure multer for file upload
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, 'uploads/'),
  filename: (req, file, cb) => cb(null, `${Date.now()}-${file.originalname}`)
});
const upload = multer({ storage });

// Route to serve the HTML page
app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, 'views', 'index.html')); // Adjusted for your setup
});

// Route to handle file upload and reading Excel files
app.post("/uploads", upload.single('uploadfile'), (req, res) => {
  if (req.file && ['application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'].includes(req.file.mimetype)) {
    const filePath = path.join(__dirname, req.file.path);

    try {
      const workbook = XLSX.readFile(filePath);
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(sheet);

      fs.unlinkSync(filePath); // Delete file after processing
      return res.json({ success: true, data });
    } catch (error) {
      console.error("Error reading Excel file:", error);
      return res.status(500).json({ success: false, message: 'Error processing file' });
    }
  }

  return res.status(400).json({ success: false, message: 'Invalid file type. Please upload an Excel file.' });
});

// Start the server
app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});
