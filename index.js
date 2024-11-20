const express = require('express');
const multer = require('multer');
const path = require('path');
const cors = require('cors');
const fs = require('fs');
const XLSX = require('xlsx'); // Import xlsx package to read .xls files

const app = express();
const port = 5000;

// Enable CORS for frontend to communicate with backend
app.use(cors());

// Configure multer for file upload
const uploads = multer({ dest: "uploads/" });

// Set up EJS as the view engine
app.set("view engine", "ejs");
app.set("views", path.resolve("./views"));

// Middleware to parse JSON data
app.use(express.json());
app.use(express.urlencoded({ extended: false }));

// Route for the homepage
app.get("/", (req, res) => {
  return res.render("homepage");
});

const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, 'uploads/');
  },
  filename: function (req, file, cb) {
    cb(null, `${Date.now()}-${file.originalname}`);
  }
});

// Set up multer with the storage configuration
const upload = multer({ storage });

// Post route to handle file upload and reading Excel file
app.post("/uploads", uploads.single('uploadfile'), (req, res) => {
  console.log(req.body);
  console.log(req.file); // The uploaded file info

  // Check if file exists and if it is an .xls or .xlsx file
  if (req.file && (req.file.mimetype === 'application/vnd.ms-excel' || req.file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')) {
    // Read the uploaded file
    const filePath = path.join(__dirname, req.file.path); // Path to the uploaded file

    try {
      // Use xlsx package to read the Excel file
      const workbook = XLSX.readFile(filePath);
      const sheetName = workbook.SheetNames[0]; // Get the first sheet
      const sheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(sheet); // Convert sheet to JSON

      console.log(data); // Log the JSON data from the Excel file

      // Delete the uploaded file after reading it
      fs.unlinkSync(filePath);

      // Send the JSON data back to the client
      res.json({ success: true, data: data });
    } catch (error) {
      console.error("Error reading Excel file:", error);
      res.status(500).json({ success: false, message: 'Error reading the file' });
    }
  } else {
    res.status(400).json({ success: false, message: 'Invalid file type. Please upload an Excel file.' });
  }
});

// Start the server
app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});
