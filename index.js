const express = require('express');
const multer = require('multer');
const { exec } = require('child_process');
const fs = require('fs');
const path = require('path');
const cors = require('cors'); // Import cors
const docxPdf = require('docx-pdf');

const app = express();
const port = 9040;

// Enable CORS for all origins
app.use(cors());

// Configure multer for file uploads
const upload = multer({ dest: 'uploads/' });

// Endpoint to upload DOCX and convert to PDF
app.post('/convert', upload.single('file'), (req, res) => {
    if (!req.file || path.extname(req.file.originalname).toLowerCase() !== '.docx') {
        return res.status(400).send('Please upload a valid DOCX file.');
    }

    const docxPath = path.join(__dirname, req.file.path);
    const outputPdfPath = path.join(__dirname, 'uploads', `${req.file.filename}.pdf`);

    docxPdf(docxPath, outputPdfPath, (err) => {
        if (err) {
            console.error(`Conversion error: ${err}`);
            return res.status(500).send('Failed to convert DOCX to PDF.');
        }

        res.download(outputPdfPath, 'converted-file.pdf', () => {
            fs.unlinkSync(docxPath);
            fs.unlinkSync(outputPdfPath);
        });
    });
});

// Start the server
app.listen(port, () => {
    console.log(`Server is running on http://localhost:${port}`);
});
