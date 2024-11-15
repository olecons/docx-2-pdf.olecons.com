const express = require('express');
const multer = require('multer');
const { exec } = require('child_process');
const fs = require('fs');
const path = require('path');
const cors = require('cors'); // Import cors

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
    // const command = `libreoffice --headless --convert-to pdf:"writer_pdf_Export" "${docxPath}" --outdir "${path.dirname(outputPdfPath)}"`;
    const command = `pandoc "${docxPath}" -o "${outputPdfPath}" --pdf-engine=wkhtmltopdf`;
    console.log({command});

    exec(command, (error, stdout, stderr) => {
        if (error) {
            console.error(`Conversion error: ${stderr}`);
            return res.status(500).send('Failed to convert DOCX to PDF.');
        }

        fs.readFile(outputPdfPath, (err, data) => {
            if (err) {
                console.error(`File read error: ${err}`);
                return res.status(500).send('Failed to read converted PDF.');
            }

            res.setHeader('Content-Type', 'application/pdf');
            res.setHeader('Content-Disposition', 'attachment; filename="converted-file.pdf"');
            res.send(data);

            fs.unlinkSync(docxPath);
            fs.unlinkSync(outputPdfPath);
        });
    });
});

// Start the server
app.listen(port, () => {
    console.log(`Server is running on http://localhost:${port}`);
});
