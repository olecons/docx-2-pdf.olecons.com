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

app.post('/convert2', upload.single('file'), async (req, res) => {
    if (!req.file || path.extname(req.file.originalname).toLowerCase() !== '.docx') {
        return res.status(400).send('Please upload a valid DOCX file.');
    }

    try {
        // Upload the DOCX file to Google Drive
        const fileMetadata = {
            name: req.file.originalname,
            mimeType: 'application/vnd.google-apps.document', // This will convert the file to Google Docs format
        };
        const media = {
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            body: fs.createReadStream(req.file.path),
        };

        const uploadResponse = await drive.files.create({
            resource: fileMetadata,
            media: media,
            fields: 'id',
        });

        const fileId = uploadResponse.data.id;

        // Export the Google Docs file to PDF
        const pdfResponse = await drive.files.export(
            {
                fileId: fileId,
                mimeType: 'application/pdf',
            },
            { responseType: 'stream' }
        );

        const outputPdfPath = path.join(__dirname, 'uploads', `${req.file.filename}.pdf`);
        const writeStream = fs.createWriteStream(outputPdfPath);

        pdfResponse.data.pipe(writeStream);

        writeStream.on('finish', async () => {
            // Respond with the PDF file
            res.download(outputPdfPath, 'converted-file.pdf', async () => {
                // Clean up files
                fs.unlinkSync(req.file.path);
                fs.unlinkSync(outputPdfPath);

                // Optionally delete the file from Google Drive
                await drive.files.delete({ fileId: fileId });
            });
        });
    } catch (error) {
        console.error(error);
        res.status(500).send('Failed to convert DOCX to PDF.');
    }
});

// Endpoint to upload DOCX and convert to PDF
app.post('/convert', upload.single('file'), (req, res) => {
    if (!req.file || path.extname(req.file.originalname).toLowerCase() !== '.docx') {
        return res.status(400).send('Please upload a valid DOCX file.');
    }

    const docxPath = path.join(__dirname, req.file.path);
    const outputPdfPath = path.join(__dirname, 'uploads', `${req.file.filename}.pdf`);
    const command = `libreoffice --headless --convert-to pdf "${docxPath}" --outdir "${path.dirname(outputPdfPath)}"`;

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
