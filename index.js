const express = require('express');
const multer = require('multer');
const { exec } = require('child_process');
const fs = require('fs');
const path = require('path');
const cors = require('cors'); // Import cors
const bodyParser = require('body-parser');
const Docxtemplater = require('docxtemplater');
const PizZip = require('pizzip');
const DocxMerger = require('docx-merger');

const app = express();
const port = 9040;

// Enable CORS for all origins
app.use(cors());
app.use(bodyParser.json());

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

// Utility to replace placeholders in a DOCX
async function replacePlaceholders(fileUrl, dataValues) {
    const response = await fetch(fileUrl);
    const fileBuffer = await response.arrayBuffer();

    const zip = new PizZip(fileBuffer);
    const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
        delimiters: { start: '###', end: '###' }
    });
    // Create the replaceData object with formatted keys
    const replaceData = {};
    for (const key in dataValues) {
        if (dataValues.hasOwnProperty(key)) {
            // Use the key without prefixes and surround with ###
            const placeholder = `${key.replace(/^(initial_|variation_|end_)/, '')}`;
            replaceData[placeholder] = dataValues[key];
        }
    }
    // Set the data for replacement
    // doc.setData(replaceData);
    doc.render(replaceData);
    return doc.getZip().generate({ type: 'nodebuffer' });
}

// Convert a DOCX file to PDF using LibreOffice
async function convertDocxToPdf(docxBuffer, outputDir) {
    const tempDocxPath = path.join(outputDir, `temp-${Date.now()}.docx`);
    const tempPdfPath = tempDocxPath.replace('.docx', '.pdf');

    fs.writeFileSync(tempDocxPath, docxBuffer);

    return new Promise((resolve, reject) => {
        const command = `libreoffice --headless --convert-to pdf "${tempDocxPath}" --outdir "${outputDir}"`;
        exec(command, (error, stdout, stderr) => {
            if (error) {
                console.error(`LibreOffice conversion error: ${stderr}`);
                return reject(error);
            }

            fs.unlinkSync(tempDocxPath); // Clean up the DOCX file
            resolve(tempPdfPath);
        });
    });
}

// Merge multiple PDF files into one
async function mergePdfFiles(pdfPaths, outputFilePath) {
    const mergedPdf = await PDFDocument.create();

    for (const pdfPath of pdfPaths) {
        const pdfBytes = fs.readFileSync(pdfPath);
        const pdfDoc = await PDFDocument.load(pdfBytes);
        const copiedPages = await mergedPdf.copyPages(pdfDoc, pdfDoc.getPageIndices());
        copiedPages.forEach((page) => mergedPdf.addPage(page));
    }

    const mergedPdfBytes = await mergedPdf.save();
    fs.writeFileSync(outputFilePath, mergedPdfBytes);

    // Clean up individual PDFs
    pdfPaths.forEach((pdfPath) => fs.unlinkSync(pdfPath));
}

app.post('/merge-pdf', async (req, res) => {
    const { initial, vacations, end } = req.body;

    try {
        const outputDir = path.join(__dirname, 'temp');
        if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir);

        const pdfPaths = [];

        // Replace placeholders in the initial file and convert to PDF
        const initialBuffer = await replacePlaceholders(initial.file_url, initial.data_values);
        const initialPdfPath = await convertDocxToPdf(initialBuffer, outputDir);
        pdfPaths.push(initialPdfPath);

        // Replace placeholders in each vacation file and convert to PDF
        for (let vacation of vacations) {
            const vacationBuffer = await replacePlaceholders(vacation.file_url, vacation.data_values);
            const vacationPdfPath = await convertDocxToPdf(vacationBuffer, outputDir);
            pdfPaths.push(vacationPdfPath);
        }

        // Replace placeholders in the end file and convert to PDF
        const endBuffer = await replacePlaceholders(end.file_url, end.data_values);
        const endPdfPath = await convertDocxToPdf(endBuffer, outputDir);
        pdfPaths.push(endPdfPath);

        // Merge all PDFs
        const mergedPdfPath = path.join(outputDir, `merged-${Date.now()}.pdf`);
        await mergePdfFiles(pdfPaths, mergedPdfPath);

        // Send the merged PDF as a response
        res.sendFile(mergedPdfPath, () => {
            // Clean up the merged PDF
            pdfPaths.map(fs.unlinkSync);
            fs.unlinkSync(mergedPdfPath);
        });
    } catch (error) {
        console.error('Error processing files:', error);
        res.status(500).send('An error occurred while processing the files.');
    }
});

// Start the server
app.listen(port, () => {
    console.log(`Server is running on http://localhost:${port}`);
});
