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
const PDFDocument = require('pdf-lib').PDFDocument;

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

async function replacePlaceholders(fileUrl, dataValues, replaceLink = false) {
    try {
        // Fetch the DOCX file from the provided URL
        const response = await fetch(fileUrl);
        if (!response.ok) {
            throw new Error(`Failed to fetch file from URL: ${fileUrl}, Status: ${response.status}`);
        }

        const fileBuffer = await response.arrayBuffer();

        // Validate if the fetched file is a valid ZIP (DOCX)
        try {

            // Initialize Docxtemplater with delimiters and other configurations
            const doc = new Docxtemplater(new PizZip(fileBuffer), {
                paragraphLoop: true,
                linebreaks: true,
                delimiters: { start: '###', end: '###' } // Example delimiter
            });

            // Prepare the replacement data by removing prefixes and formatting keys
            const replaceData = {};
            for (const key in dataValues) {
                if (dataValues.hasOwnProperty(key)) {
                    const placeholder = key.replace(/^(initial_|variation_|end_)/, ''); // Remove prefixes
                    if (dataValues[key] && placeholder == 'CXNAME') {
                        replaceData[placeholder] = " "+dataValues[key];
                    } else if (dataValues[key] && placeholder == 'EVENTDATE') {
                        replaceData[placeholder] = " on "+dataValues[key];
                    } else if (dataValues[key] && placeholder == 'EVENTNAME') {
                        replaceData[placeholder] = (dataValues[key] || '').toUpperCase() + ' - ';
                    } else {
                        replaceData[placeholder] = dataValues[key];
                    }
                }
            }

            // Render the DOCX file with the replacement data
            doc.render(replaceData);

            if(replaceLink && replaceData['SUMMARYLINK']) {
                console.log("replacing link", {replaceData});
                const zip = doc.getZip();
                // Handle hyperlinks in the document.xml.rels
                const relsXmlPath = 'word/_rels/document.xml.rels';
                if (zip.file(relsXmlPath)) {
                    const relsXml = zip.file(relsXmlPath).asText();
                    console.log({relsXml})
                    // Replace all hyperlinks in document.xml.rels
                    const updatedRelsXml = relsXml.replace(
                        /<Relationship[^>]*Target="http[s]?:\/\/[^"]*"[^>]*\/>/g,
                        (match) => {
                            return match.replace(/Target="http[s]?:\/\/[^"]*"/, `Target="${replaceData['SUMMARYLINK']}"`);
                        }
                    );

                    // Update the relationships file in the ZIP
                    zip.file(relsXmlPath, updatedRelsXml);
                }
                return zip.generate({ type: 'nodebuffer' });
            }
            // Return the modified DOCX as a buffer
            return doc.getZip().generate({ type: 'nodebuffer' });
        } catch (zipError) {
            throw new Error(`Error processing DOCX file: ${zipError.message}`);
        }
    } catch (error) {
        console.error('Error in replacePlaceholders:', error.message);
        throw error; // Rethrow the error to be handled by the calling function
    }
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
        const endBuffer = await replacePlaceholders(end.file_url, end.data_values, true);
        const endPdfPath = await convertDocxToPdf(endBuffer, outputDir);
        pdfPaths.push(endPdfPath);

        // Merge all PDFs
        const mergedPdfPath = path.join(outputDir, `merged-${Date.now()}.pdf`);
        await mergePdfFiles(pdfPaths, mergedPdfPath);

        // Send the merged PDF as a response
        res.sendFile(mergedPdfPath, () => {
            // Clean up the merged PDF
            try {
                pdfPaths.map(fs.unlinkSync);
                fs.unlinkSync(mergedPdfPath);
            } catch ($e) {

            }
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
