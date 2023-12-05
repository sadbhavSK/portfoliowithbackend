const express = require('express');
const path = require('path');
const bodyParser = require('body-parser');
const fs = require('fs');
const ExcelJS = require('exceljs');

const app = express();
const port = 80;

app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static('public'));

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.post('/submit', async (req, res) => {
    console.log('done')
    try {
        const formData = req.body;
        console.log(formData);

        // Load the existing Excel file if it exists, or create a new one if not
        const filePath = path.join(__dirname, 'file', 'contactData.xlsx');
        let workbook;

        if (fs.existsSync(filePath)) {
            workbook = new ExcelJS.Workbook();
            await workbook.xlsx.readFile(filePath);
        } else {
            workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('ContactResponses');

            // Add headers with formatting
            const headerRow = worksheet.addRow(['Name', 'Email', 'Message']);
            headerRow.font = { bold: true };
            headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
            headerRow.height = 20;

            // Apply cell border formatting to headers
            headerRow.eachCell((cell) => {
                cell.border = {
                    top: { style: 'thin' },
                    bottom: { style: 'thin' },
                    left: { style: 'thin' },
                    right: { style: 'thin' },
                };
            });
        }

        const worksheet = workbook.getWorksheet('ContactResponses');

        const newRow = worksheet.addRow([
            formData.name,
            formData.email,
            formData.message,
        ]);

        newRow.eachCell((cell) => {
            cell.border = {
                top: { style: 'thin' },
                bottom: { style: 'thin' },
                left: { style: 'thin' },
                right: { style: 'thin' },
            };
        });
        res.sendFile(path.join(__dirname, 'public', 'index.html'));

        // Save the updated Excel file
        await workbook.xlsx.writeFile(filePath);
        res.sendFile(path.join(__dirname, 'public', 'index.html'));

    } catch (error) {
        console.error(error);
        res.status(500).send('An error occurred.');
    }
});

app.listen(port, () => {
    console.log(`App is successfully serving at port ${port}`);
});
