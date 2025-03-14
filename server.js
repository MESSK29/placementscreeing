const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const app = express();
const port = 3000;
const upload = multer({ storage: multer.memoryStorage() });

app.use(express.static('public'));

app.post('/process-file', upload.single('excelFile'), (req, res) => {
    if (req.file.originalname !== 'givendata.xlsx') {
        return res.status(400).send('Please upload a file named "givendata.xlsx"');
    }

    const workbook = XLSX.read(req.file.buffer);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);

    const results = jsonData.map(student => {
        const attendance = parseFloat(student['Attendance (%)']);
        const internships = parseInt(student['Number of Internships']);
        const sgpa = parseFloat(student['Last SGPA (%)']);
        const rollNumber = student['Roll Number'];

        let eligibility = '';
        if (attendance > 80 && internships > 2) {
            if (sgpa > 8) {
                eligibility = 'Eligible for Product Based Company';
            } else {
                eligibility = 'Eligible for Service Based Company';
            }
        } else {
            eligibility = 'Not eligible for company selection';
        }

        return {
            'Roll Number': rollNumber,
            'Eligibility': eligibility,
            'Resume Suggestion': 'Would you like to assume your resume for preferred job role?',
            'Link': 'https://jobsinline.netlify.app/' // Separate field for the link
        };
    });

    // Create a new worksheet manually
    const outputWorksheet = {};
    const headers = ['Roll Number', 'Eligibility', 'Resume Suggestion'];

    // Add headers
    outputWorksheet['A1'] = { v: 'Roll Number', t: 's' };
    outputWorksheet['B1'] = { v: 'Eligibility', t: 's' };
    outputWorksheet['C1'] = { v: 'Resume Suggestion', t: 's' };

    // Add data with hyperlink
    results.forEach((row, index) => {
        const rowNum = index + 2; // Start from row 2 (after header)
        outputWorksheet[`A${rowNum}`] = { v: row['Roll Number'], t: 's' };
        outputWorksheet[`B${rowNum}`] = { v: row['Eligibility'], t: 's' };
        outputWorksheet[`C${rowNum}`] = {
            v: row['Resume Suggestion'], // Display text
            t: 's', // String type
            l: { Target: row['Link'] } // Hyperlink
        };
    });

    // Set the range of the worksheet
    outputWorksheet['!ref'] = `A1:C${results.length + 1}`;

    // Create and write the workbook
    const outputWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(outputWorkbook, outputWorksheet, 'Results');
    
    const buffer = XLSX.write(outputWorkbook, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Disposition', 'attachment; filename=resultdata.xlsx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buffer);
});

app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});