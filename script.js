console.log('script.js loaded successfully');

document.getElementById('processBtn').addEventListener('click', async function() {
    console.log('Process button clicked');
    
    const fileInput = document.getElementById('inputFile');
    const file = fileInput.files[0];
    
    if (!file) {
        console.error('No file uploaded');
        showResult("Please upload 'givendata.xlsx' file!");
        return;
    }

    if (file.name !== 'givendata.xlsx') {
        console.error('Wrong filename:', file.name);
        showResult("Please upload a file named 'givendata.xlsx'!");
        return;
    }

    try {
        console.log('Reading file:', file.name);
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data);
        const sheetName = workbook.SheetNames[0];
        console.log('Sheet name:', sheetName);
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        console.log('Data extracted:', jsonData);

        // Process data
        const results = jsonData.map(student => {
            const attendance = parseFloat(student['Attendance (%)']);
            const internships = parseInt(student['Number of Internships']);
            const sgpa = parseFloat(student['Last SGPA (%)']);
            const rollNumber = student['Roll Number'];

            console.log(`Processing Roll ${rollNumber}: SGPA=${sgpa}, Internships=${internships}, Attendance=${attendance}`);

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
                'Link': 'https://jobsinline.netlify.app/'
            };
        });

        console.log('Results generated:', results);

        // Create worksheet manually
        const outputWorksheet = {};
        const headers = ['Roll Number', 'Eligibility', 'Resume Suggestion'];
        
        // Add headers
        outputWorksheet['A1'] = { v: 'Roll Number', t: 's' };
        outputWorksheet['B1'] = { v: 'Eligibility', t: 's' };
        outputWorksheet['C1'] = { v: 'Resume Suggestion', t: 's' };

        // Add data with hyperlink
        results.forEach((row, index) => {
            const rowNum = index + 2;
            outputWorksheet[`A${rowNum}`] = { v: row['Roll Number'], t: 's' };
            outputWorksheet[`B${rowNum}`] = { v: row['Eligibility'], t: 's' };
            outputWorksheet[`C${rowNum}`] = {
                v: row['Resume Suggestion'],
                t: 's',
                l: { Target: row['Link'] }
            };
            console.log(`Row ${rowNum} Resume Suggestion:`, outputWorksheet[`C${rowNum}`]);
        });

        // Set range
        outputWorksheet['!ref'] = `A1:C${results.length + 1}`;
        console.log('Worksheet data:', outputWorksheet);

        // Create and write workbook
        const outputWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(outputWorkbook, outputWorksheet, 'Results');
        
        console.log('Writing file: resultdata.xlsx');
        XLSX.writeFile(outputWorkbook, 'resultdata.xlsx');
        
        showResult("Processing complete! Results have been downloaded as 'resultdata.xlsx'");
    } catch (error) {
        console.error('Error processing file:', error);
        showResult("An error occurred while processing the file. Check console for details.");
    }
});

function showResult(message) {
    const resultDiv = document.getElementById('result');
    resultDiv.style.display = 'block';
    resultDiv.innerHTML = `<p>${message}</p>`;
}

console.log('script.js initialization complete');