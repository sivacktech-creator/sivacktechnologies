const express = require('express');
const bodyParser = require('body-parser');
const PDFDocument = require('pdfkit');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();

// Middleware to parse form data and serve static HTML/CSS files
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(__dirname)); 

// POST Route to handle form submission
app.post('/submit-registration', (req, res) => {
    const data = req.body;

    // ==========================================
    // 1. SAVE TO EXCEL SHEET (Backend Storage)
    // ==========================================
    const excelFilePath = path.join(__dirname, 'Registrations.xlsx');
    let workbook;
    let worksheet;

    // Check if Excel file already exists
    if (fs.existsSync(excelFilePath)) {
        // Read existing file
        workbook = XLSX.readFile(excelFilePath);
        worksheet = workbook.Sheets[workbook.SheetNames[0]];
    } else {
        // Create new workbook if it doesn't exist
        workbook = XLSX.utils.book_new();
        worksheet = XLSX.utils.json_to_sheet([]);
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Students');
    }

    // Format the new data row
    const newRow = [{
        "Date Registered": new Date().toLocaleString(),
        "Full Name": data.fullName,
        "Email": data.email,
        "Phone": data.phone,
        "Date of Birth": data.dob,
        "Qualification": data.qualification,
        "Institution": data.institution,
        "Course Selected": data.course,
        "Address": data.address
    }];

    // Add new row to Excel and save it to the server folder
    XLSX.utils.sheet_add_json(worksheet, newRow, { skipHeader: true, origin: -1 });
    XLSX.writeFile(workbook, excelFilePath);


    // ==========================================
    // 2. GENERATE PDF RECEIPT (Download for User)
    // ==========================================
    const doc = new PDFDocument({ margin: 50 });

    // Set headers so the browser downloads the PDF automatically
    res.setHeader('Content-disposition', 'attachment; filename="Sivack_Registration_Receipt.pdf"');
    res.setHeader('Content-type', 'application/pdf');

    // Pipe the PDF directly to the user's browser
    doc.pipe(res);

    // --- PDF Design ---
    // Header
    doc.fontSize(24).fillColor('#001838').text('SIVACK Technologies', { align: 'center' });
    doc.fontSize(12).fillColor('#32c36c').text('Innovation for a bright future', { align: 'center' });
    doc.moveDown();
    doc.moveTo(50, 110).lineTo(550, 110).strokeColor('#dddddd').stroke();
    doc.moveDown();

    // Title
    doc.fontSize(18).fillColor('#001838').text('Registration Confirmation Receipt', { align: 'center' });
    doc.moveDown(2);

    // Details Box
    doc.fontSize(12).fillColor('#333333');
    doc.text(`Registration Date: ${new Date().toLocaleDateString()}`);
    doc.moveDown();
    
    doc.font('Helvetica-Bold').text('Applicant Details:');
    doc.font('Helvetica').text(`Name: ${data.fullName}`);
    doc.text(`Email: ${data.email}`);
    doc.text(`Phone: ${data.phone}`);
    doc.text(`Date of Birth: ${data.dob}`);
    doc.moveDown();

    doc.font('Helvetica-Bold').text('Educational Details:');
    doc.font('Helvetica').text(`Qualification: ${data.qualification}`);
    doc.text(`Institution: ${data.institution}`);
    doc.moveDown();

    doc.font('Helvetica-Bold').text('Program Details:');
    doc.font('Helvetica').text(`Course Selected: ${data.course}`);
    doc.moveDown(2);

    // Footer
    doc.moveTo(50, doc.y).lineTo(550, doc.y).strokeColor('#dddddd').stroke();
    doc.moveDown();
    doc.fontSize(10).fillColor('#777777').text('Thank you for choosing Sivack Technologies. Our team will contact you shortly.', { align: 'center' });

    // Finalize PDF
    doc.end();
});

// Start the server
const PORT = 3000;
app.listen(PORT, () => {
    console.log(`Server is running! Open http://localhost:${PORT}/register.html in your browser.`);
});