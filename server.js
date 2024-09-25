const express = require('express');
const mysql = require('mysql');
const XLSX = require('xlsx');
const cors = require('cors');
const app = express();
const port = 5000;

// Middleware
app.use(cors());

// MySQL database connection
const connection = mysql.createConnection({
    host: 'localhost',
    user: 'root',
    password: '',
    database: 'excel',
    port:3308
});
connection.connect((err) => {
    if (err) {
        console.error('Error connecting to the database:', err);
        return;
    }
    console.log('Connected to the database.');
});


// Route to export data to Excel
app.get('/export', (req, res) => {
    connection.query('SELECT * FROM employees', (error, results) => {
        if (error) {
            return res.status(500).send('Error fetching data from the database.');
        }

        // Convert data to a worksheet
        const worksheet = XLSX.utils.json_to_sheet(results);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Employees');

        // Create a buffer to send the file
        const filePath = 'employees.xlsx';
        XLSX.writeFile(workbook, filePath);

        // Set the response headers to download the file
        res.download(filePath, (err) => {
            if (err) {
                res.status(500).send('Error downloading the file.');
            }
        });
    });
});

// Start the server
app.listen(port, () => {
    console.log(`Server is running on http://localhost:${port}`);
});
