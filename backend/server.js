const express = require("express");
const bodyParser = require("body-parser");
const cors = require("cors");
const fs = require("fs");
const XLSX = require("xlsx");

const app = express();
app.use(cors());
app.use(bodyParser.json());

const FILE_PATH = "data.xlsx";
const SHEET_NAME = "GetSetPy";

// Function to read Excel file
const readExcel = () => {
    if (!fs.existsSync(FILE_PATH)) return [];

    const workbook = XLSX.readFile(FILE_PATH);
    const worksheet = workbook.Sheets[SHEET_NAME];

    return worksheet ? XLSX.utils.sheet_to_json(worksheet) : [];
};

// Function to write data to Excel
const writeExcel = (data) => {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(workbook, worksheet, SHEET_NAME);
    XLSX.writeFile(workbook, FILE_PATH);
};

// Function to save initial payment details (with "pending" status)
app.post("/savePStatus", (req, res) => {
    const { name, email, rollno, paymentStatus } = req.body;

    if (!name || !email || !rollno) {
        return res.status(400).json({ message: "Name, email, and roll number are required" });
    }

    let data = readExcel();

    // Check if the roll number already exists
    const existingEntry = data.find(entry => (entry.RollNo === rollno && entry.paymentStatus==="success"));

    if (existingEntry) {
        return res.send({ message: "Payment already initiated for this roll number" });
    }

    // Add new entry with "pending" payment status
    const newEntry = { Name: name, Email: email, RollNo: rollno, paymentStatus: "pending" };
    data.push(newEntry);
    writeExcel(data); 

    res.json({ message: "Payment status set to pending!" });
});

// Function to update payment status to "success"
app.post("/updatePayment", (req, res) => {
    const { rollno, paymentId } = req.body;

    if (!rollno || !paymentId) {
        return res.status(400).json({ message: "Roll number and payment ID are required" });
    }

    let data = readExcel();

    // Find the entry by roll number
    const entryIndex = data.findIndex(entry => entry.RollNo === rollno);

    if (entryIndex === -1) {
        return res.status(404).json({ message: "Roll number not found" });
    }

    // Update payment status and ID
    data[entryIndex].paymentStatus = "success";
    data[entryIndex].paymentId = paymentId;
    writeExcel(data);

    res.json({ message: "Payment status updated to success!" });
});

// Start the server
app.listen(5000, () => console.log("Server running on port 5000"));
