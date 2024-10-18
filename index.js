require("dotenv").config();
const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const nodemailer = require("nodemailer");
const path = require("path");
const fs = require("fs");

const app = express();
const port = process.env.PORT || 3000;

const upload = multer({ dest: "uploads/" });

app.use(express.static("public"));
app.use(express.urlencoded({ extended: true }));

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public/index.html"));
});

app.post("/upload", upload.single("file"), async (req, res) => {
  const filePath = path.join(__dirname, req.file.path);
  const workbook = xlsx.readFile(filePath);
  const sheet_name_list = workbook.SheetNames;
  const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

  const emailTemplate = req.body.emailTemplate;
  let sentCount = 0;
  let failedCount = 0;

  const transporter = nodemailer.createTransport({
    service: "Gmail",
    auth: {
      user: process.env.EMAIL_USER,
      pass: process.env.EMAIL_PASS,
    },
  });

  let emailReport = [];

  for (const row of data) {
    const personalizedTemplate = emailTemplate.replace("{{name}}", row.name);

    const mailOptions = {
      from: process.env.EMAIL_USER,
      to: row.email,
      subject: "Happy Durga Puja 2024 - Hansaria Food Private Limited",
      html: personalizedTemplate,
    };

    try {
      await transporter.sendMail(mailOptions);
      sentCount++;
      console.log(`Email sent to ${row.email}`);
      emailReport.push({
        Name: row.name,
        Email: row.email,
        Status: "Sent",
      });
    } catch (error) {
      failedCount++;
      console.error(`Error sending email to ${row.email}:`, error);
      emailReport.push({
        Name: row.name,
        Email: row.email,
        Status: "Failed",
        Error: error.message,
      });
    }
  }

  fs.unlinkSync(filePath);

  const newWorkbook = xlsx.utils.book_new();
  const reportSheet = xlsx.utils.json_to_sheet(emailReport);
  xlsx.utils.book_append_sheet(newWorkbook, reportSheet, "Email Report");

  const reportFilePath = path.join(__dirname, "uploads/email_report.xlsx");
  xlsx.writeFile(newWorkbook, reportFilePath);

  res.json({
    sentCount,
    failedCount,
    reportLink: `http://localhost:${port}/uploads/email_report.xlsx`,
  });
});

app.use("/uploads", express.static(path.join(__dirname, "uploads")));

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
