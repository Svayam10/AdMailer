const fs = require("fs");
const nodemailer = require("nodemailer");
const XLSX = require("xlsx");
require("dotenv").config();

// Configuration
const EMAIL_ADDRESS = process.env.EMAIL_ADDRESS;
const EMAIL_PASSWORD = process.env.EMAIL_PASSWORD;

const EXCEL_FILE = "clients.xlsx";
const HTML_TEMPLATE_FILE = "email_template.html";
const EMAIL_COLUMN = "Email";
const NAME_COLUMN = "Name";

const SUBJECT = "🌟 Special Offer for You!";

// Read recipients from Excel
function readRecipients(filePath) {
  try {
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet);

    // Filter rows that have both email and name
    return data.filter(
      (row) => row[EMAIL_COLUMN] && row[NAME_COLUMN]
    );
  } catch (error) {
    console.error(`Error reading Excel file: ${error.message}`);
    return [];
  }
}

// Load HTML template
function loadHtmlTemplate() {
  try {
    return fs.readFileSync(HTML_TEMPLATE_FILE, "utf-8");
  } catch (error) {
    console.error(`Error reading template: ${error.message}`);
    return "";
  }
}

// Send email
async function sendEmail(toEmail, recipientName, htmlContent) {
  // Replace placeholder with actual name
  const personalizedHtml = htmlContent.replace("{NAME}", recipientName);

  // Create transporter
  const transporter = nodemailer.createTransport({
    host: "smtp.gmail.com",
    port: 465,
    secure: true,
    auth: {
      user: EMAIL_ADDRESS,
      pass: EMAIL_PASSWORD,
    },
  });

  try {
    await transporter.sendMail({
      from: EMAIL_ADDRESS,
      to: toEmail,
      subject: SUBJECT,
      text: "This is an HTML email. Please use an email client that supports HTML.",
      html: personalizedHtml,
    });

    console.log(
      `[${new Date().toISOString()}] ✅ Sent to: ${toEmail}`
    );
  } catch (error) {
    console.error(
      `[${new Date().toISOString()}] ❌ Failed to send to ${toEmail}: ${error.message}`
    );
  }
}

// Main execution
async function main() {
  const recipients = readRecipients(EXCEL_FILE);
  const htmlTemplate = loadHtmlTemplate();

  if (recipients.length === 0) {
    console.log("❌ No recipients found!");
    return;
  }

  console.log(`📬 Sending to ${recipients.length} recipients...\n`);

  for (const recipient of recipients) {
    await sendEmail(
      recipient[EMAIL_COLUMN],
      recipient[NAME_COLUMN],
      htmlTemplate
    );
    // Wait 2 seconds between emails to avoid spam flag
    await new Promise((resolve) => setTimeout(resolve, 2000));
  }

  console.log("\n✨ All emails sent!");
}

main().catch((error) => {
  console.error("Fatal error:", error);
  process.exit(1);
});
