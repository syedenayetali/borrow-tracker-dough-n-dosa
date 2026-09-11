// ============================================
// BORROW TRACKER - GOOGLE APPS SCRIPT BACKEND
// ============================================
// Copy all of this code into Google Apps Script
// See SETUP_GUIDE.md for detailed instructions
// ============================================

// IMPORTANT: Put your Google Sheet ID here
// Get it from URL: https://docs.google.com/spreadsheets/d/SHEET_ID_HERE/edit
// Example: "1a2b3c4d5e6f7g8h9i0j1k2l3m4n5o6p"
const SHEET_ID = "1AqRl_0zxjU3uHlNzDHvcntHLERuNFe6TyeYCWNMMUG0";
const CUSTOMERS_SHEET_NAME = "Customers";
const TRANSACTIONS_SHEET_NAME = "Transactions";

// Handle POST requests from the web app
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    if (action === "saveCustomers") {
      saveCustomersToSheet(data.customers);
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    } else if (action === "getCustomers") {
      const customers = getCustomersFromSheet();
      const transactions = getTransactionsFromSheet();
      return ContentService.createTextOutput(JSON.stringify({ success: true, customers: customers, transactions: transactions }))
        .setMimeType(ContentService.MimeType.JSON);
    } else if (action === "saveTransactions") {
      console.log("DEBUG: Received transactions:", JSON.stringify(data.transactions));
      saveTransactionsToSheet(data.transactions);
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ success: false, error: "Unknown action" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Save customers to Google Sheet
function saveCustomersToSheet(customers) {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // Get or create Customers sheet
  let sheet = ss.getSheetByName(CUSTOMERS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CUSTOMERS_SHEET_NAME);
  }

  // Clear existing data (keep headers)
  const maxRows = sheet.getMaxRows();
  if (maxRows > 1) {
    sheet.deleteRows(2, maxRows - 1);
  }

  // Add headers if first time
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Customer Name", "Date", "Login Code"]);
  }

  // Add customer data (without transaction stats)
  for (const [name, data] of Object.entries(customers)) {
    sheet.appendRow([
      name,
      data.date || new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }),
      data.code || ""
    ]);
  }

  // Auto-resize columns
  sheet.autoResizeColumns(1, 3);
}

// Load customers from Google Sheet
function getCustomersFromSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(CUSTOMERS_SHEET_NAME);

  if (!sheet) return {};

  const data = sheet.getDataRange().getValues();
  const customers = {};

  // Skip header row (row 0)
  for (let i = 1; i < data.length; i++) {
    const name = data[i][0];
    const code = data[i][2];  // Column C (index 2) contains Login Code

    if (name && name.toString().trim()) {
      customers[name] = {
        totalAmount: 0,
        transactions: [],
        code: code ? code.toString().trim() : ""
      };
    }
  }

  return customers;
}

// Save transactions to Google Sheet
function saveTransactionsToSheet(transactions) {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // Get or create Transactions sheet
  let sheet = ss.getSheetByName(TRANSACTIONS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(TRANSACTIONS_SHEET_NAME);
  }

  // Clear existing data (keep headers)
  const maxRows = sheet.getMaxRows();
  if (maxRows > 1) {
    sheet.deleteRows(2, maxRows - 1);
  }

  // Add headers if first time
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Date", "Customer Name", "Customer Code", "Type", "Amount", "Order ID", "Payment Mode", "Settlement Pending", "Settled"]);
  }

  // Add transaction data
  for (const trans of transactions) {
    const rowData = [
      trans.date || new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }),
      trans.customerName || "",
      trans.customerCode || "",
      trans.type || "",
      trans.amount || 0,
      trans.orderId || "",
      trans.paymentMode || "",
      trans.settlementPending || 0,
      trans.settled || 0
    ];

    console.log("DEBUG: Appending row:", JSON.stringify(rowData));
    sheet.appendRow(rowData);
  }

  // Auto-resize columns
  sheet.autoResizeColumns(1, 9);
}

// Load transactions from Google Sheet
function getTransactionsFromSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TRANSACTIONS_SHEET_NAME);

  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const transactions = [];

  // Skip header row (row 0)
  for (let i = 1; i < data.length; i++) {
    const date = data[i][0];
    const customerName = data[i][1];
    const customerCode = data[i][2];
    const type = data[i][3];
    const amount = data[i][4];
    const orderId = data[i][5];
    const paymentMode = data[i][6];
    const settlementPending = data[i][7];
    const settled = data[i][8];

    if (date && customerName && type) {
      transactions.push({
        date: date.toString(),
        customerName: customerName.toString().trim(),
        customerCode: customerCode ? customerCode.toString().trim() : "",
        type: type.toString().trim(),
        amount: parseFloat(amount) || 0,
        orderId: orderId ? orderId.toString().trim() : "",
        paymentMode: paymentMode ? paymentMode.toString().trim() : "",
        settlementPending: parseFloat(settlementPending) || 0,
        settled: parseFloat(settled) || 0
      });
    }
  }

  return transactions;
}


// Handle GET requests (for testing deployment)
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    message: "Borrow Tracker Backend is working!",
    status: "ready",
    timestamp: new Date().toLocaleString()
  }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Test function to verify everything works
function testConnection() {
  Logger.log("Testing Borrow Tracker connection...");

  if (SHEET_ID === "YOUR_SHEET_ID_HERE") {
    Logger.log("❌ ERROR: Please set SHEET_ID in the script!");
    Logger.log("Get your Sheet ID from: https://docs.google.com/spreadsheets/d/SHEET_ID_HERE/edit");
    return;
  }

  // Get or create sheets
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let customersSheet = ss.getSheetByName(CUSTOMERS_SHEET_NAME);
  let transactionsSheet = ss.getSheetByName(TRANSACTIONS_SHEET_NAME);

  if (!customersSheet) {
    customersSheet = ss.insertSheet(CUSTOMERS_SHEET_NAME);
    Logger.log("✓ Created Customers sheet");
  } else {
    Logger.log("✓ Customers sheet exists");
  }

  if (!transactionsSheet) {
    transactionsSheet = ss.insertSheet(TRANSACTIONS_SHEET_NAME);
    Logger.log("✓ Created Transactions sheet");
  } else {
    Logger.log("✓ Transactions sheet exists");
  }

  // Add test data
  const testCustomer = {
    "Test Customer": {
      totalAmount: 0,
      transactions: [],
      code: "1234"
    }
  };

  saveCustomersToSheet(testCustomer);
  Logger.log("✓ Test customer data saved successfully");

  const testTransaction = [{
    date: new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }),
    customerName: "Test Customer",
    type: "purchase",
    amount: 500,
    orderId: "ORD-001",
    paymentMode: "Cash"
  }];

  saveTransactionsToSheet(testTransaction);
  Logger.log("✓ Test transaction data saved successfully");

  const loadedCustomers = getCustomersFromSheet();
  const loadedTransactions = getTransactionsFromSheet();
  Logger.log("✓ Data loaded successfully");
  Logger.log("Loaded customers: " + JSON.stringify(loadedCustomers));
  Logger.log("Loaded transactions: " + JSON.stringify(loadedTransactions));

  Logger.log("✅ All systems working!");
}
