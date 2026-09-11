// ============================================
// BORROW TRACKER - CLEAN GOOGLE APPS SCRIPT
// ============================================

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
    }
    else if (action === "getCustomers") {
      const customers = getCustomersFromSheet();
      const transactions = getTransactionsFromSheet();
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        customers: customers,
        transactions: transactions
      }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    else if (action === "saveTransactions") {
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

  // Add customer data
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

// Get customers from Google Sheet
function getCustomersFromSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(CUSTOMERS_SHEET_NAME);

  if (!sheet) return {};

  const data = sheet.getDataRange().getValues();
  const customers = {};

  // Skip header row (row 0)
  for (let i = 1; i < data.length; i++) {
    const name = data[i][0];
    const date = data[i][1];
    const code = data[i][2];

    if (name && name.toString().trim()) {
      customers[name.toString().trim()] = {
        date: date || "",
        code: code ? code.toString().trim() : ""
      };
    }
  }

  return customers;
}

// Save transactions to Google Sheet
function saveTransactionsToSheet(transactions) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
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
    sheet.appendRow([
      "Date",
      "Customer Name",
      "Customer Code",
      "Type",
      "Amount",
      "Order ID",
      "Payment Mode",
      "Settlement Pending",
      "Settled"
    ]);
  }

  // Add transaction data
  for (const trans of transactions) {
    sheet.appendRow([
      trans.date || "",
      trans.customerName || "",
      trans.customerCode || "",
      trans.type || "",
      trans.amount || 0,
      trans.orderId || "",
      trans.paymentMode || "",
      trans.settlementPending || 0,
      trans.settled || 0
    ]);
  }

  // Auto-resize columns
  sheet.autoResizeColumns(1, 9);
}

// Get transactions from Google Sheet
function getTransactionsFromSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TRANSACTIONS_SHEET_NAME);

  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const transactions = [];

  // Skip header row (row 0)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    if (row[1]) { // If customer name exists
      transactions.push({
        date: row[0] || "",
        customerName: row[1] ? row[1].toString().trim() : "",
        customerCode: row[2] ? row[2].toString().trim() : "",
        type: row[3] || "",
        amount: parseFloat(row[4]) || 0,
        orderId: row[5] || "",
        paymentMode: row[6] || "",
        settlementPending: parseFloat(row[7]) || 0,
        settled: parseFloat(row[8]) || 0
      });
    }
  }

  return transactions;
}
