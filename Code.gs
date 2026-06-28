// ============================================================
// EXPENSE TRACKER - Google Apps Script
// Author: Fardeen Salmani
// Description: Auto-logs Kotak Bank transactions to Google
//              Sheets via phone notifications + Automate app.
//              Includes manual entry, budget alerts, and
//              monthly email reports.
// ============================================================


// --- MENU SYSTEM ---
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('💰 Expense Manager')
      .addItem('➕ Add New Expense', 'addExpense')
      .addItem('❌ Delete Last Entry', 'deleteLastRow')
      .addToUi();
}


// --- MANUAL ENTRY FUNCTION ---
function addExpense() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  var responseCategory = ui.prompt('Enter Category (e.g., Food, Travel, Miscellaneous):');
  if (responseCategory.getSelectedButton() == ui.Button.CLOSE) return;
  var category = responseCategory.getResponseText();

  var responseAmount = ui.prompt('Enter Amount (₹):');
  if (responseAmount.getSelectedButton() == ui.Button.CLOSE) return;
  var amount = parseFloat(responseAmount.getResponseText());

  if (isNaN(amount)) {
    ui.alert('Error: Please enter a valid number for the amount.');
    return;
  }

  var responseRemark = ui.prompt('Enter Remark / Vendor Name:');
  if (responseRemark.getSelectedButton() == ui.Button.CLOSE) return;
  var remark = responseRemark.getResponseText();

  // --- 30-Day Budget Check ---
  var data = sheet.getDataRange().getValues();
  var total30DaySpending = 0;
  var today = new Date();
  var thirtyDaysAgo = new Date();
  thirtyDaysAgo.setDate(today.getDate() - 30);

  for (var i = 1; i < data.length; i++) {
    var rowDateValue = data[i][0];
    var rowAmount = parseFloat(data[i][3]);
    if (rowDateValue && !isNaN(rowAmount)) {
      var rowDate = (rowDateValue instanceof Date) ? rowDateValue : new Date(rowDateValue);
      if (rowDate >= thirtyDaysAgo && rowDate <= today) {
        total30DaySpending += rowAmount;
      }
    }
  }

  // Budget threshold: ₹4,000
  if (total30DaySpending + amount > 4000) {
    if (ui.alert('Budget Alert', 'Adding this will exceed ₹4,000 for the last 30 days. Proceed?', ui.ButtonSet.YES_NO) == ui.Button.NO) return;
  }

  sheet.appendRow([new Date(), category, "", amount, remark]);
  ui.alert('Expense Added Successfully!');
}


// --- DELETE LAST ROW ---
function deleteLastRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRow(lastRow);
    SpreadsheetApp.getUi().alert('Last entry has been deleted.');
  }
}


// --- AUTOMATED PARSING via HTTP POST (Automate App → Webhook) ---
// Triggered by the Automate flow whenever a Kotak Bank notification arrives.
// The flow captures the notification text and sends it here as JSON.
// Expected payload: { "message": "Rs.290 debited from your account to Butter Chicken on 27-Jun-2026" }
function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = JSON.parse(e.postData.contents);
  var message = data.message;

  // Extract amount: matches "Rs.290" or "Rs.1,500" etc.
  var amountMatch = message.match(/Rs\.([\d,]+\.?\d*)/i);
  // Extract payee: text between "to" and "on"
  var payeeMatch = message.match(/to\s+(.*?)\s+on/i);

  var amount = amountMatch ? parseFloat(amountMatch[1].replace(/,/g, '')) : 0;
  var payee = payeeMatch ? payeeMatch[1].trim() : "Unknown";

  // Column layout: [A: Date | B: Category | C: Blank | D: Amount | E: Remark]
  sheet.appendRow([new Date(), "Automated GPay", "", amount, payee + " (Auto)"]);

  // Check budget after auto-entry too
  checkBudgetAndAlert(sheet);

  return ContentService.createTextOutput("Success");
}


// --- BUDGET ALERT VIA EMAIL (called after auto-entry) ---
function checkBudgetAndAlert(sheet) {
  var data = sheet.getDataRange().getValues();
  var total = 0;
  var today = new Date();
  var thirtyDaysAgo = new Date();
  thirtyDaysAgo.setDate(today.getDate() - 30);

  for (var i = 1; i < data.length; i++) {
    var rowDate = data[i][0] instanceof Date ? data[i][0] : new Date(data[i][0]);
    var rowAmount = parseFloat(data[i][3]) || 0;
    if (rowDate >= thirtyDaysAgo && rowDate <= today) {
      total += rowAmount;
    }
  }

  if (total > 4000) {
    MailApp.sendEmail(
      "fardeenkhan2072@gmail.com",
      "Budget Alert: ₹4,000 Limit Exceeded",
      "Your spending in the last 30 days has crossed ₹4,000.\n\n" +
      "Total spent: ₹" + total.toFixed(2) + "\n\n" +
      "Check your Google Sheet for the full breakdown."
    );
  }
}


// --- MONTHLY REPORT (trigger via Time-based trigger in Apps Script) ---
function sendMonthlyReport() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var totalSpent = 0;
  var categoryMap = {};

  for (var i = 1; i < data.length; i++) {
    var cat = data[i][1] || "Uncategorized";
    var amt = parseFloat(data[i][3]) || 0;
    totalSpent += amt;
    categoryMap[cat] = (categoryMap[cat] || 0) + amt;
  }

  // Build category breakdown
  var breakdown = "";
  for (var key in categoryMap) {
    breakdown += key + ": ₹" + categoryMap[key].toFixed(2) + "\n";
  }

  var email = "fardeenkhan2072@gmail.com";
  var subject = "Monthly Expense Report - " + Utilities.formatDate(new Date(), "Asia/Kolkata", "MMMM yyyy");
  var body =
    "Hi Fardeen,\n\n" +
    "Here's your expense summary for " + Utilities.formatDate(new Date(), "Asia/Kolkata", "MMMM yyyy") + ":\n\n" +
    "Total Spent: ₹" + totalSpent.toFixed(2) + "\n\n" +
    "Category Breakdown:\n" + breakdown + "\n" +
    "Open your Google Sheet for the full transaction log.\n\n" +
    "Stay on budget!";

  MailApp.sendEmail(email, subject, body);
}
