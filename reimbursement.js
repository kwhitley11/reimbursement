//function checks bottom right cell for value, if null then column 12 is set to pending
function onFormSubmit(e) {
  var range = e.range;
  Logger.log('range is: ' + range);
  var column = range.getColumn();
  Logger.log('Column is: ' + column);
  var row = range.getRow();
  Logger.log('Row is: ' + row);

  var sheet = SpreadsheetApp.getActiveSheet();
  var lastCellRange = sheet.getRange(row, 12);
  var lastCell = lastCellRange.getValues();
  var emailRange = sheet.getRange(row, 7);
  var email = emailRange.getValues().toString();
  Logger.log('Email address is: ' + email)

  Logger.log(sheet.getName());
  Logger.log('The last cell value is ' + lastCell);
  if(lastCell) {
    sheet.getRange(row, 12).setValue('Pending');
    let message = 'Your reimbursement request has been received. Please wait for approval before moving forward with any purchases or reimbursements.';
    var cc = 'kyle.p.whitley@vanderbilt.edu';
    var subject = 'Your SyBBURE Reimbursement Request';
    GmailApp.sendEmail(email + ', ' + cc, subject, message);
  };
};

function onEdit(e) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = e.range;
  Logger.log('range is: ' + range);
  var column = range.getColumn();
  Logger.log('Column is: ' + column);
  var row = range.getRow();
  Logger.log('Row is: ' + row);


  var sheet = SpreadsheetApp.getActiveSheet();
  Logger.log('The last cell is: ' + lastCell);
  var idRange = sheet.getRange(row, 8);
  var id = idRange.getValues().toString().toLowerCase();
  Logger.log('Id is: ' + id);
  var emailRange = sheet.getRange(row, 7);
  var email = emailRange.getValues();
  Logger.log('Email is: ' + email);
  var nameRange = sheet.getRange(row, 2);
  var name = nameRange.getValues();
  Logger.log('Name is: ' + name);
  var amountRange = sheet.getRange(row, 5);
  var amount = parseFloat(amountRange.getValues());
  Logger.log('Amount requested is: ' + amount);
  var categoryRange = sheet.getRange(row, 6);
  var category = categoryRange.getValues();
  Logger.log('Spending category is: ' + category);
  var termRange = sheet.getRange(row, 3);
  var yearRange = sheet.getRange(row, 4);
  var term = termRange.getValues() + ' ' + yearRange.getValues();
  Logger.log('Term is: ' + term);
  var typeRange = sheet.getRange(row, 9);
  var type = typeRange.getValues();
  Logger.log('Expense type is: ' + type);
  var roleRange = sheet.getRange(row, 10);
  var role = roleRange.getValues();
  Logger.log('Role is: ' + role);
  var lastCellRange = sheet.getRange(row, 13);
  var lastCell = lastCellRange.getValue();
  Logger.log('Last cell is: ' + lastCellRange);


    //selects which tab based on whether or not expense type selection is = conference related
  if (type == 'Conference Related') {
     var allowanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Conference allowances');
     var disbursementSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Conference disbursements');
  } else {
     var allowanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Other allowances');
     var disbursementSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Other disbursements');
  } 

    //finds the column based on term 
  var allowanceSheetColumn = 0;
  var allowanceSheetColumnValue = '';
  while (allowanceSheetColumnValue !== term) {
    allowanceSheetColumn++;
    allowanceSheetColumnValue = allowanceSheet.getRange(1, allowanceSheetColumn).getValue();
  }
  Logger.log('Allowance sheet value is: ' + allowanceSheetColumnValue);
  
  // finds the row based on id
  var allowanceSheetRow = 0;
  var allowanceSheetRowValue = '';
  while (allowanceSheetRowValue != id && allowanceSheetRow !== allowanceSheet.getLastRow()) {
    allowanceSheetRow++;
    allowanceSheetRowValue = allowanceSheet.getRange(allowanceSheetRow, 1).getValue();
  } 
  Logger.log('Allowance sheet row value: ' + allowanceSheetRowValue);


// if 'approved' = defined value = executes 
  if(lastCell == 'Denied') {
    let message = 'Your reimbursement has been denied. Please reach out to Kyle for further instructions.'
    GmailApp.sendEmail(email, 'SyBBURE Reimbursement Request: Rejected', message);
    sheet.getRange(row, 12).setValue('Denied');
  } 
  var allowance = parseFloat(allowanceSheet.getRange(allowanceSheetRow, allowanceSheetColumn).getValue());
  Logger.log(allowance);
  if (allowance === 0) {
    sheet.getRange(row, 12).setValue('Not Eligible');
    var approval = 'has been DENIED';
    var message = 'You are currently not eligible for reimbursement for ' + term + '. Do not submit to Oracle.';
    }
  var previousTotal = parseFloat(disbursementSheet.getRange(allowanceSheetRow, allowanceSheetColumn).getValue());
  Logger.log('Previous total is: ' + previousTotal);
  var newTotal = previousTotal + amount;
  Logger.log('New total is: ' + parseFloat(newTotal));
  if (lastCell == 'Approved') {
    var message = 'This approval email should be attached to your expense report in Oracle Cloud, along with your receipts. ' +
    'Note that the expense report reviewers may request additional information regarding your purchase in order to approve the ' + 
    'reimbursement on their end.';     
    sheet.getRange(row, 12).setValue('Approved');
    disbursementSheet.getRange(allowanceSheetRow, allowanceSheetColumn).setValue(newTotal);
     var newMessage = message + ' You have requested a total of $' + newTotal.toString() + ' out of $' + allowance.toString() +
     ' for ' + term + '.';
     Logger.log(newMessage);
  

  if (role == 'Undergraduate Student') {
    instructionMessage = 'You are a student so follow these guidelines on how to fill out the boxes in an expense report for the categories below: <br /> For all categories, the COA should auto populate as you fill in the first few boxes <br /> COA: 270.05.27770.xxxx.076.000.000.0.0 <br /> (if you are using the POET GE_101577, then you must use 40 as the NAC i.e. 270.40.27770.xxxx.076.000.000.0.0) with the xxxx above changing based on the category based on the category. You may need to change some of the other numbers to match the above string <br /> <br /> For these categories: <br /> <br /> Conferences <br /> Fees <br /> Food and non-alchoholic beverage <br /> Lab Supplies <br /> Memberships and Dues <br /> <br /> Fill in this info: <br /> <br /> Project: GE_101577 <br /> Task: 1 <br /> Expenditure Org: 27770 - Research Operations <br /> Vanderbilt Property (may not always show up): Yes <br /> <br /> --------------------------------- <br /> <br /> For these categories: <br /> <br /> Books/Periodicals/Magazines <br /> Computer Software <br /> <br /> Fill in this info: <br /> <br /> Project: (leave blank) <br /> Task: (leave blank) <br /> Expenditure Org: 27770 - Research Operations <br /> <br /> --------------------------------- <br /> <br /> For these categories: <br /> <br /> Educational Supplies <br /> Instructional Supplies <br /> <br /> Follow these instructions: <br /> <br /> 50/50 split: Use Itemization to split charges. Create two identical itemizations at the bottom of the page. The only difference should be that in one the “project” and “task” boxes will be left blank. In the other, the “project” is ‘GE_101577’ and the “task” is ‘1’. The “expense organization” box should contain “27770 – Research Operations” in both itemizations. ' 
  }
  if (role == 'Team Sybbure') {
    instructionMessage = 'You are on Team Sybbure so follow these guidelines on how to fill out the boxes in an expense report for the categories below: <br /> For all categories, the COA should auto populate as you fill in the first few boxes <br /> COA: 270.05.27770.xxxx.076.000.000.0.0 <br /> (if you are using the POET GE_101577, then you must use 40 as the NAC i.e. 270.40.27770.xxxx.076.000.000.0.0) with the xxxx above changing based on the category based on the category. You may need to change some of the other numbers to match the above string <br /> <br /> For these categories: <br /> <br /> Conferences <br /> Fees <br /> Non-Capital Equipment (Related to research) <br /> Lab Supplies <br /> Memberships and Dues <br /> Subject Participation <br /> <br /> Fill in this info: <br /> <br /> Project: GE_101577 <br /> Task: 1 <br /> Expenditure Org: 27770 - Research Operations <br /> Vanderbilt Property (may not always show up): Yes <br /> <br /> --------------------------------- <br /> <br /> For these categories: <br /> <br /> Books/Periodicals/Magazines <br /> Advertising <br /> Office Supplies <br /> Non-Capital Equipment (Related to administrative work) <br /> Computer Software <br /> <br /> Fill in this info: <br /> <br /> Project: (leave blank) <br /> Task: (leave blank) <br /> Expenditure Org: 27770 - Research Operations <br /> <br /> --------------------------------- <br /> <br /> For these categories: <br /> <br /> Food and non-alchoholic beverage <br /> Catering <br /> Educational Supplies <br /> Instructional Supplies <br /> <br /> Follow these instructions: <br /> <br /> 50/50 split: Use Itemization to split charges. Create two identical itemizations at the bottom of the page. The only difference should be that in one the “project” and “task” boxes will be left blank. In the other, the “project” is ‘GE_101577’ and the “task” is ‘1’. The “expense organization” box should contain “27770 – Research Operations” in both itemizations. ' 
  }
  var remainingBudget = allowance - newTotal;
  var approvalMessage = 'Your SyBBURE ' + type.toString().toLowerCase() + ' reimbursement request for ' + term + ' has been approved for $' +
    amount.toString() + ' for ' + category.toString().toLowerCase() + '. Your remaining budget for the semester is: ' +  remainingBudget + '.';

  Logger.log('Approval message is: ' + approvalMessage);
  Logger.log('Message is: ' + message);
  var html = '<body><p>' + name + ' (' + id + '):</p><p>' + approvalMessage + '</p><p>' + message + '</p><p>' + instructionMessage + '</p><p>- Jonathan Ehrman</p></body>';
  var options = {};
  options.htmlBody = html;
  var cc = 'kyle.p.whitley@vanderbilt.edu'
  GmailApp.sendEmail(email + ', ' + cc, 'SyBBURE Reimbursement Request: Approved', approvalMessage, options);

  }

}