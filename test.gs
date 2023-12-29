const sheet = SpreadsheetApp.getActiveSheet();
const row = sheet.getActiveCell().getRow();
const col = sheet.getRange(row, 1, 1, 17).getValues()[0];

const customer = {
  timestamp: col[0],
  email: col[1],
  name: col[2],
  whatsapp: col[3],
  ticketType: col[4],
  matric: col[5],
  inasis: col[6],
  price: col[12],
  description: col[13],
  id: col[14],
  status: col[15],
  row: row
}

function onOpen() {
  let ui = SpreadsheetApp.getUi()
  ui.createMenu('Approve').addItem('Send Email', 'approve').addToUi()
  ui.createMenu('Reject').addItem('Reject User', 'reject').addToUi()
  ui.createMenu('Re-Send'.addItem('Re-SendEmail', 'resendEmail')).addToUi();
}

function onFormSubmit(e) {
  var responses = e.namedValues;
  const ss = SpreadsheetApp.openById('11DG4LQcQfF1agXcONE6tzSN8bLqa68raKKNH8B6xxNs')
  const sheet = ss.getSheetByName('Form Responses 1')
  setStatus('SUBMITTED', 'grey')
  var price, discount, totalPrice, description;
  if (responses['Ticket Type'][0].trim() == 'Jakarta') {
    price = 55;
    description = 'Regular Seat, Wristband Ticket, Snack & Drink Hamper, Lightstick, Merchandise';
  } else if (responses['Ticket Type'][0].trim() == 'Batavia') {
    price = 45;
    description = 'Wristband, Regular Seat, Snack, and Drink Hamper';
  }

  var type = responses['Ticket Type'][0].trim();
  var ctype = type.substring(0, 2);
  var name = responses['Name'][0].trim();
  var cname = name.substring(0, 1);
  var whatsNum = responses['Whatsapp Number'][0].trim();
  var cnum = whatsNum.slice(-3);
  var email = responses['Email'][0].trim();
  var cmail = email.substring(0, 2);
  var id = 'PRI2023-' + ctype + cname + cnum + cmail + row

  var newCustomer = {
    timestamp: responses['Timestamp'][0].trim(),
    email: responses['Email'][0].trim(),
    name: responses['Name'][0].trim(),
    whatsapp: responses['Whatsapp Number'][0].trim(),
    ticketType: responses['Ticket Type'][0].trim(),
    matric: responses['Matric Ticket Holder (UUM Student Only)'][0].trim(),
    inasis: responses['Inasis (UUM Student Only)'][0].trim(),
    description: description,
    id: id,
    status: 'IN REVIEW',
    price: price,
    row: row
  }
  setNewCustomer(newCustomer)
  sendNotificationEmail(newCustomer)
}

function setNewCustomer(newCustomer) {
  const lastCol = sheet.getLastColumn()
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
  const priceIndex = headers.indexOf('Price') + 1
  sheet.getRange(newCustomer.row, priceIndex).setValue(newCustomer.price)
  const descriptionIndex = headers.indexOf('Description') + 1
  sheet.getRange(newCustomer.row, descriptionIndex).setValue(newCustomer.description)
  const idIndex = headers.indexOf('ID') + 1
  sheet.getRange(newCustomer.row, idIndex).setValue(newCustomer.id)
  const statusIndex = headers.indexOf('Status') + 1
  sheet.getRange(newCustomer.row, statusIndex).setValue(newCustomer.status)
  sheet.getRange(newCustomer.row, 1, 1, statusIndex).setBackground('yellow')
}

function sendNotificationEmail(newCustomer) {
  var pdfFolder = DriveApp.getFolderById('1bYabYitJzv0Ylu1KceP6VhBZgcemAcU3');
  var eTicketTemplateFolder = DriveApp.getFolderById('1bYabYitJzv0Ylu1KceP6VhBZgcemAcU3');
  var eTicketTemplateFile = DriveApp.getFileById('1UNTVl1br5afH1b0DtVV5Tbc-GOaIwBbNdvCXTTV3duo');
  var newETicketTemplate = eTicketTemplateFile.makeCopy(eTicketTemplateFolder);
  var openDoc = DocumentApp.openById(newETicketTemplate.getId());
  var body = openDoc.getBody();
  var header = openDoc.getHeader();
  header.replaceText("{ticketID}", newCustomer.id);
  body.replaceText("{name}", newCustomer.name);
  body.replaceText("{residential}", newCustomer.inasis);
  body.replaceText("{phoneNumber}", newCustomer.whatsapp);
  body.replaceText("{email}", newCustomer.email);
  body.replaceText("{description}", newCustomer.description);
  body.replaceText("{ticketType}", newCustomer.ticketType);
  body.replaceText("{ticketPrice}", newCustomer.price);
  openDoc.saveAndClose();
  var blobPDF = newETicketTemplate.getAs(MimeType.PDF);

  var pdfFile = pdfFolder.createFile(blobPDF).setName('PRI 2023 Receipt');
  eTicketTemplateFolder.removeFile(newETicketTemplate);
  let attachments = []
  let template = HtmlService.createTemplateFromFile('index')
  let subject = 'PRI2023 Booking'
  // createTicket(attachments)
  template.newCustomer = newCustomer
  let message = template.evaluate().getContent()
  MailApp.sendEmail({
    name: 'PRI Committee',
    to: newCustomer.email,
    subject: subject,
    htmlBody: message,
    attachments: pdfFile
  })
}

function reject() {
  let ui = SpreadsheetApp.getUi()
  let confirm = ui.alert(`REJECT ${customer.name} Registration?`, ui.ButtonSet.YES_NO)
  if (confirm == ui.Button.YES) {
    sendDecHTMLEmail();
    setStatus('REJECTED', 'red')
  }
}

function approve() {
  let ui = SpreadsheetApp.getUi()
  let confirm = ui.alert(`APPROVE ${customer.name} Registration?`, ui.ButtonSet.YES_NO)
  if (confirm == ui.Button.YES) {
    sendAccHTMLEmail(customer)
    setStatus('APPROVED', '#77c7a4')
  }
}

function sendAccHTMLEmail(customer) {
  var pdfFolder = DriveApp.getFolderById('1bYabYitJzv0Ylu1KceP6VhBZgcemAcU3');
  var eTicketTemplateFolder = DriveApp.getFolderById('1bYabYitJzv0Ylu1KceP6VhBZgcemAcU3');
  var eTicketTemplateFile = DriveApp.getFileById('1UNTVl1br5afH1b0DtVV5Tbc-GOaIwBbNdvCXTTV3duo');
  var newETicketTemplate = eTicketTemplateFile.makeCopy(eTicketTemplateFolder);
  var openDoc = DocumentApp.openById(newETicketTemplate.getId());
  var body = openDoc.getBody();
  var header = openDoc.getHeader();
  header.replaceText("{ticketID}", customer.id);
  body.replaceText("{name}", customer.name);
  body.replaceText("{residential}", customer.inasis);
  body.replaceText("{phoneNumber}", customer.whatsapp);
  body.replaceText("{email}", customer.email);
  body.replaceText("{description}", customer.description);
  body.replaceText("{ticketType}", customer.ticketType);
  body.replaceText("{ticketPrice}", customer.price);
  openDoc.saveAndClose();
  var blobPDF = newETicketTemplate.getAs(MimeType.PDF);

  var pdfFile = pdfFolder.createFile(blobPDF).setName('PRI 2023 Receipt');
  eTicketTemplateFolder.removeFile(newETicketTemplate);
  let attachments = []
  let template = HtmlService.createTemplateFromFile('confirm')
  let subject = 'PRI2023 E-Ticket'
  // createTicket(attachments)
  template.customer = customer
  let message = template.evaluate().getContent()
  MailApp.sendEmail({
    name: 'PRI2023 Committee',
    to: customer.email,
    subject: subject,
    htmlBody: message,
    attachments: pdfFile
  })
}

function sendDecHTMLEmail() {
  let attachments = []
  let template = HtmlService.createTemplateFromFile('decline')
  let subject = 'Issue upon your PRI 2023 E-Ticket'
  // createTicket(attachments)
  template.customer = customer
  let message = template.evaluate().getContent()
  MailApp.sendEmail({
    name: 'PRI2023 Committee',
    to: customer.email,
    subject: subject,
    htmlBody: message,
    // attachments: attachments
  })
}

function setStatus(status, color) {
  const lastCol = sheet.getLastColumn()
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
  const statusIndex = headers.indexOf('Status') + 1
  sheet.getRange(customer.row, statusIndex).setValue(status)
  sheet.getRange(customer.row, 1, 1, statusIndex).setBackground(color)
}

function checkEmailQuota() {
  // var service = Gmail.Users.getProfile("me");
  // Logger.log('Emails sent today: ' + service.messagesTotal);
  // Logger.log('Emails remaining today: ' + (service.messagesMax - service.messagesTotal));
  // console.log(MailApp.getRemainingDailyQuota());
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  Logger.log("Remaining email quota: " + emailQuotaRemaining);
}

function resendEmail() {
  
}