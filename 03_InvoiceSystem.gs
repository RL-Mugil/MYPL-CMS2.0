/**
 * ============================================================
 * 03_InvoiceSystem.gs - Company Invoice Generation & Management
 * ============================================================
 */

var COMPANY_INVOICE_ACCOUNT_EMAIL_ = "accounts@myipstrategy.com";
var COMPANY_INVOICE_SENDER_NAME_ = "Metayage Accounts";
var COMPANY_INVOICE_HSN_SAC_ = "998213";

function showGenerateInvoiceDialog() {
  SpreadsheetApp.getUi().alert("Use Staff Portal -> Finance -> New Invoice.");
}

function getInvoiceById_(invoiceId) {
  var lookup = String(invoiceId || "").trim();
  if (!lookup) return null;
  var invoices = getRecords_("INVOICE");
  for (var i = 0; i < invoices.length; i++) {
    var item = invoices[i];
    var visibleId = getInvoiceVisibleId_(item);
    var taxInvoiceNumber = String(item["Tax Invoice Number"] || "").trim();
    if (String(item.INVOICE_ID || "").trim() === lookup || visibleId === lookup || taxInvoiceNumber === lookup) {
      return item;
    }
  }
  return null;
}

function getClientInvoices(clientId) {
  var lookup = String(clientId || "").trim().toUpperCase();
  return getRecords_("INVOICE").filter(function(invoice) {
    return String(invoice.CLIENT_ID || "").trim().toUpperCase() === lookup;
  });
}

function generateInvoice(invoiceData) {
  return upsertCompanyInvoice_(invoiceData || {});
}

function upsertCompanyInvoice_(invoiceData) {
  return withScriptLock_(function() {
    var existing = invoiceData.INVOICE_ID ? getInvoiceById_(invoiceData.INVOICE_ID) : null;
    var client = resolveInvoiceClient_(invoiceData, existing);
    if (!client) return { success: false, error: "Client not found." };

    var caseRecord = resolveInvoiceCase_(invoiceData, existing);
    var built = buildCompanyInvoiceRecord_(invoiceData, existing, client, caseRecord);
    built.record.INVOICE_PDF_LINK = createInvoicePDF_(built.record, client, caseRecord);

    writeInvoiceRecord_(built.record, existing);

    logActivity_(
      existing ? "UPDATE_INVOICE" : "CREATE_INVOICE",
      "INVOICE",
      built.record.INVOICE_ID,
      "Invoice " + getInvoiceVisibleId_(built.record) + " saved for " + (built.record.CLIENT_NAME || built.record["Name of Client"] || "")
    );

    return {
      success: true,
      invoiceId: built.record.INVOICE_ID,
      invoiceNumber: getInvoiceVisibleId_(built.record),
      pdfLink: built.record.INVOICE_PDF_LINK,
      invoice: sanitizeDataForFrontend_(built.record),
      message: existing ? "Invoice updated successfully." : "Invoice created successfully."
    };
  });
}

function updateInvoicePayment_(invoiceId, paymentData) {
  return withScriptLock_(function() {
    var existing = getInvoiceById_(invoiceId);
    if (!existing) return { success: false, error: "Invoice not found." };

    var invoiceTotal = getInvoiceTotalAmount_(existing);
    var paidAmount = Math.max(0, parseInvoiceNumber_(paymentData && paymentData.paymentAmount));
    if (!paidAmount && String(paymentData && paymentData.status || "").trim() === "Paid") {
      paidAmount = invoiceTotal;
    }

    var updates = {
      PAYMENT_AMOUNT: paidAmount,
      "Amount due": Math.max(invoiceTotal - paidAmount, 0),
      PAYMENT_STATUS: deriveInvoicePaymentStatus_(invoiceTotal, paidAmount),
      PAYMENT_DATE: normalizeDateForSheet_(paymentData && paymentData.paymentDate),
      PAYMENT_MODE: String(paymentData && paymentData.paymentMode || existing.PAYMENT_MODE || "").trim()
    };

    var ok = updateRecordById_("INVOICE", "INVOICE_ID", existing.INVOICE_ID, updates);
    if (!ok) return { success: false, error: "Invoice not found." };

    clearPortalCaches_(["invoices", "dashboard", "dashboardSummary", "dashboardDetails"]);
    logActivity_("UPDATE_PAYMENT", "INVOICE", existing.INVOICE_ID, "Payment updated for " + getInvoiceVisibleId_(existing));
    return {
      success: true,
      invoiceId: existing.INVOICE_ID,
      paymentStatus: updates.PAYMENT_STATUS,
      paymentAmount: paidAmount,
      amountDue: updates["Amount due"],
      message: "Payment updated successfully."
    };
  });
}

function updatePaymentStatus(invoiceId, status, paymentDate, paymentMode) {
  var existing = getInvoiceById_(invoiceId);
  if (!existing) return { success: false, message: "Invoice not found" };
  var invoiceTotal = getInvoiceTotalAmount_(existing);
  var paidAmount = String(status || "") === "Paid" ? invoiceTotal : 0;
  return updateInvoicePayment_(invoiceId, {
    status: status,
    paymentAmount: paidAmount,
    paymentDate: paymentDate,
    paymentMode: paymentMode
  });
}

function sendInvoiceById_(invoiceId, manualEmail) {
  var invoice = getInvoiceById_(invoiceId);
  if (!invoice) return { success: false, error: "Invoice not found." };

  var client = resolveInvoiceClient_(invoice, invoice) || null;
  if (!invoice.INVOICE_PDF_LINK) {
    var caseRecord = resolveInvoiceCase_(invoice, invoice);
    var generatedPdfLink = createInvoicePDF_(invoice, client || {}, caseRecord);
    if (generatedPdfLink) {
      invoice.INVOICE_PDF_LINK = generatedPdfLink;
      updateRecordById_("INVOICE", "INVOICE_ID", invoice.INVOICE_ID, {
        INVOICE_PDF_LINK: generatedPdfLink
      });
    }
  }
  var recipient = String(manualEmail || "").trim();
  if (!recipient && client) recipient = String(client.EMAIL || "").trim();
  if (!recipient) {
    return {
      success: false,
      needsEmail: true,
      message: "No client email found for this invoice."
    };
  }

  sendInvoiceEmail_(invoice, client, recipient);
  logActivity_("SEND_INVOICE", "INVOICE", invoice.INVOICE_ID, "Invoice emailed to " + recipient);
  return {
    success: true,
    message: "Invoice sent successfully.",
    recipient: recipient
  };
}

function resolveInvoiceClient_(invoiceData, existing) {
  var allClients = getAllClients();
  var candidates = [
    invoiceData.CLIENT_ID,
    invoiceData.clientId,
    invoiceData.ClientCode,
    invoiceData.clientCode,
    existing && existing.CLIENT_ID,
    existing && existing.ClientCode
  ];

  for (var i = 0; i < candidates.length; i++) {
    var raw = String(candidates[i] || "").trim();
    if (!raw) continue;
    for (var j = 0; j < allClients.length; j++) {
      var client = allClients[j];
      if (
        String(client.CLIENT_ID || "").trim().toUpperCase() === raw.toUpperCase() ||
        String(client.CLIENT_CODE || "").trim().toUpperCase() === raw.toUpperCase()
      ) {
        return client;
      }
    }
  }
  return null;
}

function resolveInvoiceCase_(invoiceData, existing) {
  var lookup = firstNonEmpty_(
    invoiceData.CASE_ID,
    invoiceData.caseId,
    invoiceData["Docket Number"],
    existing && existing.CASE_ID,
    existing && existing["Docket Number"]
  );
  if (!lookup) return null;
  var cases = getAllCases();
  for (var i = 0; i < cases.length; i++) {
    if (String(cases[i].CASE_ID || "").trim() === String(lookup).trim()) return cases[i];
  }
  return null;
}

function buildCompanyInvoiceRecord_(invoiceData, existing, client, caseRecord) {
  var sheet = getSheet_("INVOICE");
  var headers = getSheetHeaders_("INVOICE");
  var internalId = existing && existing.INVOICE_ID
    ? existing.INVOICE_ID
    : generateId_("INV", sheet, Math.max(headers.indexOf("INVOICE_ID"), 0));
  var visibleId = String(firstNonEmpty_(
    invoiceData["Docket# [Invoice UIN]"],
    invoiceData.invoiceUin,
    existing && existing["Docket# [Invoice UIN]"],
    internalId
  ) || "").trim();

  if (!visibleId) visibleId = internalId;

  var caseId = String(firstNonEmpty_(
    invoiceData.CASE_ID,
    invoiceData.caseId,
    caseRecord && caseRecord.CASE_ID,
    existing && existing.CASE_ID,
    existing && existing["Docket Number"]
  ) || "").trim();

  var patentTitle = String(firstNonEmpty_(
    invoiceData["Title Identfication (Only key words, Enter in lower case}"],
    caseRecord && caseRecord.PATENT_TITLE,
    existing && existing["Title Identfication (Only key words, Enter in lower case}"],
    ""
  ) || "");

  var serviceCode1 = String(firstNonEmpty_(invoiceData["Main ServiceCode"], existing && existing["Main ServiceCode"], "") || "").trim();
  var serviceCode2 = String(firstNonEmpty_(invoiceData["Service Code 2"], existing && existing["Service Code 2"], "") || "").trim();
  var serviceCode3 = String(firstNonEmpty_(invoiceData["Service Code 3"], existing && existing["Service Code 3"], "") || "").trim();
  var description = String(firstNonEmpty_(
    invoiceData["Additional information on service (Do not repeat what is in code description)"],
    invoiceData.DESCRIPTION,
    invoiceData.description,
    existing && existing["Additional information on service (Do not repeat what is in code description)"],
    existing && existing.DESCRIPTION,
    ""
  ) || "").trim();

  var poFee = parseInvoiceNumber_(firstNonEmpty_(invoiceData.POFee, existing && existing.POFee, 0));
  var servFee = parseInvoiceNumber_(firstNonEmpty_(invoiceData.ServFee, existing && existing.ServFee, 0));
  var expenses = parseInvoiceNumber_(firstNonEmpty_(invoiceData.Expenses, existing && existing.Expenses, 0));
  var attorneyFee = parseInvoiceNumber_(firstNonEmpty_(invoiceData["Attorney Fee"], existing && existing["Attorney Fee"], 0));
  var consultantFee = parseInvoiceNumber_(firstNonEmpty_(invoiceData["Consultant fee"], existing && existing["Consultant fee"], 0));
  var referralFee = parseInvoiceNumber_(firstNonEmpty_(invoiceData["Referral fee"], existing && existing["Referral fee"], 0));
  var taxType = normalizeTaxType_(firstNonEmpty_(invoiceData.TAX_TYPE, existing && existing.TAX_TYPE, "IGST"));
  var taxBreakup = buildTaxBreakup_(servFee, taxType);
  var paymentAmount = parseInvoiceNumber_(firstNonEmpty_(invoiceData.PAYMENT_AMOUNT, existing && existing.PAYMENT_AMOUNT, 0));
  var invoiceTotal = poFee + servFee + expenses + taxBreakup.igst + taxBreakup.sgst + taxBreakup.cgst;
  var amountDue = Math.max(invoiceTotal - paymentAmount, 0);
  var paymentStatus = deriveInvoicePaymentStatus_(invoiceTotal, paymentAmount, firstNonEmpty_(invoiceData.PAYMENT_STATUS, existing && existing.PAYMENT_STATUS, ""));

  var record = {
    "Client Manager": String(firstNonEmpty_(invoiceData["Client Manager"], existing && existing["Client Manager"], "") || "").trim(),
    "Docket Number": String(firstNonEmpty_(invoiceData["Docket Number"], caseId, existing && existing["Docket Number"], "") || "").trim(),
    "Docket# [Invoice UIN]": visibleId,
    "Invoice Date": normalizeDateForSheet_(firstNonEmpty_(invoiceData["Invoice Date"], invoiceData.INVOICE_DATE, existing && existing["Invoice Date"], new Date())),
    "Tax Invoice Date": normalizeDateForSheet_(firstNonEmpty_(invoiceData["Tax Invoice Date"], existing && existing["Tax Invoice Date"], invoiceData["Invoice Date"], new Date())),
    "Tax Serial Number": String(firstNonEmpty_(invoiceData["Tax Serial Number"], existing && existing["Tax Serial Number"], "") || "").trim(),
    "Tax Invoice Number": String(firstNonEmpty_(invoiceData["Tax Invoice Number"], existing && existing["Tax Invoice Number"], visibleId) || "").trim(),
    "Client Reference": String(firstNonEmpty_(invoiceData["Client Reference"], existing && existing["Client Reference"], "") || "").trim(),
    "ClientCode": String(firstNonEmpty_(invoiceData.ClientCode, client.CLIENT_CODE, existing && existing.ClientCode, client.CLIENT_ID) || "").trim(),
    "Invention#": String(firstNonEmpty_(invoiceData["Invention#"], existing && existing["Invention#"], caseId) || "").trim(),
    "Patent Office": String(firstNonEmpty_(invoiceData["Patent Office"], caseRecord && caseRecord.COUNTRY, existing && existing["Patent Office"], "") || "").trim(),
    "First Inventor": String(firstNonEmpty_(invoiceData["First Inventor"], existing && existing["First Inventor"], "") || "").trim(),
    "Title Identfication (Only key words, Enter in lower case}": patentTitle ? String(patentTitle).toLowerCase() : "",
    "Entity Status": String(firstNonEmpty_(invoiceData["Entity Status"], existing && existing["Entity Status"], "") || "").trim(),
    "PO Application # e.g., US 15339876": String(firstNonEmpty_(invoiceData["PO Application # e.g., US 15339876"], caseRecord && caseRecord.APPLICATION_NUMBER, existing && existing["PO Application # e.g., US 15339876"], "") || "").trim(),
    "Main ServiceCode": serviceCode1,
    "Service Code 2": serviceCode2,
    "Service Code 3": serviceCode3,
    "Additional information on service (Do not repeat what is in code description)": description,
    "State of Supply": String(firstNonEmpty_(invoiceData["State of Supply"], existing && existing["State of Supply"], "") || "").trim(),
    "Name of Client": String(firstNonEmpty_(invoiceData["Name of Client"], client.CLIENT_NAME, existing && existing["Name of Client"], "") || "").trim(),
    "POFeeAck": String(firstNonEmpty_(invoiceData.POFeeAck, existing && existing.POFeeAck, "") || "").trim(),
    "POFee": poFee,
    "ServFee": servFee,
    "TAX_TYPE": taxType,
    "IGST": taxBreakup.igst,
    "SGST": taxBreakup.sgst,
    "CGST": taxBreakup.cgst,
    "Expenses": expenses,
    "InvAmount": invoiceTotal,
    "Amount due": amountDue,
    "Attorney Fee": attorneyFee,
    "Consultant fee": consultantFee,
    "Referral fee": referralFee,
    "Net Revenue": Math.max(servFee - attorneyFee - consultantFee - referralFee, 0),
    "PAYMENT_STATUS": paymentStatus,
    "PAYMENT_DATE": normalizeDateForSheet_(firstNonEmpty_(invoiceData.PAYMENT_DATE, existing && existing.PAYMENT_DATE, "")),
    "PAYMENT_MODE": String(firstNonEmpty_(invoiceData.PAYMENT_MODE, existing && existing.PAYMENT_MODE, "") || "").trim(),
    "PAYMENT_AMOUNT": paymentAmount,
    "INVOICE_PDF_LINK": String(firstNonEmpty_(existing && existing.INVOICE_PDF_LINK, invoiceData.INVOICE_PDF_LINK, "") || "").trim(),
    "CLIENT_ID": String(firstNonEmpty_(invoiceData.CLIENT_ID, client.CLIENT_ID, existing && existing.CLIENT_ID, "") || "").trim(),
    "CLIENT_NAME": String(firstNonEmpty_(invoiceData.CLIENT_NAME, client.CLIENT_NAME, existing && existing.CLIENT_NAME, "") || "").trim(),
    "CASE_ID": caseId,
    "ORG_ID": String(firstNonEmpty_(invoiceData.ORG_ID, client.ORG_ID, existing && existing.ORG_ID, "") || "").trim(),
    "NOTES": String(firstNonEmpty_(invoiceData.NOTES, invoiceData.notes, existing && existing.NOTES, "") || "").trim(),
    "INVOICE_ID": internalId
  };

  if (record.PAYMENT_STATUS === "Paid" && !record.PAYMENT_DATE) {
    record.PAYMENT_DATE = normalizeDateForSheet_(new Date());
  }

  return {
    record: record,
    client: client,
    caseRecord: caseRecord
  };
}

function writeInvoiceRecord_(record, existing) {
  var sheet = getSheet_("INVOICE");
  var headers = getSheetHeaders_("INVOICE");
  var row = headers.map(function(header) {
    return record.hasOwnProperty(header) ? record[header] : "";
  });

  if (existing && existing.INVOICE_ID) {
    var data = sheet.getDataRange().getValues();
    var idIndex = headers.indexOf("INVOICE_ID");
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idIndex] || "") === String(existing.INVOICE_ID || "")) {
        sheet.getRange(i + 1, 1, 1, headers.length).setValues([row]);
        clearRequestSheetCache_("INVOICE");
        return;
      }
    }
  }

  sheet.appendRow(row);
  clearRequestSheetCache_("INVOICE");
}

function createInvoicePDF_(invoice, client, caseRecord) {
  var invoiceDate = formatInvoiceDate_(invoice["Invoice Date"], "dd-MMM-yyyy");
  var dueAmount = parseInvoiceNumber_(invoice["Amount due"]);
  var serviceSubtotal = parseInvoiceNumber_(invoice.ServFee) + parseInvoiceNumber_(invoice.IGST) + parseInvoiceNumber_(invoice.SGST) + parseInvoiceNumber_(invoice.CGST);
  var grandTotal = parseInvoiceNumber_(invoice.InvAmount);
  var taxRowsHtml = buildPdfTaxRows_(invoice);
  var reimbursementRowsHtml = buildPdfReimbursementRows_(invoice);

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
    'body{font-family:Arial,sans-serif;font-size:10pt;color:#000;margin:0;padding:18px;}' +
    'table{width:100%;border-collapse:collapse;}td,th{border:1px solid #000;padding:4px 6px;vertical-align:top;}' +
    '.nb td,.nb th,.nb{border:none!important;}.logo{font-size:28px;font-weight:700;color:#2a9fd6;letter-spacing:.5px;}' +
    '.title{font-size:18pt;font-weight:700;text-align:center;}.right{text-align:right;}.center{text-align:center;}.bold{font-weight:700;}' +
    '.small{font-size:8.5pt;}.wording{font-size:9pt;font-weight:700;}.muted{color:#444;}.empty-box{width:96px;height:96px;border:1px dashed #777;display:flex;align-items:center;justify-content:center;font-size:8pt;color:#666;}' +
    '</style></head><body>' +
    '<table class="nb" style="margin-bottom:8px;"><tr><td width="42%" class="logo">Metayage</td><td width="58%" class="title">Sales Proforma</td></tr></table>' +
    '<table>' +
      '<tr>' +
        '<td width="36%" rowspan="2"><div class="bold" style="font-size:11pt;">METAYAGE PRIVATE LIMITED</div>207 EGPI Arcadia 32 Banaswadi Main Road<br>Jai Bharath Nagar Bengaluru<br><br><span class="bold">GSTIN/UIN:</span> 29AAICG1194L1Z4<br><span class="bold">State Name :</span> Karnataka, Code : 29</td>' +
        '<td width="12%"><span class="small">Invoice No.</span><br><span class="bold">' + escapeInvoiceHtml_(invoice["Docket# [Invoice UIN]"]) + '</span></td>' +
        '<td width="12%"><span class="small">Dated</span><br><span class="bold">' + escapeInvoiceHtml_(invoiceDate) + '</span></td>' +
        '<td width="40%"><span class="small">Invoice due (Including tds):</span><br><span class="bold" style="font-size:18pt;">' + escapeInvoiceHtml_(formatInvoiceCurrency_(dueAmount)) + '</span></td>' +
      '</tr>' +
      '<tr>' +
        '<td colspan="3"><span class="small">Other Reference:</span><br><span class="bold">' + escapeInvoiceHtml_(invoice["Tax Invoice Number"] || "") + '</span><br><span class="small">Client Reference:</span><br><span class="bold">' + escapeInvoiceHtml_(invoice["Client Reference"] || "") + '</span></td>' +
      '</tr>' +
      '<tr>' +
        '<td><div class="bold" style="font-size:11pt;">' + escapeInvoiceHtml_(invoice["Name of Client"] || client.CLIENT_NAME || "") + '</div>' +
          (client && client.ADDRESS ? escapeInvoiceHtml_(client.ADDRESS).replace(/\n/g, "<br>") + "<br>" : "") +
          (client && client.EMAIL ? escapeInvoiceHtml_(client.EMAIL) + "<br>" : "") +
          (invoice["State of Supply"] ? '<span class="bold">State of Supply:</span> ' + escapeInvoiceHtml_(invoice["State of Supply"]) : '') +
        '</td>' +
        '<td colspan="3" style="padding:0;border:none;"><table class="nb" style="width:100%;height:100%;"><tr><td style="border:1px solid #000;width:28%;" class="bold">Title:</td><td style="border:1px solid #000;">' + escapeInvoiceHtml_(invoice["Title Identfication (Only key words, Enter in lower case}"] || "") + '</td></tr><tr><td style="border:1px solid #000;" class="bold">Application No:</td><td style="border:1px solid #000;">' + escapeInvoiceHtml_(invoice["PO Application # e.g., US 15339876"] || "") + '</td></tr><tr><td style="border:1px solid #000;" class="bold">Patent Office:</td><td style="border:1px solid #000;">' + escapeInvoiceHtml_(invoice["Patent Office"] || "") + '</td></tr></table></td>' +
      '</tr>' +
    '</table>' +
    '<table style="border-top:none;"><tr><th width="5%">Sl No.</th><th width="55%" class="center">Description of Services</th><th width="10%" class="center">HSN/SAC</th><th width="10%" class="center">Rate</th><th width="20%" class="center">Amount</th></tr>' +
      '<tr style="height:120px;"><td class="right">1</td><td>' + escapeInvoiceHtml_(buildInvoiceServiceLabel_(invoice)) + '<br><span class="muted">' + escapeInvoiceHtml_(invoice["Additional information on service (Do not repeat what is in code description)"] || "") + '</span>' + taxRowsHtml.description + '</td><td class="center">' + escapeInvoiceHtml_(COMPANY_INVOICE_HSN_SAC_) + '</td><td>' + taxRowsHtml.rate + '</td><td class="right">' + escapeInvoiceHtml_(formatInvoiceMoney_(invoice.ServFee)) + taxRowsHtml.amount + '</td></tr>' +
      '<tr><td colspan="4" class="right bold">Sub Total</td><td class="right bold">' + escapeInvoiceHtml_(formatInvoiceMoney_(serviceSubtotal)) + '</td></tr></table>' +
    '<div class="small" style="margin-top:4px;">Amount Chargeable (in words)</div><div class="wording" style="margin-bottom:8px;">' + escapeInvoiceHtml_(invoiceAmountInWords_(serviceSubtotal)) + '</div>' +
    '<table style="border-top:none;"><tr><th colspan="2" class="center">Debit Note for Reimbursement of Government Fees & Stamp Charges</th></tr><tr><th width="80%" class="center">Description</th><th width="20%" class="center">Amount</th></tr>' +
      reimbursementRowsHtml +
      '<tr><td class="right bold">Grand Total</td><td class="right bold">' + escapeInvoiceHtml_(formatInvoiceCurrency_(grandTotal)) + '</td></tr></table>' +
    '<div class="small" style="margin-top:4px;">Amount Chargeable (in words)</div><div class="wording" style="margin-bottom:8px;">' + escapeInvoiceHtml_(invoiceAmountInWords_(grandTotal)) + '</div>' +
    '<table style="border-top:none;"><tr><td width="26%" style="vertical-align:top;"><div class="empty-box">UPI QR</div><div class="small bold" style="margin-top:8px;">UPI Scan to pay</div></td><td width="74%"><div class="bold">Company&#39;s Bank Details</div><table class="nb" style="margin-top:6px;"><tr><td width="24%">Bank Name:</td><td class="bold">HDFC Bank</td></tr><tr><td>A/c No.:</td><td class="bold">50200070931727</td></tr><tr><td>Branch & IFSC:</td><td class="bold">Vadavalli-2, Coimbatore & HDFC0007237</td></tr><tr><td>SWIFT Code:</td><td class="bold">HDFCINBBCMB</td></tr></table><div style="margin-top:10px;"><span class="bold">Declaration</span><br>Micro Entity# KR030204347<br>Thank you for the opportunity to be of service</div></td></tr><tr><td colspan="2" class="center small">Electronic Document. No signature needed</td></tr></table>' +
    '</body></html>';

  var htmlOutput = HtmlService.createHtmlOutput(html);
  var pdfBlob = Utilities.newBlob(htmlOutput.getContent(), "text/html", "Invoice_" + getInvoiceVisibleId_(invoice) + ".html").getAs("application/pdf");
  pdfBlob.setName("Invoice_" + getInvoiceVisibleId_(invoice) + ".pdf");

  var pdfFile;
  try {
    var clientFolder = DriveApp.getFolderById(client.CLIENT_FOLDER_ID);
    var invoiceFolders = clientFolder.getFoldersByName("INVOICES");
    pdfFile = invoiceFolders.hasNext() ? invoiceFolders.next().createFile(pdfBlob) : clientFolder.createFile(pdfBlob);
  } catch (e) {
    pdfFile = getFolder_("04_INVOICE_SYSTEM").createFile(pdfBlob);
  }

  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return pdfFile.getUrl();
}

function sendInvoiceEmail_(invoice, client, recipientEmail) {
  var subject = "Invoice " + getInvoiceVisibleId_(invoice) + " from Metayage Accounts";
  var pdfLink = invoice.INVOICE_PDF_LINK || "";
  var total = getInvoiceTotalAmount_(invoice);
  var dueAmount = parseInvoiceNumber_(invoice["Amount due"]);
  var invoiceDate = formatInvoiceDate_(invoice["Invoice Date"], "dd-MMM-yyyy");
  var plainBody =
    "Dear " + String((client && (client.CONTACT_PERSON || client.CLIENT_NAME)) || invoice["Name of Client"] || "Client") + ",\n\n" +
    "Please find your invoice details below.\n\n" +
    "Invoice: " + getInvoiceVisibleId_(invoice) + "\n" +
    "Invoice Date: " + invoiceDate + "\n" +
    "Invoice Amount: " + formatInvoiceCurrency_(total) + "\n" +
    "Amount Due: " + formatInvoiceCurrency_(dueAmount) + "\n\n" +
    (pdfLink ? ("Invoice PDF: " + pdfLink + "\n\n") : "") +
    "Regards,\nMetayage Accounts";
  var htmlBody = '<div style="font-family:Arial,sans-serif;max-width:640px;margin:0 auto;"><div style="background:#111425;color:#fff;padding:20px 24px;border-radius:12px 12px 0 0;"><div style="font-size:24px;font-weight:700;">Metayage Accounts</div><div style="opacity:.8;margin-top:6px;">Invoice ' + escapeInvoiceHtml_(getInvoiceVisibleId_(invoice)) + '</div></div><div style="padding:24px;border:1px solid #e6e8ef;border-top:none;border-radius:0 0 12px 12px;background:#fff;"><p>Dear ' + escapeInvoiceHtml_((client && (client.CONTACT_PERSON || client.CLIENT_NAME)) || invoice["Name of Client"] || "Client") + ',</p><p>Your invoice is ready for review.</p><table style="width:100%;border-collapse:collapse;background:#f7f8fc;border-radius:10px;overflow:hidden;"><tr><td style="padding:10px 14px;font-weight:700;">Invoice No.</td><td style="padding:10px 14px;">' + escapeInvoiceHtml_(getInvoiceVisibleId_(invoice)) + '</td></tr><tr><td style="padding:10px 14px;font-weight:700;">Invoice Date</td><td style="padding:10px 14px;">' + escapeInvoiceHtml_(invoiceDate) + '</td></tr><tr><td style="padding:10px 14px;font-weight:700;">Invoice Amount</td><td style="padding:10px 14px;">' + escapeInvoiceHtml_(formatInvoiceCurrency_(total)) + '</td></tr><tr><td style="padding:10px 14px;font-weight:700;">Amount Due</td><td style="padding:10px 14px;">' + escapeInvoiceHtml_(formatInvoiceCurrency_(dueAmount)) + '</td></tr></table>' + (pdfLink ? '<p style="margin-top:18px;"><a href="' + escapeInvoiceHtml_(pdfLink) + '" style="display:inline-block;padding:10px 18px;background:#6366f1;color:#fff;border-radius:999px;text-decoration:none;">Open Invoice PDF</a></p>' : '') + '<p style="margin-top:20px;color:#59607a;">Reply to this email for billing queries.</p></div></div>';

  var options = { htmlBody: htmlBody, replyTo: COMPANY_INVOICE_ACCOUNT_EMAIL_, name: COMPANY_INVOICE_SENDER_NAME_ };
  var aliases = [];
  try { aliases = GmailApp.getAliases(); } catch (e) {}
  if (aliases.indexOf(COMPANY_INVOICE_ACCOUNT_EMAIL_) > -1) options.from = COMPANY_INVOICE_ACCOUNT_EMAIL_;

  var attachment = getInvoicePdfBlob_(pdfLink);
  if (attachment) options.attachments = [attachment];

  GmailApp.sendEmail(recipientEmail, subject, plainBody, options);
}

function getInvoicePdfBlob_(pdfLink) {
  var fileId = extractDriveFileId_(pdfLink);
  if (!fileId) return null;
  try { return DriveApp.getFileById(fileId).getBlob(); } catch (e) { return null; }
}

function extractDriveFileId_(link) {
  var match = String(link || "").match(/\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : "";
}

function buildTaxBreakup_(servFee, taxType) {
  var amount = parseInvoiceNumber_(servFee);
  if (normalizeTaxType_(taxType) === "IGST") {
    return { igst: roundInvoiceNumber_(amount * 0.18), sgst: 0, cgst: 0 };
  }
  return { igst: 0, sgst: roundInvoiceNumber_(amount * 0.09), cgst: roundInvoiceNumber_(amount * 0.09) };
}

function deriveInvoicePaymentStatus_(invoiceTotal, paymentAmount, explicitStatus) {
  var total = parseInvoiceNumber_(invoiceTotal);
  var paid = parseInvoiceNumber_(paymentAmount);
  if (paid >= total && total > 0) return "Paid";
  if (paid > 0) return "Partially Paid";
  if (String(explicitStatus || "").trim() === "Deleted") return "Deleted";
  return "Unpaid";
}

function normalizeTaxType_(value) {
  return String(value || "").toUpperCase().indexOf("IGST") > -1 ? "IGST" : "CGST/SGST";
}

function parseInvoiceNumber_(value) {
  if (typeof value === "number") return isNaN(value) ? 0 : value;
  var normalized = String(value || "").replace(/,/g, "").replace(/[^\d.-]/g, "").trim();
  var parsed = parseFloat(normalized);
  return isNaN(parsed) ? 0 : parsed;
}

function roundInvoiceNumber_(value) {
  return Math.round(parseInvoiceNumber_(value) * 100) / 100;
}

function normalizeDateForSheet_(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) return value;
  var raw = String(value || "").trim();
  if (!raw) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return new Date(raw + "T00:00:00");
  var parsed = new Date(raw);
  return isNaN(parsed.getTime()) ? raw : parsed;
}

function formatInvoiceDate_(value, pattern) {
  var normalized = normalizeDateForSheet_(value);
  if (!normalized || typeof normalized === "string") return String(value || "-");
  return Utilities.formatDate(normalized, Session.getScriptTimeZone(), pattern || "dd-MMM-yyyy");
}

function formatInvoiceMoney_(value) {
  return formatInvoiceCurrency_(parseInvoiceNumber_(value)).replace(/^₹/, "");
}

function formatInvoiceCurrency_(value) {
  return "₹" + parseInvoiceNumber_(value).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function getInvoiceTotalAmount_(invoice) {
  return parseInvoiceNumber_(firstNonEmpty_(invoice.InvAmount, invoice.TOTAL, 0));
}

function getInvoiceVisibleId_(invoice) {
  return String(firstNonEmpty_(invoice["Docket# [Invoice UIN]"], invoice["Tax Invoice Number"], invoice.INVOICE_ID, "") || "").trim();
}

function buildInvoiceServiceLabel_(invoice) {
  return [invoice["Main ServiceCode"], invoice["Service Code 2"], invoice["Service Code 3"]]
    .filter(function(part) { return !!String(part || "").trim(); })
    .join(" / ");
}

function buildPdfTaxRows_(invoice) {
  var parts = [];
  if (parseInvoiceNumber_(invoice.IGST) > 0) parts.push({ label: "IGST", rate: "18%", amount: formatInvoiceMoney_(invoice.IGST) });
  if (parseInvoiceNumber_(invoice.CGST) > 0) parts.push({ label: "CGST", rate: "9%", amount: formatInvoiceMoney_(invoice.CGST) });
  if (parseInvoiceNumber_(invoice.SGST) > 0) parts.push({ label: "SGST", rate: "9%", amount: formatInvoiceMoney_(invoice.SGST) });
  if (!parts.length) return { description: "", rate: "", amount: "" };
  return {
    description: '<div style="margin-top:18px;text-align:right;">' + parts.map(function(part) { return '<div class="bold">' + escapeInvoiceHtml_(part.label) + "</div>"; }).join("") + "</div>",
    rate: '<div style="margin-top:18px;text-align:right;">' + parts.map(function(part) { return "<div>" + escapeInvoiceHtml_(part.rate) + "</div>"; }).join("") + "</div>",
    amount: '<div style="margin-top:18px;">' + parts.map(function(part) { return "<div>" + escapeInvoiceHtml_(part.amount) + "</div>"; }).join("") + "</div>"
  };
}

function buildPdfReimbursementRows_(invoice) {
  var rows = [];
  if (parseInvoiceNumber_(invoice.POFee) > 0) rows.push({ description: "Professional / Patent Office Fee", amount: formatInvoiceMoney_(invoice.POFee) });
  if (parseInvoiceNumber_(invoice.Expenses) > 0) rows.push({ description: "Reimbursable expenses", amount: formatInvoiceMoney_(invoice.Expenses) });
  if (!rows.length) rows.push({ description: "", amount: "" });
  var html = rows.map(function(row) {
    return '<tr style="height:32px;"><td>' + escapeInvoiceHtml_(row.description) + '</td><td class="right">' + escapeInvoiceHtml_(row.amount) + "</td></tr>";
  }).join("");
  var subtotal = parseInvoiceNumber_(invoice.POFee) + parseInvoiceNumber_(invoice.Expenses);
  html += '<tr><td class="right bold">Sub Total</td><td class="right bold">' + escapeInvoiceHtml_(formatInvoiceMoney_(subtotal)) + "</td></tr>";
  return html;
}

function invoiceAmountInWords_(amount) {
  var value = roundInvoiceNumber_(amount);
  var rupees = Math.floor(value);
  var paise = Math.round((value - rupees) * 100);
  var out = numberToIndianWords_(rupees) + " INR";
  if (paise > 0) out += " and " + numberToIndianWords_(paise) + " Paise";
  return out;
}

function numberToIndianWords_(num) {
  num = Math.floor(parseInvoiceNumber_(num));
  if (num === 0) return "Zero";

  var ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"];
  var tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"];

  function twoDigits(n) {
    if (n < 20) return ones[n];
    return tens[Math.floor(n / 10)] + (n % 10 ? " " + ones[n % 10] : "");
  }

  function threeDigits(n) {
    var hundred = Math.floor(n / 100);
    var rest = n % 100;
    var text = hundred ? ones[hundred] + " Hundred" : "";
    if (rest) text += (text ? " " : "") + twoDigits(rest);
    return text;
  }

  var parts = [];
  var crore = Math.floor(num / 10000000);
  num = num % 10000000;
  var lakh = Math.floor(num / 100000);
  num = num % 100000;
  var thousand = Math.floor(num / 1000);
  num = num % 1000;
  var remainder = num;

  if (crore) parts.push(twoDigits(crore) + " Crore");
  if (lakh) parts.push(twoDigits(lakh) + " Lakh");
  if (thousand) parts.push(twoDigits(thousand) + " Thousand");
  if (remainder) parts.push(threeDigits(remainder));
  return parts.join(" ").trim();
}

function escapeInvoiceHtml_(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function firstNonEmpty_() {
  for (var i = 0; i < arguments.length; i++) {
    var value = arguments[i];
    if (value === 0) return value;
    if (value !== null && value !== undefined && String(value).trim() !== "") return value;
  }
  return "";
}
