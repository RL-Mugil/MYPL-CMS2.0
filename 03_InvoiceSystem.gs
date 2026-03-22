/**
 * ============================================================
 * 03_InvoiceSystem.gs — Invoice Generation & Management
 * ============================================================
 */

/**
 * Generate an invoice
 */
function generateInvoice(invoiceData) {
  return withScriptLock_(function() {
  var sheet = getSheet_("INVOICE");
  // Use manual invoiceId if provided, fallback to auto-generated
  var invoiceId = invoiceData.invoiceId ? invoiceData.invoiceId.trim() : generateId_("INV", sheet, 0);

  // Get client info
  Logger.log("generateInvoice CALLED WITH invoiceData: " + JSON.stringify(invoiceData));
  var searchId = invoiceData.clientId ? invoiceData.clientId.toString().trim() : "";
  Logger.log("generateInvoice LOOKING UP client: [" + searchId + "]");
  var client = getClientById(searchId);
  if (!client) {
    Logger.log("generateInvoice FAILED: Client not found for ID [" + searchId + "]");
    return { success: false, message: "Client not found: " + searchId };
  }

  var amount = parseFloat(invoiceData.amount) || 0;
  var gstRate = parseFloat(invoiceData.gstRate) || 18;
  var gstAmount = amount * (gstRate / 100);
  var total = amount + gstAmount;
  var dueDate = new Date();
  var govtFees = parseFloat(invoiceData.govtFees) || 0;
  dueDate.setDate(dueDate.getDate() + 30);

  // Create Invoice PDF
  var pdfLink = "";
  try {
    pdfLink = createInvoicePDF_(invoiceId, client, invoiceData, amount, gstRate, gstAmount, total, dueDate);
  } catch (e) {
    Logger.log("PDF generation error: " + e.message);
  }

  var notes = invoiceData.notes || "";
  if (govtFees > 0) {
    notes = (notes ? notes + " | " : "") + "Govt Fees: " + govtFees.toFixed(2);
  }

  // Add to sheet
  sheet.appendRow([
    invoiceId,
    invoiceData.clientId,
    client.CLIENT_NAME,
    invoiceData.caseId || "",
    new Date(),
    dueDate,
    invoiceData.serviceType || "Professional Fees",
    invoiceData.description || "",
    amount,
    gstRate,
    gstAmount,
    total,
    "Unpaid",
    "",
    "",
    pdfLink,
    client.ORG_ID || invoiceData.orgId || "",
    notes
  ]);

  // Email invoice to client
  if (client.EMAIL && pdfLink) {
    try {
      sendInvoiceEmail_(client, invoiceId, total, dueDate, pdfLink);
    } catch (e) {
      Logger.log("Email error: " + e.message);
    }
  }

  logActivity_("CREATE_INVOICE", "INVOICE", invoiceId, "Amount: " + total + " for " + client.CLIENT_NAME);

  return {
    success: true,
    invoiceId: invoiceId,
    total: total,
    message: "Invoice " + invoiceId + " generated. Total: ₹" + total.toFixed(2)
  };
  });
}

/**
 * Create invoice PDF using HTML template approach matching Metayage design
 */
function createInvoicePDF_(invoiceId, client, data, amount, gstRate, gstAmount, total, dueDate) {
  var govtFees = parseFloat(data.govtFees) || 0;
  var grandTotal = total + govtFees;
  
  var invoiceDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM-yy");
  
  // Extract Case details if caseId is provided
  var patentTitle = "", appNo = "", patentOffice = "", firstInventor = data.firstInventor || "N/A";
  if (data.caseId) {
    try {
      var allCases = getAllCases(); // Global function from 02_CaseManagement.gs
      for (var i = 0; i < allCases.length; i++) {
        if (allCases[i].CASE_ID === data.caseId) {
          patentTitle = allCases[i].PATENT_TITLE || "";
          appNo = allCases[i].APPLICATION_NUMBER || "";
          patentOffice = allCases[i].COUNTRY || "";
          // If the user didn't manually type a First Inventor, and we wanted to load it here we could
          break;
        }
      }
    } catch (e) {
      Logger.log("Could not load case details: " + e.message);
    }
  }

  var html = `<!DOCTYPE html>
<html>
<head>
<style>
  body { font-family: Arial, sans-serif; font-size: 10pt; color: #000; margin: 0; padding: 15px; }
  table { width: 100%; border-collapse: collapse; }
  td, th { padding: 4px 6px; border: 1px solid #000; vertical-align: top; }
  .no-border, .no-border td { border: none !important; }
  .header-logo { color: #0088cc; font-size: 24pt; font-weight: bold; font-family: 'Trebuchet MS', Arial, sans-serif; }
  .header-title { font-size: 20pt; font-weight: bold; text-align: right; }
  .bold { font-weight: bold; }
  .right { text-align: right; }
  .center { text-align: center; }
  .bill-to { min-height: 80px; }
</style>
</head>
<body>

<table class="no-border" style="margin-bottom: 10px;">
  <tr>
    <td width="60%" class="header-logo">&#x2230; Metayage</td>
    <td width="40%" class="header-title">Sales Proforma</td>
  </tr>
</table>

<table>
  <!-- Top Section : METAYAGE & Invoice Info -->
  <tr>
    <td width="50%" rowspan="2">
      <div class="bold" style="font-size: 11pt;">METAYAGE PRIVATE LIMITED</div>
      207 EGPI Arcadia 32 Banaswadi Main Road<br>
      Jai Bharath Nagar Bengaluru<br>
      <b>GSTIN/UIN: 29AAICG1194L1Z4</b><br>
      <b>State Name : Karnataka, Code : 29</b>
    </td>
    <td width="25%">
      Invoice No.<br>
      <div class="bold">${invoiceId}</div>
    </td>
    <td width="25%">
      Dated<br>
      <div class="bold">${invoiceDate}</div>
    </td>
  </tr>
  <tr>
    <td colspan="2">Client Reference:<br>${data.caseId || ''}</td>
  </tr>
  
  <!-- Middle Section : Client & Case Info -->
  <tr>
    <td class="bill-to">
      <div class="bold" style="font-size: 11pt;">${client.CLIENT_NAME}</div>
      ${client.CONTACT_PERSON ? client.CONTACT_PERSON + '<br>' : ''}
      ${client.EMAIL ? client.EMAIL + '<br>' : ''}
      ${data.gstNo ? '<b>GSTN</b> &nbsp;&nbsp;' + data.gstNo : ''}
    </td>
    <td colspan="2" style="padding: 0; border: none;">
      <table class="no-border" style="width: 100%; height: 100%;">
        <tr>
          <td width="30%" class="bold" style="border-right: 1px solid #000; border-bottom: 1px solid #000;">First Inventor:</td>
          <td width="70%" style="border-bottom: 1px solid #000;">${firstInventor || '-'}</td>
        </tr>
        <tr>
          <td class="bold" style="border-right: 1px solid #000; border-bottom: 1px solid #000;">Title:</td>
          <td style="border-bottom: 1px solid #000;">${patentTitle || '-'}</td>
        </tr>
        <tr>
          <td class="bold" style="border-right: 1px solid #000; border-bottom: 1px solid #000;">Application No:</td>
          <td style="border-bottom: 1px solid #000;">${appNo || '-'}</td>
        </tr>
        <tr>
          <td class="bold" style="border-right: 1px solid #000;">Patent Office:</td>
          <td>${patentOffice || '-'}</td>
        </tr>
      </table>
    </td>
  </tr>
</table>

<!-- Items Table -->
<table style="border-top: none;">
  <tr>
    <th width="5%">Sl No.</th>
    <th width="55%" class="center">Description of Services</th>
    <th width="10%" class="center">HSN/SAC</th>
    <th width="10%" class="center">Rate</th>
    <th width="20%" class="center">Amount</th>
  </tr>
  <tr style="height: 120px;">
    <td class="right">1</td>
    <td>
      ${data.serviceType}<br>
      <span style="font-style: italic;">${data.description || ''}</span>
      <br><br>
      <div style="text-align: right; width: 100%;">
         ${gstRate > 0 ? (data.gstType === 'IGST' ? '<br><br><b>IGST</b>' : '<br><br><b>CGST</b><br><b>SGST</b>') : ''}
      </div>
    </td>
    <td class="center">998213</td>
    <td>
      <br><br><br>
      <div style="text-align: right;">
         ${gstRate > 0 ? (data.gstType === 'IGST' ? gstRate + '%' : (gstRate/2) + '%<br>' + (gstRate/2) + '%') : ''}
      </div>
    </td>
    <td class="right">
      ${amount.toFixed(2)}<br><br><br>
      ${gstRate > 0 ? (data.gstType === 'IGST' ? gstAmount.toFixed(2) : (gstAmount/2).toFixed(2) + '<br>' + (gstAmount/2).toFixed(2)) : ''}
    </td>
  </tr>
  <tr>
    <td colspan="4" class="right bold">Sub Total</td>
    <td class="right bold">${total.toFixed(2)}</td>
  </tr>
</table>

<!-- Debit Note / Govt Fees if applicable -->
<table style="border-top: none;">
  <tr>
    <th colspan="2" class="center">Debit Note for Reimbursement of Government Fees & Stamp Charges</th>
  </tr>
  <tr>
    <th width="80%" class="center">Description</th>
    <th width="20%" class="center">Amount</th>
  </tr>
  <tr style="height: 40px;">
    <td class="right">${govtFees > 0 ? 'Patent office fee (Govt Fee)' : ''}</td>
    <td class="right">${govtFees > 0 ? govtFees.toFixed(2) : ''}</td>
  </tr>
  <tr>
    <td class="right bold">SubTotal</td>
    <td class="right bold">${govtFees.toFixed(2)}</td>
  </tr>
</table>

<!-- Grand Total -->
<table style="border-top: none;">
  <tr>
    <td width="80%" class="right bold">Grand Total</td>
    <td width="20%" class="right bold" style="font-size: 12pt;">&#8377;${grandTotal.toFixed(2)}</td>
  </tr>
</table>

<!-- Bottom Section: Bank & QR -->
<table style="border-top: none;">
  <tr>
    <td width="40%" class="center" style="vertical-align: middle;">
      ${data.qrCodeBase64 
        ? '<img src="' + data.qrCodeBase64 + '" style="max-width: 100px; max-height: 100px;" />' 
        : '<div style="width: 100px; height: 100px; border: 1px dashed #999; display: inline-block; line-height: 100px; color: #555;">[QR CODE]</div>'}
    </td>
    <td width="60%">
      Company's Bank Details<br>
      <table class="no-border" style="margin-top: 5px;">
        <tr><td width="30%">Bank Name:</td><td class="bold">HDFC Bank</td></tr>
        <tr><td>A/c No.:</td><td class="bold">50200070931727</td></tr>
        <tr><td>Branch & IFSC:</td><td class="bold">Vadavalli-2, Coimbatore & HDFC0007237</td></tr>
        <tr><td>SWIFT Code:</td><td class="bold">HDFC0007237</td></tr>
      </table>
    </td>
  </tr>
</table>

<table style="border-top: none;">
  <tr>
    <td class="bold" style="background: #f0f0f0;">UPI Scan to pay</td>
  </tr>
  <tr>
    <td>
      <span style="text-decoration: underline;">Declaration</span><br>
      Micro Entity# KR030204347<br>
      Thank you for the opportunity to be of service
    </td>
  </tr>
  <tr>
    <td class="center" style="font-size: 8pt; color: #555; border-top: 1px solid #000;">
      Electronic Document. No signature needed
    </td>
  </tr>
</table>

</body>
</html>`;

  // Convert to PDF
  var htmlOutput = HtmlService.createHtmlOutput(html);
  var pdfBlob = Utilities.newBlob(htmlOutput.getContent(), 'text/html', "Invoice_" + invoiceId + ".html").getAs('application/pdf');
  pdfBlob.setName("Invoice_" + invoiceId + ".pdf");

  // Save PDF to client's INVOICES folder
  var pdfFile;
  try {
    var clientFolder = DriveApp.getFolderById(client.CLIENT_FOLDER_ID);
    var invoiceFolders = clientFolder.getFoldersByName("INVOICES");
    if (invoiceFolders.hasNext()) {
      pdfFile = invoiceFolders.next().createFile(pdfBlob);
    } else {
      pdfFile = clientFolder.createFile(pdfBlob);
    }
  } catch (e) {
    // Fallback: save in invoice system folder
    pdfFile = getFolder_("04_INVOICE_SYSTEM").createFile(pdfBlob);
  }

  return pdfFile.getUrl();
}

/**
 * Send invoice email to client
 */
function sendInvoiceEmail_(client, invoiceId, total, dueDate, pdfLink) {
  var subject = "Invoice " + invoiceId + " from Your IP Patent Firm";
  var htmlBody = '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">' +
      '<div style="background: #1a237e; color: white; padding: 20px; text-align: center;">' +
        '<h1 style="margin: 0;">Invoice ' + invoiceId + '</h1>' +
      '</div>' +
      '<div style="padding: 24px; background: #f5f5f5;">' +
        '<p>Dear ' + client.CONTACT_PERSON + ',</p>' +
        '<p>Please find your invoice details below:</p>' +
        '<table style="width: 100%; background: white; border-radius: 8px; padding: 16px;">' +
          '<tr><td style="padding: 8px; font-weight: bold;">Invoice Number:</td><td style="padding: 8px;">' + invoiceId + '</td></tr>' +
          '<tr><td style="padding: 8px; font-weight: bold;">Total Amount:</td><td style="padding: 8px; font-size: 18px; color: #1a237e;"><strong>₹' + total.toFixed(2) + '</strong></td></tr>' +
          '<tr><td style="padding: 8px; font-weight: bold;">Due Date:</td><td style="padding: 8px;">' + Utilities.formatDate(dueDate, Session.getScriptTimeZone(), "dd-MMM-yyyy") + '</td></tr>' +
        '</table>' +
        '<p style="margin-top: 16px;"><a href="' + pdfLink + '" style="background: #1a237e; color: white; padding: 10px 24px; border-radius: 4px; text-decoration: none;">Download Invoice PDF</a></p>' +
        '<p style="color: #666; font-size: 12px; margin-top: 24px;">' +
          'If you have any questions, please reply to this email or contact us through your client portal.' +
        '</p>' +
      '</div>' +
    '</div>';

  MailApp.sendEmail({
    to: client.EMAIL,
    subject: subject,
    htmlBody: htmlBody
  });
}

/**
 * Get invoices for a client
 */
function getClientInvoices(clientId) {
  var sheet = getSheet_("INVOICE");
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var invoices = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === clientId) {
      var inv = {};
      headers.forEach(function(h, idx) { inv[h] = data[i][idx]; });
      invoices.push(inv);
    }
  }
  return invoices;
}

/**
 * Update payment status
 */
function updatePaymentStatus(invoiceId, status, paymentDate, paymentMode) {
  var sheet = getSheet_("INVOICE");
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === invoiceId) {
      var statusIdx = headers.indexOf("PAYMENT_STATUS");
      var dateIdx = headers.indexOf("PAYMENT_DATE");
      var modeIdx = headers.indexOf("PAYMENT_MODE");

      sheet.getRange(i + 1, statusIdx + 1).setValue(status);
      if (paymentDate) sheet.getRange(i + 1, dateIdx + 1).setValue(paymentDate);
      if (paymentMode) sheet.getRange(i + 1, modeIdx + 1).setValue(paymentMode);

      logActivity_("UPDATE_PAYMENT", "INVOICE", invoiceId, "Status: " + status);
      return { success: true };
    }
  }
  return { success: false, message: "Invoice not found" };
}

// ── UI Dialog ──────────────────────────────────────────────
function showGenerateInvoiceDialog() {
  var clients = getAllClients();
  var clientOptions = clients.map(function(c) {
    return '<option value="' + c.CLIENT_ID + '">' + c.CLIENT_ID + ' - ' + c.CLIENT_NAME + '</option>';
  }).join("");

  var html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: 'Google Sans', Arial, sans-serif; padding: 16px; }
      h2 { color: #1a237e; }
      label { display: block; margin-top: 10px; font-weight: 500; }
      input, select, textarea { width: 100%; padding: 8px; margin-top: 4px; border: 1px solid #ddd;
        border-radius: 6px; font-size: 14px; box-sizing: border-box; }
      .btn { background: #1a237e; color: white; padding: 10px 24px; border: none;
        border-radius: 6px; cursor: pointer; margin-top: 16px; }
      .btn-cancel { background: #757575; margin-left: 8px; }
      .row { display: flex; gap: 12px; }
      .row > div { flex: 1; }
      #status { margin-top: 12px; padding: 8px; border-radius: 4px; display: none; }
    </style>
    <h2>🧾 Generate Invoice</h2>
    <div class="row">
      <div>
        <label>Invoice Number *</label>
        <input id="invoiceId" type="text" placeholder="e.g. 616M001EPRNF" />
      </div>
      <div>
        <label>Client *</label>
        <select id="clientId" onchange="loadCases()">${clientOptions}</select>
      </div>
    </div>
    
    <div class="row">
      <div>
        <label>Client GST No.</label>
        <input id="gstNo" type="text" placeholder="e.g. 29AADCO3132B1ZW" />
      </div>
      <div>
        <label>Case (optional)</label>
        <select id="caseId"><option value="">-- General --</option></select>
      </div>
    </div>
    
    <div class="row">
      <div>
        <label>First Inventor</label>
        <input id="firstInventor" type="text" placeholder="e.g. Adarsh Malagouda Patil" />
      </div>
    </div>
    
    <div class="row">
      <div>
        <label>Service Type</label>
        <select id="serviceType">
          <option>Patent Filing</option><option>Patent Prosecution</option>
          <option>Patent Search</option><option>Legal Opinion</option>
          <option>Drafting Fees</option><option>Government Fees</option>
          <option>Consultation</option><option>Annual Maintenance</option>
          <option>Other</option>
        </select>
      </div>
      <div>
        <label>Description</label>
        <textarea id="description" rows="1" placeholder="Service description"></textarea>
      </div>
    </div>
    
    <div class="row">
      <div><label>Amount (₹) *</label><input id="amount" type="number" step="0.01" /></div>
      <div><label>GST Rate (%)</label><input id="gstRate" type="number" value="18" /></div>
    </div>
    
    <div class="row">
      <div>
        <label>GST Type</label>
        <select id="gstType">
          <option value="CGST_SGST">CGST & SGST (Split)</option>
          <option value="IGST">IGST (Single)</option>
        </select>
      </div>
      <div>
        <label>Government Fees (₹)</label>
        <input id="govtFees" type="number" step="0.01" placeholder="e.g. 106305" />
      </div>
    </div>
    
    <div class="row">
      <div>
        <label>QR Code Image</label>
        <input type="file" id="qrCode" accept="image/*" />
      </div>
    </div>
    
    <label>Notes</label>
    <textarea id="notes" rows="2"></textarea>
    <div>
      <button class="btn" onclick="submitFiles()">Generate Invoice</button>
      <button class="btn btn-cancel" onclick="google.script.host.close()">Cancel</button>
    </div>
    <div id="status"></div>
    <script>
      function loadCases() {
        var clientId = document.getElementById('clientId').value;
        google.script.run.withSuccessHandler(function(cases) {
          var sel = document.getElementById('caseId');
          sel.innerHTML = '<option value="">-- General --</option>';
          cases.forEach(function(c) {
            sel.innerHTML += '<option value="'+c.CASE_ID+'">'+c.CASE_ID+' - '+c.PATENT_TITLE+'</option>';
          });
        }).getClientCases(clientId);
      }
      
      function submitFiles() {
        var qrFileInput = document.getElementById('qrCode');
        var qrFile = qrFileInput.files[0];
        if (qrFile) {
          showMsg('Reading image file...', 'info');
          var reader = new FileReader();
          reader.onload = function(e) {
            submitData(e.target.result);
          };
          reader.readAsDataURL(qrFile);
        } else {
          submitData("");
        }
      }

      function submitData(qrBase64) {
        var data = {
          invoiceId: document.getElementById('invoiceId').value,
          clientId: document.getElementById('clientId').value,
          gstNo: document.getElementById('gstNo').value,
          caseId: document.getElementById('caseId').value,
          firstInventor: document.getElementById('firstInventor').value,
          serviceType: document.getElementById('serviceType').value,
          description: document.getElementById('description').value,
          amount: document.getElementById('amount').value,
          gstRate: document.getElementById('gstRate').value,
          gstType: document.getElementById('gstType').value,
          govtFees: document.getElementById('govtFees').value,
          qrCodeBase64: qrBase64,
          notes: document.getElementById('notes').value
        };
        if (!data.invoiceId) { showMsg('Enter Invoice Number','error'); return; }
        if (!data.amount) { showMsg('Enter amount','error'); return; }
        showMsg('Generating invoice & PDF...','info');
        google.script.run
          .withSuccessHandler(function(r){showMsg(r.message,'success');setTimeout(function(){google.script.host.close();},2500);})
          .withFailureHandler(function(e){showMsg('Error: '+e.message,'error');})
          .generateInvoice(data);
      }
      function showMsg(m,t){var e=document.getElementById('status');e.style.display='block';
        e.style.background=t==='error'?'#ffebee':t==='success'?'#e8f5e9':'#e3f2fd';
        e.style.color=t==='error'?'#c62828':t==='success'?'#2e7d32':'#1565c0';e.textContent=m;}
      loadCases();
    </script>
  `).setWidth(550).setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, "Generate Invoice");
}
