function myFunction() {
  // This is where you will call the functions to process the email body and send the formatted email.
  // Example:
  Logger.log("Starting myFunction..."); 
  var emailBody = getFormattedReportBody(get_report_body());
  Logger.log("Formatted email body: " + emailBody); 
  send_email(get_report_subject(), emailBody);
  Logger.log("Email sent successfully."); 
}

function get_report_body() {
  Logger.log("Getting report body..."); 
  var latest_report = GmailApp.search('"active bod report" from:noreply@npnoc.verizon.com', 0, 1)[0];
  Logger.log("Latest report found: " + latest_report); 
  return latest_report.getMessages()[0].getBody();
}

function get_report_subject() {
  var latest_report = GmailApp.search('"active bod report" from:noreply@npnoc.verizon.com', 0, 1)[0];
  return latest_report.getMessages()[0].getSubject();
}

function get_report_time() {
  var latest_report = GmailApp.search('"active bod report" from:noreply@npnoc.verizon.com', 0, 1)[0];
  return latest_report.getMessages()[0].getDate();
}

function send_email(email_subject, email_body) {
  var latest_report = GmailApp.search('"active bod report" from:noreply@npnoc.verizon.com', 0, 1)[0]
  var latest_report_thread = latest_report.getMessages();
  var first_message = latest_report_thread[0];
  var first_message_plain = first_message.getPlainBody();
  var first_message_html = first_message.getBody();

  var email_body_html = email_body.toString().replace(/\n/g, "<br>");
  var email_body_html = email_body_html.toString().replace(/<br><br><br>/, "<br><br>");
  var email_body_html = email_body_html.toString().replace(/(<br>|^)([a-z]{6}[a-z0-9]{0,2} -.*?)<br>/gm, "$1<b><u>$2<\/b><\/u><br>");
  var email_body_html = email_body_html.toString().replace(/<br>cary network management center - power\/transport/, "---<br>cary network management center - power\/transport");

  return first_message.reply(email_body + '\n\n' + first_message_plain, { 'cc': 'teonmoore@verizon.com', 'htmlBody': email_body_html + '<br><br>' + first_message_html })
}

function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
}

function getFormattedReportBody(reportBody) {
  Logger.log("Formatting report body...");
  var siteReports = extractSiteReportsFromHtml(reportBody);
  Logger.log("Site reports: " + siteReports); 

  var formattedReports = siteReports.map(function (siteReport) {
    Logger.log("Processing site report: " + siteReport);
    return formatSiteReport(siteReport);
  });

  Logger.log("Formatted reports: " + formattedReports); 
  return formattedReports.join("\n\n");
}

function extractSiteReportsFromHtml(htmlBody) {
  // Use XmlService to parse the HTML body
  var xml = XmlService.parse(htmlBody);
  var root = xml.getRootElement();

  // Find all 'pre' tags that contain the site information
  var preElements = root.getDescendants()
    .filter(function(d) { 
      return d.getType() === 'ELEMENT' && d.getName() === 'pre'; 
    });

  // Extract the text content from each 'pre' tag
  var siteReports = preElements.map(function(preElement) {
    return preElement.getText();
  });

  return siteReports;
}

function formatSiteReport(siteReport) {
  var siteInfo = extractSiteInfo(siteReport);
  
  if (!siteInfo.site) {
    // Skip reports that don't have valid site information
    return "";
  }

  Logger.log("Fetching address for site: " + siteInfo.site); 
  siteInfo.address = getAddressFromSheet(siteInfo.site);
  Logger.log("Address found: " + siteInfo.address);
  
  // Fetch site manager
  siteInfo.manager = getManagerFromSheet(siteInfo.site);
  Logger.log("Manager found: " + siteInfo.manager);

  var formattedReport = 
    siteInfo.site + "\n" +
    "Address: " + siteInfo.address + "\n" +
    "PB Level " + siteInfo.pb + "\n" + 
    "Battery Reserve time/remain battery: " + siteInfo.batteryReserve + "hrs" +
    (siteInfo.manager ? "\nSite Manager: " + siteInfo.manager : "") + 
    "\n" + 
    (siteInfo.tickets ? siteInfo.tickets + "\n" : "") +
    (siteInfo.flashInfo ? siteInfo.flashInfo + "\n" : "") +
    (siteInfo.message ? siteInfo.message : "");

  return formattedReport;
}

function extractSiteInfo(siteReport) {
  var siteInfo = {
    site: "",
    pb: "",
    batteryReserve: "",
    address: "",
    manager: "",
    tickets: "",
    flashInfo: "",
    message: ""
  };
  
  // Remove HTML tags for easier text processing
  var cleanedReport = siteReport.replace(/<\/?[^>]+(>|$)/g, " ").replace(/\s+/g, " ");
  
  // Extract site information using regex patterns
  var siteMatch = /Location\(Site\)\s+(\w+)/.exec(siteReport);
  if (siteMatch) {
    siteInfo.site = siteMatch[1];
  }
  
  var pbMatch = /PB Level\s+(\d+)/.exec(siteReport);
  if (pbMatch) {
    siteInfo.pb = pbMatch[1];
  }
  
  var batteryMatch = /Est Batt Status\s+R=(\d+\.\d+)\//.exec(siteReport);
  if (batteryMatch) {
    siteInfo.batteryReserve = batteryMatch[1];
  }
  
  // Extract ticket information
  var etmsMatch = /ETMS Ticket#\s+(\S+)/.exec(siteReport);
  if (etmsMatch && etmsMatch[1].trim() !== "") {
    siteInfo.tickets = "ETMS Ticket: " + etmsMatch[1];
  }
  
  var vrepairMatch = /VRepair Ticket#\s+(\S+)/.exec(siteReport);
  if (vrepairMatch && vrepairMatch[1].trim() !== "") {
    if (siteInfo.tickets) {
      siteInfo.tickets += ", VRepair Ticket: " + vrepairMatch[1];
    } else {
      siteInfo.tickets = "VRepair Ticket: " + vrepairMatch[1];
    }
  }
  
  // Extract Flash information
  var flashMatch = /Flash ID:(\S+).*?Flash Status: (\S+)/.exec(siteReport);
  if (flashMatch) {
    siteInfo.flashInfo = "Flash ID: " + flashMatch[1] + ", Status: " + flashMatch[2];
    
    // Look for flash comment
    var commentMatch = /timeStamp \(GMT\).*?comment.*?<td>(.*?)<\/td>/.exec(siteReport);
    if (commentMatch && commentMatch[1].trim() !== "") {
      siteInfo.flashInfo += "\nComment: " + commentMatch[1].trim();
    }
  }
  
  return siteInfo;
}

function getAddressFromSheet(site) {
  Logger.log("Getting address from sheet for site: " + site); 
  var spreadsheet = SpreadsheetApp.openById("1i-hry21wuunsvqp5y87dl2xsakwfviipijhwhfzvzaa").getSheetByName("PBL Domestic");
  var data = spreadsheet.getDataRange().getValues();

  var searchColumn = site.length > 6 ? 1 : 2; 

  for (var i = 1; i < data.length; i++) { 
    if (data[i][searchColumn].trim() === site) {
      var addr1 = data[i][16] || ""; 
      var addr2 = data[i][17] || ""; 
      var city = data[i][18] || "";
      var state = data[i][19] || ""; 
      return `${addr1} ${addr2}, ${city}, ${state}`.trim();
    }
  }

  Logger.log("Site not found in the spreadsheet: " + site); 
  return "Site not found in spreadsheet";
}

function getManagerFromSheet(site) {
  Logger.log("Getting manager from sheet for site: " + site); 
  var spreadsheet = SpreadsheetApp.openById("1i-hry21wuunsvqp5y87dl2xsakwfviipijhwhfzvzaa").getSheetByName("PBL Domestic");
  var data = spreadsheet.getDataRange().getValues();

  // Find the header row index first to locate the Ops Mngr column
  var headerRow = 0;
  var opsManagerCol = -1;
  
  for (var col = 0; col < data[headerRow].length; col++) {
    if (data[headerRow][col] === "Ops Mngr (AD)") {
      opsManagerCol = col;
      break;
    }
  }
  
  if (opsManagerCol === -1) {
    Logger.log("Could not find 'Ops Mngr (AD)' column"); 
    return "";
  }
  
  var searchColumn = site.length > 6 ? 1 : 2; 

  for (var i = 1; i < data.length; i++) { 
    if (data[i][searchColumn].trim() === site) {
      return data[i][opsManagerCol] || "";
    }
  }

  Logger.log("Site manager not found for: " + site); 
  return "";
}