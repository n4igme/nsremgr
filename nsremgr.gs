var sstatus = "https://docs.google.com/spreadsheets/d/1jEm329I6Yqwn-h-aHfniR8iyrkchRvB_SrdLcSC3tWo/edit";
var supdate = "https://docs.google.com/spreadsheets/d/1wWgvjlAJsrVi_vXEnSZxc_JMmqewcFP9Z6HTUmaWiTg/edit";
var ss = SpreadsheetApp.openByUrl(sstatus);
var su = SpreadsheetApp.openByUrl(supdate);
var hostList = ss.getSheetByName("Host Summary");
var baseline = ss.getSheetByName("Vulnerability List");
var newReport = su.getSheetByName("Report_13/10/2023");
var reportDate = "13 Oct 2023";

// Assuming your data starts from the second row (header in the first row)
var ss1 = baseline.getRange(2, 1, baseline.getLastRow(), baseline.getLastColumn()).getValues();
var ss2 = newReport.getRange(2, 1, newReport.getLastRow(), newReport.getLastColumn()).getValues();
var hosts = hostList.getRange(2, 1, hostList.getLastRow(), hostList.getLastColumn()).getValues();

function updateData() {
  // Loop through data in the baseline sheet
  for (var i = 1; i < ss1.length; i++) {
    var host1 = baseline.getRange(i+1, 10).getValue();
    var port1 = baseline.getRange(i+1, 16).getValue();
    var nessID1 = baseline.getRange(i+1, 17).getValue();
    var foundInSS2 = false;
    
    // Compare baseline in newReport to update the baseline sheet
    for (var x = 1; x < ss2.length; x++) {
      var host2 = newReport.getRange(x+1, 5).getValue();
      var port2 = newReport.getRange(x+1, 7).getValue();
      var nessID2 = newReport.getRange(x+1, 1).getValue();

      // Compare rows based on host, port, and pluginID
      if (host1 === host2 && port1 === port2 && nessID1 === nessID2) {
        foundInSS2 = true;
        break;
      }
    }
    
    if (!foundInSS2) {
      baseline.getRange(i+1, 1).setValue("Closed");
      baseline.getRange(i+1, 5).setValue(reportDate);
    }
    else {
      baseline.getRange(i+1, 1).setValue("Open");
    }

    var plan = baseline.getRange(i+1, 3).getValue();
    var initiate = baseline.getRange(i+1, 2).getValue();
    var severity = baseline.getRange(i+1, 6).getValue();
    if (plan === ""){
      var x = targetDate(initiate, severity);
      baseline.getRange(i+1, 3).setValue(x);
    }

    var host = baseline.getRange(i+1, 8).getValue();
    var env = baseline.getRange(i+1, 9).getValue();
    if (host === "") {
      baseline.getRange(i+1, 8).setValue(getName(host1));
    }

    if (env === "") {
      baseline.getRange(i+1, 9).setValue(getEnv(host1));
    }
  }
  
  // Check data for new vulnerabilities in newReport
  for (var i = 1; i < ss2.length; i++) {
    var host1 = newReport.getRange(i+1, 5).getValue();
    var port1 = newReport.getRange(i+1, 7).getValue();
    var nessID1 = newReport.getRange(i+1, 1).getValue();
    var vulner = newReport.getRange(i+1, 8).getValue();
    var risk = newReport.getRange(i+1, 4).getValue();
    var solusi = newReport.getRange(i+1, 11).getValue();
    var sinopsis = newReport.getRange(i+1, 9).getValue();
    var deskripsi = newReport.getRange(i+1, 10).getValue();
    var protokol = newReport.getRange(i+1, 6).getValue();
    var cvss = newReport.getRange(i+1, 3).getValue();
    var foundInSS1 = false;
    
    // Compare newReport to update data to the baseline sheet
    for (var x = 1; x < ss1.length; x++) {
      var host2 = baseline.getRange(x+1, 10).getValue();
      var port2 = baseline.getRange(x+1, 16).getValue();
      var nessID2 = baseline.getRange(x+1, 17).getValue();

      // Compare rows based on host, port, and vulnerability name
      if (host1 === host2 && port1 === port2 && nessID1 === nessID2) {
        foundInSS1 = true;
        break;
      } 
    }

    // Compare rows based on host, port, and pluginID
    if(!foundInSS1){
      baseline.appendRow(["Open", reportDate, targetDate(reportDate, risk), "", "", risk, "",
      getName(host1), getEnv(host1), host1, vulner, solusi, sinopsis, deskripsi, protokol, port1, nessID1, cvss]);
    }
  }

  Logger.log("Update completed. Please refer to the sheet.");
}

function getName(ipaddress) {
  var name;
  for (var i = 1; i < hosts.length; i++) {
    var host = hostList.getRange(i+2, 4).getValue();
    if(host === ipaddress) {
      name = hostList.getRange(i+2, 2).getValue();
      break;
    }
  }
  return name;
}

function getEnv(ipaddress) {
  var env;
  for (var i = 1; i < hosts.length; i++) {
    var host = hostList.getRange(i+2, 4).getValue();
    if(host === ipaddress) {
      env = hostList.getRange(i+2, 3).getValue();
      break;
    }
  }
  return env;
}

function targetDate(initiateDate, severity) {
  //var currentDate = new Date();
  var targetDate;

  //Vulnerability treatment based on the SLA in each company
  switch (severity) {
    case "Critical": targetDate=30;break;
    case "High": targetDate=45;break;
    case "Medium" : targetDate=60;break;
    case "Low" : targetDate=90;break;
  }
  var calculatedDate = new Date(initiateDate);
  calculatedDate.setDate(calculatedDate.getDate() + targetDate);
  var resultDate = Utilities.formatDate(calculatedDate, Session.getScriptTimeZone(), "dd MM yyyy");
  return resultDate;
}