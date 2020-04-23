function refreshPrinciples() {
  process_all_docs_p();
}

function process_all_docs_p() {
  var driveid = '1pJdoPTIrghW7o2yEX5Oeb9HH35r_4ecs'; // Where assessments are stored
  var template_skip_id = '1ybxztOpl537WDE5bARZmDtmS5x1bqYGlbbEZzOIy8gc'; // the template, don't process that
  var s = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = s.getSheetByName('Principles Assessment');
  var folder = DriveApp.getFolderById(driveid);
  var files = folder.getFiles();

  //Start fresh
  sheet.clearContents();
  // Headers
  sheet.appendRow(['Link', 'Name', 'Do not expose unnecessary services', 'Do not grant or retain permissions that are no longer needed', 'Do not allow lateral movement', 'Isolate environments', 'Patch systems', 'Meet web standards', 'Guarantee data integrity and confidentiality', 'Fraud detection and forensics', 'Are you at risk?', 'Inventory the landscape', 'KISS - Keep It Simple and thus Secure', 'Require two-factor authentication', 'Use central identity management (Single Sign-On)', 'Require strong authentication']);

  while (files.hasNext()) {
    var file = files.next();
    if (file.getId() == template_skip_id) {
      continue;
    }
    s.toast("Importing assessment: "+file.getName()+"...");
    var results = import_p(file.getId());
    insert_p('https://docs.google.com/document/d/'+file.getId(), file.getName(), sheet, results);
  }
  s.toast("All done!");
}

// Import a principles assessment doc to register
function import_p(fid) {
  var doc = DocumentApp.openById(fid);
  var docid = doc.getId();
  var tables = doc.getBody().getTables();


  // Export all data
  // Conversion table (values are slightly arbitrary, losely mapped on percentage
  var grades = {'Principle is consistently followed (>90% of the time)': 90,
                'Principle is generally followed (60-90% of the time)': 60, 
                'Principle is occasionally followed (15-60% of the time)': 15, 
                'Principle is rarely followed (<15% of the time)': 0,
                'N/A': -1,
                '': -1
               };
  var results = [];

  for (var i = 0; i < tables.length; i++) {
    var t = tables[i];

    if (t.getNumRows() > 1) {
      var c1 = t.getCell(0, 0);
      //Find principle
      var principle = c1.getText().split('\n')[0];
      if (principle == "Operational Group") {
        continue; //skip that cruft
      }

      //Find grade
      var c2 = t.getCell(0, 1);
      var g = null;
      for (y in grades) {
        if (y == c2.getText()) {
          g = grades[c2.getText()];
          break;
        }
      }
      results.push([principle, g]);
    }
  }
  return results
}

// Insert results in register
function insert_p(docid, fname, sheet, results) {
  var s = SpreadsheetApp.getActiveSpreadsheet()
  var row = [docid, fname.split(' - ')[1]];
  var valid = true;

  Logger.log(results);
  for (var y = 0; y < results.length; y++) {
    row.push(results[y][1]);
    if (results[y][1] == '') {
      valid = false;
    }
  }
  //Logger.log(row);
  sheet.appendRow(row);
  if (!valid) {
    Logger.log("Row is missing elements: "+row);
  }



  // Below is code to insert by column instead of by row. We don't currently use that, but in case we decide to revert, I'm keeping this around.

  //var r = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  //Logger.log("Scanning entries 1 to "+sheet.getLastColumn()+" range is 1 to "+r.getNumColumns());
  //var found = false;

  /*
  for (var i = 1; i <= r.getNumColumns(); i++) {
    var c = r.getCell(1,i);
    if (c.getValue() == docid) {
      Logger.log("Found previous entry: "+docid);
      found = true;
      //i=i+1;
      break;
    }
  }
  // Add new column since we don't have any

  if (!found) {
    Logger.log("No previous entry found for: "+docid);
    sheet.insertColumns(i);
    sheet.getRange(1, i).setValue(docid);
  }
  Logger.log("Selected col "+i+" for data insertion");
  // Strip first part of filename for readability
  sheet.getRange(2, i).setValue(fname.split(' - ')[1]);
  for (var y = 0; y < results.length; y++) {
    //Logger.log("inserting data row x col: "+i+"x"+(y+2)+" "+results[y][1]);
    sheet.getRange(y+3, i).setValue(results[y][1]);
  }
   */
}
