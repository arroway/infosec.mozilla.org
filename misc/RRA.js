function refreshRRA() {
  process_all_rra_docs();
}

function process_all_rra_docs() {
  var driveid = '1HaEyQY97DOJhx8g56nTDai6lwjf7_B-x'; // Where assessments are stored
  var template_skip_id = ''; // the template, don't process that
  var s = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = s.getSheetByName('RRA3');
  var folder = DriveApp.getFolderById(driveid);
  var files = folder.getFiles();

  //Start fresh
  sheet.clearContents();
  // Headers
  sheet.appendRow(['Link', 'Name', 'Service Owner', 'Director', 'Service Data Classification', 'Highest Risk Impact', 'Recommendations', 'Highest Recommendation', 'Reviewer', 'Creation date', 'Modification date']);

  while (files.hasNext()) {
    var file = files.next();
    if (file.getId() == template_skip_id) {
      continue;
    }

    // Skip if file owner is not from FoxSec team
    if (!is_foxsec(file.getId())) {
      continue;
    }

    var rra_name = clean_fname(file.getName());
    s.toast("Importing RRA: "+rra_name+"...");
    var results = import_rra(file.getId());
    if (results == undefined) {
      continue;
    }
    insert_rra('https://docs.google.com/document/d/'+file.getId(), rra_name, sheet, results);
  }
  s.toast("All done!");
}

// Test is file owner is not from FoxSec team
// TDB: take into account Mozilla Archive cases (fallback to last editor); could have a FoxSec tag too.
function is_foxsec(fid) {
  var owner = DriveApp.getFileById(fid).getOwner().getName();
  var foxsec_owners = ["AJ Bahnken", "Aaron Meihm", "Gregory Guthe", "Hal Wine", "Julien Vehent", "Simon Bennetts", "Stephanie Ouillon"];
  for (i = 0; i < foxsec_owners.length; i++) {
    var o = foxsec_owners[i];
    if (o.indexOf(owner) !== -1){
      return true;
    }
  }
  return false;
}

// Import RRA doc to register
function import_rra(fid) {
  var doc = DocumentApp.openById(fid);
  var docid = doc.getId();
  var tables = doc.getBody().getTables();
  var paragraphs = doc.getBody().getParagraphs();
  // These are not the manual "marked as reviewed" dates, but actual modification/creation date of the document
  var creation_date = DriveApp.getFileById(fid).getDateCreated();
  var modification_date = DriveApp.getFileById(fid).getLastUpdated();
  var file_owner = DriveApp.getFileById(fid).getOwner().getName();
  var levels = ['UNKNOWN', 'LOW', 'MEDIUM', 'HIGH', 'MAXIMUM']; //order matters
  var recs = [];
  var highest_rec_rank = 0;
  var results = [];

  // First table has metadata
  var meta_table = tables[0];
  var i =0;

  if (meta_table.getNumRows() > 1) {
    // do we have a service name cell? if yes skip it, we already have the name in the file name
    if (meta_table.getCell(0,0).getText().split('\n')[0] == 'Service Name') {
      i = i+1
    }
    var serviceowner =   meta_table.getCell(i,1).getText().split('\n')[0];
    results.push(['Service Owner', serviceowner]); i=i+1;
    var director =       meta_table.getCell(i,1).getText().split('\n')[0];
    results.push(['Director', director]); i=i+1;
    var classification = meta_table.getCell(i,1).getText().split('\n')[0];
    results.push(['Service Data Classfication', classification]); i=i+1;
    var impact =         meta_table.getCell(i,1).getText().split('\n')[0];
    results.push(['Highest Risk Impact', impact]); i=i+1;
  }

  // Find recommendations (this loop is a little hackish, but i couldnt find a good way to iterate without changing the original docs/adding bookmarks f.e.)
  for (var p=0;p<paragraphs.length;p++) {
    var current = paragraphs[p];
    if (current.getText() == 'Recommendations') {
      for (var p1=p;p1<paragraphs.length;p1++) {
        var line = paragraphs[p1];
        if (line.getType() == DocumentApp.ElementType.LIST_ITEM) {
          //Find if we have a recommendation level associated with the recommendation list item
          var rec_level = 'UNKNOWN';
          for (var l=0;l<levels.length;l++) {
            if (line.getText().split(' ')[0] == levels[l]) {
              rec_level = levels[l];
              // Find highest rec
              if (l > highest_rec_rank) {
                highest_rec_rank = l;
              }
              break;
            }
          }
          // Check that the rec is not solved/strikedout
          // And that it isn't just a sub list item (i.e. it has a rec_level)
          var attrs = line.getAttributes();
          if (attrs[DocumentApp.Attribute.STRIKETHROUGH] != true && rec_level != 'UNKNOWN') {
            recs.push([highest_rec_rank, rec_level, line.getText()]);
          } else {
            Logger.log('Recommendation already striked out, skipping');
          }
        }
      }
      break;
    }
  }
  results.push(['Recommendations', recs.length]);
  results.push(['Highest Recommendation', levels[highest_rec_rank]]);
  results.push(['Reviewer', file_owner]);
  results.push(['Creation date', creation_date]);
  results.push(['Modification date', modification_date]);
  return results
}

// clean up filename a bit
function clean_fname(fname) {
  var clean_name = fname.split(' - ')[1];
  if (clean_name === undefined) {
    clean_name = fname.split('RRA ')[1];
  }
  if (clean_name === undefined) {
    clean_name = fname;
  }
  return clean_name
}

// Insert results in register
function insert_rra(docid, fname, sheet, results) {
  var s = SpreadsheetApp.getActiveSpreadsheet()

  var row = [docid, fname];
  var valid = true;

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
}
