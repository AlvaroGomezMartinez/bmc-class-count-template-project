// Utility function: Copies campus names and sets data validation in campus spreadsheets
// See main script for usage and documentation

function copyValidationToAllMonthsAllCampuses() {
  var masterId = '1iIkKYUMsc7Lo8CZXBryOBRccIFtMcOdJP4aANeKejgs';
  var master = SpreadsheetApp.openById(masterId);
  var campusInfoSheet = master.getSheetByName('CampusBMCSheetInfo');
  var data = campusInfoSheet.getRange(2, 2, campusInfoSheet.getLastRow() - 1, 4).getValues(); // B2:E

  var months = ['AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER', 'JANUARY', 'FEBRUARY', 'MARCH', 'APRIL/ MAY PROJECTIONS'];

  var campusList = [
    "Adams Hill","Aue","Behlau","Blattman","Boldt","Boone","Brauchle","Braun Station","Burke","Cable","Carlos Coon","Carnahan","Carson","Chumbley","Cody","Cole","Colonies North","Driggers","Ellison","Elrod","Evers","Fields","Fisher","Forester","Franklin","Galm","Glass","Glenn","Glenoaks","Hatchett","Henderson","Hoffmann","Howsman","Kallison","Krueger","Kuentz","Langley","Lewis","Lieck","Linton","Locke Hill","Los Reyes","Martin","Mary Hull","May","McDermott","Mead","Meadow Village","Michael","Mireles","Murnin","Northwest Crossing","Oak Hills Terrace","Ott","Passmore","Powell","Raba","Reed","Rhodes","Scarborough","Scobee","Steubing","Thornton","Timberwilde","Tomlinson","Valley Hi","Wanke","Ward","Wernli","WWT","Bernal","Briscoe","Connally","Folks","Garcia","Hobby","Holmgreen MS","Jefferson","Jones","Jordan","Luna","Neff","Northside Alternative MS","Pease","Rawlinson","Rayburn","Ross","Rudder","Stevenson","Stinson","Straus","Vale","Zachry","Brandeis","Brennan","Clark","Harlan","Holmes","Jay","Marshall","NAHS","O’Connor","Sotomayor","Stevens","Taft","Warren"
  ];
  var validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(campusList, true)
    .setAllowInvalid(false)
    .build();

  var log = [];
  var props = PropertiesService.getScriptProperties();
  var batchSize = 10;
  var cursorKey = 'CAMPUS_FILL_IDX';
  var startIndex = parseInt(props.getProperty(cursorKey) || '0', 10);
  var endIndex = Math.min(data.length, startIndex + batchSize);
  log.push('Processing campuses ' + (startIndex + 1) + ' to ' + endIndex + ' of ' + data.length);
  for (var i = startIndex; i < endIndex; i++) {
    var campusName = data[i][0]; // column B
    var spreadsheetId = data[i][3]; // column E
    if (!campusName || !spreadsheetId) continue;
    log.push('---');
    log.push('About to open spreadsheet: ' + spreadsheetId + ' for campus: ' + campusName);
    try {
      var ss = SpreadsheetApp.openById(spreadsheetId);
      log.push('Opened spreadsheet: ' + spreadsheetId);
      for (var j = 0; j < months.length; j++) {
        var month = months[j];
        log.push('  Checking sheet: ' + month);
        var sheet = ss.getSheetByName(month);
        if (sheet) {
          if (month === 'APRIL/ MAY PROJECTIONS') {
            // Remove data validation in D4:D1000
            var dRange = sheet.getRange(4, 4, 997, 1); // D4:D1000
            dRange.clearDataValidations();
            // Fill D4:D1000 with campus name
            var dValues = Array(997).fill([campusName]);
            dRange.setValues(dValues);
            // Set data validation in E4:E1000
            var eRange = sheet.getRange(4, 5, 997, 1); // E4:E1000
            eRange.setDataValidation(validation);
            log.push('    Cleared validation in D4:D1000, filled with campus name, and applied validation to E4:E1000 in sheet: ' + month);
          } else {
            // Remove data validation in D3:D1000
            var dRange = sheet.getRange(3, 4, 998, 1); // D3:D1000
            dRange.clearDataValidations();
            // Fill D3:D1000 with campus name
            var dValues = Array(998).fill([campusName]);
            dRange.setValues(dValues);
            log.push('    Cleared validation in D3:D1000 and filled with campus name in sheet: ' + month);
          }
        } else {
          log.push('    Sheet not found: ' + month);
        }
      }
    } catch (e) {
      log.push('  Error opening spreadsheet: ' + spreadsheetId + ' - ' + e);
    }
  }
  props.setProperty(cursorKey, String(endIndex));
  if (endIndex >= data.length) {
    log.push('All campuses processed. Resetting cursor.');
    props.deleteProperty(cursorKey);
  } else {
    log.push('Batch complete. Next run will process campuses ' + (endIndex + 1) + ' to ' + Math.min(data.length, endIndex + batchSize));
  }
  Logger.log(log.join('\n'));
  SpreadsheetApp.getUi().alert('Campus name fill complete. See View > Logs for details.');
}
