function doNothing(){}
function trueFilter() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('I:I').createFilter();
  var criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['', 'FALSE'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(7, criteria);
};
