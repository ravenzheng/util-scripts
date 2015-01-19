/**
 * Author: raven
 * Date: 2015/01/19 14:40
 * Converting excel files to word files, using Windows JScript.
 */

var fso = new ActiveXObject("Scripting.FileSystemObject");
var curDir = ".";
var fileArr = [];
listExcelFiles(curDir);
convertExcelToWord(fileArr);

function listExcelFiles(dir) {
  var folder = fso.GetFolder(dir);
  var fileEnumerator = new Enumerator(folder.Files);
  while (!fileEnumerator.atEnd()) {
    var file = fileEnumerator.item();
    var ext = fso.GetExtensionName(file);
    var fileInfo = {};
    if (ext == "xls" || ext == "xlsx") {
      fileArr.push(file.Path);
    }
    fileEnumerator.moveNext()
  }
  var folderEnumerator = new Enumerator(folder.SubFolders);
  while (!folderEnumerator.atEnd()) {
    var folder = folderEnumerator.item();
    listExcelFiles(folder.Path);
    folderEnumerator.moveNext();
  }
}

function convertExcelToWord(fileList) {
  var length = fileList.length;
  if (length == 0) {
    return;
  }
  var excelApp = new ActiveXObject("Excel.Application");
  excelApp.Visible = true;
  excelApp.DisplayAlerts = false;
  var wordApp = new ActiveXObject("Word.Application");
  wordApp.Visible = true;
  var doc;
  var docRange;
  var sheetCount;
  var sheet;
  for (var i = 0; i < length; i ++) {
    var excelPath = fileList[i];
    var workbook = excelApp.Workbooks.Open(excelPath);
    doc = wordApp.Documents.Add();
    docRange = doc.Range();
    var docPath = excelPath.substr(0, excelPath.length - 4) + ".doc";
    sheetCount = workbook.WorkSheets.Count;
    for(var j = 0; j< sheetCount; j ++) {
      sheet = workbook.WorkSheets(j + 1);
      sheet.UsedRange.Copy();
      docRange.PasteExcelTable(false, true, false);
      docRange.Collapse(0);
      docRange.InsertParagraphAfter();
      docRange.Collapse(0);
    }
    doc.SaveAs(docPath);
    doc.Close();
    workbook.Close();
  }
  excelApp.Quit();
  wordApp.Quit();
}

