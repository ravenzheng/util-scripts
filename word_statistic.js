/**
 * Author: raven
 * Date: 2014/12/29 12:50
 * Words Count Statistic, using Windows JScript.
 */

var fso = new ActiveXObject('Scripting.FileSystemObject');
var column1_filename = '文件名';
var column2_words = '字数';
var column3_chars = '字符数(不计空格)';
var column4_chars_without_space = '字符数(计空格)';
var column5_non_chinese = '非中文单词';
var column6_fareastchars = '中文和朝鲜语单词';
var wdStatisticCharacters = 3; //from enum WdStatistic
var wdStatisticCharactersWithSpaces = 5; //from enum WdStatistic
var wdStatisticFarEastCharacters = 6; //from enum WdStatistic
// var wdStatisticLines = 1; //from enum WdStatistic
// var wdStatisticPages = 2; //from enum WdStatistic
// var wdStatisticParagraphs = 4; //from enum WdStatistic
var wdStatisticWords = 0; //from enum WdStatistic
var curDir = '.';
var fileArr = [];
scanDocFiles(curDir);
var statisticList = statistic(fileArr);
generateExcelSheet(statisticList);

function scanDocFiles(dir) {
  var folder = fso.GetFolder(dir);
  var fileEnumerator = new Enumerator(folder.Files);
  while (!fileEnumerator.atEnd()) {
    var file = fileEnumerator.item();
    var ext = fso.GetExtensionName(file);
    if (ext === 'doc' || ext === 'docx') {
      fileArr.push(file.Path);
    }
    fileEnumerator.moveNext();
  }
  var folderEnumerator = new Enumerator(folder.SubFolders);
  while (!folderEnumerator.atEnd()) {
    var subFolder = folderEnumerator.item();
    scanDocFiles(subFolder.Path);
    folderEnumerator.moveNext();
  }
}

function statistic(fileList) {
  var length = fileList.length;
  if (length === 0) {
    return [];
  }
  var folder = fso.GetFolder(curDir);
  var wordApp = new ActiveXObject('Word.Application');
  wordApp.Visible = true;
  var statistic_list = [];
  for (var i = 0; i < length; i ++) {
    var docPath = fileList[i];
    var doc = wordApp.Documents.Open(docPath);
    var wordsCount = doc.ComputeStatistics(wdStatisticWords, true);
    var charsCount = doc.ComputeStatistics(wdStatisticCharacters, true);
    var charsWithSpaceCount = doc.ComputeStatistics(wdStatisticCharactersWithSpaces, true);
    var fareastCharsCount = doc.ComputeStatistics(wdStatisticFarEastCharacters, true);
    var statisticData = {};
    statisticData.filename = docPath.replace(folder.Path + '\\', '');
    statisticData.wordsCount = wordsCount;
    statisticData.charsCount = charsCount;
    statisticData.charsWithSpaceCount = charsWithSpaceCount;
    statisticData.fareastCharsCount = fareastCharsCount;
    statisticData.nonChineseCount = wordsCount - fareastCharsCount;
    statistic_list.push(statisticData);
    doc.Close();
  }
  wordApp.Quit();
  return statistic_list;
}

function generateExcelSheet(statistiList) {
  var excelApp = new ActiveXObject('Excel.Application');
  excelApp.Visible = true;
  var workbook = excelApp.Workbooks.Add();
  var sheet = workbook.Sheets(1);
  sheet.Columns('A:A').ColumnWidth = 40;
  sheet.Columns('B:B').ColumnWidth = 8;
  sheet.Columns('C:C').ColumnWidth = 15;
  sheet.Columns('D:D').ColumnWidth = 15;
  sheet.Columns('E:E').ColumnWidth = 15;
  sheet.Columns('F:F').ColumnWidth = 15;
  sheet.Cells(1, 1).Value = column1_filename;
  sheet.Cells(1, 2).Value = column2_words;
  sheet.Cells(1, 3).Value = column3_chars;
  sheet.Cells(1, 4).Value = column4_chars_without_space;
  sheet.Cells(1, 5).Value = column5_non_chinese;
  sheet.Cells(1, 6).Value = column6_fareastchars;
  var length = statistiList.length;
  for (var i = 0; i< length; i ++) {
    var dataItem = statistiList[i];
    sheet.Cells(i + 2, 1).Value = dataItem.filename;
    sheet.Cells(i + 2, 2).Value = dataItem.wordsCount;
    sheet.Cells(i + 2, 3).Value = dataItem.charsCount;
    sheet.Cells(i + 2, 4).Value = dataItem.charsWithSpaceCount;
    sheet.Cells(i + 2, 5).Value = dataItem.nonChineseCount;
    sheet.Cells(i + 2, 6).Value = dataItem.fareastCharsCount;
  }
  var curFolder = fso.GetFolder(curDir);
  var excelPath = fso.BuildPath(curFolder.Path, fso.GetBaseName(curFolder.Name) + '字数统计.xls');
  if (fso.FileExists(excelPath)) {
    fso.DeleteFile(excelPath);
  }
  sheet.SaveAs(excelPath);
}
