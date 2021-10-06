function onFormSubmit(e) {

  const lock = LockService.getScriptLock();
  try {
    // ロックを取得する
    if (lock.tryLock(10 * 1000)) {

      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("顧客名簿");

      console.log(e.values)

      for (var i = 12; i <= 16; i++) {
        if (e.values[i] !== "") {
          console.log(i)
          var mama = KanriSheet(e.values[i], e.values[2])
          ClassMeibo(e.values[i], e.values[2], e.values[3])

          // 該当ライン取得
          if (mama) {
            var targetRow = sheet.getLastRow() + 1;
          } else {
            var targetRow = SearchMamaLine(sheet)
            console.log(targetRow)
            sheet.insertRowAfter(targetRow);
          }

          console.log("targetRow:" + targetRow)
          // 名前
          sheet.getRange(targetRow, 1).setValue(e.values[2]);
          // ふりがな
          sheet.getRange(targetRow, 2).setValue(e.values[3]);
          // 生年月日
          sheet.getRange(targetRow, 3).setValue(e.values[4]);
          sheet.getRange(targetRow, 4).setFormula('=DATEDIF(C' + targetRow + ',TODAY(),"Y")');
          // 住所
          sheet.getRange(targetRow, 5).setValue(e.values[5]);
          // 保護者名
          sheet.getRange(targetRow, 6).setValue(e.values[6]);
          // 続柄
          sheet.getRange(targetRow, 7).setValue(e.values[7]);
          // 電話番号
          sheet.getRange(targetRow, 8).setValue(e.values[8]);
          // 備考
          sheet.getRange(targetRow, 9).setValue(e.values[10]);
        }
      };
    }
  } catch (error) {
    console.error(printError(error));
  } finally {
    // ロック開放
    lock.releaseLock();
  }
}

function KanriSheet(oneClass, name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("管理シート");
  var allClasses = sheet.getRange('C1:1').getValues();
  var allClassList = Array.prototype.concat.apply([], allClasses);
  // console.log(allClassList)

  var trueLength = SearchCount(allClassList)
  console.log('trueLength: ' + trueLength)

  if (oneClass == "Jazz Yukana  一般(ママ)") {
    var trgRow = sheet.getLastRow() + 1;
    console.log('trgRow: ' + trgRow)

    sheet.getRange(trgRow, 1).setValue(name);
    sheet.getRange(trgRow, 2).setFormula("=sum(Y" + trgRow + ":AC" + trgRow + ")");
    sheet.getRange(trgRow, trueLength + 3).setFormula("=COUNTA(C" + trgRow + ":W" + trgRow + ")");
    // sheet.getRange(trgRow, trueLength + 4).setFormula("=sum(C" + trgRow + ":U" + trgRow + ")*5000");
    // sheet.getRange(trgRow, trueLength + 5).setFormula("=sum(V" + trgRow + ":W" + trgRow + ")*4000");
    // sheet.getRange(trgRow, trueLength + 6).setFormula("=if(X" + trgRow + ">=3,-1000*(X" + trgRow + "-2),0)");
    sheet.getRange(trgRow, trueLength + 8).setFormula("=sum(C" + trgRow + ":U" + trgRow + ")*2000");
    sheet.insertRowAfter(trgRow);

    for (var i = 0; i < trueLength; i++) {
      // console.log('i:' + i)
      if (allClassList[i] == oneClass) {
        sheet.getRange(trgRow, i + 3).setValue(1)
      }
    }
    return true;

  } else {
    var names = sheet.getRange('A1:A').getValues();
    var nameList = Array.prototype.concat.apply([], names);

    var trgRow = nameList.indexOf("ママさんクラス"); // returns number
    console.log('trgRow: ' + trgRow)


    sheet.getRange(trgRow, 1).setValue(name)
    sheet.getRange(trgRow, 2).setFormula("=sum(Y" + trgRow + ":AC" + trgRow + ")");
    sheet.getRange(trgRow, trueLength + 3).setFormula("=COUNTA(C" + trgRow + ":W" + trgRow + ")");
    sheet.getRange(trgRow, trueLength + 4).setFormula("=sum(C" + trgRow + ":U" + trgRow + ")*5000");
    sheet.getRange(trgRow, trueLength + 5).setFormula("=sum(V" + trgRow + ":W" + trgRow + ")*4000");
    sheet.getRange(trgRow, trueLength + 6).setFormula("=if(X" + trgRow + ">=3,-1000*(X" + trgRow + "-2),0)");
    sheet.insertRowAfter(trgRow);

    for (var i = 0; i < trueLength; i++) {
      // console.log('i:' + i)
      if (allClassList[i] == oneClass) {
        sheet.getRange(trgRow, i + 3).setValue(1)
      }
    }
    return false;
  }
}

function SearchCount(allClassList) {
  for (var i = 0; i < allClassList.length; i++) {
    // console.log('allClassList[i]: ' + i)
    if (allClassList[i] === "合計") {
      console.log('合計　＝ ' + i)
      return i;
    }
  }
}

function ClassMeibo(oneClass, name, furigana) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("クラス名簿");
  // -1
  var allClasses = sheet.getRange('B2:B').getValues();
  var allClassList = Array.prototype.concat.apply([], allClasses);
  // console.log(allClassList)
  var allClassLength = allClassList.length
  // console.log(allClassLength)

  for (var i = 0; i < allClassLength; i++) {
    // console.log('ClassMeibo ' + i)
    if (allClassList[i] === oneClass) {
      console.log(allClassList[i])
      PutName(sheet, i + 2, name, furigana)
    }
  }
  return;
}

function PutName(sheet, trgRow, name, furigana) {
  // console.log("trgRow:" + trgRow)
  var lastRow = sheet.getLastRow() + 1;
  console.log("PutName:lastRow = " + lastRow)

  for (var i = trgRow; i <= lastRow; i++) {
    console.log('PutName:i = ' + i)
    // console.log(sheet.getRange(i, 3).getValue())
    if (sheet.getRange(i, 3).isBlank() === true) {
      sheet.getRange(i, 3).setValue(name);
      sheet.getRange(i, 4).setValue(furigana);
      sheet.insertRowAfter(i);
      // console.log('Done:PutName')
      return;
    }
  }
}

function SearchMamaLine(sheet) {
  var nameList = sheet.getRange("A1:A").getValues();
  console.log(nameList)
  var listLength = nameList.length;
  console.log(listLength)

  for (var i = 0; i < listLength; i++) {
    if (nameList[i] == "ママさんクラス") {
      return i;
    }
  }
}


function printError(error) {

  var mailTitle = "エラー通知";
  var mailText = "Bridge Over dance studio 顧客名簿\n\n" +
    "[名前] " + error.name + "\n\n" +
    "[場所] " + error.fileName + "(" + error.lineNumber + "行目)\n\n" +
    "[メッセージ]" + error.message + "\n\n" +
    "[StackTrace]\n" + error.stack;

  GmailApp.sendEmail('yuma.tanaka@openstore-japan.com', mailTitle, mailText);

  return "[名前] " + error.name + "\n" +
    "[場所] " + error.fileName + "(" + error.lineNumber + "行目)\n" +
    "[メッセージ]" + error.message + "\n" +
    "[StackTrace]\n" + error.stack;
}

function demo() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("顧客名簿");
  var nameList = sheet.getRange("A1:A").getValues();
  console.log(nameList)
  var listLength = nameList.length;
  console.log(listLength)

  for (var i = 0; i < listLength; i++) {
    if (nameList[i] == "ママさんクラス") {
      var mamaLine = i;
      console.log(mamaLine)
      return mamaLine;
    }
  }
}

