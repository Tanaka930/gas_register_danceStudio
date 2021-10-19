function onFormSubmit(e) {

  //openstoreアカウントのアドレス
  COMPANY_MAIL = PropertiesService.getScriptProperties().getProperty("COMPANY_MAIL");

  const lock = LockService.getScriptLock();
  try {
    // ロックを取得する
    if (lock.tryLock(10 * 1000)) {

      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("顧客名簿");

      console.log(e.values)

      for (var i = 10; i <= 14; i++) {
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
          sheet.getRange(targetRow, 9).setValue(e.values[15]);
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

  var totalNumber = trueLength + 2

  var BaseAbc = NumberToAbc(String(totalNumber))
  console.log("BaseAbc:" + BaseAbc)

  var graylist = CountGray(allClassList)

  var mamaNum = graylist[2] - graylist[1] -1;
  console.log("mama" + mamaNum)
  var littleNum = graylist[1] - graylist[0] -1;
  console.log("littleNum" + littleNum)

  if (oneClass == "Jazz Yukana  一般(ママ)") {
    var trgRow = sheet.getLastRow() + 1;
    console.log('trgRow: ' + trgRow)

    sheet.getRange(trgRow, 1).setValue(name);
    // sheet.getRange(trgRow, 2).setFormula("=sum(Z" + trgRow + ":AD" + trgRow + ")");
    sheet.getRange(trgRow, 2).setFormula("=sum(" + NumberToAbc(String(totalNumber + 1)) + trgRow + ":" + NumberToAbc(String(totalNumber + 5)) + trgRow + ")");
    // sheet.getRange(trgRow, trueLength + 3).setFormula("=COUNTA(C" + trgRow + ":X" + trgRow + ")");
    sheet.getRange(trgRow, trueLength + 3).setFormula("=COUNTA(C" + trgRow + ":" + NumberToAbc(String(totalNumber - 2)) + trgRow + ")");
    // sheet.getRange(trgRow, trueLength + 8).setFormula("=sum(C" + trgRow + ":V" + trgRow + ")*2000");
    sheet.getRange(trgRow, trueLength + 8).setFormula("=sum(" + NumberToAbc(String(totalNumber - mamaNum -1)) + trgRow + ":" + NumberToAbc(String(totalNumber - 2)) + trgRow + ")*2000");
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
    // sheet.getRange(trgRow, 2).setFormula("=sum(Z" + trgRow + ":AC" + trgRow + ")");
    sheet.getRange(trgRow, 2).setFormula("=sum(" + NumberToAbc(String(totalNumber + 1)) + trgRow + ":" + NumberToAbc(String(totalNumber + 4)) + trgRow + ")");

    // sheet.getRange(trgRow, trueLength + 3).setFormula("=COUNTA(C" + trgRow + ":X" + trgRow + ")");
    sheet.getRange(trgRow, trueLength + 3).setFormula("=COUNTA(C" + trgRow + ":" + NumberToAbc(String(totalNumber - 2)) + trgRow + ")");

    // sheet.getRange(trgRow, trueLength + 4).setFormula("=sum(C" + trgRow + ":V" + trgRow + ")*5000");
    sheet.getRange(trgRow, trueLength + 4).setFormula("=sum(C" + trgRow + ":" + NumberToAbc(String(totalNumber - littleNum - mamaNum - 4)) + trgRow + ")*5000");

    // sheet.getRange(trgRow, trueLength + 5).setFormula("=sum(W" + trgRow + ":X" + trgRow + ")*4000");
    sheet.getRange(trgRow, trueLength + 5).setFormula("=sum(" + NumberToAbc(String(totalNumber - littleNum - mamaNum - 2)) + trgRow + ":" + NumberToAbc(String(totalNumber - mamaNum - 3)) + trgRow + ")*4000");

    // sheet.getRange(trgRow, trueLength + 6).setFormula("=if(Y" + trgRow + ">=3,-1000*(Y" + trgRow + "-2),0)");
    sheet.getRange(trgRow, trueLength + 6).setFormula("=if("+ NumberToAbc(String(totalNumber)) + trgRow + ">=3,-1000*(" + NumberToAbc(String(totalNumber)) + trgRow + "-2),0)");

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


function CountGray(allClassList) {
  var ary = new Array(); 

  for (var i = 0; i <= allClassList.length; i++) {
    if (allClassList[i] == "A" || allClassList[i] == "B" || allClassList[i] == "C") {
      ary.push(i)
    }
  }
  console.log("ary:" + ary)
  return ary;
}

// 数値をアルファベットに変更する関数
function NumberToAbc(iCol) {
  var str = "";
  var iAlpha = 0;
  var iRemainder = 0;

  iAlpha = parseInt((iCol / 26), 10);
  iRemainder = iCol - (iAlpha * 26);
  if (iAlpha > 0) {
    str = String.fromCharCode(iAlpha + 64);
  }
  if (iRemainder >= 0) {
    str = str + String.fromCharCode(iRemainder + 65);
  }
  return str;
}


function printError(error) {

  var mailTitle = "エラー通知";
  var mailText = "Bridge Over dance studio 顧客名簿\n\n" +
    "[名前] " + error.name + "\n\n" +
    "[場所] " + error.fileName + "(" + error.lineNumber + "行目)\n\n" +
    "[メッセージ]" + error.message + "\n\n" +
    "[StackTrace]\n" + error.stack;

  GmailApp.sendEmail(COMPANY_MAIL, mailTitle, mailText);

  return "[名前] " + error.name + "\n" +
    "[場所] " + error.fileName + "(" + error.lineNumber + "行目)\n" +
    "[メッセージ]" + error.message + "\n" +
    "[StackTrace]\n" + error.stack;
}

