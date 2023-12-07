const threadsInPasma = 8
const alertBackgroundColor = "red"

const rawNumberCell = 0
const rawLengthCell = 1
const resultNumberCell = "D"
const resultLengthCell = "E"
const resultPasmasCell = "F"

function countLength() {
  var result = {}
  var clearColors = []

  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();

  for (var i = 0; i < data.length; i++) {
    var thread_number = data[i][rawNumberCell].toString().trim()
    var thread_length = parseInt(parseFloat(data[i][rawLengthCell]) * 10)

    clearDataRow(sheet, i+1)

    if (isBlend(thread_number)) {
      var colors = parseBlendNumbers(thread_number)
      var colors_amount = colors.length

      while (thread_length % colors_amount != 0)
        thread_length += 1

      color_length = thread_length/colors_amount

      for (var j = 0; j < colors_amount; j++) {
        thr = colors[j]
        if (!(thr in result))
          result[thr] = 0
        result[thr] += color_length
      }
    }
    else {
      if (!(thread_number in result))
        result[thread_number] = 0
      result[thread_number] += thread_length

      clearColors.push(thread_number)
    }
  }

  // print result
  var i = 1;
  for (var key in result) {
    setCellValue(sheet, resultNumberCell + i, key)
    if (!clearColors.includes(key))
      setCellAlertColor(sheet, resultNumberCell + i)

    var threads_float = parseFloat(result[key])/10
    setCellValue(sheet, resultLengthCell + i, threads_float)

    var amount = 1 + parseInt(threads_float/threadsInPasma)
    setCellValue(sheet, resultPasmasCell + i, amount)

    i++
  }
}

function clearDataRow(sheet, i) {
    sheet.getRange(resultNumberCell + (i)).setValue(null).setBackground(null);
    sheet.getRange(resultLengthCell + (i)).setValue(null);
    sheet.getRange(resultPasmasCell + (i)).setValue(null);
}

function isBlend(thread_number) {
  return thread_number.indexOf('+') > -1
}

function parseBlendNumbers(thread_number) {
  return thread_number.split("+").map(function(item) {
    return item.trim();
  })
}

function setCellValue(sheet, cell, value) {
  sheet.getRange(cell).setValue(value);
}

function setCellAlertColor(sheet, cell) {
  sheet.getRange(cell).setBackground(alertBackgroundColor)
}
