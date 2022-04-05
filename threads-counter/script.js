function countLength() {
  var threads = {}
  var blends_only = []

  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();

  for (var i = 0; i < data.length; i++) {
    var thread_n = data[i][0].toString().trim()
    var thread_count = parseInt(parseFloat(data[i][1]) * 10)

    sheet.getRange("D" + (i+1)).setValue(null).setBackground(null);
    sheet.getRange("E" + (i+1)).setValue(null);
    sheet.getRange("F" + (i+1)).setValue(null);

    if (thread_n.indexOf('+') > -1) {
      // this is blend
      var arr = thread_n.split("+").map(function(item) {
        return item.trim();
      })

      if (thread_count % 2 != 0) {
        thread_count += 1
      }

      for (var j = 0; j < arr.length; j++) {
        thr = arr[j]
        if (!(thr in threads)) {
          threads[thr] = 0
          blends_only.push(thr)
        }
        threads[thr] += thread_count/2
      }
    }
    else {
      // this is pure color
      if (!(thread_n in threads))
        threads[thread_n] = 0
      threads[thread_n] += thread_count
    }
  }

  var i = 1;
  for (var key in threads) {
    sheet.getRange("D" + i).setValue(key);
    var threads_float = parseFloat(threads[key])/10
    sheet.getRange("E" + i).setValue(threads_float);
    if (blends_only.includes(key))
      sheet.getRange("D" + i).setBackground("red")
    var amount = 1
    amount += parseInt(threads_float/8)
    sheet.getRange("F" + i).setValue(amount);
    i++
  }
}
