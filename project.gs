function myFunction() {

  // Gathering names and bdays from my 'Birthday Reminders' spreadsheet
  var sheet = SpreadsheetApp.getActiveSheet()
  var nameCol = 1
  var dateCol = 0
  var startRow = 2

  var dataRange = sheet.getRange(startRow, dateCol+1, 50, 2)
  var data = dataRange.getValues()  

  // Looping through each row
  for (var rowNum in data) {
    var row = data[rowNum]

    // Skipping a row if there is no date (for ex. if multiple prezzies are purchased)
    if (!row[dateCol]){continue}

    // Figure out how many days until a birthday
    var today = new Date()
    var bday = row[dateCol]
    bday.setFullYear(today.getFullYear())

    var dayAfterBday = new Date(bday.getTime()+1000*3600*24)
    if (today > dayAfterBday) {
      bday.setFullYear(today.getFullYear()+1)
    }
    var daysBetween = Math.ceil((bday - today)/(24*3600*1000))

    // Figure out which cute message, if any, to send to myself
    var message = null
    var subject = null
    switch (daysBetween){
      case 14:
        const bdayformatted=bday.toLocaleDateString('en-us',{weekday:"long",month:"numeric",day:"numeric"})
        var message = (
          "Hi, cutie!\n\n" +
          `${row[nameCol]} birthday is coming up in two weeks on ${bdayformatted}.\n\n` +
          "Start thinking about a gift.\n\n" +
          "Love you!"
        )
        var subject = "A loved one's bday is in two weeks!"
        break
      case 10:
        var message = (
          "Hi, little one -\n\n" +
          `${row[nameCol]} birthday is in ten days...\n\n` +
          "Better buy that card / gift!\n\n" +
          "xo"
        )
        var subject = "Somebody's bday is in ten short days!"
        break
      case 7:
        var message = (
          "Allyson -\n\n" +
          `One week until ${row[nameCol]} birthday.\n\n` +
          "Put treats in the mail ASAP!"
        )
        var subject = "One week til a special day..."
        break
      case 3:
        var message = (
          "Henlo,\n\n" +
          "Time is running short. Last call for card or gift. Pls advise.\n\n"
        )
        var subject = `${row[nameCol]} bday is in three short days.`
        break
      case 0:
        var message = (
          "Hi, babelah,\n\n" +
          `It's ${row[nameCol]} birthday today!\n\n` +
          "Don't forget to tell them how you love them.\n\n" +
          "Love you, too."
        )
        var subject = `It's ${row[nameCol]} birthday! Yay!`
        break
      default:
    }

    // To whomst to send the message
    if(message){
      MailApp.sendEmail("ENTER PREFERRED EMAIL HERE", subject, message)
    }
  }
}
