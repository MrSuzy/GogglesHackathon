function parseTime(timeInput) { 

  if (!timeInput) { 
    // Log the error or handle it by returning a default time, such as "00:00:00" 
    Logger.log('parseTime was called with undefined or null input. Returning default time "00:00:00".'); 
    return "00:00:00";  // Provide a default return that does not disrupt the function's callers 

  } 

  let timeStr = String(timeInput); 

  let formattedTime; 

  if (timeStr.match(/\d{1,2}:\d{2}:\d{2} \w{2}/)) { 

    let parts = timeStr.match(/(\d{1,2}):(\d{2}):(\d{2}) (\w{2})/); 

    let hours = parseInt(parts[1], 10); 

    let minutes = parseInt(parts[2], 10); 

    let period = parts[4]; 

 

    if (period === 'PM' && hours < 12) hours += 12; 

    else if (period === 'AM' && hours === 12) hours = 0; 

 

    formattedTime = hours.toString().padStart(2, '0') + ':' + minutes.toString().padStart(2, '0') + ':00'; 

  } else { 

    let match = timeStr.match(/(\d{2}:\d{2}:\d{2})/); 

    if (match) { 

      formattedTime = match[1]; 

    } else { 

      // Handle unexpected format by logging and returning a default value 

      Logger.log('Received an unexpected time format: ' + timeStr + '. Returning default time "00:00:00".'); 

      return "00:00:00";  // Default safe value if format is unrecognized 

    } 

  } 

  return formattedTime; 

} 

function getOrdinalSuffix(number) {
  var j = number % 10,
    k = number % 100;
  if (j == 1 && k != 11) {
    return number + "st";
  }
  if (j == 2 && k != 12) {
    return number + "nd";
  }
  if (j == 3 && k != 13) {
    return number + "rd";
  }
  return number + "th";
}

function sendInvoice() { 

  var responseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Response'); 

  var data = responseSheet.getDataRange().getValues(); 

  var cateringPricesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Catering'); 

  var decorPricesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Decorations'); 

  var musicPricesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Music'); 

  var accountingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Accounting'); 

  var availabilitySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Availability'); 

 

  var multipliers = { 

    'Basic': 1.1, 
    'Silver': 1.5, 
    'Gold': 2, 
    'Platinum': 2.5 

  }; 

 

  function getPrice(sheet, detail) { 

    var prices = sheet.getDataRange().getValues(); 

    for (var i = 1; i < prices.length; i++) { 
      if (prices[i][0] === detail) { 
        return prices[i][1]; 
      } 
    } 
    return 0; 
  } 

 

  function formatTime(timeStr) { 

    if (typeof timeStr !== 'string') { 
      throw new Error('formatTime expects a string input, got ' + typeof timeStr); 

    } 

 

    var parts = timeStr.match(/(\d{1,2}):(\d{2}):(\d{2}) (\w{2})/); 

    if (!parts) { 
      throw new Error('Invalid time format'); 

    } 

    var hours = parseInt(parts[1], 10); 
    var minutes = parseInt(parts[2], 10); 
    var period = parts[4]; 

 

    if (period === 'PM' && hours < 12) { 
      hours += 12; 

    } else if (period === 'AM' && hours === 12) { 
      hours = 0; 
    } 

 

    var date = new Date(); 

    date.setHours(hours); 
    date.setMinutes(minutes); 
    date.setSeconds(0);

    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'HH:mm:ss'); 

  } 

 

  function isRecordInAvailabilitySheet(eventDate, startTime, endTime) { 
    var records = availabilitySheet.getDataRange().getValues(); 

 

    // Check if the sheet is empty (only the header row is present) 

    if (records.length <= 1) { 
      return false; // No records to compare, no clashing events exist 

    } 

 

    for (var i = 1; i < records.length; i++) { 

      var row = records[i];    
      var recordStartTime = parseTime(row[1]);
      var recordEventTime= formatEventDate(row[0]); 

      Logger.log('Comparing event date and time: %s %s with record: %s time: %s', recordEventTime, startTime, row[0], recordStartTime); 

      if ( startTime === recordStartTime) { 
        return true; // Clashing event exists 
      } 

    } 
    return false; // No clashing event exists 

  } 

 

  function formatEventDate(dateStr) { 

    var date = new Date(dateStr); 
    var options = { year: 'numeric', month: 'long', day: 'numeric' }; 

    return date.toLocaleDateString('en-GB', options); 

  } 

 

function sendUnavailableEmail(email, name, eventDate, startTime, endTime) { 
    var formattedEventDate = formatEventDate(eventDate); 
    var subject = "Booking Unavailability Notice"; 

    var messageBody = "<p>Hi " + name + ",</p>" + 
                      "<p>Unfortunately, our services are already booked for the time slot you requested from " + startTime + " to " + endTime + " on " + formattedEventDate + ".</p>" + 
                      "<p><a href=\"https://forms.gle/SmKNQVjuxhzjfR3o7\">Please visit our booking page again to select a new time</a> or contact us directly for further assistance.</p>" +
                      "<p>Thank you for your understanding.</p>" + 
                      "<p>Best regards,<br>Snorkels Events Crew.</p>"; 

    MailApp.sendEmail(email, subject, "", {htmlBody: messageBody}); 
} 


 

  function formatField(field) { 
    return (field === "No" || field === "None" || field === "" || field === null) ? "-" : field; 
  } 

 

  for (var i = 1; i < data.length; i++) { 

    var row = data[i]; 
    var timestamp = row[0]; 
    var email = row[1]; 
    var name = row[2]; 
    var phoneNumber = row[3]; 
    var eventDate = row[4]; 
    var startTime = row[5]; 
    var endTime = row[6]; 
    var eventType = row[7]; 
    var eventVenue = row[8]; 
    var numberOfGuests = row[9]; 
    var cateringService = row[10]; 
    var cateringPackage = row[11]; 
    var serviceStyle = row[12]; 
    var menuOptions = row[13]; 
    var cuisineCategory = row[14]; 
    var additionalServicesCatering = row[16]; 
    var decorationService = row[17]; 
    var decorationPackage = row[18]; 
    var seatingArrangements = row[20]; 
    var floralArrangements = row[21]; 
    var lighting = row[22]; 
    var staging = row[23]; 
    var otherDecorations = row[24]; 
    var musicService = row[25]; 
    var musicPackage = row[26]; 
    var musicGenre = row[27]; 
    var eventHighlights = row[28];  
    var status = row[30]; 

    if (!row[5] || !row[6]) { 
      Logger.log('Start time or end time is undefined for row: ' + i); 
      continue; // Skip to the next iteration if times are undefined 
    } 

    var formattedStartTime = formatTime(String(row[5])); 
    var formattedEndTime = formatTime(String(row[6])); 

    if (status === "Confirmation Sent" || status === "Not Available") { 
      continue; // Skip rows that are already marked 
    } else if (status !== "Confirmation Sent" && status !== "Not Available") { 

      var formattedStartTime = formatTime(String(startTime)); 
      var formattedEndTime = formatTime(String(endTime)); 

      Logger.log('Formatted start time: %s, end time: %s', formattedStartTime, formattedEndTime); 

      if (isRecordInAvailabilitySheet(eventDate, formattedStartTime, formattedEndTime)) { 
        Logger.log('This date and time are booked, an unavailable notice will be sent'); 
        sendUnavailableEmail(email, name, eventDate, formattedStartTime, formattedEndTime); 
        responseSheet.getRange(i + 1, 31).setValue('Not Available'); 
        continue; 

      } else { 
        availabilitySheet.appendRow([eventDate, formattedStartTime, formattedEndTime, 'Confirmed']); 

        var totalPrice = 0; 
        var itemPrices = []; 
        // Catering price 

        if (cateringService !== "No") { 
          var cateringStylePrice = getPrice(cateringPricesSheet, serviceStyle) * multipliers[cateringPackage]; 
          if (cateringStylePrice > 0) { 
            totalPrice += cateringStylePrice; 
            itemPrices.push(["Service Style", serviceStyle, cateringStylePrice.toFixed(2)]); 
          } 

          var menuItems = menuOptions.split(', '); 
          var cuisineItems = cuisineCategory.split(', '); 
          var additionalItems = additionalServicesCatering.split(', '); 

          menuItems.forEach(function(item) { 
            var itemPrice = getPrice(cateringPricesSheet, item) * multipliers[cateringPackage]; 
            if (itemPrice > 0) { 
              totalPrice += itemPrice; 
              itemPrices.push(["Menu Option", item, itemPrice.toFixed(2)]); 
            } 
          }); 

          cuisineItems.forEach(function(item) { 
            var itemPrice = getPrice(cateringPricesSheet, item) * multipliers[cateringPackage]; 
            if (itemPrice > 0) { 
              totalPrice += itemPrice; 
              itemPrices.push(["Cuisine Category", item, itemPrice.toFixed(2)]); 
            } 
          }); 
          additionalItems.forEach(function(item) { 
            var itemPrice = getPrice(cateringPricesSheet, item) * multipliers[cateringPackage]; 
            if (itemPrice > 0) { 
              totalPrice += itemPrice; 
              itemPrices.push(["Additional Service", item, itemPrice.toFixed(2)]); 
            } 
          }); 

        } 
        // Decoration price 

        if (decorationService !== "No") { 
          var decorPackagePrice = getPrice(decorPricesSheet, decorationPackage) * multipliers[decorationPackage]; 
          if (decorPackagePrice > 0) { 
            totalPrice += decorPackagePrice; 
            itemPrices.push(["Decoration Package", decorationPackage, decorPackagePrice.toFixed(2)]); 
          } 
          var decorItems = [ 
            { name: "Seating Arrangements", value: seatingArrangements },  
            { name: "Floral Arrangements", value: floralArrangements },  
            { name: "Lighting", value: lighting },  
            { name: "Staging", value: staging },  
            { name: "Other Decorations", value: otherDecorations } 
          ]; 
          decorItems.forEach(function(item) { 
            if (item.value !== "-" && item.value !== "") { 
              var itemPrice = getPrice(decorPricesSheet, item.value) * multipliers[decorationPackage]; 
              if (itemPrice > 0) { 
                totalPrice += itemPrice; 
                itemPrices.push([item.name, item.value, itemPrice.toFixed(2)]); 
              } 

            } 

          }); 

        } 
        // Music price 
        if (musicService !== "No") { 
          var musicPackagePrice = getPrice(musicPricesSheet, musicPackage) * multipliers[musicPackage]; 
          if (musicPackagePrice > 0) { 
            totalPrice += musicPackagePrice; 
            itemPrices.push(["Music Package", musicPackage, musicPackagePrice.toFixed(2)]); 

          } 
          var musicItems = [ 
            { name: "Music Genre", value: musicGenre }, 
            { name: "Event Highlights", value: eventHighlights } 

          ]; 

          musicItems.forEach(function(item) { 
            if (item.value !== "-" && item.value !== "") { 
              var itemPrice = getPrice(musicPricesSheet, item.value) * multipliers[musicPackage]; 
              if (itemPrice > 0) { 
                totalPrice += itemPrice; 
                itemPrices.push([item.name, item.value, itemPrice.toFixed(2)]); 
              } 
            } 
          }); 

        } 
        // Format timestamp to just date (YYYY-MM-DD) 
        var date = new Date(timestamp); 
        var formattedDate = date.getFullYear() + '-' + ('0' + (date.getMonth() + 1)).slice(-2) + '-' + ('0' + date.getDate()).slice(-2); 
        var eventDateObj = new Date(eventDate); 
        var eventDateObj = new Date(eventDate); 
        var year = eventDateObj.getFullYear(); 
        var month = ('0' + (eventDateObj.getMonth() + 1)).slice(-2); 
        var day = ('0' + eventDateObj.getDate()).slice(-2); 
        var formattedEventDate = year + '-' + month + '-' + day; 
        var subject = "Your Event Booking Confirmation & Receipt"; 
        var body = "<div style='color: black; font-family: Arial, sans-serif;'>" + 

                  "Dear " + name + ",<br><br>" + 

                  "Thank you for choosing our event planning services. Here are the details of your request:<br><br>" + 

                  "<table border='1' style='border-collapse: collapse; width: 100%;'>" + 

                  "<tr><th style='padding: 8px; text-align: left; background-color: #f2f2f2;'>Request Date</th><td style='padding: 8px;' colspan='2'>" + formattedDate + "</td></tr>" + 

                  "<tr><th style='padding: 8px; text-align: left; background-color: #f2f2f2;'>Event Type</th><td style='padding: 8px;' colspan='2'>" + formatField(eventType) + "</td></tr>" + 

                  "<tr><th style='padding: 8px; text-align: left; background-color: #f2f2f2;'>Event Date</th><td style='padding: 8px;' colspan='2'>" + formattedEventDate + "</td></tr>" + 

                  "<tr><th style='padding: 8px; text-align: left; background-color: #f2f2f2;'>Start Time</th><td style='padding: 8px;' colspan='2'>" + formattedStartTime + "</td></tr>" + 

                  "<tr><th style='padding: 8px; text-align: left; background-color: #f2f2f2;'>End Time</th><td style='padding: 8px;' colspan='2'>" + formattedEndTime + "</td></tr>" + 

                  "<tr><th style='padding: 8px; text-align: left; background-color: #f2f2f2;'>Event Venue</th><td style='padding: 8px;' colspan='2'>" + formatField(eventVenue) + "</td></tr>" + 

                  "<tr><th style='padding: 8px; text-align: left; background-color: #f2f2f2;'>Number of Guests</th><td style='padding: 8px;' colspan='2'>" + formatField(numberOfGuests) + "</td></tr>" + 

                  "<tr><th style='padding: 8px; text-align: left; background-color: #f2f2f2;'>Item</th><th style='padding: 8px; text-align: left; background-color: #f2f2f2;'>Description</th><th style='padding: 8px; text-align: left; background-color: #f2f2f2;'>Price</th></tr>"; 
        // Add item prices to the email body 
        itemPrices.forEach(function(item) { 
          body += "<tr><td style='padding: 8px;'>" + item[0] + "</td><td style='padding: 8px;'>" + item[1] + "</td><td style='padding: 8px;'>MYR " + item[2] + "</td></tr>"; 

        }); 
        // Add total price to the email body 
        body += "<tr><th colspan='2' style='padding: 8px; text-align: left; background-color: #f2f2f2;'>Total Price</th><td style='padding: 8px;'>MYR " + totalPrice.toFixed(2) + "</td></tr>"; 
        body += "</table><br><br>" + 

                  "We have confirmed the availability of the time slot and items you requested.<br><br>" + 

                  "Best regards,<br>Snorkels Events Crew.</div>";
        // Send email 
        Logger.log('Attempting to send email to: ' + email); 

        MailApp.sendEmail({ 
          to: email, 
          subject: subject, 
          htmlBody: body 
        }); 
        responseSheet.getRange(i + 1, 31).setValue('Confirmation Sent');  
        Logger.log('Email sent for ' + email); 
        // Add event to Accounting sheet to track due payment details 
        accountingSheet.appendRow([formattedEventDate, email, totalPrice.toFixed(2), "0", "Unpaid", formattedStartTime]); 
        setupMinuteTrigger();
        
      } 

    } 

  }     

} 

 

function sendPaymentReminder() { 
  var accountingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Accounting'); 
  var responseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Response'); 
  var responseRows = responseSheet.getDataRange().getValues(); 
  var rows = accountingSheet.getDataRange().getValues(); 
  var now = new Date(); 

  Logger.log('Current date and time: ' + now); 

  function getCustomerName(email) { 
    for (var j = 1; j < responseRows.length; j++) { 
      if (responseRows[j][1] === email) { 
        return responseRows[j][2]; 
      } 
    } 
    return "Customer"; // Default if no name found 
  } 

  function sendReminderEmail(email, name, amountDue, reminderCount) { 
    var ordinalReminderCount = getOrdinalSuffix(reminderCount);
    var subject = ordinalReminderCount + " Payment Reminder";
    var body = "<div style='color: black; font-family: Arial, sans-serif;'>" + 
               "Dear " + name + ",<br><br>" + 
               "This is the " + ordinalReminderCount + " reminder that your payment of MYR " + amountDue + " is due soon. Please ensure your payment is made promptly to avoid any disruptions to your event.<br><br>";

    body += "Thank you!<br><br>Best regards,<br>Snorkels Events Crew.</div>"; 

    MailApp.sendEmail({ 
      to: email, 
      subject: subject, 
      htmlBody: body 
    }); 

    Logger.log('Reminder sent to: ' + email + ' for amount: ' + amountDue + ' (Reminder count: ' + reminderCount + ')'); 
  } 

  for (var i = 1; i < rows.length; i++) { 
    var row = rows[i]; 
    var email = row[1]; 
    var amountDue = row[2]; 
    var remindersSent = parseInt(row[3]); 
    var amountReceived = row[4]; 
    var startingTime = row[5]; 

    Logger.log('Row ' + (i + 1) + ': ' + JSON.stringify(row)); 

    var name = getCustomerName(email); 
    Logger.log('Customer name: ' + name); 

    if (amountReceived === "Unpaid" && remindersSent < 5) { 
      var eventStartTime = new Date(startingTime); 
      var timeDiff = eventStartTime - now; 
      var hoursDiff = timeDiff / (1000 * 60 * 60); 
  
      Logger.log('Sending reminder ' + (remindersSent+1) + ' for ' + name); 
      sendReminderEmail(email, name, amountDue, remindersSent + 1); 
      accountingSheet.getRange(i + 1, 4).setValue(remindersSent + 1); 
      Logger.log('Reminder updated for row ' + (i + 1)); 
    } else { 
      Logger.log('No reminder needed for row ' + (i + 1) + ' for ' + email); 
    } 
  } 
} 

function setupMinuteTrigger() { 
  var allTriggers = ScriptApp.getProjectTriggers(); 
    if (allTriggers.length === 0) { 
    ScriptApp.newTrigger('sendPaymentReminder')
      .timeBased()
      .everyMinutes(1)
      .create(); 
  }
} 


 

function onOpen() { 
  var ui = SpreadsheetApp.getUi(); 
  ui.createMenu('Event Planner') 
    .addItem('Send Invoice', 'sendInvoice') 
    .addItem('Send Payment Reminder', 'sendPaymentReminder') 
    .addToUi(); 
  setupMinuteTrigger(); 
} 

 


