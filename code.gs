/**
 * Main function to loop through contacts (people) that have a birthday
 * and create/update an all-day event for each birthday.
 */
function loopThroughContacts() {
  // Get all contacts with names and birthdays.
  var connections = getAllContacts()
  
  // Iterate over each contact.
  for (var i = 0; i < connections.length; i++) {
    var person = connections[i];
    
    // Check if the person has at least one birthday entry.
    if (person.birthdays && person.birthdays.length > 0) {
      var birthdayData = null;
      
      // Look for a birthday entry that has a 'date' object.
      for (var j = 0; j < person.birthdays.length; j++) {
        if (person.birthdays[j].date) {
          birthdayData = person.birthdays[j].date;
          break;
        }
      }
      
      if (birthdayData) {
        updateOrCreateBirthDayEvent(person, birthdayData);
      }
    }
  }
}

/**
 * Get all contacts, using pagination.
 */
function getAllContacts() {
  var connections = [];
  var nextPageToken = undefined;
  
  do {
    var reqOpts = {
      personFields: 'names,birthdays',
      sortOrder: 'LAST_NAME_ASCENDING',
      pageSize: 100,
    }

    if (nextPageToken != undefined) {
      reqOpts.pageToken = nextPageToken
    }

    var response = People.People.Connections.list('people/me', reqOpts);

    connections.push(...response.connections)
    nextPageToken = response.nextPageToken

  } while (nextPageToken != undefined)
  
  return connections
}

/**
 * Creates or updates a calendar event for the contact’s upcoming birthday.
 *
 * @param {Object} person The People API person object.
 * @param {Object} birthdayDate An object with day, month, and optionally year.
 */
function updateOrCreateBirthDayEvent(person, birthdayDate) {
  // Get the contact's name from the People API data.
  var contactName = "Unknown";
  if (person.names && person.names.length > 0) {
    // Use displayName if available, otherwise combine given and family names.
    contactName = person.names[0].displayName ||
                  (person.names[0].givenName + " " + person.names[0].familyName);
  }

  // Replace with your calendar ID.
  var calendarId = "";
  // Added warning if calendarId is not filled in.
  if (!calendarId) {
    Logger.log("Warning: calendarId is not filled in. Please change in script.");
    return;
  }
  
  var calendar = CalendarApp.getCalendarById(calendarId);
  
  // Calculate the next birthday date.
  var nextBirthday = calculateNextBirthday(birthdayDate);

  // Calculate the age at the next birthday if a birth year was provided.
  var ageAtNextBirthday = "";
  if (birthdayDate.year) {
    ageAtNextBirthday = nextBirthday.getFullYear() - birthdayDate.year;
  }
  
  // Build the event title.
  var suffix = (ageAtNextBirthday !== "" ? ("'s birthday (" + ageAtNextBirthday + ")") : "'s birthday");
  var title = contactName + suffix;

  var oneDayReminder = 1 * 24 * 60; // minutes in 1 day

  // See if an event already exists on the birthday date.
  var birthdayEvent = findBirthdayEvent(calendar, contactName, nextBirthday);
  
  if (birthdayEvent) {
    // If the title and all‑day setting are correct, no update is needed.
    if (birthdayEvent.getTitle() === title && birthdayEvent.isAllDayEvent()) {
      return;
    }
    // Update the existing event.
    birthdayEvent.setTitle(title);
    birthdayEvent.setAllDayDate(new Date(nextBirthday.getFullYear(), nextBirthday.getMonth(), nextBirthday.getDate()));
    
    // Add a popup reminder (1 day before) if not already present.
    var reminders = birthdayEvent.getPopupReminders();
    if (reminders.indexOf(oneDayReminder) === -1) {
      birthdayEvent.addPopupReminder(oneDayReminder);
    }
    Logger.log("Updated event: " + title);
  } else {
    // Create a new all‑day event.
    var startTime = new Date(nextBirthday.getFullYear(), nextBirthday.getMonth(), nextBirthday.getDate());
    var allDayEvent = calendar.createAllDayEvent(title, startTime);
    // Add a popup reminder (1 day before).
    allDayEvent.addPopupReminder(oneDayReminder);
    Logger.log("Created new event: " + title);
  }
}

/**
 * Searches for an existing birthday event on the calendar.
 *
 * @param {Calendar} calendar The CalendarApp calendar.
 * @param {string} contactName The contact's name to search for.
 * @param {Date} nextBirthday The date of the next birthday.
 * @return {CalendarEvent|null} The matching event, if found.
 */
function findBirthdayEvent(calendar, contactName, nextBirthday) {
  // Define the start and end of the birthday day.
  var startDate = new Date(nextBirthday.getFullYear(), nextBirthday.getMonth(), nextBirthday.getDate());
  var endDate = new Date(startDate);
  endDate.setDate(startDate.getDate() + 1);

  var events = calendar.getEvents(startDate, endDate);
  
  // Look for an event with the contact's name in its title.
  for (var i = 0; i < events.length; i++) {
    var eventTitle = events[i].getTitle();
    if (eventTitle.indexOf(contactName) !== -1) {
      return events[i];
    }
  }
  return null;
}

/**
 * Calculates the next birthday date given a People API birthday date object.
 *
 * Note: The People API returns months as 1–12 while JavaScript Date months are 0–11.
 *
 * @param {Object} birthdayDate An object with properties: day, month, (and optionally year).
 * @return {Date} The Date object for the next birthday.
 */
function calculateNextBirthday(birthdayDate) {
  var day = birthdayDate.day;
  var month = birthdayDate.month - 1; // Adjust because JavaScript Date months are 0-indexed.
  var today = new Date();
  var currentYear = today.getFullYear();
  
  // Create a date for this year's birthday.
  var thisYearsBirthday = new Date(currentYear, month, day);
  if (today >= thisYearsBirthday) {
    currentYear++;
  }
  
  return new Date(currentYear, month, day);
}
