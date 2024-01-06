// Loop through contacts that have a birthday
function loopThroughContacts() {
  var contacts = ContactsApp.getContacts();
  for (var i = 0; i < contacts.length; i++) {
    var contact = contacts[i];
    var birthdays = contact.getDates(ContactsApp.Field.BIRTHDAY);
    if (birthdays.length > 0) {
      var birthdayField = birthdays[0];
      updateOrCreateBirthDayEvent(contact, birthdayField);
    }
  }
}

function updateOrCreateBirthDayEvent(contact, birthdayField) {
  var contactName = contact.getFullName();
  var calendar = CalendarApp.getCalendarById("calendarID");
  var nextBirthday = calculateNextBirthday(birthdayField);

  var ageAtNextBirthday = nextBirthday.getFullYear() - birthdayField.getYear();
  var suffix = ageAtNextBirthday !== nextBirthday.getFullYear() ? "'s birthday (" + ageAtNextBirthday + ")" : "'s birthday";
  var title = contactName + suffix;

  var birthdayEvent = findBirthdayEvent(calendar, contactName, nextBirthday);
  
  if (birthdayEvent) {
    if (birthdayEvent.getTitle() == title && birthdayEvent.isAllDayEvent()) {
      return;
    }

    // Update the existing event
    birthdayEvent.setTitle(title);
    birthdayEvent.setAllDayDate(new Date(nextBirthday.getFullYear(), nextBirthday.getMonth(), nextBirthday.getDate()));

    var reminders = birthdayEvent.getPopupReminders();
    var oneWeekReminder = 7 * 24 * 60; // 7 days * 24 hours * 60 minutes
    if (!reminders.includes(oneWeekReminder)) {
      birthdayEvent.addPopupReminder(oneWeekReminder);
    }

    Logger.log("Updated event: " + title);
  } else {
    // Create new event
    var startTime = new Date(nextBirthday.getFullYear(), nextBirthday.getMonth(), nextBirthday.getDate());
    var allDayEvent = calendar.createAllDayEvent(title, startTime);

    // Add a reminder a week in advance
    allDayEvent.addPopupReminder(7 * 24 * 60); // 7 days * 24 hours * 60 minutes
    Logger.log("Created new event: " + title);
  }
}

function findBirthdayEvent(calendar, contactName, nextBirthday) {
  var startDate = new Date(nextBirthday.getFullYear(), nextBirthday.getMonth(), nextBirthday.getDate());
  var endDate = new Date(startDate);
  endDate.setDate(startDate.getDate() + 1);

  var events = calendar.getEvents(startDate, endDate);
  var searchPattern = contactName + "'s birthday";

  for (var i = 0; i < events.length; i++) {
    var eventTitle = events[i].getTitle();
    if (eventTitle.includes(searchPattern)) {
      return events[i];
    }
  }

  return null;
}

function calculateNextBirthday(birthdayField) {
  var monthNames = ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"];
  var day = birthdayField.getDay();
  var monthName = birthdayField.getMonth().toString(); // Convert to uppercase
  var month = monthNames.indexOf(monthName);

  var today = new Date();
  var nextBirthdayYear = today.getFullYear();
  
  var thisYearsBirthday = new Date(today.getFullYear(), month, day);
  if (today >= thisYearsBirthday) {
    nextBirthdayYear++;
  }

  var nextBirthDay = new Date(nextBirthdayYear, month, day);
  return nextBirthDay;
}// Loop through contacts that have a birthday
function loopThroughContacts() {
  var contacts = ContactsApp.getContacts();
  for (var i = 0; i < contacts.length; i++) {
    var contact = contacts[i];
    var birthdays = contact.getDates(ContactsApp.Field.BIRTHDAY);
    if (birthdays.length > 0) {
      var birthdayField = birthdays[0];
      updateOrCreateBirthDayEvent(contact, birthdayField);
    }
  }
}

function updateOrCreateBirthDayEvent(contact, birthdayField) {
  var contactName = contact.getFullName();
  var calendar = CalendarApp.getCalendarById("81d36cdf357fac88d8b93e80ce01bd929b8461b5c45b961a871a5fdf3522ee6a@group.calendar.google.com");
  var nextBirthday = calculateNextBirthday(birthdayField);
  var ageAtNextBirthday = nextBirthday.getFullYear() - birthdayField.getYear();
  var title = contactName + "'s birthday (" + ageAtNextBirthday + ")";

  var birthdayEvent = findBirthdayEvent(calendar, contactName, nextBirthday);
  
  if (birthdayEvent) {
    // Update the existing event
    if (birthdayEvent.getTitle() == title && birthdayEvent.isAllDayEvent()) {
      return;
    }

    birthdayEvent.setTitle(title);
    birthdayEvent.setAllDayDate(new Date(nextBirthday.getFullYear(), nextBirthday.getMonth(), nextBirthday.getDate()));

    var reminders = birthdayEvent.getPopupReminders(); // or getEmailReminders(), depending on your preference
    var oneWeekReminder = 7 * 24 * 60; // in minutes
    if (!reminders.includes(oneWeekReminder)) {
      birthdayEvent.addPopupReminder(oneWeekReminder);
    }

    Logger.log("Updated event: " + title);
  } else {
    // Create new event
    var startTime = new Date(nextBirthday.getFullYear(), nextBirthday.getMonth(), nextBirthday.getDate());
    var allDayEvent = calendar.createAllDayEvent(title, startTime);

    // Add a reminder a week in advance
    allDayEvent.addPopupReminder(7 * 24 * 60); // 7 days * 24 hours * 60 minutes
    Logger.log("Created new event: " + title);
  }// Loop through contacts that have a birthday
function loopThroughContacts() {
  var contacts = ContactsApp.getContacts();
  for (var i = 0; i < contacts.length; i++) {
    var contact = contacts[i];
    var birthdays = contact.getDates(ContactsApp.Field.BIRTHDAY);
    if (birthdays.length > 0) {
      var birthdayField = birthdays[0];
      updateOrCreateBirthDayEvent(contact, birthdayField);
    }
  }
}

function updateOrCreateBirthDayEvent(contact, birthdayField) {
  var contactName = contact.getFullName();
  var calendar = CalendarApp.getCalendarById("calendarID");
  var nextBirthday = calculateNextBirthday(birthdayField);

  var ageAtNextBirthday = nextBirthday.getFullYear() - birthdayField.getYear();
  var suffix = ageAtNextBirthday !== nextBirthday.getFullYear() ? "'s birthday (" + ageAtNextBirthday + ")" : "'s birthday";
  var title = contactName + suffix;

  var birthdayEvent = findBirthdayEvent(calendar, contactName, nextBirthday);
  
  if (birthdayEvent) {
    if (birthdayEvent.getTitle() == title && birthdayEvent.isAllDayEvent()) {
      return;
    }

    // Update the existing event
    birthdayEvent.setTitle(title);
    birthdayEvent.setAllDayDate(new Date(nextBirthday.getFullYear(), nextBirthday.getMonth(), nextBirthday.getDate()));

    var reminders = birthdayEvent.getPopupReminders();
    var oneWeekReminder = 7 * 24 * 60; // 7 days * 24 hours * 60 minutes
    if (!reminders.includes(oneWeekReminder)) {
      birthdayEvent.addPopupReminder(oneWeekReminder);
    }

    Logger.log("Updated event: " + title);
  } else {
    // Create new event
    var startTime = new Date(nextBirthday.getFullYear(), nextBirthday.getMonth(), nextBirthday.getDate());
    var allDayEvent = calendar.createAllDayEvent(title, startTime);

    // Add a reminder a week in advance
    allDayEvent.addPopupReminder(7 * 24 * 60); // 7 days * 24 hours * 60 minutes
    Logger.log("Created new event: " + title);
  }
}

function findBirthdayEvent(calendar, contactName, nextBirthday) {
  var startDate = new Date(nextBirthday.getFullYear(), nextBirthday.getMonth(), nextBirthday.getDate());
  var endDate = new Date(startDate);
  endDate.setDate(startDate.getDate() + 1);

  var events = calendar.getEvents(startDate, endDate);
  var searchPattern = contactName + "'s birthday";

  for (var i = 0; i < events.length; i++) {
    var eventTitle = events[i].getTitle();
    if (eventTitle.includes(searchPattern)) {
      return events[i];
    }
  }

  return null;
}

function calculateNextBirthday(birthdayField) {
  var monthNames = ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"];
  var day = birthdayField.getDay();
  var monthName = birthdayField.getMonth().toString(); // Convert to uppercase
  var month = monthNames.indexOf(monthName);

  var today = new Date();
  var nextBirthdayYear = today.getFullYear();
  
  var thisYearsBirthday = new Date(today.getFullYear(), month, day);
  if (today >= thisYearsBirthday) {
    nextBirthdayYear++;
  }

  var nextBirthDay = new Date(nextBirthdayYear, month, day);
  return nextBirthDay;
}
}

function findBirthdayEvent(calendar, contactName, nextBirthday) {
  var startDate = new Date(nextBirthday.getFullYear(), nextBirthday.getMonth(), nextBirthday.getDate());
  var endDate = new Date(startDate);
  endDate.setDate(startDate.getDate() + 1);

  var events = calendar.getEvents(startDate, endDate);
  var searchPattern = contactName + "'s birthday";

  for (var i = 0; i < events.length; i++) {
    var eventTitle = events[i].getTitle();
    if (eventTitle.includes(searchPattern)) {
      return events[i];
    }
  }

  return null;
}

function calculateNextBirthday(birthdayField) {
  var monthNames = ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"];
  var year = birthdayField.getYear();
  var day = birthdayField.getDay();
  var monthName = birthdayField.getMonth().toString(); // Convert to uppercase
  var month = monthNames.indexOf(monthName);

  var today = new Date();
  var nextBirthdayYear = today.getFullYear();
  
  var thisYearsBirthday = new Date(today.getFullYear(), month, day);
  if (today >= thisYearsBirthday) {
    nextBirthdayYear++;
  }

  var nextBirthDay = new Date(nextBirthdayYear, month, day);
  return nextBirthDay;
}