/******************************
 * CONFIGURATION
 ******************************/
const CONFIG = {
  calendarId: 'primary',            // 'primary' or your calendar ID
  useEmoji: true,                   // Add 🎂 emoji to titles
  useRecurrence: true,              // true = recurring yearly; false = one-time
  showYear: true,                   // true = (*1988); false = (36)
  reminderMinutesBefore: 1 * 24 * 60, // Popup 1 day before
  recurrenceYears: 50,              // How many years to repeat recurring events

  useTrigger: true,                 // Automatically run on a schedule
  triggerFrequency: 'daily',        // 'daily' or 'hourly'
  triggerHour: 4                    // When to run if 'daily' (0–23)
};

/**
 * Main function to create/update birthday events.
 * Also manages the optional time-based trigger.
 */
function loopThroughContacts() {
  if (CONFIG.useTrigger) {
    ensureTriggerExists();
  } else {
    removeTriggerIfExists();
  }

  const connections = getAllContacts();
  for (let i = 0; i < connections.length; i++) {
    const person = connections[i];
    if (person.birthdays && person.birthdays.length > 0) {
      let birthdayData = null;
      for (let j = 0; j < person.birthdays.length; j++) {
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
 * Get all Google Contacts using pagination.
 */
function getAllContacts() {
  const connections = [];
  let nextPageToken;

  do {
    const reqOpts = {
      personFields: 'names,birthdays',
      sortOrder: 'LAST_NAME_ASCENDING',
      pageSize: 100,
      pageToken: nextPageToken
    };

    const response = People.People.Connections.list('people/me', reqOpts);
    connections.push(...(response.connections || []));
    nextPageToken = response.nextPageToken;
  } while (nextPageToken);

  return connections;
}

/**
 * Create or update a calendar event for a contact's birthday.
 */
function updateOrCreateBirthDayEvent(person, birthdayDate) {
  const contactName = getContactName(person);
  const calendar = CalendarApp.getCalendarById(CONFIG.calendarId);

  if (!calendar) {
    Logger.log("⚠️ Calendar not found.");
    return;
  }

  const nextBirthday = calculateNextBirthday(birthdayDate);
  const startDate = new Date(nextBirthday.getFullYear(), nextBirthday.getMonth(), nextBirthday.getDate());
  const age = birthdayDate.year ? nextBirthday.getFullYear() - birthdayDate.year : null;

  let title = '';
  if (CONFIG.useEmoji) title += '🎂 ';
  title += contactName;
  if (birthdayDate.year) {
    title += CONFIG.showYear ? ` (*${birthdayDate.year})` : ` (${age})`;
  }

  const existingEvent = findBirthdayEvent(calendar, contactName, nextBirthday);
  if (existingEvent) {
    const isRecurring = existingEvent.isRecurringEvent && existingEvent.isRecurringEvent();
    if (
      existingEvent.getTitle() === title &&
      existingEvent.isAllDayEvent() &&
      (!CONFIG.useRecurrence || isRecurring)
    ) {
      Logger.log(`✅ No change needed for: ${title}`);
      return;
    }
    existingEvent.deleteEvent();
    Logger.log(`🗑️ Deleted outdated event: ${existingEvent.getTitle()}`);
  }

  if (CONFIG.useRecurrence) {
    const recurrence = CalendarApp.newRecurrence()
      .addYearlyRule()
      .until(new Date(nextBirthday.getFullYear() + CONFIG.recurrenceYears, 11, 31));

    const eventSeries = calendar.createAllDayEventSeries(
      title,
      startDate,
      recurrence,
      { description: `🎂 Happy Birthday ${contactName}` }
    );
    eventSeries.addPopupReminder(CONFIG.reminderMinutesBefore);
    Logger.log(`🎉 Created RECURRING event: ${title}`);
  } else {
    const event = calendar.createAllDayEvent(
      title,
      startDate,
      { description: `🎂 Happy Birthday ${contactName}` }
    );
    event.addPopupReminder(CONFIG.reminderMinutesBefore);
    Logger.log(`🎁 Created ONE-TIME event: ${title}`);
  }
}

/**
 * Extract name from contact.
 */
function getContactName(person) {
  if (person.names && person.names.length > 0) {
    return person.names[0].displayName ||
      `${person.names[0].givenName} ${person.names[0].familyName}`.trim();
  }
  return "Unknown";
}

/**
 * Check if birthday event already exists.
 */
function findBirthdayEvent(calendar, contactName, birthday) {
  const startDate = new Date(birthday.getFullYear(), birthday.getMonth(), birthday.getDate());
  const endDate = new Date(startDate);
  endDate.setDate(startDate.getDate() + 1);

  const events = calendar.getEvents(startDate, endDate);
  for (let i = 0; i < events.length; i++) {
    if (events[i].getTitle().includes(contactName)) {
      return events[i];
    }
  }
  return null;
}

/**
 * Get upcoming birthday (this year or next).
 */
function calculateNextBirthday(birthdayDate) {
  const day = birthdayDate.day;
  const month = birthdayDate.month - 1;
  const today = new Date();
  let year = today.getFullYear();
  const birthdayThisYear = new Date(year, month, day);
  if (today >= birthdayThisYear) {
    year++;
  }
  return new Date(year, month, day);
}

/**
 * Ensure a correct time-based trigger exists for loopThroughContacts.
 * Removes old one and creates a new one according to CONFIG.
 */
function ensureTriggerExists() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'loopThroughContacts') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log("🗑️ Removed outdated trigger for loopThroughContacts");
    }
  }

  const builder = ScriptApp.newTrigger('loopThroughContacts').timeBased();
  if (CONFIG.triggerFrequency === 'hourly') {
    builder.everyHours(1);
    Logger.log("✅ Created hourly trigger");
  } else {
    builder.everyDays(1).atHour(CONFIG.triggerHour);
    Logger.log(`✅ Created daily trigger at ${CONFIG.triggerHour}:00`);
  }

  builder.create();
}

/**
 * Remove time-based trigger if disabled in config.
 */
function removeTriggerIfExists() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'loopThroughContacts') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log("🗑️ Removed existing trigger");
    }
  }
}
