/******************************
 * CONFIGURATION
 ******************************/
const CONFIG = {
  calendarId: 'primary',           // Replace with your calendar ID or 'primary'
  useEmoji: true,                  // Add 🎂 to title
  useRecurrence: true,            // Create yearly recurring events
  showYear: true,                 // true: (*1988), false: (36)
  reminderMinutesBefore: 1 * 24 * 60, // 1 day before
  recurrenceYears: 50,            // For how many years should recurrence go
};

/******************************
 * MAIN SCRIPT
 ******************************/

function loopThroughContacts() {
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
 * Main logic: create or update a birthday event.
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
  if (CONFIG.useEmoji) {
    title += '🎂 ';
  }
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
 * Get best name available.
 */
function getContactName(person) {
  if (person.names && person.names.length > 0) {
    return person.names[0].displayName ||
      `${person.names[0].givenName} ${person.names[0].familyName}`.trim();
  }
  return "Unknown";
}

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
