/******************************
 * CONFIGURATION
 ******************************/
const CONFIG = {
  calendarId: 'primary',            // 'primary' or your calendar ID

  // Title customization
  useEmoji: true,                    // Add 🎂 emoji to event titles
  showYearOrAge: true,               // Recurrence on: shows (*YYYY), off: shows (age)

  // Recurrence
  useRecurrence: true,               // Create recurring yearly events
  futureYears: 10,                   // Recurring events end this many years in the future
  pastYears: 1,                      // Recurring events start this many years in the past

  // Reminder settings
  reminderMinutesBefore: 1440,       // Popup reminder (in minutes); 1440 = 1 day before

  // Cleanup
  cleanupEvents: false,              // ⚠️⚠️⚠️ Deletes all matching birthday events between ±100 years

  // Trigger options
  useTrigger: true,                  // Automatically run on a schedule
  triggerFrequency: 'daily',         // 'daily' or 'hourly'
  triggerHour: 4,                    // If 'daily', the hour of day to run (0–23)

  // Script identification
  scriptKey: 'CREATED_BY_Auto-Birthdays' // Unique identifier for events created by this script; customize if desired
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
  const calendar = CalendarApp.getCalendarById(CONFIG.calendarId);

  if (!calendar) {
    Logger.log("⚠️ Calendar not found.");
    return;
  }

  if (CONFIG.cleanupEvents) {
    cleanupOldBirthdayEvents(calendar, connections);
  }

  const currentYear = new Date().getFullYear();
  const startDate = new Date(currentYear - CONFIG.pastYears - 1, 0, 1);
  const endDate = new Date(currentYear + CONFIG.futureYears + 1, 11, 31);
  const allEvents = calendar.getEvents(startDate, endDate);
  const eventIndex  = buildBirthdayIndex(allEvents);

  for (const person of connections) {
    const birthdayData = person.birthdays?.find(b => b.date);
    if (birthdayData) {
      updateOrCreateBirthDayEvent(person, birthdayData, calendar, allEvents, eventIndex);
    }
  }
}

/**
 * Build a quick‑lookup map keying events by “title|month|day”.
 * Significantly reduces runtime from O(contacts × events) to O(contacts + events).
 */
function buildBirthdayIndex(events) {
  const map = new Map();
  for (const ev of events) {
    if (!ev.isAllDayEvent()) continue;
    const d   = ev.getStartTime();
    const key = `${ev.getTitle()}|${d.getMonth()}|${d.getDate()}`;
    map.set(key, ev);
  }
  return map;
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
function updateOrCreateBirthDayEvent(person, birthdayRaw, calendar, allEvents, eventIndex) {
  const contactName = getContactName(person);

  // Extract actual date from birthday object
  const birthdayDate = birthdayRaw.date;
  if (!birthdayDate || typeof birthdayDate.day !== 'number' || typeof birthdayDate.month !== 'number') {
    Logger.log(`⚠️ Skipping ${contactName} due to invalid birthdayDate: ${JSON.stringify(birthdayRaw)}`);
    return;
  }

  const currentYear = new Date().getFullYear();
  const startYear = currentYear - CONFIG.pastYears;

  const month = parseInt(birthdayDate.month, 10) - 1;
  const day = parseInt(birthdayDate.day, 10);

  const birthdayStartDate = new Date(startYear, month, day);
  const birthdayDateThisYear = new Date(currentYear, month, day);

  if (isNaN(birthdayStartDate.getTime()) || isNaN(birthdayDateThisYear.getTime())) {
    Logger.log(`⚠️ Invalid date for ${contactName}: ${birthdayDate.month}-${birthdayDate.day}`);
    return;
  }
  
  // Build expected title
  let expectedTitle = '';
  const age = birthdayDate.year ? currentYear - birthdayDate.year : null;
  if (CONFIG.useEmoji) expectedTitle += '🎂 ';
  expectedTitle += contactName;
  if (birthdayDate.year && CONFIG.showYearOrAge) {
    expectedTitle += CONFIG.useRecurrence ? ` (*${birthdayDate.year})` : ` (${age})`;
  }

  // FAST existence check
  const key = `${expectedTitle}|${month}|${day}`;
  const existingQuick = eventIndex.get(key);
  if (existingQuick &&
      existingQuick.isAllDayEvent() &&
      isEventCreatedByScript(existingQuick) &&
      (CONFIG.useRecurrence === (existingQuick.isRecurringEvent && existingQuick.isRecurringEvent()))) {
    return; // nothing to do
  }

  // Find and update related events
  const relatedEvents = findBirthdayEvents(allEvents, contactName, month, day);
  const deletedSeriesIds = new Set();
  let correctEventExists = false;

  for (const event of relatedEvents) {
    const title = event.getTitle();
    const isTitleOutdated = title !== expectedTitle;
    const isNotAllDay = !event.isAllDayEvent();
    const isRecurrenceMismatch = CONFIG.useRecurrence !== (event.isRecurringEvent && event.isRecurringEvent());
    const isNotFromScript = !isEventCreatedByScript(event);

    if (isTitleOutdated || isNotAllDay || isRecurrenceMismatch || isNotFromScript) {
      try {
        if (event.isRecurringEvent && event.isRecurringEvent()) {
          
          const series = event.getEventSeries();
          const seriesId = series.getId();
          if (deletedSeriesIds.has(seriesId)) {
            break;
          }
          deletedSeriesIds.add(seriesId);
          series.deleteEventSeries();
          Logger.log(`🗑️ Deleted outdated recurring series: ${title}`);
        } else {
          event.deleteEvent();
          Logger.log(`🗑️ Deleted outdated event: ${title}`);
        }
      } catch (e) {
        Logger.log(`❌ Error deleting event: ${title} → ${e}`);
      }
    } else {
      correctEventExists = true;
    }
  }

  if (correctEventExists) {
    return;
  }

  // Create event description with script key
  const eventDescription = `🎂 Happy Birthday ${contactName}\n\n[${CONFIG.scriptKey}]`;

  // Create one new event — recurring or one-time
  if (CONFIG.useRecurrence) {
    const recurrence = CalendarApp.newRecurrence()
      .addYearlyRule()
      .until(new Date(currentYear + CONFIG.futureYears, 11, 31));

    const eventSeries = calendar.createAllDayEventSeries(
      expectedTitle,
      birthdayStartDate,
      recurrence,
      { description: eventDescription }
    );
    eventSeries.addPopupReminder(CONFIG.reminderMinutesBefore);
    Logger.log(`🎉 Created RECURRING event: ${expectedTitle} [starts ${birthdayStartDate.toDateString()}]`);
  } else {
    const event = calendar.createAllDayEvent(
      expectedTitle,
      birthdayDateThisYear,
      { description: eventDescription }
    );
    event.addPopupReminder(CONFIG.reminderMinutesBefore);
    Logger.log(`🎁 Created ONE-TIME event: ${expectedTitle} [${birthdayDateThisYear.toDateString()}]`);
  }
}

/**
 * Check if an event was created by this script by looking for the script key in description.
 */
function isEventCreatedByScript(event) {
  try {
    const description = event.getDescription();
    return description && description.includes(`[${CONFIG.scriptKey}]`);
  } catch (e) {
    return false;
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
function findBirthdayEvents(allEvents, contactName, month, day) {
  return allEvents.filter(ev => {
    const title = ev.getTitle();
    const d     = ev.getStartTime();
    return ev.isAllDayEvent() &&
           (title.includes(contactName) && (d.getMonth() === month && d.getDate() === day));
  });
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
 * Cleans birthdayEvents in the future and past 100 years
 */
function cleanupOldBirthdayEvents(calendar, allContacts) {
  const currentYear = new Date().getFullYear();
  const startYear = currentYear - 100;
  const endYear = currentYear + 100;

  const startDate = new Date(startYear, 0, 1);
  const endDate = new Date(endYear, 11, 31);
  const allEvents = calendar.getEvents(startDate, endDate);
  
  Logger.log(`🧹 Cleanup started between: ${startDate.toDateString()} - ${endDate.toDateString()}`);

  // Collect valid contacts with birthday info
  const contactBirthdays = allContacts
    .map(person => {
      if (!person.names || !person.birthdays) return null;
      const name = getContactName(person);
      const bday = person.birthdays.find(b => b.date);
      if (!bday) return null;

      return {
        name,
        day: bday.date.day,
        month: bday.date.month - 1,
      };
    })
    .filter(Boolean); // remove nulls

  const deletedSeriesIds = new Set();
  const eventsToDelete = [];

  for (const event of allEvents) {
    const title = event.getTitle();
    const start = event.getStartTime();
    const isAllDay = event.isAllDayEvent();
    const startsWithCakeEmoji = title.trim().startsWith('🎂');
    const isFromScript = isEventCreatedByScript(event);

    for (const contact of contactBirthdays) {
      const isNameMatch = title.includes(contact.name);
      const isBirthdayDateMatch =
        start.getDate() === contact.day &&
        start.getMonth() === contact.month;

      const shouldDelete =
        isFromScript && // Only delete events created by this script
        ((startsWithCakeEmoji && isNameMatch) ||
         (isAllDay && isNameMatch && isBirthdayDateMatch));

      if (!shouldDelete) continue;

      // Check if it's a recurring event and already handled
      if (event.isRecurringEvent && event.isRecurringEvent()) {
        try {
          const series = event.getEventSeries();
          const seriesId = series.getId();
          if (deletedSeriesIds.has(seriesId)) {
            break;
          }
          deletedSeriesIds.add(seriesId);
          eventsToDelete.push({ event: series, isSeries: true });
          break;
        } catch (e) {
          Logger.log(`⚠️ Failed to resolve recurring series: ${title} — ${e}`);
          Logger.log(`⚠️ Waiting for rate limiter and safely retrying`);
          Utilities.sleep(5000);
          try {
            const series = event.getEventSeries();
            const seriesId = series.getId();
            if (deletedSeriesIds.has(seriesId)) {
              Logger.log(`✅ Second attempt succeeded for: ${title}`);
              break;
            }
            deletedSeriesIds.add(seriesId);
            eventsToDelete.push({ event: series, isSeries: true });
            Logger.log(`✅ Second attempt succeeded for: ${title}`);
            break;
          } catch (e) {
            Logger.log(`⚠️ Second attempt failed — skipping entry ${e}`);
          }
        }
      } else {
        eventsToDelete.push({ event, isSeries: false });
        break;
      }
    }
  }

  // Now delete all marked events
  for (const { event, isSeries } of eventsToDelete) {
    try {
      if (isSeries) {
        event.deleteEventSeries(); // Deletes the entire series
        Logger.log(`🧹 Deleted recurring series → ${event.getTitle()}`);
      } else {
        event.deleteEvent();
        Logger.log(`🧹 Deleted single event → ${event.getTitle()} on ${event.getStartTime().toDateString()}`);
      }
    } catch (e) {
      Logger.log(`❌ Failed to delete event → ${event.getTitle()} | Reason: ${e}`);
    }
  }

  Logger.log(`✅ Total deleted events: ${eventsToDelete.length}`);
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