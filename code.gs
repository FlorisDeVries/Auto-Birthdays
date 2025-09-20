/******************************
 * CONFIGURATION
 ******************************/
const CONFIG = {
  calendarId: 'primary',            // 'primary' or your calendar ID

  // Title customization
  useEmoji: true,                    // Add 🎂 emoji to event titles
  showYearOrAge: true,               // Recurrence on: shows (*YYYY), off: shows (age)
  showAgeOnRecurring: false,         // If true, shows (age) on recurring events instead of (*YYYY)
  
  // Language and localization
  language: 'en',                    // Language code: 'en' (English), 'it' (Italian), 'fr' (French), 'de' (German), 'es' (Spanish)
  titleFormat: '',                   // Title format template (see LANGUAGE_CONFIG for placeholders), e.g. {emoji}{name} ({ageOrYear})

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
 * Language configuration for localized event titles and descriptions
 */
const LANGUAGE_CONFIG = {
  en: {
    titleFormats: {
      'default': '{emoji}{name} ({ageOrYear})',
      'birthday': '{emoji}{name}\'s birthday - {age} years',
      'simple': '{emoji}{name} - {age}'
    },
    terms: {
      'age': 'age',
      'years': 'years',
      'year': 'year',
      'birthday': 'birthday',
      'happyBirthday': 'Happy Birthday'
    }
  },
  it: {
    titleFormats: {
      'default': '{emoji}Compleanno di {name} - {age} anni',
      'birthday': '{emoji}Compleanno di {name} - {age} anni', 
      'simple': '{emoji}{name} - {age} anni'
    },
    terms: {
      'age': 'età',
      'years': 'anni',
      'year': 'anno',
      'birthday': 'compleanno',
      'happyBirthday': 'Buon Compleanno'
    }
  },
  fr: {
    titleFormats: {
      'default': '{emoji}Anniversaire de {name} - {age} ans',
      'birthday': '{emoji}Anniversaire de {name} - {age} ans',
      'simple': '{emoji}{name} - {age} ans'
    },
    terms: {
      'age': 'âge',
      'years': 'ans',
      'year': 'an',
      'birthday': 'anniversaire',
      'happyBirthday': 'Joyeux Anniversaire'
    }
  },
  de: {
    titleFormats: {
      'default': '{emoji}Geburtstag von {name} - {age} Jahre',
      'birthday': '{emoji}Geburtstag von {name} - {age} Jahre',
      'simple': '{emoji}{name} - {age} Jahre'
    },
    terms: {
      'age': 'Alter',
      'years': 'Jahre',
      'year': 'Jahr',
      'birthday': 'Geburtstag',
      'happyBirthday': 'Alles Gute zum Geburtstag'
    }
  },
  es: {
    titleFormats: {
      'default': '{emoji}Cumpleaños de {name} - {age} años',
      'birthday': '{emoji}Cumpleaños de {name} - {age} años',
      'simple': '{emoji}{name} - {age} años'
    },
    terms: {
      'age': 'edad',
      'years': 'años',
      'year': 'año',
      'birthday': 'cumpleaños',
      'happyBirthday': 'Feliz Cumpleaños'
    }
  }
};

/**
 * Language configuration for localized event titles and descriptions.
 * Supports: English (en), Italian (it), French (fr), German (de), Spanish (es)
 * 
 * For configuration examples and full documentation, see README.md
 */

/**
 * Generate a localized title for a birthday event
 * @param {string} contactName - The name of the person
 * @param {number|null} age - The age of the person (null if unknown)
 * @param {number|null} birthYear - The birth year (null if unknown)  
 * @param {boolean} showYear - Whether to show birth year instead of age
 * @param {boolean} isRecurring - Whether this is a recurring event
 * @returns {string} The formatted title
 */
function generateLocalizedTitle(contactName, age, birthYear, showYear, isRecurring) {
  const langConfig = LANGUAGE_CONFIG[CONFIG.language] || LANGUAGE_CONFIG['en'];
  
  // Use titleFormat if provided, otherwise fall back to language default
  const formatTemplate = CONFIG.titleFormat || langConfig.titleFormats['default'];
  
  // Build replacement values
  const emoji = CONFIG.useEmoji ? '🎂 ' : '';
  let ageOrYear = '';
  let ageText = '';
  
  if (birthYear && CONFIG.showYearOrAge) {
    if (showYear) {
      ageOrYear = `*${birthYear}`;
      ageText = `*${birthYear}`;
    } else if (age !== null) {
      const yearWord = age === 1 ? langConfig.terms.year : langConfig.terms.years;
      ageOrYear = age.toString();
      ageText = `${age} ${yearWord}`;
    }
  }
  
  // Replace placeholders in the format template
  let title = formatTemplate
    .replace(/{emoji}/g, emoji)
    .replace(/{name}/g, contactName)
    .replace(/{ageOrYear}/g, ageOrYear)
    .replace(/{age}/g, age !== null ? age.toString() : '')
    .replace(/{ageText}/g, ageText)
    .replace(/{birthYear}/g, birthYear ? birthYear.toString() : '')
    .replace(/{years}/g, langConfig.terms.years)
    .replace(/{year}/g, langConfig.terms.year)
    .replace(/{birthday}/g, langConfig.terms.birthday);
  
  // Clean up empty age/year information - remove parentheses, dashes, etc. around empty values
  title = title
    .replace(/\s*\(\s*\)\s*/g, '')        // Remove empty parentheses like " () "
    .replace(/\s*\(\s*\*\s*\)\s*/g, '')   // Remove empty year parentheses like " (*) "
    .replace(/\s*-\s*\*?\s*$/g, '')       // Remove trailing " - " or " - *"
    .replace(/\s+/g, ' ');                // Clean up multiple spaces
    
  return title.trim();
}

/**
 * Generate a localized description for a birthday event
 * @param {string} contactName - The name of the person
 * @returns {string} The formatted description
 */
function generateLocalizedDescription(contactName) {
  const langConfig = LANGUAGE_CONFIG[CONFIG.language] || LANGUAGE_CONFIG['en'];
  const happyBirthdayText = langConfig.terms.happyBirthday;
  
  return `🎂 ${happyBirthdayText} ${contactName}\n\n[${CONFIG.scriptKey}]`;
}

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
  
  // Build expected title using localized function
  const age = birthdayDate.year ? currentYear - birthdayDate.year : null;
  const showYear = CONFIG.useRecurrence && !CONFIG.showAgeOnRecurring;
  const expectedTitle = generateLocalizedTitle(
    contactName,
    age,
    birthdayDate.year,
    showYear,
    CONFIG.useRecurrence
  );

  // FAST existence check
  const key = `${expectedTitle}|${month}|${day}`;
  const existingQuick = eventIndex.get(key);
  
  // For individual age events, we need to check differently
  const usingIndividualEvents = CONFIG.useRecurrence && CONFIG.showAgeOnRecurring && birthdayDate.year;
  
  if (!usingIndividualEvents && existingQuick &&
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

    // When using individual events for age display, we need to delete any existing recurring events
    const needsConversionToIndividual = usingIndividualEvents && event.isRecurringEvent && event.isRecurringEvent();

    if (isTitleOutdated || isNotAllDay || isRecurrenceMismatch || isNotFromScript || needsConversionToIndividual) {
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

  // For individual age events, check if all years already exist
  if (usingIndividualEvents) {
    let allYearsExist = true;
    for (let year = startYear; year <= currentYear + CONFIG.futureYears; year++) {
      const yearAge = year - birthdayDate.year;
      
      // Generate title for this specific year using localized function
      const yearTitle = generateLocalizedTitle(
        contactName,
        yearAge,
        birthdayDate.year,
        false, // Always show age for individual events
        false  // Individual events are not recurring
      );
      
      const yearKey = `${yearTitle}|${month}|${day}`;
      const yearEvent = eventIndex.get(yearKey);
      if (!yearEvent || !isEventCreatedByScript(yearEvent)) {
        allYearsExist = false;
        break;
      }
    }
    
    if (allYearsExist) {
      return; // All individual events already exist
    }
  }

  // Create event description using localized function
  const eventDescription = generateLocalizedDescription(contactName);

  // Create events - individual yearly events if showing age on recurring, otherwise use recurrence
  if (CONFIG.useRecurrence && CONFIG.showAgeOnRecurring && birthdayDate.year) {
    // Create individual events for each year to show correct age
    for (let year = startYear; year <= currentYear + CONFIG.futureYears; year++) {
      const yearAge = year - birthdayDate.year;
      const yearBirthdayDate = new Date(year, month, day);
      
      // Build title with age for this specific year using localized function
      const yearTitle = generateLocalizedTitle(
        contactName,
        yearAge,
        birthdayDate.year,
        false, // Always show age for individual events
        false  // Individual events are not recurring
      );
      
      const event = calendar.createAllDayEvent(
        yearTitle,
        yearBirthdayDate,
        { description: eventDescription }
      );
      event.addPopupReminder(CONFIG.reminderMinutesBefore);
      Logger.log(`🎁 Created individual event: ${yearTitle} [${yearBirthdayDate.toDateString()}]`);
    }
  } else if (CONFIG.useRecurrence) {
    // Create standard recurring event
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
    // Create single event for this year
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