/******************************
 * CONFIGURATION
 * Version: 1.0
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
  useReminders: true,                // Enable/disable reminders for birthday events
  reminderMinutesBefore: 1440,       // Popup reminder time (in minutes) - only used if useReminders is true
                                     // Common values: 0 = at event time, 60 = 1 hour before, 1440 = 1 day before, 10080 = 1 week before

  // Cleanup
  cleanupEvents: false,              // ⚠️⚠️⚠️ Deletes all matching birthday events between ±100 years

  // Trigger options
  useTrigger: true,                  // Automatically run on a schedule
  triggerFrequency: 'daily',         // 'daily' or 'hourly'
  triggerHour: 4,                    // If 'daily', the hour of day to run (0–23)

  // Script identification
  scriptKey: 'CREATED_BY_Auto-Birthdays', // Unique identifier for events created by this script; customize if desired

  // Contact label filtering (optional)
  useLabels: false,                  // Enable filtering contacts by labels
  contactLabels: [],                 // Array of contact label IDs to include (e.g. ['abc123'])

  // Month filtering (optional)
  useMonthFilter: false,             // Enable filtering contacts by birth month
  filterMonths: []                   // Array of months to include (1-12), e.g. [1, 2, 4] for Jan, Feb, Apr
};

/**
 * Language configuration for localized event titles and descriptions
 */
const LANGUAGE_CONFIG = {
  en: {
    titleFormats: {
      'default': '{emoji}{name} ({ageOrYear})'
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
      'default': '{emoji}Compleanno di {name} - {age} anni'
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
      'default': '{emoji}Anniversaire de {name} - {age} ans'
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
      'default': '{emoji}Geburtstag von {name} - {age} Jahre'
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
      'default': '{emoji}Cumpleaños de {name} - {age} años'
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
  // Validate reminder configuration
  if (CONFIG.useReminders && (typeof CONFIG.reminderMinutesBefore !== 'number' || CONFIG.reminderMinutesBefore < 0)) {
    Logger.log("⚠️ Invalid reminder configuration: reminderMinutesBefore must be a non-negative number when useReminders is true.");
    return;
  }

  const calendar = CalendarApp.getCalendarById(CONFIG.calendarId);

  if (!calendar) {
    Logger.log("⚠️ Calendar not found.");
    return;
  }

  if (CONFIG.cleanupEvents) {
    Logger.log("🧹 Starting cleanup of old birthday events...");
    cleanupOldBirthdayEvents(calendar, connections);
    Logger.log(""); // Empty line for separation
  }

  Logger.log("📊 Starting birthday event processing...");

  const currentYear = new Date().getFullYear();
  const startDate = new Date(currentYear - CONFIG.pastYears - 1, 0, 1);
  const endDate = new Date(currentYear + CONFIG.futureYears + 1, 11, 31);
  const allEvents = calendar.getEvents(startDate, endDate);
  const eventIndex  = buildBirthdayIndex(allEvents);

  // Initialize counters for processing report
  let totalContacts = connections.length;
  let contactsWithBirthdays = 0;
  let processedContacts = 0;
  let eventsCreated = 0;
  let eventsUpdated = 0;
  let skippedByLabelFilter = 0;
  let skippedByMonthFilter = 0;
  let skippedInvalidBirthdays = 0;

  for (const person of connections) {
    // Check if label filtering is enabled and if this contact has the required labels
    if (CONFIG.useLabels && !hasRequiredLabel(person, CONFIG.contactLabels)) {
      skippedByLabelFilter++;
      continue; // Skip this contact if it doesn't have the required labels
    }

    // Check if month filtering is enabled and if this contact's birthday month matches
    if (CONFIG.useMonthFilter && !hasMatchingBirthMonth(person, CONFIG.filterMonths)) {
      skippedByMonthFilter++;
      continue; // Skip this contact if its birthday month doesn't match the filter
    }

    const birthdayData = person.birthdays?.find(b => b.date);
    if (birthdayData) {
      contactsWithBirthdays++;
      try {
        const result = updateOrCreateBirthDayEvent(person, birthdayData, calendar, allEvents, eventIndex);
        if (result === 'created') {
          processedContacts++;
          eventsCreated++;
        } else if (result === 'updated') {
          processedContacts++;
          eventsUpdated++;
        } else if (result === 'skipped_existing') {
          processedContacts++;
          // Event already exists and is correct - no action needed
        } else if (result === 'skipped_invalid') {
          skippedInvalidBirthdays++;
        }
      } catch (error) {
        const contactName = getContactName(person);
        Logger.log(`❌ Error processing ${contactName}: ${error}`);
        skippedInvalidBirthdays++;
      }
    }
  }

  // Log processing summary report
  Logger.log("📊 PROCESSING SUMMARY REPORT:");
  Logger.log(`📞 Total contacts retrieved: ${totalContacts}`);
  Logger.log(`🎂 Contacts with birthday data: ${contactsWithBirthdays}`);
  Logger.log(`✅ Successfully processed contacts: ${processedContacts}`);
  Logger.log(`🆕 Events created: ${eventsCreated}`);
  Logger.log(`🔄 Events updated: ${eventsUpdated}`);
  
  const eventsAlreadyCorrect = processedContacts - eventsCreated - eventsUpdated;
  if (eventsAlreadyCorrect > 0) {
    Logger.log(`✓ Events already up-to-date: ${eventsAlreadyCorrect}`);
  }
  if (CONFIG.useLabels) {
    Logger.log(`🏷️ Contacts skipped by label filter: ${skippedByLabelFilter}`);
  }
  
  if (CONFIG.useMonthFilter) {
    Logger.log(`📅 Contacts skipped by month filter: ${skippedByMonthFilter}`);
  }
  
  if (skippedInvalidBirthdays > 0) {
    Logger.log(`⚠️ Contacts with invalid birthday data: ${skippedInvalidBirthdays}`);
  }
  
  const contactsWithoutBirthdays = totalContacts - contactsWithBirthdays - skippedByLabelFilter - skippedByMonthFilter;
  if (contactsWithoutBirthdays > 0) {
    Logger.log(`📝 Contacts without birthday data: ${contactsWithoutBirthdays}`);
  }
  
  // Show reminder configuration info
  if (CONFIG.useReminders) {
    const reminderText = CONFIG.reminderMinutesBefore === 0 ? "at event time" :
                        CONFIG.reminderMinutesBefore === 60 ? "1 hour before" :
                        CONFIG.reminderMinutesBefore === 1440 ? "1 day before" :
                        CONFIG.reminderMinutesBefore === 10080 ? "1 week before" :
                        `${CONFIG.reminderMinutesBefore} minutes before`;
    Logger.log(`⏰ Reminders enabled: ${reminderText}`);
  } else {
    Logger.log(`⏰ Reminders disabled`);
  }
  
  Logger.log("🎉 Processing completed successfully!");
}

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
    .replace(/{age}/g, '${age}')
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
  
  return `🎂 ${happyBirthdayText} ${contactName}!\n\n[${CONFIG.scriptKey}]`;
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
 * Check if a contact has any of the specified labels.
 * @param {object} person - The contact person object
 * @param {string[]} labelIds - Array of contact label IDs to check for (just the ID part, not the full resource name)
 * @returns {boolean} - True if contact has any of the specified labels
 */
function hasRequiredLabel(person, labelIds) {
  if (!labelIds || labelIds.length === 0) {
    return true; // No labels specified, so all contacts match
  }

  if (!person.memberships || person.memberships.length === 0) {
    return false; // Contact has no group memberships
  }

  // Check if any membership matches any of the specified label IDs
  for (const membership of person.memberships) {
    if (membership.contactGroupMembership && membership.contactGroupMembership.contactGroupId) {
      const contactGroupId = membership.contactGroupMembership.contactGroupId;
      
      // Check if this contactGroupId contains any of our target labels
      // contactGroupId format is typically "contactGroups/abc122"
      for (const labelId of labelIds) {
        if (contactGroupId.includes(labelId)) {
          return true;
        }
      }
    }
  }

  return false;
}

/**
 * Check if a contact's birthday month matches any of the specified filter months.
 * @param {object} person - The contact person object
 * @param {number[]} filterMonths - Array of months to include (1-12)
 * @returns {boolean} - True if contact's birthday month matches any filter month
 */
function hasMatchingBirthMonth(person, filterMonths) {
  if (!filterMonths || filterMonths.length === 0) {
    return true; // No month filter specified, so all contacts match
  }

  const birthdayData = person.birthdays?.find(b => b.date);
  if (!birthdayData || !birthdayData.date || typeof birthdayData.date.month !== 'number') {
    return false; // No valid birthday data
  }

  const birthMonth = birthdayData.date.month;
  return filterMonths.includes(birthMonth);
}

/**
 * Get all Google Contacts using pagination.
 */
function getAllContacts() {
  const connections = [];
  let nextPageToken;

  do {
    const reqOpts = {
      personFields: 'names,birthdays,memberships',
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
 * @returns {string} - Operation result: 'created', 'updated', 'skipped_existing', or 'skipped_invalid'
 */
function updateOrCreateBirthDayEvent(person, birthdayRaw, calendar, allEvents, eventIndex) {
  const contactName = getContactName(person);

  // Extract actual date from birthday object
  const birthdayDate = birthdayRaw.date;
  if (!birthdayDate || typeof birthdayDate.day !== 'number' || typeof birthdayDate.month !== 'number') {
    Logger.log(`⚠️ Skipping ${contactName} due to invalid birthdayDate: ${JSON.stringify(birthdayRaw)}`);
    return 'skipped_invalid';
  }

  const currentYear = new Date().getFullYear();
  const startYear = currentYear - CONFIG.pastYears;

  const month = parseInt(birthdayDate.month, 10) - 1;
  const day = parseInt(birthdayDate.day, 10);

  const birthdayStartDate = new Date(startYear, month, day);
  const birthdayDateThisYear = new Date(currentYear, month, day);

  if (isNaN(birthdayStartDate.getTime()) || isNaN(birthdayDateThisYear.getTime())) {
    Logger.log(`⚠️ Invalid date for ${contactName}: ${birthdayDate.month}-${birthdayDate.day}`);
    return 'skipped_invalid';
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

  // Generate expected description using localized function
  const expectedDescription = generateLocalizedDescription(contactName);

  // FAST existence check
  const key = `${expectedTitle}|${month}|${day}`;
  const existingQuick = eventIndex.get(key);
  
  // For individual age events, we need to check differently
  const usingIndividualEvents = CONFIG.useRecurrence && CONFIG.showAgeOnRecurring && birthdayDate.year;
  
  if (!usingIndividualEvents && existingQuick &&
      existingQuick.isAllDayEvent() &&
      isEventCreatedByScript(existingQuick) &&
      (CONFIG.useRecurrence === (existingQuick.isRecurringEvent && existingQuick.isRecurringEvent())) &&
      (existingQuick.getDescription() || '') === expectedDescription &&
      hasCorrectReminders(existingQuick) &&
      hasScriptTag(existingQuick)) {
    return 'skipped_existing'; // Event already exists and is correct
  }

  // Find and update related events
  const relatedEvents = findBirthdayEvents(allEvents, contactName, month, day);
  const deletedSeriesIds = new Set();
  let correctEventExists = false;
  let eventsWereDeleted = false;

  for (const event of relatedEvents) {
    const title = event.getTitle();
    let description;
    try {
      description = event.getDescription() || '';
    } catch (e) {
      description = '';
    }
    const isTitleOutdated = title !== expectedTitle;
    const isDescriptionOutdated = description !== expectedDescription;
    const isNotAllDay = !event.isAllDayEvent();
    const isRecurrenceMismatch = CONFIG.useRecurrence !== (event.isRecurringEvent && event.isRecurringEvent());
    const isNotFromScript = !isEventCreatedByScript(event);
    const hasIncorrectReminders = !hasCorrectReminders(event);
    const isMissingTag = !hasScriptTag(event);  // Check if event is missing the new tag

    // When using individual events for age display, we need to delete any existing recurring events
    const needsConversionToIndividual = usingIndividualEvents && event.isRecurringEvent && event.isRecurringEvent();

    if (isTitleOutdated || isDescriptionOutdated || isNotAllDay || isRecurrenceMismatch || isNotFromScript || hasIncorrectReminders || isMissingTag || needsConversionToIndividual) {
      eventsWereDeleted = true;
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
    return 'skipped_existing'; // Correct event already exists
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
      if (!yearEvent || !isEventCreatedByScript(yearEvent) || (yearEvent.getDescription() || '') !== expectedDescription || !hasCorrectReminders(yearEvent) || !hasScriptTag(yearEvent)) {
        allYearsExist = false;
        break;
      }
    }
    
    if (allYearsExist) {
      return 'skipped_existing'; // All individual events already exist
    }
  }

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
        { description: expectedDescription }
      );
      if (CONFIG.useReminders) {
        event.addPopupReminder(CONFIG.reminderMinutesBefore);
      }
      // Add script identification tag
      addScriptTag(event);
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
      { description: expectedDescription }
    );
    if (CONFIG.useReminders) {
      eventSeries.addPopupReminder(CONFIG.reminderMinutesBefore);
    }
    // Add script identification tag
    addScriptTag(eventSeries);
    Logger.log(`🎉 Created RECURRING event: ${expectedTitle} [starts ${birthdayStartDate.toDateString()}]`);
  } else {
    // Create single event for this year
    const event = calendar.createAllDayEvent(
      expectedTitle,
      birthdayDateThisYear,
      { description: expectedDescription }
    );
    if (CONFIG.useReminders) {
      event.addPopupReminder(CONFIG.reminderMinutesBefore);
    }
    // Add script identification tag
    addScriptTag(event);
    Logger.log(`🎁 Created ONE-TIME event: ${expectedTitle} [${birthdayDateThisYear.toDateString()}]`);
  }
  
  return eventsWereDeleted ? 'updated' : 'created'; // Successfully processed the birthday event
}

/**
 * Add script identification tag to an event using extended properties.
 * @param {CalendarEvent|CalendarEventSeries} event - The event to tag
 */
function addScriptTag(event) {
  try {
    // Use Calendar API to add extended properties
    const calendarId = event.getOriginalCalendarId ? event.getOriginalCalendarId() : CONFIG.calendarId;
    const eventId = event.getId();
    
    // Get the event using Calendar API to add extended properties
    const apiEvent = Calendar.Events.get(calendarId, eventId);
    
    // Add our script identifier as an extended property
    if (!apiEvent.extendedProperties) {
      apiEvent.extendedProperties = {};
    }
    if (!apiEvent.extendedProperties.private) {
      apiEvent.extendedProperties.private = {};
    }
    
    apiEvent.extendedProperties.private[CONFIG.scriptKey] = 'true';
    apiEvent.extendedProperties.private.scriptVersion = '1.0';
    
    // Update the event with the extended properties
    Calendar.Events.update(apiEvent, calendarId, eventId);
  } catch (e) {
    Logger.log(`⚠️ Failed to add script tag to event: ${e}`);
    // Fallback: we'll still rely on description for identification
  }
}

/**
 * Get events created by this script using advanced Calendar API query.
 * This is more efficient than scanning all events when doing cleanup operations.
 * @param {Calendar} calendar - The calendar to search
 * @param {Date} timeMin - Start time for search
 * @param {Date} timeMax - End time for search
 * @returns {CalendarEvent[]} - Array of events created by this script
 */
function getScriptEvents(calendar, timeMin, timeMax) {
  try {
    const calendarId = calendar.getId();
    
    // Try to use Calendar API to search by extended properties
    // Note: This is more efficient but may not be supported in all Apps Script versions
    const events = [];
    let pageToken;
    
    do {
      const response = Calendar.Events.list(calendarId, {
        timeMin: timeMin.toISOString(),
        timeMax: timeMax.toISOString(),
        singleEvents: true,
        maxResults: 250,
        pageToken: pageToken
      });
      
      if (response.items) {
        // Filter events that have our script tag
        const scriptEvents = response.items.filter(apiEvent => {
          return apiEvent.extendedProperties &&
                 apiEvent.extendedProperties.private &&
                 apiEvent.extendedProperties.private[CONFIG.scriptKey] === 'true';
        });
        
        // Convert API events back to CalendarEvent objects
        for (const apiEvent of scriptEvents) {
          try {
            const calendarEvent = calendar.getEventById(apiEvent.id);
            if (calendarEvent) {
              events.push(calendarEvent);
            }
          } catch (e) {
            // Event might be part of a recurring series, skip individual instances
            Logger.log(`⚠️ Skipped event ${apiEvent.summary}: ${e}`);
          }
        }
      }
      
      pageToken = response.nextPageToken;
    } while (pageToken);
    
    return events;
  } catch (e) {
    Logger.log(`⚠️ Failed to get script events via API: ${e}`);
    // Fallback to scanning all events with the old method
    return [];
  }
}

/**
 * Check if an event has the script tag (extended properties).
 * @param {CalendarEvent} event - The event to check
 * @returns {boolean} - True if event has the script tag
 */
function hasScriptTag(event) {
  try {
    const calendarId = event.getOriginalCalendarId ? event.getOriginalCalendarId() : CONFIG.calendarId;
    const eventId = event.getId();
    
    const apiEvent = Calendar.Events.get(calendarId, eventId);
    return apiEvent.extendedProperties && 
           apiEvent.extendedProperties.private && 
           apiEvent.extendedProperties.private[CONFIG.scriptKey] === 'true';
  } catch (e) {
    // If we can't check the tag, assume it doesn't have it
    return false;
  }
}

/**
 * Check if an event was created by this script by looking for script tags first, then description as fallback.
 * @param {CalendarEvent} event - The event to check
 * @returns {boolean} - True if event was created by this script
 */
function isEventCreatedByScript(event) {
  try {
    // First try to check extended properties (new tag-based approach)
    if (hasScriptTag(event)) {
      return true;
    }
    
    // Fallback to description-based checking for backward compatibility
    const description = event.getDescription();
    return description && description.includes(`[${CONFIG.scriptKey}]`);
  } catch (e) {
    return false;
  }
}

/**
 * Check if an event has the correct reminder settings according to current configuration.
 * @param {CalendarEvent} event - The calendar event to check
 * @returns {boolean} - True if reminders match current CONFIG settings
 */
function hasCorrectReminders(event) {
  try {
    const reminders = event.getPopupReminders();
    
    if (!CONFIG.useReminders) {
      // If reminders are disabled, event should have no reminders
      return reminders.length === 0;
    } else {
      // If reminders are enabled, event should have exactly one reminder with correct timing
      return reminders.length === 1 && reminders[0] === CONFIG.reminderMinutesBefore;
    }
  } catch (e) {
    // If we can't get reminders info, assume it's incorrect to be safe
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
  
  Logger.log(`🧹 Cleanup started between: ${startDate.toDateString()} - ${endDate.toDateString()}`);

  // Initialize cleanup counters
  let totalEventsScanned = 0;
  let candidateEventsFound = 0;
  let recurringSeriesDeleted = 0;
  let singleEventsDeleted = 0;
  let deleteFailures = 0;
  let retryAttempts = 0;
  let retrySuccesses = 0;
  let usingTagBasedQuery = false;

  // Collect valid contacts with birthday info
  const contactBirthdays = allContacts
    .filter(person => {
      // Apply label filtering if enabled
      if (CONFIG.useLabels && !hasRequiredLabel(person, CONFIG.contactLabels)) {
        return false;
      }
      return true;
    })
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

  // Phase 1: Try to get script events efficiently using tags, fallback to scanning all events
  Logger.log(`🔍 Looking for birthday events created by this script...`);
  
  let candidateEvents = [];
  
  // Try to use efficient tag-based query first
  try {
    candidateEvents = getScriptEvents(calendar, startDate, endDate);
    if (candidateEvents.length > 0) {
      usingTagBasedQuery = true;
      Logger.log(`✅ Found ${candidateEvents.length} tagged events using efficient query`);
    }
  } catch (e) {
    Logger.log(`⚠️ Tag-based query failed, falling back to full scan: ${e}`);
  }
  
  // If tag-based query didn't work or returned no results, fall back to scanning all events
  if (candidateEvents.length === 0) {
    Logger.log(`🔍 Performing full event scan for cleanup (fallback mode)...`);
    const allEvents = calendar.getEvents(startDate, endDate);
    totalEventsScanned = allEvents.length;
    
    // Filter to only events created by this script
    candidateEvents = allEvents.filter(event => isEventCreatedByScript(event));
    Logger.log(`📋 Found ${candidateEvents.length} script events out of ${totalEventsScanned} total events`);
  } else {
    totalEventsScanned = candidateEvents.length; // We only scanned script events
  }

  // Phase 2: Process candidate events and mark for deletion if they match contacts
  for (const event of candidateEvents) {
    const title = event.getTitle();
    const start = event.getStartTime();
    const isAllDay = event.isAllDayEvent();
    const startsWithCakeEmoji = title.trim().startsWith('🎂');

    for (const contact of contactBirthdays) {
      const isNameMatch = title.includes(contact.name);
      const isBirthdayDateMatch =
        start.getDate() === contact.day &&
        start.getMonth() === contact.month;

      const shouldDelete =
        (startsWithCakeEmoji && isNameMatch) ||
        (isAllDay && isNameMatch && isBirthdayDateMatch);

      if (!shouldDelete) continue;

      candidateEventsFound++;

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
          Logger.log(`📋 Marked recurring series for deletion: ${title}`);
          break;
        } catch (e) {
          retryAttempts++;
          Logger.log(`⚠️ Failed to resolve recurring series: ${title} — ${e}`);
          Logger.log(`⚠️ Waiting for rate limiter and safely retrying`);
          Utilities.sleep(5000);
          try {
            const series = event.getEventSeries();
            const seriesId = series.getId();
            if (deletedSeriesIds.has(seriesId)) {
              retrySuccesses++;
              Logger.log(`✅ Retry successful - series marked for deletion: ${title}`);
              break;
            }
            deletedSeriesIds.add(seriesId);
            eventsToDelete.push({ event: series, isSeries: true });
            retrySuccesses++;
            Logger.log(`✅ Retry successful - marked recurring series for deletion: ${title}`);
            break;
          } catch (e) {
            Logger.log(`⚠️ Second attempt failed — skipping entry: ${e}`);
          }
        }
      } else {
        eventsToDelete.push({ event, isSeries: false });
        Logger.log(`📋 Marked single event for deletion: ${title} on ${event.getStartTime().toDateString()}`);
        break;
      }
    }
  }

  // Phase 3: Execute deletions for all marked events
  Logger.log(`🗑️ Starting deletion of ${eventsToDelete.length} marked events (${deletedSeriesIds.size} recurring series, ${eventsToDelete.length - deletedSeriesIds.size} single events)`);
  
  for (const { event, isSeries } of eventsToDelete) {
    try {
      if (isSeries) {
        event.deleteEventSeries(); // Deletes the entire series
        recurringSeriesDeleted++;
        Logger.log(`🧹 Deleted recurring series → ${event.getTitle()}`);
      } else {
        event.deleteEvent();
        singleEventsDeleted++;
        Logger.log(`🧹 Deleted single event → ${event.getTitle()} on ${event.getStartTime().toDateString()}`);
      }
    } catch (e) {
      deleteFailures++;
      Logger.log(`❌ Failed to delete event → ${event.getTitle()} | Reason: ${e}`);
    }
  }

  // Log comprehensive cleanup summary report
  Logger.log("🧹 CLEANUP SUMMARY REPORT:");
  if (usingTagBasedQuery) {
    Logger.log(`✅ Used efficient tag-based query to find script events`);
  } else {
    Logger.log(`📅 Events scanned in date range: ${totalEventsScanned}`);
  }
  Logger.log(`🎯 Birthday events found for cleanup: ${candidateEventsFound}`);
  Logger.log(`🔄 Recurring series deleted: ${recurringSeriesDeleted}`);
  Logger.log(`📅 Single events deleted: ${singleEventsDeleted}`);
  
  const totalSuccessfulDeletions = recurringSeriesDeleted + singleEventsDeleted;
  Logger.log(`✅ Total successful deletions: ${totalSuccessfulDeletions}`);
  
  if (deleteFailures > 0) {
    Logger.log(`❌ Failed deletions: ${deleteFailures}`);
  }
  
  if (retryAttempts > 0) {
    Logger.log(`🔄 Retry attempts made: ${retryAttempts}`);
    Logger.log(`✅ Successful retries: ${retrySuccesses}`);
  }
  
  Logger.log(`👥 Contacts with birthdays considered: ${contactBirthdays.length}`);
  Logger.log("🎉 Cleanup completed successfully!");
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