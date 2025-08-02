# 🎉 Auto-Birthdays

This Google Apps Script creates birthday events in your Google Calendar based on your Google Contacts.

---

## 🚀 Getting Started

This project uses **Google Apps Script** and the **Google People API**.

1. Open [Google Apps Script](https://script.google.com/) and create a new project.
2. Paste the contents of `code.gs` into the editor.
3. In the left sidebar, go to **Services** and add the **People API**.
4. Modify the `CONFIG` section at the top of the script (see below).
5. Run `loopThroughContacts()` once to:
   - Generate birthday events
   - Optionally clean up outdated ones
   - Set up a time-based trigger (if enabled)

> **⚠️ DATA-LOSS WARNING**  
> The script **deletes any all-day event whose date and title look like a birthday** so it can rebuild a clean set each run.  
> That can accidentally match anniversaries, retirements, “Remembrance”, etc.  
> **Protect yourself:**
> 1. **Use a separate calendar** (create one called “Birthdays”) and point `CONFIG.calendarId` to it.  
> 2. Keep `CONFIG.cleanupEvents = false` until you have verified everything on a test calendar.
> 3. If something disappears, restore it from **Calendar Trash/Bin** within 30 days.

---

## ⚙️ Configuration

The CONFIG object gives you full control over how events are created and managed:

```javascript
const CONFIG = {
  calendarId: 'primary',             // 'primary' or your custom calendar ID

  // Title customization
  useEmoji: true,                    // Add 🎂 emoji to event titles
  showYearOrAge: true,               // Recurrence on: shows (*YYYY), off: shows (age)
  showAgeOnRecurring: false,         // If true, shows (age) on recurring events instead of (*YYYY)

  // Recurrence
  useRecurrence: true,               // Create recurring yearly events
  futureYears: 20,                   // Recurring events end this many years in the future
  pastYears: 2,                      // Recurring events start this many years in the past

  // Reminder settings
  reminderMinutesBefore: 1440,       // Popup reminder (in minutes); 1440 = 1 day before

  // Cleanup
  cleanupEvents: false,               // ⚠️⚠️⚠️ Deletes all matching birthday events between ±100 years

  // Trigger options
  useTrigger: true,                  // Automatically run on a schedule
  triggerFrequency: 'daily',         // 'daily' or 'hourly'
  triggerHour: 4,                    // If 'daily', the hour of day to run (0–23)

  // Script identification
  scriptKey: 'CREATED_BY_Auto-Birthdays' // Unique identifier for events created by this script
};
```

---

## 🖋️ Event Title Formats

Depending on your configuration, birthday events will appear with different formats:

| `useEmoji` | `useRecurrence` | `showYearOrAge` | `showAgeOnRecurring` | `Event Title Example`     |
|------------|-----------------|-----------------|---------------------|---------------------------|
| true       | true            | true            | false               | 🎂 John Doe (*1988)       |
| true       | true            | true            | true                | 🎂 John Doe (36)          |
| true       | false           | true            | N/A                 | 🎂 John Doe (36)          |
| false      | true            | false           | N/A                 | John Doe                  |
| false      | false           | true            | N/A                 | John Doe (36)             |
| false      | false           | false           | N/A                 | John Doe                  |

**Notes:**
- `showAgeOnRecurring` only applies when `useRecurrence=true` and `showYearOrAge=true`
- If no birth year is provided, age or year is omitted
- When `showAgeOnRecurring=true`, individual events are created for each year instead of recurring events:
  - 2025: 🎂 John Doe (30)
  - 2026: 🎂 John Doe (31) 
  - 2027: 🎂 John Doe (32)
- This allows you to see all future birthdays with correct ages when browsing your calendar
- Individual events span from `pastYears` to `futureYears` relative to the current year

---

## 🧹 Automatic Cleanup

If `CONFIG.cleanupEvents` is enabled:

- Search your calendar between 100 years in the past and future
- Find outdated or duplicate birthday events **created by this script only**
- Delete them safely, including recurring series
- Manual birthday events are never touched

**Safety Note**: The cleanup only affects events containing the script's unique identifier, ensuring your manually created events remain safe.

---

## ⏰ Trigger Behavior

If `CONFIG.useTrigger` is enabled:

- Automatically create a **time-based trigger**
- Run either:
  - **Hourly** every hour
  - Or **Daily** at `triggerHour`

🛠️ If you change the trigger settings, the script will:
- **Remove any existing trigger**
- **Install a new one with the updated config**

This ensures only **one correct trigger** is active.

---

## 🔑 Script Identification

The script uses a unique key system to identify events it has created:

- **`scriptKey`**: A unique identifier embedded in each event's description
- **Safe operation**: Only modifies events it has created, leaving manual birthday events untouched
- **Easy identification**: You can search for events created by the script using the key
- **Version tracking**: Update the key for different script versions if needed

**How it works:**
- Each created event includes `[CREATED_BY_Auto-Birthdays]` in its description
- The script only deletes/updates events containing this identifier
- Manual birthday events are completely safe from script operations

---

## 🗓️ Examples

You'll see all-day birthday events appear in your calendar like these:

- 🎂 John Doe (*1988)
- 🎂 Jane Smith (36)
- John Appleseed (no year provided)

All events appear as **all-day events** on the person’s birthday.

![Example Event](img/example.png)

---

## 🧰 Troubleshooting

If you encounter issues:

- Check the **Apps Script execution log**
- Run the script **manually the first time** to authorize permissions
- Ensure you've **enabled the People API** under **Services**
- Make sure you've granted **Calendar and Contacts permissions**
- Restore missing items via Calendar → Trash/Bin.

---

## 🤝 Contributing

Contributions are welcome! Fork this repo, improve it, and submit a pull request.

---

## 📄 License

This script is released under the **MIT License**.

---

## 💬 Contact

For support or feedback, please [file an issue](https://github.com/FlorisDeVries/Auto-Birthdays/issues).