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

---

## ⚙️ Configuration

The CONFIG object gives you full control over how events are created and managed:

```javascript
const CONFIG = {
  calendarId: 'primary',             // 'primary' or your custom calendar ID

  // Title customization
  useEmoji: true,                    // Add 🎂 emoji to event titles
  showYearOrAge: true,               // Recurrence on: shows (*YYYY), off: shows (age)

  // Recurrence
  useRecurrence: true,               // Create recurring yearly events
  futureYears: 20,                   // Recurring events end this many years in the future
  pastYears: 2,                      // Recurring events start this many years in the past

  // Reminder settings
  reminderMinutesBefore: 1440,       // Popup reminder (in minutes); 1440 = 1 day before

  // Cleanup
  cleanupEvents: true,               // ⚠️ Deletes all matching birthday events between ±100 years

  // Trigger options
  useTrigger: true,                  // Automatically run on a schedule
  triggerFrequency: 'daily',         // 'daily' or 'hourly'
  triggerHour: 4                     // If 'daily', the hour of day to run (0–23)
};
```

---

## 🖋️ Event Title Formats

Depending on your configuration, birthday events will appear with different formats:

| `useEmoji` | `useRecurrence` | `showYearOrAge` | `Event Title Example`     |
|------------|-----------------|-----------------|---------------------------|
| true       | true            | true            | 🎂 John Doe (*1988)       |
| true       | false           | false           | 🎂 John Doe (36)          |
| false      | true            | false           | John Doe                  |
| false      | false           | true            | John Doe (36)             |
| false      | false           | false           | John Doe                  |

If no birth year is provided, age or year is omitted.

---

## 🧹 Automatic Cleanup

If `CONFIG.cleanupEvents` is enabled:

- Search your calendar between 100 years in the past and future
- Find outdated or duplicate birthday events
- Delete them safely, including recurring series

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

---

## 🤝 Contributing

Contributions are welcome! Fork this repo, improve it, and submit a pull request.

---

## 📄 License

This script is released under the **MIT License**.

---

## 💬 Contact

For support or feedback, please [file an issue](https://github.com/FlorisDeVries/Auto-Birthdays/issues).