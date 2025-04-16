# 🎉 Auto-Birthdays

This Google Apps Script creates birthday events in your Google Calendar based on your Google Contacts.

---

## 🚀 Getting Started

This project uses **Google Apps Script** and the **Google People API**.

1. Open [Google Apps Script](https://script.google.com/) and create a new script.
2. Copy the contents of `code.gs` into the editor.
3. Go to the **left bar → Services** and add the **People API** to the project.
4. Edit the `CONFIG` section at the top of the script (see below).
5. Run `loopThroughContacts()` once to:
   - Generate birthday events
   - Set up a time-based trigger (if enabled)

---

## ⚙️ Configuration

At the top of the script you'll find the `CONFIG` object to control behavior:

```javascript
const CONFIG = {
  calendarId: 'primary',
  useEmoji: true,
  useRecurrence: true,
  showYear: true,
  reminderMinutesBefore: 1 * 24 * 60,
  recurrenceYears: 50,

  useTrigger: true,
  triggerFrequency: 'daily',  // 'daily' or 'hourly'
  triggerHour: 4              // Only used if frequency is 'daily'
};
```

---

## 🖋️ Title Format Examples

| useEmoji | useRecurrence | showYear | Event Title Example     |
|----------|----------------|----------|--------------------------|
| true     | true           | true     | 🎂 John Doe (*1988)       |
| true     | true           | false    | 🎂 John Doe (36)          |
| false    | false          | false    | John Doe (36)            |
| false    | true           | true     | John Doe (*1988)         |

If no birth year is available, the title will only include the name.

---

## ⏰ Trigger Behavior

If `CONFIG.useTrigger` is `true`, the script will:

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

Here are some examples of how the events will look in your Google Calendar:

- 🎂 John Doe (*1988)
- 🎂 Jane Smith (36)
- John Appleseed (no year provided)

All events appear as **all-day events** on the person’s birthday.

![Example Event](img/example.png)

---

## 🧰 Troubleshooting

If you encounter issues:

- Check the **Apps Script execution log**
- Ensure you've **enabled the People API** under **Services**
- Make sure you've granted **Calendar and Contacts permissions**
- Run `loopThroughContacts()` manually the first time

---

## 🤝 Contributing

Contributions are welcome! Fork this repo, improve it, and submit a pull request.

---

## 📄 License

This script is released under the **MIT License**.

---

## 💬 Contact

For support or feedback, please [file an issue](https://github.com/willi84/Auto-Birthdays/issues).

---

## 🤖 Tip

This script was enhanced with help from **ChatGPT**.
