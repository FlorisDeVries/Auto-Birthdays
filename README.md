# Auto-Birthdays: Google Apps Script for Birthday Reminders

## Overview
This script automates the process of updating and creating birthday events in Google Calendar based on your Google Contacts. 
For each contact with a birthday, it adds an all-day event in your Google Calendar and sets a reminder a week in advance.

## Features
- **Automatically Update Events**: Updates titles of existing events to include age and sets them as all-day events.
- **Create New Events**: If no birthday event exists for a contact, creates a new all-day event with a reminder.
- **Daily Automation**: Designed to run daily to keep your calendar updated.

## Setup Instructions
1. **Open Google Apps Script**: Navigate to [Google Apps Script](https://script.google.com) and create a new project
2. **Copy the Script**: Copy the provided script into the script editor
3. **Customize Script**: Replace the `calendarID` on line 16 with the target CalendarID
   - Navigate to [Calendar settings](https://calendar.google.com/calendar/u/0/r/settings)
   - Select desired calendar on the left
   - Scroll down to find `Calendar ID`
4. **Set Trigger**: Set a daily trigger to run the `loopThroughContacts` function
    - Click on the clock icon (Triggers) on the left sidebar
    - Click on "+ Add Trigger" at the bottom right corner of the screen
    - Choose the function you want to run from the "Choose which function to run" dropdown (loopThroughContacts)
    - Choose "Time-driven" from the "Select event source" dropdown
    - Choose the type of time trigger (e.g., "Day timer") and specify the time range
5. **Run & Permissions**: From the script editor click on "Run" and accept permissions when asked
6. **Deploy**: Save changes

## Usage
This script automatically updates and creates birthday events in Google Calendar based on birthdays in your Google Contacts. To ensure it works correctly:

1. **Adding Birthdays to Contacts**:
   - Navigate to [Google Contacts](https://contacts.google.com) or go-to contacts on your phone
   - Select a contact to edit
   - Add or edit the birthday
   - Save the contact

The script runs daily and will check these birthdays, adding or updating events in your Google Calendar accordingly.

## Usage
The script will run automatically once every day, updating your Google Calendar with the latest birthday information from your Google Contacts.

## Troubleshooting
If you encounter issues, check the Google Apps Script execution log for error messages. Ensure that your Google Calendar and Contacts permissions are correctly set.

## License
This script is released under the MIT License.

## Contact
For support or queries, please file an issue in the project repository.