# Changelog

All notable changes to Auto-Birthdays will be documented in this file.

## [1.2.0] - 2026-02-04

### Added
- Added `setTransparency` configuration option to control calendar availability settings
- Events can now be set to show you as "available" (transparent) instead of the default "busy" (opaque)
- Transparency setting applies to all event types: recurring, individual age events, and one-time events

## [1.1.0] - 2026-02-04

### Fixed
- Fixed title generation for custom formats when `showYearOrAge: false` - age variables are now always populated when birth year data is available
- Fixed `{years}` placeholder to only show when age exists and is not 1, preventing malformed titles like "30Jahre"
- Fixed `{age}` placeholder to properly handle age 0 (newborns) by changing check from `age ?` to `age !== null`
- Fixed recurrence mismatch detection when using individual age events (`showAgeOnRecurring: true`) - script now correctly recognizes non-recurring individual events as valid
- Added trailing dash cleanup in title formatting to handle empty age information

### Changed
- Updated emoji logic: `useEmoji` config now only applies to default title formats; custom formats always have access to `{emoji}` placeholder
- Improved title cleanup regex to remove trailing dashes and multiple spaces

### Added
- Added inline comments explaining showYearOrAge behavior with custom formats
- Added deletion reason logging for better debugging of event updates

## [1.0.0] - Initial Release

### Features
- Automatic birthday event creation from Google Contacts
- Multi-language support (English, Italian, French, German, Spanish)
- Customizable title formats with placeholder system
- Recurring event support with configurable date ranges
- Individual age-based events for recurring birthdays
- Reminder configuration
- Contact filtering by labels and birth months
- Automated trigger scheduling
- Event cleanup functionality
