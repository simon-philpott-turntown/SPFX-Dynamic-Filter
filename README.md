# SPFx Dynamic Filter Web Part

Central SPFx Dynamic Data filter controller with search, time-of-day greetings, user profile dynamic pre-filtering, and Term Store synonym mapping.

## Features
- **Dynamic Data Controller**: Dispatches active search queries, selected filters, and pre-filters to connected web parts.
- **Time-of-Day Salutation**: Dynamic morning, afternoon, and generic greeting messages with optional first name appending.
- **User Profile Diagnostics**: Author-only user profile diagnostics popover with site admin verification and Microsoft Graph attribute inspection.
- **Dynamic Pre-Filtering**: Automatic pre-filtering based on logged-in user profile attributes (e.g. Department, Office Location).
- **Term Store Synonym Resolution**: Maps raw profile strings to canonical taxonomy term names across term sets and children.
- **Dismissible Applied Filter Chips**: Individual filter removal via dismiss cross icons, plus full clear-all capabilities.

## Getting Started
```bash
npm install
heft start
```
