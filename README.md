# Toastmasters Google Sheets Role Generator

A Google Apps Script to help schedule Toastmasters meeting roles in a Google Sheets spreadsheet by reading settings and availability data, then automatically filling the next empty meeting column.

## Features

- Adds a custom "Schedule Helper" menu to the spreadsheet UI.
- Static assignments for roles that should always be assigned to specific members.
- Main protected roles and unique role groups to prevent repeat assignments.
- Ignored roles for assignment (e.g., speaker roles filled separately).
- Role equivalencies to group similar roles together.
- Availability-based assignment using an availability sheet.
- Avoids assigning members the same or equivalent roles consecutively.
- Randomized assignment with fallback logic and volunteer prompts.

## Setup / Installation

1. Open your Google Sheets spreadsheet.
2. Select **Extensions > Apps Script**.
3. Copy the contents of `Code.gs` into the script editor.
4. Save the project (e.g., "Toastmasters Role Generator").
5. Grant the required permissions when prompted.
6. Reload the spreadsheet to see the new **Schedule Helper** menu.

## Configuration

### 1. Settings Sheet (Third Sheet)

Ensure the **third sheet** in the workbook (after the Schedule and Availability sheets) contains the settings data. Name it whatever you like (e.g., `Settings`). The script accesses it by position, not by name.

Add the following columns in the header row:

- `Main Protected Roles`: Roles requiring a unique assignee each meeting.
- `Ignored Roles for Assignment`: Roles to skip during auto-assignment.
- `Static Role` and `Assigned Member`: Pairs defining static assignments.
- `Equivalent Roles - Group 1`, `Equivalent Roles - Group 2`, etc.: Columns listing roles considered equivalent. Each group prevents assigning similar roles consecutively.

Populate subsequent rows with values. Empty cells are ignored.

### 2. Availability Sheet

Create a second sheet (e.g., `Availability`) where:

- The first row contains dates matching those in the schedule sheet.
- The first column lists member names.
- Cells contain `0` for **unavailable**; leave blank or enter any other value to mark members **available** (blank is treated as available by default).

### 3. Schedule Sheet

- The first sheet in the workbook should serve as the schedule.
- The first row lists meeting dates.
- The first column lists role names.
- The script fills each cell intersection of role and date.

## Usage

1. Ensure your **Schedule**, **Settings**, and **Availability** sheets are properly configured.
2. Reload the spreadsheet.
3. Click **Schedule Helper > Fill Next Empty Meeting**.
4. The script will auto-fill assignments for the next date column and notify you upon completion.

## Troubleshooting

- Alerts appear if the `Settings` or `Availability` sheets or required headers cannot be found.
- If no editable roles remain, an alert will indicate completion.
- Verify header spelling, sheet names, and data formats if errors occur.

## License

MIT License 