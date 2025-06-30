# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https.keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https.semver.org/spec/v2.0.0.html).

## [1.1.0] - YYYY-MM-DD

### Added
- New setting in the 'Settings' sheet: `Role Priority Lookback Weeks`. This allows you to define how many previous meetings the script should check to prioritize members with fewer recent roles.
- The script now creates a `CHANGELOG.md` to track versions.

### Changed
- **Smarter Role Assignment:** The script now prioritizes assigning roles to members who have had fewer assignments in recent weeks (configurable via the new setting).
- Members who had an equivalent role in the *immediately preceding* meeting are still deprioritized, but the main factor is now the historical role count.
- The assignment logic for both protected and other roles has been updated to use a scoring system, ensuring a fairer distribution of roles over time.

## [1.0.0] - YYYY-MM-DD

### Added
- Initial release of the Toastmasters Google Sheets Role Generator.
- Custom menu `Schedule Helper` with `Fill Next Empty Meeting` functionality.
- Support for:
  - Static role assignments.
  - Main protected roles to ensure unique assignments.
  - Ignored roles for assignment.
  - Role equivalency groups to avoid similar consecutive roles.
  - Member availability checks.
  - Randomized assignment with fallback logic. 