# Equipment Lending Management App

Google Apps Script web app for managing company tablet and Proud Note lending and returns.

## Main Features

- Sales staff can independently record tablet and Proud Note borrowing and return actions.
- Proud Notes are optionally assigned to specific sales staff in the sixth column of `users`.
- Admin view shows current lending status and can filter by equipment type.
- Usage monitoring records app screen views and lending/return operations.
- Admin view highlights sales staff with no use, no operation, or low use in the last 30 days.
- `usage_logs` stores raw usage events.
- `usage_summary` stores the latest monitoring summary.

## GAS Setup

Run `setupApp` from the Apps Script editor when initializing sheets manually.

For an existing installation, run `setupApp` once after deploying this version. It
adds the new headers while preserving existing rows. Enter each assigned Proud
Note number in column F (`割当プラウドノート`) of `users`; leave it blank for staff
who do not receive one. Existing status and log rows with no equipment type are
treated as tablets.

The admin dashboard also creates the usage monitoring sheets automatically when opened.
