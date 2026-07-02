# Tablet Lending Management App

Google Apps Script web app for managing company tablet lending and returns.

## Main Features

- Sales staff can record tablet borrowing and return actions.
- Admin view shows current lending status.
- Usage monitoring records app screen views and lending/return operations.
- Admin view highlights sales staff with no use, no operation, or low use in the last 30 days.
- `usage_logs` stores raw usage events.
- `usage_summary` stores the latest monitoring summary.

## GAS Setup

Run `setupApp` from the Apps Script editor when initializing sheets manually.

The admin dashboard also creates the usage monitoring sheets automatically when opened.
