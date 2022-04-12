# Release Notes

## V2.1

- Patched the smart quote replace code. Strings with smart single and double quotes are causing the email sending to fail.

## V2.0

- Removed Office 365 Service Communications API and replaced with Microsoft Graph API to read the service health events.
- Removed the `JWTDetails` module as a requirement.
- Added `-Status` parameter to filter query results based on status (`Ongoing` or `Closed`). This parameter is optional and if not used, all issues will be retrieved.
- Added `Get-MS365HealthOverview` which you can use to retrieve the health overview summary only.
- Removed `Get-MS365CurrentStatus` as it is no longer applicable. Use `Get-MS365HealthOverview` instead.

## V1.4.2

- Fixed error in reading the last run timestamp from the registry.

## V1.4.1

- Add "Classification" column to summary.
- Add On-page anchor links in summary.

## V1.4

- Add `-Consolidate` parameter (boolean) to consolidate reports in one email.

## V1.3

- Code cleanup.
- Fixed some JSON related errors.

## V1.2

- Add code to force TLS 1.2 connection [Issue #2](https://github.com/junecastillote/MS365HealthReport/issues/1)

## v1.1

- Added logic to replace smart quotes in messages [Issue #1](https://github.com/junecastillote/MS365HealthReport/issues/1)