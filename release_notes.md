# Release Notes

## v2.1.2

- Fixed error in output directory handling whenever the `$ENV:HOME` or `$ENV:HOMEPATH` does not exist.

## v2.1.1

- Minor code change to properly display the "Status" values.
  - Example: `extendedRecovery` = `Extended Recovery`
  - Example: `serviceDegradation` = `Service Degradation`.
- Added unicode space character replacement code.
- Changed `LastRunTime` location from the registry to a CSV file on the user's home folder.
  - On Windows systems -  *`Env:\HOMEPATH\MS365HealthReport\<tenant>\runHistory.csv`*
  - On non-Windows systems -  *`Env:\HOME\MS365HealthReport\<tenant>\runHistory.csv`*
- Now compatible with PowerShell on Linux.

## v2.1

- Patched the smart quote replace code. Strings with smart single and double quotes are causing the email sending to fail.

## v2.0

- Removed Office 365 Service Communications API and replaced with Microsoft Graph API to read the service health events.
- Removed the `JWTDetails` module as a requirement.
- Added `-Status` parameter to filter query results based on status (`Ongoing` or `Closed`). This parameter is optional and if not used, all issues will be retrieved.
- Added `Get-MS365HealthOverview` which you can use to retrieve the health overview summary only.
- Removed `Get-MS365CurrentStatus` as it is no longer applicable. Use `Get-MS365HealthOverview` instead.

## v1.4.2

- Fixed error in reading the last run timestamp from the registry.

## v1.4.1

- Add "Classification" column to summary.
- Add On-page anchor links in summary.

## v1.4

- Add `-Consolidate` parameter (boolean) to consolidate reports in one email.

## v1.3

- Code cleanup.
- Fixed some JSON related errors.

## v1.2

- Add code to force TLS 1.2 connection [Issue #2](https://github.com/junecastillote/MS365HealthReport/issues/1)

## v1.1

- Added logic to replace smart quotes in messages [Issue #1](https://github.com/junecastillote/MS365HealthReport/issues/1)
