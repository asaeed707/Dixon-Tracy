# Dixon and Tracy Loading Times

This repository builds and emails the weekday loading time reports for BB Dixon and BB Tracy.

The scheduled workflow runs Monday-Friday. On Mondays it reports the previous Friday; Tuesday-Friday it reports the previous calendar day.

The workflow is intended to run without manual intervention. If email settings
are missing in GitHub Actions, the run fails instead of silently creating files
that nobody receives.

## Required GitHub Secrets

Add these in `Settings` -> `Secrets and variables` -> `Actions` -> `New repository secret`.

- `GEOTAB_USERNAME`
- `GEOTAB_PASSWORD`
- `GEOTAB_DATABASE`
- `GEOTAB_SERVER`
- `SMTP_HOST`
- `SMTP_PORT`
- `SMTP_USERNAME`
- `SMTP_PASSWORD`
- `EMAIL_FROM`
- `EMAIL_TO`

For Gmail SMTP:

- `SMTP_HOST`: `smtp.gmail.com`
- `SMTP_PORT`: `587`
- `SMTP_USERNAME`: your Gmail address
- `SMTP_PASSWORD`: a Gmail app password
- `EMAIL_FROM`: your Gmail address
- `EMAIL_TO`: `ali.saeed@materialtransport.com`

Multiple recipients can be comma-separated in `EMAIL_TO`.

## Go-Live Checklist

1. Confirm this repository is pushed to GitHub.
2. Add every required secret listed above.
3. Open `Actions` -> `Dixon Tracy Loading Times` -> `Run workflow`.
4. Leave `report_date` blank for the normal business-day rule, or enter a date
   like `2026-08-06` for a backfill test.
5. Confirm the email arrives with both plant workbooks attached.
6. Confirm the run summary shows raw trip count, cleaned load counts, ignored
   under-25 stops, long-load counts, and sorted start checks.
7. Leave the schedule enabled for weekday automatic delivery.

## Manual Run

In GitHub, open `Actions` -> `Dixon Tracy Loading Times` -> `Run workflow`.

Optionally enter `REPORT_DATE` as `YYYY-MM-DD` to rebuild a specific day.

## Output Files

The emailed files are named like:

- `Dixon Loading Times for July 22.xlsx`
- `Tracy Loading Times for July 22.xlsx`

Each file contains only:

- `Device`
- `Date`
- `Start Time`
- `End Time`
- `Duration`

Rows over 1 hour 30 minutes are highlighted yellow, visits under 25 minutes are excluded, and re-entry gaps within 25 minutes are merged.

The workflow also uploads the Excel files and audit JSON as GitHub Actions
artifacts for backup.
