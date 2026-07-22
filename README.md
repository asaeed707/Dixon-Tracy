# Dixon and Tracy Loading Times

This repository builds and emails the weekday loading time reports for BB Dixon and BB Tracy.

The scheduled workflow runs Monday-Friday. On Mondays it reports the previous Friday; Tuesday-Friday it reports the previous calendar day.

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
