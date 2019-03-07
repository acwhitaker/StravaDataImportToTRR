# StravaDataImportToTRR
Google Apps Script to import activity data from Strava to the TeamRunRun Training Log Google Spreadsheet

Data imported includes these fields for Run type activities only: time, distance, and elevation gain. The script assumes the TRRv2 spreadsheet layout and hard codes certain sheet names and cell locations based on this spreadsheet layout.

Credit to:
https://gist.github.com/elifkus/09cd63b3cfbf4e070ecc83b4a4358eaa
I began with this code and modified it heavily to fit the needs of TeamRunRun data import.

Credit to:
https://github.com/gsuitedevs/apps-script-oauth2
The code relies on this Oath2 library to function.
