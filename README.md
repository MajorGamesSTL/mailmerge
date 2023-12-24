# mailmerge

Google Drive mail merger. Pulls rows from sheet into separate pages of a doc.
Based on https://console.cloud.google.com/apis/api/sheets.googleapis.com/credentials?project=mailmerge-406807&supportedpurview=project

# Usage

Create a google service with required permissions and download the oauth credentials.json.
* https://developers.google.com/apps-script/samples/automations/mail-merge

Create a base doc with variables denoted as {{varname}}, these will reference column names.
If {{varname}} with whitespace before it is the only thing on the line, the line will be removed if varname is unset for a specific row.
Use spaces for indentation, not formatting tools.

In mailmerge.py, set the doc and sheets id based on the URLs.

Run the code with python.

# Contact

Feel free to contact via email in account description or linked social medias.
