"""
Google Drive mail merger.
"""
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

DOCS_FILE_ID = ""
SHEETS_FILE_ID = ""

SCOPES = (
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/spreadsheets",
)

COLUMNS = ["Order #", "FirstName", "LastName", "Address1", "Address2", "City", "State", "PostalCode", "Country", "Order Date", "Product Weight", "Shipping Method", "Item Count", "Value of Products", "Shipping Fee Paid", "Tracking #"]


def get_data():
    return (
        SHEETS.spreadsheets()
        .values()
        .get(spreadsheetId=SHEETS_FILE_ID, range="A:P")
        .execute()
        .get("values")[1:]
    )


def merge_template(data, tmpl_id):
    data_refs = []
    data_merged = []
    for row in data:
        merged = dict(zip(COLUMNS, row))
        data_ref = str(row[1:8])
        if data_ref in data_refs:
            print(f"Duplicate removed: {data_ref}")
            continue

        data_merged.append(merged)
        data_refs.append(data_ref)

    contents_original = DRIVE.files().export(fileId=tmpl_id, mimeType="text/plain").execute()
    contents = contents_original.decode("utf-8")
    output = ""
    breakpoints = []
    for row in data_merged:
        buff = contents
        for key, value in row.items():
            buff = buff.replace("{{key}}".replace("key", key), value)
        output += buff
        breakpoints.append(len(bytes(buff, "utf-8").replace(b"\r\n", b"\n")) - 1)

    breakpos = len(contents_original.replace(b"\r\n", b"\n"))
    reqs = [
        {"deleteContentRange": {"range": {"segmentId": "", "startIndex": 1, "endIndex": breakpos - 2}}},
        {"insertText": {"text": output, "endOfSegmentLocation": {"segmentId": ""}}},
    ]
    reqs += [{"insertPageBreak": {"location": {"segmentId": "", "index": sum(breakpoints[:i+1]) + i}}} for i in range(len(data) - 1)]
    body = {"title": f"Merged"}
    copy_id = DRIVE.files().copy(body=body, fileId=tmpl_id, fields="id").execute().get("id")
    DOCS.documents().batchUpdate(body={"requests": reqs}, documentId=copy_id, fields="").execute()

    return copy_id


if __name__ == "__main__":
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    DRIVE = build("drive", "v2", credentials=creds)
    DOCS = build("docs", "v1", credentials=creds)
    SHEETS = build("sheets", "v4", credentials=creds)

    print(f"Merged: docs.google.com/document/d/{merge_template(get_data(), DOCS_FILE_ID)}/edit")
