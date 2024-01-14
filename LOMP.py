import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import numpy as np
import matplotlib.pyplot as plt

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1BOnpf3t9ldGz0tTWSXnjYyD0fghsq8RwOesOzj10Zwk"
SAMPLE_RANGE_NAME = "B2:D"


def main():
  """Shows basic usage of the Sheets API.
  Prints values from a sample spreadsheet.
  """
  creds = None
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
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

  try:
    service = build("sheets", "v4", credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = (
        sheet.values()
        .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME)
        .execute()
    )
    values = result.get("values", [])

    if not values:
      print("No data found.")
      return

    names = []
    pounds = []
    for row in values:
      # Print columns A and E, which correspond to indices 0 and 4.
      names.append(row[0])
      pounds.append(float(row[1]))

    names = np.array(names)
    pounds = np.array(pounds)

    test_names = np.unique(names)
    percentiles = []
    for t in test_names:
        percentiles.append(np.sum(pounds[names == t]))
    percentiles = np.array(percentiles)
  
    fig, ax1 = plt.subplots(figsize=(9, 6), layout='constrained')
    fig.canvas.manager.set_window_title('')
    rects = ax1.barh(test_names, percentiles, align='center', height=0.5)
    large_percentiles = []
    small_percentiles = []
    for p in percentiles:
        if p < 100000:
            large_percentiles.append('')
            small_percentiles.append(p)
        else:
            large_percentiles.append(p)
            small_percentiles.append('')

    ax1.bar_label(rects, small_percentiles,
                    padding=5, color='black', fontweight='bold')
    ax1.bar_label(rects, large_percentiles,
                    padding=-50, color='white', fontweight='bold')

    ax1.set_xlim([0, 1000000])

    xlabels = []
    for x in np.arange(0,1100000,100000):
        xlabels.append(str(x))

    ax1.set_xticks(np.arange(0,1100000,100000),
        labels = (xlabels))
    ax1.xaxis.grid(True, linestyle='--', which='major',
                    color='grey', alpha=.25)
    ax1.axvline(500000, color='grey', alpha=0.25)  # median position


    # Set the right-hand Y-axis ticks and labels
    ax2 = ax1.twinx()
    # Set equal limits on both yaxis so that the ticks line up
    ax2.set_ylim(ax1.get_ylim())
    # Set the tick locations and labels
    percentage_completion = np.array(percentiles)/10000.
    percent_labels = []
    for p in percentage_completion:
        label = '{:.2f} %'.format(p)
        percent_labels.append(label)

    ax2.set_yticks(
        np.arange(len(percentiles)),
        labels = (percent_labels))

    ax2.set_ylabel('Percentage Completion')
    plt.savefig("LOMP.png", dpi = 300)
  except HttpError as err:
    print(err)


if __name__ == "__main__":
  main()