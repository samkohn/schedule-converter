import argparse
from dataclasses import dataclass
from datetime import datetime
import json

import gspread

import sav_shifts


@dataclass
class SpreadsheetLocation:
    url: str
    tab: str


def open_spreadsheet(url):
    connection = gspread.oauth()
    spreadsheet = connection.open_by_url(url)
    return spreadsheet


def open_worksheet(location):
    spreadsheet = open_spreadsheet(location.url)
    worksheet = spreadsheet.worksheet(location.tab)
    return worksheet


def scan_gsheet(in_location, config):
    spreadsheet = open_spreadsheet(in_location.url)
    worksheet = spreadsheet.worksheet(in_location.tab)
    all_values = worksheet.get_all_values()
    signups = []
    for row_index, row in enumerate(all_values):
        row_number = row_index + 1
        if row_number < min(config["rows"].keys()):
            continue
        for column_number in sav_shifts.good_columns(config):
            column_index = column_number - 1
            name = row[column_index]
            email_phone_string = row[column_index + 1]
            signup = sav_shifts.parse_cell(
                name, email_phone_string, row_index, column_index, config
            )
            if signup is not None:
                signups.append(signup)
    return signups


def load_grid_schedule(in_location, config):
    signups = scan_gsheet(in_location, config)
    people = sav_shifts.aggregate_signups(signups)
    return sorted(people)


def write_schedule(out_location, people):
    spreadsheet = open_spreadsheet(out_location.url)
    if out_location.tab in [ws.title for ws in spreadsheet.worksheets()]:
        worksheet = spreadsheet.worksheet(out_location.tab)
        if worksheet.frozen_row_count == 0:
            worksheet.freeze(rows=1)
    else:
        worksheet = spreadsheet.add_worksheet(out_location.tab, rows=1000, cols=26)
        worksheet.freeze(rows=1)
    first_person = people[0]
    headers = first_person.list_headers()
    update = []
    update.append(
        {
            "range": "A1:Z1",
            "values": [headers],
        }
    )
    update.append(
        {
            "range": "A2:Z",
            "values": [person.to_list() for person in people],
        }
    )
    worksheet.clear()
    worksheet.batch_update(update)


def scan_mailmerge_sheet(location):
    spreadsheet = open_spreadsheet(location.url)
    worksheet = spreadsheet.worksheet(location.tab)
    num_standard_columns = len(sav_shifts.PersonSchedule.list_headers())
    people = {}
    for i, row in enumerate(worksheet.get_values()):
        if i == 0:
            header = row
            additional_columns = row[num_standard_columns:]
            continue
        person_row = sav_shifts.parse_mailmerge_row(row, additional_columns)
        people[person_row.full_name.lower()] = person_row
    return people


def update_schedule(existing_location, update_location, out_location, config):
    """Update the records from existing_location with the current signup sheet
    at update_location and put the result in out_location.
    """
    existing_people = scan_mailmerge_sheet(existing_location)
    new_version_people = load_grid_schedule(update_location, config)
    updated_people = sav_shifts.update_with_new_shifts(
        existing_people, new_version_people
    )
    write_schedule(out_location, list(updated_people.values()))


def daily_shifts(date_str, in_location, out_location):
    if out_location.tab is None:
        now = datetime.now()
        new_tab_name = now.strftime(f"Shifts for {date_str} as of %m/%d %I:%M%p")
        out_location.tab = new_tab_name
    if in_location.tab is None:
        in_sheet = open_spreadsheet(out_location.url)
        worksheets = in_sheet.worksheets()
        in_location.tab = worksheets[-1].title
    existing_people = list(scan_mailmerge_sheet(in_location).values())
    specific_date_people = sav_shifts.filter_daily_shifts(date_str, existing_people)
    write_schedule(out_location, specific_date_people)


def process_calendar(config, in_location, out_location, update):
    """Convert from signup calendar spreadsheet to new mail merge spreadsheet.

    If out_location.tab is None, compute a new tab name based on the current time.

    If update is False, don't update from anywhere.
    If update is None, update from the rightmost tab in out_location.url.
    If update is a string, update from the tab in out_location.url with that name,
    or error out if the tab doesn't exist.
    """
    now = datetime.now()
    new_tab_name = now.strftime("Shifts as of %m/%d %I:%M%p")
    out_location.tab = new_tab_name
    if update is False:
        people = load_grid_schedule(in_location, config)
        write_schedule(out_location, people)
        return
    if update is None:
        out_sheet = open_spreadsheet(out_location.url)
        worksheets = out_sheet.worksheets()
        update = worksheets[-1].title
    existing_location = SpreadsheetLocation(out_location.url, update)
    update_schedule(existing_location, in_location, out_location, config)


def parse_setup(url):
    spreadsheet = open_spreadsheet(url)
    setup_worksheet = spreadsheet.worksheet("Setup")
    signups_url = setup_worksheet.acell("B1").value
    signups_tab_name = setup_worksheet.acell("B2").value
    return SpreadsheetLocation(signups_url, signups_tab_name)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("url", help="The URL of the output spreadsheet with Setup tab")
    parser.add_argument("--daily")
    parser.add_argument(
        "--update",
        nargs="?",
        default=False,
        help="Carry over custom fields from rightmost tab, "
        "or provide a specific tab name to use",
    )  # If flag not present, will be False; if flag present with no value, will be None
    parser.add_argument("--config", default="config.json")
    args = parser.parse_args()
    with open(args.config, "r") as configfile:
        config = json.load(configfile)
        for sub_dict in config:
            for k, v in config[sub_dict].copy().items():
                config[sub_dict][int(k)] = v
                del config[sub_dict][k]
    signups_location = parse_setup(args.url)
    print(signups_location)
    if args.daily is None:
        process_calendar(
            config, signups_location, SpreadsheetLocation(args.url, None), args.update
        )
    else:
        in_location = SpreadsheetLocation(args.url, None)
        out_location = SpreadsheetLocation(args.url, None)
        daily_shifts(args.daily, in_location, out_location)
