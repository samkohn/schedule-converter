import argparse
import json

import gspread

import sav_shifts


def scan_gsheet(url, tab_name, config):
    connection = gspread.oauth()
    spreadsheet = connection.open_by_url(url)
    worksheet = spreadsheet.worksheet(tab_name)
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

def load_grid_schedule(url, tab_name, config):
    signups = scan_gsheet(url, tab_name, config)
    people = sav_shifts.aggregate_signups(signups)
    return sorted(people)


def write_schedule(location, people):
    pass


def update_schedule(existing_location, update_location, out_location, config):
    pass


def daily_shifts(date_str, in_location, out_location):
    pass


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("in_url")
    parser.add_argument("tab_name")
    parser.add_argument("outfile")
    parser.add_argument("--config", default="config.json")
    args = parser.parse_args()
    with open(args.config, "r") as configfile:
        config = json.load(configfile)
        for sub_dict in config:
            for k, v in config[sub_dict].copy().items():
                config[sub_dict][int(k)] = v
                del config[sub_dict][k]

    people = load_grid_schedule(args.in_url, args.tab_name, config)
    sav_shifts.write_csv(args.outfile, people)
