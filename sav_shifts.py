import argparse
import copy
import csv
from dataclasses import dataclass, field as dc_field
import json
import re

class SpreadsheetLocationError(Exception):
    pass

@dataclass
class SignupCell:
    content: list
    row: int
    column: int
    date: str
    time: str
    shift_type: str
    name: str = None
    phone: str = None
    email: str = None
    turf: str = None
    hq: str = None


@dataclass(order=True)
class PersonSchedule:
    name: str
    phone: str
    walkthrough_shifts: list = dc_field(default_factory=list)
    phonebank_shifts: list = dc_field(default_factory=list)

    def to_list(self):
        return [
            self.name,
            self.first_name(),
            self.last_name(),
            self.phone,
            self.shifts_to_str(self.walkthrough_shifts),
            self.shifts_to_str(self.phonebank_shifts),
        ]

    @staticmethod
    def list_headers():
        return [
            "Full name",
            "First name",
            "Last name",
            "cell",
            "Walkthrough shifts",
            "Phonebank shifts",
        ]

    @staticmethod
    def shifts_to_list(shifts_str):
        if shifts_str == "":
            return []
        shifts_str_list = shifts_str.split("\n")
        return [MailMergeRow.parse_shiftstring(x) for x in shifts_str_list]
    
    @staticmethod
    def parse_shiftstring(shift_string):
        date = re.search(r".*(?= from )", shift_string)
        time = re.search(r"(?<= from ).*?(?=( for )|( \[)|$)", shift_string)
        if not (date and time):
            raise ValueError(f"String '{shift_string}' is not a valid shift (requires a date and time)")
        turf = re.search(r"(?<= for ).*?(?=( \[)|$)", shift_string)
        hq = re.search(r"(?<=\[Report to: ).*(?=\])", shift_string)
        return [m[0] if m else None for m in ([date, time, turf, hq] if turf or hq else [date,time])]

    @staticmethod
    def shifts_to_str(shift_list):
        return "\n".join(
            sorted(
                [MailMergeRow._shift_to_str(shift) for shift in shift_list],
                key=MailMergeRow._extract_date
            )
        )
    
    @staticmethod
    def _shift_to_str(shift):
        string = f"{shift[0]} from {shift[1]}" 
        if len(shift)>2:
            if shift[2]: string += f" for {shift[2]}"
            if shift[3]: string += f" [Report to: {shift[3]}]"
        return string
    
    @staticmethod
    def _extract_date(shift_string):
        return re.search(r".*(?= from )", shift_string)[0].split(", ")[1]

    def first_name(self):
        return self.name.split()[0]

    def last_name(self):
        return " ".join(self.name.split()[1:])


@dataclass
class MailMergeRow:
    full_name: str
    first_name: str
    last_name: str
    phone: str
    email: str
    walkthrough_shifts: list
    phonebank_shifts: list
    other_columns: dict

    def to_list(self):
        standard_columns = [
            self.full_name,
            self.first_name,
            self.last_name,
            self.phone,
            self.email,
            self.shifts_to_str(self.walkthrough_shifts),
            self.shifts_to_str(self.phonebank_shifts),
        ]
        additional_columns = list(self.other_columns.values())
        return standard_columns + additional_columns

    @staticmethod
    def shifts_to_list(shifts_str):
        if shifts_str == "":
            return []
        shifts_str_list = shifts_str.split("\n")
        return [MailMergeRow.parse_shiftstring(x) for x in shifts_str_list]
    
    @staticmethod
    def parse_shiftstring(shift_string):
        date = re.search(r".*(?= from )", shift_string)
        time = re.search(r"(?<= from ).*?(?=( for )|( \[)|$)", shift_string)
        if not (date and time):
            raise ValueError(f"String '{shift_string}' is not a valid shift (requires a date and time)")
        turf = re.search(r"(?<= for ).*?(?=( \[)|$)", shift_string)
        hq = re.search(r"(?<=\[Report to: ).*(?=\])", shift_string)
        return [m[0] if m else None for m in ([date, time, turf, hq] if turf or hq else [date,time])]

    @staticmethod
    def shifts_to_str(shift_list):
        return "\n".join(
            sorted(
                [MailMergeRow._shift_to_str(shift) for shift in shift_list],
                key=MailMergeRow._extract_date
            )
        )
    
    @staticmethod
    def _shift_to_str(shift):
        string = f"{shift[0]} from {shift[1]}" 
        if len(shift)>2:
            if shift[2]: string += f" for {shift[2]}"
            if shift[3]: string += f" [Report to: {shift[3]}]"
        return string
    
    @staticmethod
    def _extract_date(shift_string):
        return re.search(r".*(?= from )", shift_string)[0].split(", ")[1]

    def list_headers(self):
        standard_columns = [
            "Full name",
            "First name",
            "Last name",
            "cell",
            "Recipient",
            "Walkthrough shifts",
            "Phonebank shifts",
        ]
        additional_columns = list(self.other_columns.keys())
        return standard_columns + additional_columns

    def process_shifts_list(self):
        self.walkthrough_shifts = self.shifts_to_list(self.walkthrough_shifts)
        self.phonebank_shifts = self.shifts_to_list(self.phonebank_shifts)


def columns_lookup(config, column_number):
    if column_number not in config["columns"]:
        column_number = column_number - 1
    return config["columns"][column_number]
    # first_day = 8
    # remainder = index % 5
    # if remainder in (1, 2):
        # base = index // 5
        # day = first_day + base
        # return f"{day_lookup[day]}, 11/{day:02d}"


def weekday_columns(config):
    columns = []
    for column in config["columns"]:
        columns.append(column)
        columns.append(column + 2)
    return columns

def weekend_columns(config):
    columns = []
    for column in config["weekend_columns"]:
        columns.append(column)
        columns.append(column + 2)
    return columns

def good_columns(config):
    return weekday_columns(config) + weekend_columns(config)

def schedule_lookup(config, row_number, column_number):
    """Convert from a row and column number to event date, time and shift type.

    row_number and column_number should be 1-based!!!
    """
    if column_number in weekday_columns(config):
        row_key = "rows"
        column_key = "columns"
    elif column_number in weekend_columns(config):
        row_key = "weekend_rows"
        column_key = "weekend_columns"
    else:
        raise SpreadsheetLocationError(f"Column {column_number} isn't a cell for someone's name")
    # For columns, could be Organizer 1 or Organizer 2 so have to check
    if column_number in config[column_key]:
        date = config[column_key][column_number]
    elif column_number - 2 in config[column_key]:
        date = config[column_key][column_number - 2]
    else:
        raise SpreadsheetLocationError(f"Column {column_number} isn't conforming to the config!")

    # Look up the shift type and time from the row
    time_block_defs = config[row_key]
    last_block = [None, None]
    for time_block_start_row in sorted(time_block_defs):
        if time_block_start_row > row_number:
            if last_block is None:
                # Then this row should be ignored ...
                return None, None, None
            time, shift_type = last_block
        else:
            last_block = time_block_defs[time_block_start_row]
    return date, time, shift_type

def parse_turfHQ(turf_string):
    split = turf_string.split("//")
    # Replace blank turf ("") with None
    turf = split[0].strip() if split[0].strip() else None
    if len(split)<2:
        return turf, None
    else:
        return turf, split[1].strip()


def hour_24_to_12(hour_24):
    return (hour_24 - 1) % 12 + 1


def ampm(hour):
    if hour < 12:
        return "AM"
    if hour >= 12:
        return "PM"


### Regexes
NAME_REGEX = r"^[^0-9(\n]+[^0-9,?(\n- ]"
PHONE_REGEX = r"((\(?[0-9]\)?[-.]?){10})"
EMAIL_REGEX = r"\S+@\S+"  # An @ surrounded by 1 or more non-whitespace characters


def extract_phone_email(content):
    phone = re.search(PHONE_REGEX, content)
    if phone:
        phone = phone.group(0)
        phone = "".join([s for s in phone if s in "0123456789"])
    email = re.search(EMAIL_REGEX, content)
    if email:
        email = email.group(0)
    return phone, email


def scan_csv(filename, config):
    signups = []
    with open(filename, "r") as infile:
        csv_reader = csv.reader(infile)
        for row_index, row in enumerate(csv_reader):
            row_number = row_index + 1
            if row_number < min(config["rows"].keys()):
                continue
            for column_number in good_columns(config):
                column_index = column_number - 1
                name = row[column_index]
                email_phone_string = row[column_index + 1]
                turf_string = row[column_index + 4 if column_index+2 in good_columns(config) else column_index + 2]
                signup = parse_cell(
                    name, email_phone_string, turf_string, row_index, column_index, config
                )
                if signup is not None:
                    signups.append(signup)
    return signups


def parse_cell(name, email_phone_string, turf_string, row_index, column_index, config):
    """Create a SignupCell object based on the content and row/column
    of a spreadsheet cell.

    row_index and column_index should be 0-based!!!
    """
    if len(name) <= 5:
        return
    column_number = column_index + 1
    row_number = row_index + 1
    phone, email = extract_phone_email(email_phone_string)
    date, time, shift_type = schedule_lookup(config, row_number, column_number)
    turf, hq = parse_turfHQ(turf_string)
    if time is not None:
        return (
            SignupCell(
                [name, email_phone_string],
                row_number,
                column_number,
                date,
                time,
                shift_type,
                name,
                phone,
                email,
                turf,
                hq
            )
        )


def aggregate_signups(signups):
    people = {}
    for signup in signups:
        if signup.name in people:
            if signup.shift_type == "walkthrough":
                people[signup.name].walkthrough_shifts.append(
                    (signup.date, signup.time, signup.turf, signup.hq)
                )
            elif signup.shift_type == "phonebank":
                people[signup.name].phonebank_shifts.append((signup.date, signup.time))
            else:
                raise ValueError(f"Invalid shift_type: {signup.shift_type}")
        else:
            if signup.shift_type == "walkthrough":
                people[signup.name] = PersonSchedule(
                    signup.name,
                    signup.phone,
                    walkthrough_shifts=[(signup.date, signup.time, signup.turf, signup.hq)],
                )
            elif signup.shift_type == "phonebank":
                people[signup.name] = PersonSchedule(
                    signup.name,
                    signup.phone,
                    phonebank_shifts=[(signup.date, signup.time)],
                )
            else:
                raise ValueError(f"Invalid shift_type: {signup.shift_type}")
    return list(people.values())


def write_csv(filename, people):
    first_person = people[0]
    with open(filename, "w") as outfile:
        csv_writer = csv.writer(outfile)
        csv_writer.writerow(first_person.list_headers())
        for person in people:
            csv_writer.writerow(person.to_list())


def load_grid_schedule(filename, config):
    signup_cells = scan_csv(filename, config)
    people = aggregate_signups(signup_cells)
    return sorted(people)


def scan_mailmerge_csv(filename):
    people = {}
    with open(filename, "r") as infile:
        csv_reader = csv.reader(infile)
        for i, row in enumerate(csv_reader):
            print(row)
            if i == 0:
                header = row
                additional_columns = row[6:]
                continue
            additional_values = dict(zip(additional_columns, row[6:]))
            person_row = MailMergeRow(*row[:6], additional_values)
            person_row.process_shifts_list()
            people[person_row.full_name.lower()] = person_row
    return people


def daily_shifts(date_str, existing_mailmerge_filename, output_filename):
    existing_people = list(scan_mailmerge_csv(existing_mailmerge_filename).values())
    specific_date_people = []
    for person in existing_people:
        walkthroughs = [
            shift for shift in person.walkthrough_shifts if date_str in shift[0]
        ]
        phonebanks = [
            shift for shift in person.phonebank_shifts if date_str in shift[0]
        ]
        if walkthroughs or phonebanks:  # i.e. if either is not empty
            new_person = copy.deepcopy(person)
            new_person.walkthrough_shifts = walkthroughs
            new_person.phonebank_shifts = phonebanks
            specific_date_people.append(new_person)
    write_csv(output_filename, specific_date_people)


def update_csv(grid_filename, existing_mailmerge_filename, output_filename, config):
    new_version_people = load_grid_schedule(grid_filename, config)
    existing_people = scan_mailmerge_csv(existing_mailmerge_filename)
    for new_version_person in new_version_people:
        if new_version_person.name.lower() in existing_people:
            existing_row = existing_people[new_version_person.name.lower()]
            # update phonebank and walkthrough shifts
            existing_row.walkthrough_shifts = new_version_person.walkthrough_shifts
            existing_row.phonebank_shifts = new_version_person.phonebank_shifts
        else:
            new_mailmergerow = MailMergeRow(
                new_version_person.name,
                new_version_person.first_name(),
                new_version_person.last_name(),
                new_version_person.phone,
                new_version_person.walkthrough_shifts,
                new_version_person.phonebank_shifts,
                {},
            )
            print(f"New person: {new_mailmergerow}")
            existing_people[new_mailmergerow.full_name.lower()] = new_mailmergerow
    new_version_names = {person.name.lower() for person in new_version_people}
    to_delete = []
    for existing_person_name in existing_people:
        if existing_person_name not in new_version_names:
            to_delete.append(existing_person_name)
    for name in to_delete:
        del existing_people[name]
    write_csv(output_filename, list(existing_people.values()))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="""Convert the SAV organizing shift signup schedule grid
into a consolidated spreadsheet with 1 row per person
listing all their shifts.

Suggested workflow is to run this once to obtain a starting spreadsheet,
then add people's email addresses in a new column to the right.
Then when you run the script again, include "--update <existing_output.csv>"
so that the email addresses and other custom columns are preserved
and copied appropriately into the new output.
Do not rearrange the first 6 columns of the output!
"""
    )
    parser.add_argument("infile")
    parser.add_argument("outfile")
    parser.add_argument("--update")
    parser.add_argument("--daily")
    parser.add_argument("--config", default="config.json")
    args = parser.parse_args()
    with open(args.config, "r") as configfile:
        config = json.load(configfile)
        for sub_dict in config:
            for k, v in config[sub_dict].copy().items():
                config[sub_dict][int(k)] = v
                del config[sub_dict][k]
    if args.update is None and args.daily is None:
        people = load_grid_schedule(args.infile, config)
        write_csv(args.outfile, people)
    elif args.daily is None:
        update_csv(args.infile, args.update, args.outfile, config)
    else:
        daily_shifts(args.daily, args.infile, args.outfile)
