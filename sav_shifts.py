import argparse
import copy
import csv
from dataclasses import dataclass, field as dc_field
import re

import pandas as pd


@dataclass
class SignupCell:
    content: str
    row: int
    column: int
    date: str
    time: str
    shift_type: str
    name: str = None
    phone: str = None


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
    def shifts_to_str(shift_list):
        return "\n".join(
            sorted(
                [f"{shift[0]} from {shift[1]}" for shift in shift_list],
                key=lambda shift_str: shift_str.split(" from ")[0].split(", ")[1],
            )
        )

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
    walkthrough_shifts: list
    phonebank_shifts: list
    other_columns: dict

    def to_list(self):
        standard_columns = [
            self.full_name,
            self.first_name,
            self.last_name,
            self.phone,
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
        return [x.split(" from ") for x in shifts_str_list]

    @staticmethod
    def shifts_to_str(shift_list):
        return "\n".join(
            sorted(
                [f"{shift[0]} from {shift[1]}" for shift in shift_list],
                key=lambda shift_str: shift_str.split(" from ")[0].split(", ")[1],
            )
        )

    def list_headers(self):
        standard_columns = [
            "Full name",
            "First name",
            "Last name",
            "cell",
            "Walkthrough shifts",
            "Phonebank shifts",
        ]
        additional_columns = list(self.other_columns.keys())
        return standard_columns + additional_columns

    def process_shifts_list(self):
        self.walkthrough_shifts = self.shifts_to_list(self.walkthrough_shifts)
        self.phonebank_shifts = self.shifts_to_list(self.phonebank_shifts)


def load_5year(filename):
    full_list = pd.read_csv(filename)
    required_fields = ["FullName", "Email"]
    return full_list[required_fields]


def columns_lookup(index):
    first_day = 8
    remainder = index % 5
    if remainder in (1, 2):
        base = index // 5
        day = first_day + base
        return f"{day_lookup[day]}, 11/{day:02d}"


good_columns = [x for x in range(59) if (x % 5 in (1, 2)) and (x < 26 or x > 34)]
weekend_columns = [26, 27, 31, 32]
day_lookup = {
    8: "Monday",
    9: "Tuesday",
    10: "Wednesday",
    11: "Thursday",
    12: "Friday",
    13: "Saturday",
    14: "Sunday",
    15: "Monday",
    16: "Tuesday",
    17: "Wednesday",
    18: "Thursday",
    19: "Friday",
}


def rows_lookup(index):
    first_hour = 10
    walkthrough_start_row = 6
    num_rows_slot_10to11 = 10
    num_rows_slot_11to12 = 11
    num_rows_slot_12to1 = 11
    num_rows_slot_1to2 = 8
    num_rows_slot_2to3 = 9
    num_rows_slot_3to430 = 8
    num_rows_slot_debrief = 4
    num_rows_slot_5to6 = 5
    num_rows_slot_6to7 = 6
    num_rows_slot_7to8 = 4

    row_lengths = [
        walkthrough_start_row,
        num_rows_slot_10to11,
        num_rows_slot_11to12,
        num_rows_slot_12to1,
        num_rows_slot_1to2,
        num_rows_slot_2to3,
        num_rows_slot_3to430,
        num_rows_slot_debrief,
        num_rows_slot_5to6,
        num_rows_slot_6to7,
        num_rows_slot_7to8,
    ]

    row_boundaries = [sum(row_lengths[:i]) for i in range(1, len(row_lengths) + 1)]

    if index < walkthrough_start_row:
        return (None, None)
    for i, boundary in enumerate(row_boundaries):
        if index < boundary:
            hour = i - 1 + first_hour
            break
    else:  # if no break
        raise ValueError("Couldn't find appropriate hour slot")
    if hour == 15:
        time_str = "3:00PM - 4:30PM"
    else:
        time_str = (
            f"{hour_24_to_12(hour)}:00{ampm(hour)} - "
            f"{hour_24_to_12(hour + 1)}:00{ampm(hour + 1)}"
        )
    if hour < 16:
        return ("walkthrough", time_str)
    elif hour == 16:
        return (None, None)
    else:
        return ("phonebank", time_str)


def weekend_rows_lookup(index):
    walkthrough_start_row = 6
    num_rows_slot_10to1 = 10
    num_rows_slot_2to5 = 5
    num_rows_slot_phonebanklabel = 6
    num_rows_slot_2to3 = 11
    num_rows_slot_3to4 = 8
    num_rows_slot_4to5 = 9
    num_rows_slot_5to6 = 11

    row_lengths = [
        walkthrough_start_row,
        num_rows_slot_10to1,
        num_rows_slot_2to5,
        num_rows_slot_phonebanklabel,
        num_rows_slot_2to3,
        num_rows_slot_3to4,
        num_rows_slot_4to5,
        num_rows_slot_5to6,
    ]

    row_boundaries = [sum(row_lengths[:i]) for i in range(1, len(row_lengths) + 1)]

    if index < walkthrough_start_row:
        return (None, None)
    for i, boundary in enumerate(row_boundaries):
        if index < boundary:
            shift_slot = i
            break
    else:  # if no break
        raise ValueError("Couldn't find appropriate hour slot")
    shift_slots = [
        "10:00AM - 1:00PM",
        "2:00PM - 5:00PM",
        None,
        "2:00PM - 3:00PM",
        "3:00PM - 4:00PM",
        "4:00PM - 5:00PM",
        "5:00PM - 6:00PM",
    ]
    time_str = shift_slots[shift_slot]
    if shift_slot < 2:
        return ("walkthrough", time_str)
    elif shift_slot == 2:
        return (None, None)
    else:
        return ("phonebank", time_str)


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


def extract_name_phone(content):
    name = re.search(NAME_REGEX, content)
    if name:
        name = name.group(0)
    phone = re.search(PHONE_REGEX, content)
    if phone:
        phone = phone.group(0)
        phone = "".join([s for s in phone if s in "0123456789"])
    return (name, phone)


def scan_csv(filename):
    signups = []
    with open(filename, "r") as infile:
        csv_reader = csv.reader(infile)
        for row_number, row in enumerate(csv_reader):
            if row_number < 6:
                continue
            for column in good_columns:
                if len(row[column]) > 5:
                    content = row[column]
                    name, phone = extract_name_phone(content)
                    date = columns_lookup(column)
                    shift_type, time = rows_lookup(row_number)
                    if time is not None:
                        signups.append(
                            SignupCell(
                                content,
                                row_number,
                                column,
                                date,
                                time,
                                shift_type,
                                name,
                                phone,
                            )
                        )
            for column in weekend_columns:
                if len(row[column]) > 5:
                    content = row[column]
                    name, phone = extract_name_phone(content)
                    date = columns_lookup(column)
                    shift_type, time = weekend_rows_lookup(row_number)
                    if time is not None:
                        signups.append(
                            SignupCell(
                                content,
                                row_number,
                                column,
                                date,
                                time,
                                shift_type,
                                name,
                                phone,
                            )
                        )
    return signups


def aggregate_signups(signups):
    people = {}
    for signup in signups:
        if signup.name in people:
            if signup.shift_type == "walkthrough":
                people[signup.name].walkthrough_shifts.append(
                    (signup.date, signup.time)
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
                    walkthrough_shifts=[(signup.date, signup.time)],
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


def load_grid_schedule(filename):
    signup_cells = scan_csv(filename)
    people = aggregate_signups(signup_cells)
    return sorted(people)


def scan_mailmerge_csv(filename):
    people = {}
    with open(filename, "r") as infile:
        csv_reader = csv.reader(infile)
        for i, row in enumerate(csv_reader):
            if i == 0:
                header = row
                additional_columns = row[6:]
                continue
            additional_values = dict(zip(additional_columns, row[6:]))
            person_row = MailMergeRow(*row[:6], additional_values)
            person_row.process_shifts_list()
            people[person_row.full_name] = person_row
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


def update_csv(grid_filename, existing_mailmerge_filename, output_filename):
    new_version_people = load_grid_schedule(grid_filename)
    existing_people = scan_mailmerge_csv(existing_mailmerge_filename)
    for new_version_person in new_version_people:
        if new_version_person.name in existing_people:
            existing_row = existing_people[new_version_person.name]
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
            existing_people[new_mailmergerow.full_name] = new_mailmergerow
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
    args = parser.parse_args()
    if args.update is None and args.daily is None:
        people = load_grid_schedule(args.infile)
        write_csv(args.outfile, people)
    elif args.daily is None:
        update_csv(args.infile, args.update, args.outfile)
    else:
        daily_shifts(args.daily, args.infile, args.outfile)
