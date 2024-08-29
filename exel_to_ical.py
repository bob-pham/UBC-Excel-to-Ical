from datetime import datetime, timezone
import argparse
from typing import List
import icalendar
from pathlib import Path
import pandas as pd
import calendar

STR_TO_CAL_DATE = {
    "Mon": calendar.MONDAY,
    "Tue": calendar.TUESDAY,
    "Wed": calendar.WEDNESDAY,
    "Thu": calendar.THURSDAY,
    "Fri": calendar.FRIDAY,
    "Sat": calendar.SATURDAY,
    "Sun": calendar.SUNDAY,
}

STR_TO_RFC_DATE = {
    "Mon": "MO",
    "Tue": "TU",
    "Wed": "WE",
    "Thu": "TH",
    "Fri": "FR",
    "Sat": "SA",
    "Sun": "SU",
}


def excel_to_ical(excel_name: Path, output_name: str):
    df = pd.read_excel(excel_name)
    cal = icalendar.Calendar()

    # Read dates
    for _, row in df.iterrows():
        parse_excel_row(cal, row)

    with open(f"{output_name}.ics", "wb") as fp:
        fp.write(cal.to_ical())


def excel_to_ical_recurring(excel_name: Path, output_name: str):
    df = pd.read_excel(excel_name)
    cal = icalendar.Calendar()

    # Read dates
    for _, row in df.iterrows():
        parse_excel_row_recurring(cal, row)

    with open(f"{output_name}.ics", "wb") as fp:
        fp.write(cal.to_ical())


def parse_excel_row(current_cal, row):
    instruction_format = row["Instructional Format"]
    section = row["Section"]

    dates = row["Meeting Patterns"]
    for local_date in dates.split("\n"):
        if len(local_date) <= 0:
            continue

        [duration, days, time, location] = local_date.split(" | ")
        [start, end] = duration.split(" - ")
        [start_time, end_time] = time.split(" - ")

        date_objs = get_all_dates_in_interval(start, end, days.split(" "))

        for date_obj in date_objs:
            current_cal.add_component(
                create_event(
                    f"{section} {instruction_format}",
                    format_time(start_time),
                    format_time(end_time),
                    location,
                    date_obj,
                )
            )


def parse_excel_row_recurring(current_cal, row):
    instruction_format = row["Instructional Format"]
    section = row["Section"]

    dates = row["Meeting Patterns"]
    for local_date in dates.split("\n"):
        if len(local_date) <= 0:
            continue

        [duration, days, time, location] = local_date.split(" | ")
        [start, end] = duration.split(" - ")
        [start_time, end_time] = time.split(" - ")

        start_date = datetime.strptime(start, "%Y-%m-%d")
        end_date = datetime.strptime(end, "%Y-%m-%d")

        # end_date = end_date.replace(tzinfo=timezone.utc)
        # end_str = end_date.strftime('%Y%m%dT%H%M%SZ')

        days = list(map(lambda day: STR_TO_RFC_DATE[day], days.split(" ")))

        event = create_event(
            f"{section} {instruction_format}",
            format_time(start_time),
            format_time(end_time),
            location,
            start_date,
        )

        event.add("rrule", {"FREQ": "WEEKLY", "BYDAY": days, "UNTIL": end_date})

        current_cal.add_component(event)


def get_all_dates_in_interval(start: str, end: str, days: List[str]):
    events = []
    start_date = datetime.strptime(start, "%Y-%m-%d")
    end_date = datetime.strptime(end, "%Y-%m-%d")

    current_date = start_date

    while current_date <= end_date:
        month_calendar = calendar.monthcalendar(current_date.year, current_date.month)

        # Loop through the weeks in the month calendar
        for week in month_calendar:
            for day in days:
                if (
                    week[STR_TO_CAL_DATE[day]] != 0
                ):  # Check if there is a Monday in the week
                    found_date = datetime(
                        current_date.year,
                        current_date.month,
                        week[STR_TO_CAL_DATE[day]],
                    )
                    if start_date <= found_date <= end_date:
                        events.append(found_date)

        # Move to the next month
        next_month = current_date.month + 1 if current_date.month < 12 else 1
        next_year = (
            current_date.year if current_date.month < 12 else current_date.year + 1
        )
        current_date = datetime(next_year, next_month, 1)

    return events


def create_event(section, start_time, end_time, location, date_obj):
    event = icalendar.Event()

    event.add("name", section)
    event.add("summary", section)

    time_obj = datetime.strptime(start_time, "%I:%M %p").time()
    start_obj = date_obj.replace(
        hour=time_obj.hour, minute=time_obj.minute, second=time_obj.second
    )

    time_obj = datetime.strptime(end_time, "%I:%M %p").time()
    end_obj = date_obj.replace(
        hour=time_obj.hour, minute=time_obj.minute, second=time_obj.second
    )

    event.add("dtstart", start_obj)
    event.add("dtend", end_obj)
    event["location"] = location

    return event


def format_time(time: str):
    return time.replace(".", "")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        prog="exel_to_ical.py",
        description="Script that reads an excel file, and outputs an ICal with corresponding dates",
    )

    parser.add_argument(
        "-i", "--input", help="The input Excel file name", required=True
    )
    parser.add_argument(
        "-d", "--destination", help="The destination file name", required=False
    )
    parser.add_argument(
        "-e",
        "--events",
        help="Saves as individual events instead of recurring ones",
        required=False,
        action="store_true",
    )

    args = parser.parse_args()

    input_name = Path(args.input)
    if not input_name.is_file():
        raise Exception("Invalid Input")

    if input_name.suffix != ".xlsx":
        raise Exception("Requires Excel File Input")

    output_name = args.destination if args.destination else input_name.stem

    if args.events:
        excel_to_ical(input_name, output_name)
    else:
        excel_to_ical_recurring(input_name, output_name)
