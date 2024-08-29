# Excel To ICS

Simple script that parses UBC's workday generated excel files into `.ics` files

### Requires:
- python
- pipenv
- pandas
- icalendar
- openpyxl


### Install depedencies:

```
pipenv install
```

# Usage:

#### Basic Usage
```shell
pipenv run python3 exel_to_ical.py -i <excel-name>.xlsx
```

### Optionally specify name of the calendar
```shell
pipenv run python3 exel_to_ical.py -i <excel-name>.xlsx -d <output-calendar-name>
```

### Return all events as individual events instead of a recurring series
```shell
pipenv run python3 exel_to_ical.py -i <excel-name>.xlsx -d <output-calendar-name> -e
```

### Help

```shell
pipenv run python excel_to_ical.py -h
```
