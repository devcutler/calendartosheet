# calendartosheet

### Usage
```
./calendartosheet.exe --help

calendartosheet 1.0.0
Copyright (C) 2024 calendartosheet

  -f, --format       (Default: name description start duration end comments status location organizer) Output column values.
                     Usable values: name, summary, description, start, end, duration, length, comments, status, location, geolocation, organizer, creator, 
                     calendar, type, attendees, created, modified, lastmodified, id, uid, uuid, priority, url, link

  -x, --xlsx         (Default: false) Output XLSX instead of CSV.

  -h, --header       (Default: true) Output a header row with the names of the columns.

  --help             Display this help screen.

  --version          Display version information.

  input (pos. 0)     Required. Input file to convert.

  output (pos. 1)    Required. File path to output to.

```
#### Examples
```
./calendartosheet.exe calendar.ics events.csv
./calendartosheet.exe calendar.ics events.xlsx --xlsx
./calendartosheet.exe calendar.ics events.csv --format name,description,start,duration,end
```
