# bulky-ol-cal
Bulk Import of Outlook Calendar Entries from a JSON file

## About

This program reads appointments/meetings from json file and uses Outlook Interop (Outlook.Application) to create meetings.

https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.appointmentitem?view=outlook-pia


## Data Structure

The following json structure shows all AppointmentItem/RecurrencePattern fields which are supported in this tool.

In practice don't mix all the fields. The API documentation for AppointmentItem and RecurrencePatterns will establish the fields which can be used in combination.  The project includes a larger sample json that shows various RecurrencePatterns.

```
[
    {
        "Recipients": [
            {"email": "person1@example.com","type": "Required"},
            {"email": "person2@example.com","type": "Optional"},
            {"email": "confroom@example.com", "type": "Resource"}],
        "Subject": "Meeting Subject",
        "Body": "Meeting Body",
        "Location": "Meeting Location",
        "Importance": "Low|Normal|High",
        "Sensitivity": "Normal|Confidential|Personal|Private",
        "AllDayEvent": false,
        "BusyStatus": "Free|Tentative|Busy|OutOfOffice|WorkingElsewhere",
        "Start": "2023-01-01T14:30:00-05:00",
        "End": "2023-01-01T15:00:00-05:00",
        "Duration": 30,
        "ReminderSet": true,
        "ReminderMinutesBeforeStart": 15,
        "ResponseRequested": false,
        "RecurrencePattern":
        {
            "RecurrenceType": "Daily|Weekly|Monthly|MonthlyNth|Yearly|YearlyNth",
            "Interval": 1,
            "MonthOfYear": 1,
            "DayOfMonth": 1,
            "DayOfWeekMask": ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","WeekDays"],
            "Instance": 1,
            "PatternStartDate": "2023-03-01T00:00-05:00",
            "PatternEndDate": "2023-03-31T00:00-05:00",
            "Occurrences": 1,
            "NoEndDate": false,
            "StartTime": "1980-01-01T09:00-05:00",
            "EndTime": "1980-01-01T09:15-05:00",
            "Duration": 15
        }
   }
]
```

## Example Usage

Test an input file.  Checks json file, string/enum attribute mappings and creates/deletes meetings

```
bulky-ol-cal --test mymeetings.json
```

Test an input file & visually inspect each meeting one at a time

```
bulky-ol-cal --test --display --modal mymeetings.json
```

Create the meetings

```
bulky-ol-cal mymeetings.json
```

Create the meetings and display them all to make additional edits or manually send

```
bulky-ol-cal mymeetings.json
```

Create the meetings and display them all one at a time to make additional edits [manually save the meeting] then send

```
bulky-ol-cal --display --modal --send mymeetings.json
```

## Troubleshooting

* For RecurrenceType `Yearly` & `YearlyNth`, the Interval value must be the number of months, not years.  For example, Interval = 12 for 1 year.  Interval = 36 for 3 years.

## Todo

* Consider adding O365 cloud support
