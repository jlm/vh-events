Introduction
============

This utility reads a [Vcalendar](https://icalendar.org) feed and a template Excel spreadsheet,
and produces an Excel spreadsheet populated with selected events from the calendar.  It is intended
for parsing the feed produced by [Hallmaster](https://www.hallmaster.co.uk) for publication in a village
magazine.  The output table has a
section for weekly events followed by a list of other events (which are typically one-off events or
events on a non-weekly cycle). The code doesn't concern itself with text styles, fonts or other
presentation issues, instead preserving those from the template spreadsheet.

Usage
=====
Create the Excel template file (which is in XLSX format) and place it in the `template` directory.
An example is provided as `template/table-only.xlsx`.

Typical usage:
```aiignore
ruby vh-events.rb --excel tmp/event-table.xlsx
```

By default:
* the Vcalendar event feed is read from the URL specified in the configuration file. This can
be overridden with the `--read-file` flag.
* only events from the month following the current month are selected.  This can be overridden with the
`--year` and `--month` flags.
```aiignore
usage: vh_events.rb [options]
    -c, --config                configuration YAML file name (default: config.yml)
    -C, --config-from-template  take rewrite rules from the template file
    -t, --template              name of the Excel template file to read
    -r, --read-file             Read iCal data from file (default: read from URL in config file)
    -m, --month                 the (numeric) month to search (default: next month)
    -y, --year                  the year to search
    -d, --debug                 debug mode
    -v, --verbose               be verbose: list extra detail
    -a, --print-all             print all events
    -j, --json                  output results in JSON to the named file
    -x, --excel                 output results in XLSX to the named file
    -l, --list                  list events to standard output
    --help
```
Weekly events
=============

The substance of the code is involved with recognising and grouping weekly events
to make a useful table.  Although Hallmaster can classify events as weekly,
people booking events often don't book them as such. Instead, a set of rules can be provided
which specify which events are weekly, and allow re-writing of the event names and modification
of their times.  These rules can be provided in a 
configuration file (typically `config.yml`) or in the Excel template file in a special 'Config' tab.
Re-writing the event names allows the published name of an event to differ from the way it is shown
in the event feed.  Re-writing event times allows the published times to be different from the booked
times, to allow for set-up and clearing up time.

License and Copyright
=======
This program is licensed under the Apache License.  See the `LICENSE` file for details.

Copyright :copyright: 2025 John Messenger.
