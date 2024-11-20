# MZOutlookAppointmentTools

This library simplifies the creation of recurring Outlook AppointmentItems using standard iCalendar recurrence patterns. It also enables the retrieval of recurrence pattern strings from existing Outlook recurring AppointmentItems. Currently the library provides some tools to parse,generate and compare iCalendar recurrence rules for Outlook appointment items

## Limitations
  1. The library currently just supports these recurrence types:  "DAILY" / "WEEKLY" / "MONTHLY" / "YEARLY". This is because these are the only supported types by Outlook. As a consequence all the related keywords (e.g. "BYHOUR","BYMINUTE","BYSECOND") are not supported.  
  2. The "WKST" is not supported as it seems that there are no way to handle it in Outlook.  
  3. Patterns containing "BYWEEKNO" are not supported. 

## Supported Samples
  -  "FREQ=WEEKLY;BYDAY=SA"
  -  "FREQ=WEEKLY;BYDAY=MO,TU,WE"
  -  "FREQ=WEEKLY;BYDAY=MO,TU,WE;INTERVAL=2"
  -  "FREQ=MONTHLY;BYMONTHDAY=12"
  -  "FREQ=MONTHLY;BYDAY=MO,TU,WE,TH,FR;BYSETPOS=1"
  -  "FREQ=MONTHLY;BYDAY=MO,TU,WE,TH,FR;BYSETPOS=-1"
  -  "FREQ=YEARLY;BYMONTHDAY=7;BYMONTH=3"
  -  "FREQ=YEARLY;BYDAY=FR;BYMONTH=11;BYSETPOS=2"
  -  "FREQ=YEARLY;BYDAY=FR;BYMONTH=11;BYSETPOS=3;INTERVAL=3"
  -  "FREQ=YEARLY;BYDAY=MO,TU,WE,TH,FR;BYMONTH=9;BYSETPOS=1"
