# MZOutlookAppointmentTools

This library simplifies the creation of recurring Outlook AppointmentItems using standard iCalendar recurrence patterns. It also enables the retrieval of recurrence pattern strings from existing Outlook recurring AppointmentItems. Currently the library provides some tools to parse,generate and compare iCalendar recurrence rules for Outlook appointment items

## Limitations
  1. The library currently just supports these recurrence types:  "DAILY" / "WEEKLY" / "MONTHLY" / "YEARLY". This is because these are the only supported types by Outlook. As a consequence all the related keywords (e.g. "BYHOUR","BYMINUTE","BYSECOND") are not supported.  
  2. The "WKST" is not supported as it seems that there are no way to handle it in Outlook.  
  3. Patterns containing "BYWEEKNO" are not supported.
  4. Patterns containing BYDAY with numbers such as "FREQ=MONTHLY;INTERVAL=1;BYDAY=1FR,3MO" Are not supported

## Supported Pattern Samples
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

## Usage:
### SetRecurrencePattern:

```cs
 AppointmentItem aItem = (AppointmentItem)ApplicationInstance.CreateItem(OlItemType.olAppointmentItem);
 string pattern = "FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR";
 RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
 var recPattern = aItem.GetRecurrencePattern();
//Expected values:
//  recPattern.RecurrenceType == OlRecurrenceType.olRecursWeekly
//  recPattern.DayOfWeekMask == OlDaysOfWeek.olMonday | OlDaysOfWeek.olTuesday | OlDaysOfWeek.olWednesday | OlDaysOfWeek.olThursday | OlDaysOfWeek.olFriday
//  recPattern.Instance == 0
//  recPattern.Interval == 1
```

### GetRecurrencePattern

```cs
 AppointmentItem aItem = (AppointmentItem)ApplicationInstance.CreateItem(OlItemType.olAppointmentItem);
 var occ = aItem.GetRecurrencePattern();
 occ.RecurrenceType = OlRecurrenceType.olRecursDaily;
 occ.Interval = 1;
 aItem.Save();
 var item = RecurrenceStringTools.GetRecurrenceString(aItem); // "FREQ=DAILY;INTERVAL=1"
```

### AreEqual
```cs
  string pattern1 = "FREQ=WEEKLY;BYDAY=MO,TU,WE";
  string pattern2 = "FREQ=WEEKLY;BYDAY=TU,WE,MO";
  var areEqual = RecurrenceStringTools.AreEqual(pattern1, pattern2);// true
```

More samples could be found in the test project. 

## Discution
1. This pattern "FREQ=MONTHLY;INTERVAL=1;BYDAY=FR" would be converted to a `olRecursWeekly` as `olRecursMonthNth` requires the `instance` field to be set to a value from 1 to 5. So when it is converted back to recurrence rule the value would be simply "FREQ=WEEKLY;INTERVAL=1;BYDAY=FR" (Check these tests: `Combo_MonthlyNthSetPos_Existence1` and `Combo_MonthlyNthSetPos_Existence2` in `RecurrenceStringToolsTest`). Moreover, whenever we get the recurrence pattern from an AppointmentItem with  `olRecursMonthNth`, we always include the "BYSETPOS".
