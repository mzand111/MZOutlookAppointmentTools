using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;

namespace MZOutlookAppointmentTools.iCalendarTools
{
    public partial class RecurrenceStringTools
    {
        //Returns standard recurrence pattern string from RecurrencePattern of an Outlook AppoingmentItem
        public static string GetRecurrenceString(AppointmentItem myItem)
        {
            // Returns a properly formatted recurrence string for the item.
            RecurrencePattern pattern = myItem.GetRecurrencePattern();
            string str = "RRULE:";
            try
            {
                switch (pattern.RecurrenceType)
                {
                    case OlRecurrenceType.olRecursDaily:
                        str += "FREQ=DAILY";
                        if (!pattern.NoEndDate)
                        {
                            str += ";UNTIL=" + FormatICalDateTime(pattern.PatternEndDate);
                            // End datetime issue fix to be from 12:00am to 11:59:59pm.
                            str = str.Replace("T000000", "T235959");
                        }
                        str += ";INTERVAL=" + pattern.Interval;
                        break;

                    case OlRecurrenceType.olRecursMonthly:
                        str += "FREQ=MONTHLY";
                        if (!pattern.NoEndDate)
                        {
                            str += ";UNTIL=" + FormatICalDateTime(pattern.PatternEndDate);
                        }
                        str += ";INTERVAL=" + pattern.Interval;
                        str += ";BYMONTHDAY=" + pattern.DayOfMonth;
                        break;

                    case OlRecurrenceType.olRecursMonthNth:
                        str += "FREQ=MONTHLY";
                        if (!pattern.NoEndDate)
                        {
                            str += ";UNTIL=" + FormatICalDateTime(pattern.PatternEndDate);
                        }
                        str += ";INTERVAL=" + pattern.Interval;
                        if (pattern.Instance == 5)
                        {
                            str += ";BYWEEK=-1";
                            str += ";BYDAY=" + DaysOfWeek("", pattern);
                        }
                        else
                        {
                            str += ";BYDAY=" + DaysOfWeek(WeekNum(pattern.Instance), pattern);
                        }
                        break;

                    case OlRecurrenceType.olRecursWeekly:
                        str += "FREQ=WEEKLY";
                        if (!pattern.NoEndDate)
                        {
                            str += ";UNTIL=" + FormatICalDateTime(pattern.PatternEndDate);
                        }
                        str += ";INTERVAL=" + pattern.Interval;
                        str += ";BYDAY=" + DaysOfWeek("", pattern);
                        break;

                    case OlRecurrenceType.olRecursYearly:
                        str += "FREQ=YEARLY";
                        if (!pattern.NoEndDate)
                        {
                            str += ";UNTIL=" + FormatICalDateTime(pattern.PatternEndDate);
                        }
                        str += ";INTERVAL=1";  // Outlook does not support, every nth year in 
                        str += ";BYDAY=" + DaysOfWeek("", pattern);
                        break;

                    case OlRecurrenceType.olRecursYearNth:
                        str += "FREQ=YEARLY";
                        if (!pattern.NoEndDate)
                        {
                            str += ";UNTIL=" + FormatICalDateTime(pattern.PatternEndDate);
                        }
                        str += ";BYMONTH=" + MonthNum(pattern.MonthOfYear);
                        str += ";BYDAY=" + DaysOfWeek(WeekNum(pattern.Instance), pattern);
                        break;
                }

                return str;
            }
            finally
            {
                if (pattern != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(pattern);
            }
        }

        private static string FormatICalDateTime(DateTime date)
        {
            // Converts DateTime to iCalendar date-time format
            return date.ToString("yyyyMMdd\\THHmmss\\Z");
        }
        // Returns the days of the week for the recurrence rule
        private static string DaysOfWeek(string weekNum, RecurrencePattern pattern)
        {

            List<string> dayOfWeeks = new List<string>();

            if (pattern.DayOfWeekMask.HasFlag(OlDaysOfWeek.olSunday))
                dayOfWeeks.Add("SU");
            if (pattern.DayOfWeekMask.HasFlag(OlDaysOfWeek.olMonday))
                dayOfWeeks.Add("MO");
            if (pattern.DayOfWeekMask.HasFlag(OlDaysOfWeek.olTuesday))
                dayOfWeeks.Add("TU");
            if (pattern.DayOfWeekMask.HasFlag(OlDaysOfWeek.olWednesday))
                dayOfWeeks.Add("WE");
            if (pattern.DayOfWeekMask.HasFlag(OlDaysOfWeek.olThursday))
                dayOfWeeks.Add("TH");
            if (pattern.DayOfWeekMask.HasFlag(OlDaysOfWeek.olFriday))
                dayOfWeeks.Add("FR");
            if (pattern.DayOfWeekMask.HasFlag(OlDaysOfWeek.olSaturday))
                dayOfWeeks.Add("SA");
            return string.Join(",", dayOfWeeks);
        }
        // Returns the week number for the recurrence rule
        private static string WeekNum(int instance)
        {
            return instance.ToString();
        }
        // Converts month number to the correct format for iCalendar
        private static string MonthNum(int monthOfYear)
        {
            return monthOfYear.ToString("D2");
        }
    }
}
