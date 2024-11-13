using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;



namespace MZOutlookAppointmentTools.iCalendarTools
{
    public class RecurrenceStringTools
    {
        public static RecurrencePattern ParseRecurrencePattern(string recurrenceString, AppointmentItem appointmentItem, DateTime? start)
        {

            // Split the recurrence string into key-value pairs
            string[] parts = recurrenceString.Split(';');

            bool rt_set = false;
            OlRecurrenceType rt = OlRecurrenceType.olRecursDaily;
            bool bd_set = false;
            OlDaysOfWeek bd = OlDaysOfWeek.olSunday;
            bool bm_set = false;
            int bm = 0;
            bool bsp_set = false;
            int bsp = 0;
            bool interval_set = false;
            int interval = 0;
            bool endDate_set = false;
            DateTime endDate = DateTime.MinValue;
            Dictionary<String, String> ruleBook = new Dictionary<string, string>();

            foreach (string part in parts)
            {
                string[] kvp = part.Split('=');
                if (kvp.Length != 2)
                    continue;

                string key = kvp[0];
                string value = kvp[1];
                ruleBook.Add(key, value);

                switch (key)
                {
                    case "FREQ":
                        rt_set = true;
                        switch (value)
                        {
                            case "DAILY":
                                rt = OlRecurrenceType.olRecursDaily;
                                break;
                            case "WEEKLY":
                                rt = OlRecurrenceType.olRecursWeekly;
                                break;
                            case "MONTHLY":
                                rt = OlRecurrenceType.olRecursMonthly;
                                break;
                            case "YEARLY":
                                rt = OlRecurrenceType.olRecursYearly;
                                break;
                        }
                        break;

                    case "BYDAY":
                        bd_set = true;
                        bd = ParseDayOfWeekMask(value);
                        break;

                    case "BYMONTH":
                        bm_set = true;
                        bm = int.Parse(value);
                        break;

                    case "BYSETPOS":
                        bsp_set = true;
                        bsp = ParseBySetPos(value);
                        break;

                    case "INTERVAL":
                        interval_set = true;
                        interval = int.Parse(value);
                        break;

                    case "UNTIL":
                        endDate_set = true;
                        endDate = DateTime.ParseExact(value, "yyyyMMdd\\THHmmss\\Z", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.AssumeUniversal | System.Globalization.DateTimeStyles.AdjustToUniversal);
                        break;

                    case "WKST":
                        // Can not set FirstDayOfWeek in Outlook API
                        break;

                }
            }
            RecurrencePattern pattern = appointmentItem.GetRecurrencePattern();
            try
            {
                if (rt_set)
                {
                    if (rt == OlRecurrenceType.olRecursMonthly)
                    {
                        if (bsp_set)
                        //If BYSETPOS is set the recurrence is olRecursMonthNth
                        {
                            pattern.RecurrenceType = OlRecurrenceType.olRecursMonthNth;
                            pattern.Instance = bsp;
                            if (bd_set)
                            {
                                pattern.DayOfWeekMask = bd;
                                if (pattern.DayOfWeekMask == (OlDaysOfWeek)127 && bsp == 5 &&
                                   start.Value.Day > 28)
                                {
                                    //In Outlook this is simply a monthly recurring
                                    pattern.RecurrenceType = OlRecurrenceType.olRecursMonthly;
                                }
                                bd_set = false;
                            }
                            bsp_set = false;
                        }
                        else if (ruleBook.ContainsKey("BYDAY"))
                        //If BYDAY is set to -1, the recurrence is olRecursMonthNth
                        {
                            if (ruleBook["BYDAY"].StartsWith("-1"))
                            {
                                pattern.RecurrenceType = OlRecurrenceType.olRecursMonthNth;
                                pattern.Instance = 5;
                                pattern.DayOfWeekMask = ParseDayOfWeekMask(ruleBook["BYDAY"].TrimStart("-1".ToCharArray()));
                            }
                        }
                        else if (ruleBook.ContainsKey("BYMONTHDAY"))
                        {
                            pattern.RecurrenceType = OlRecurrenceType.olRecursMonthly;
                            if (bsp_set)
                            {
                                pattern.Instance = bsp;
                                bsp_set = false;
                            }
                            pattern.DayOfMonth = Int16.Parse(ruleBook["BYMONTHDAY"]);

                        }
                    }
                    else if (rt == OlRecurrenceType.olRecursYearly)
                    {

                        if (ruleBook.ContainsKey("BYSETPOS"))
                        {
                            pattern.RecurrenceType = OlRecurrenceType.olRecursYearNth;
                            int gInstance = Convert.ToInt16(ruleBook["BYSETPOS"]);
                            pattern.Instance = (gInstance == -1) ? 5 : gInstance;

                            pattern.DayOfWeekMask = ParseDayOfWeekMask(ruleBook["BYDAY"]);
                            if (ruleBook.ContainsKey("BYMONTH"))
                            {
                                pattern.MonthOfYear = Convert.ToInt16(ruleBook["BYMONTH"]);
                            }
                        }
                        else
                        {
                            pattern.RecurrenceType = rt;
                        }

                        if (ruleBook.ContainsKey("INTERVAL") && Convert.ToInt16(ruleBook["INTERVAL"]) > 1)
                        {
                            pattern.Interval = Convert.ToInt16(ruleBook["INTERVAL"]) * 12;
                            interval_set = false;
                        }
                        if (ruleBook.ContainsKey("BYMONTH"))
                        {
                            pattern.MonthOfYear = Convert.ToInt16(ruleBook["BYMONTH"]);
                        }
                        if (ruleBook.ContainsKey("BYMONTHDAY"))
                        {
                            //pattern.RecurrenceType = OlRecurrenceType.olRecursMonthly;
                            //if (bsp_set)
                            //{
                            //    pattern.Instance = bsp;
                            //    bsp_set = false;
                            //}
                            pattern.DayOfMonth = Int16.Parse(ruleBook["BYMONTHDAY"]);

                        }
                    }
                    else
                    {
                        pattern.RecurrenceType = rt;
                        pattern.DayOfWeekMask = bd;
                    }
                }
                if (interval_set)
                {
                    pattern.Interval = interval;
                }

                if (endDate_set)
                {
                    pattern.PatternEndDate = endDate;
                }
                return pattern;
            }
            finally
            {

            }
        }

        private static OlDaysOfWeek ParseDayOfWeekMask(string byDay)
        {
            OlDaysOfWeek mask = 0;
            string[] days = byDay.Split(',');

            foreach (string day in days)
            {
                switch (day)
                {
                    case "SU":
                        mask |= OlDaysOfWeek.olSunday;
                        break;
                    case "MO":
                        mask |= OlDaysOfWeek.olMonday;
                        break;
                    case "TU":
                        mask |= OlDaysOfWeek.olTuesday;
                        break;
                    case "WE":
                        mask |= OlDaysOfWeek.olWednesday;
                        break;
                    case "TH":
                        mask |= OlDaysOfWeek.olThursday;
                        break;
                    case "FR":
                        mask |= OlDaysOfWeek.olFriday;
                        break;
                    case "SA":
                        mask |= OlDaysOfWeek.olSaturday;
                        break;
                }
            }

            return mask;
        }

        private static int ParseBySetPos(string bySetPos)
        {
            int value = int.Parse(bySetPos); // For BYSETPOS=-1, set Instance to 5 to indicate the last instance of the specified day in the month
            return value == -1 ? 5 : value;
        }


        public static string GenRecurString(AppointmentItem myItem)
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
                            // Fixing the end date/time issue from 12:00am to 11:59:59pm.
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
                        str += ";INTERVAL=1";  // Can't do every nth year in Outlook
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

        private static string DaysOfWeek(string weekNum, RecurrencePattern pattern)
        {
            // Returns the days of the week for the recurrence rule
            // Implement your logic here
            // This function is a placeholder and needs to be implemented based on your requirements
            return "FR"; // Example placeholder return
        }

        private static string WeekNum(int instance)
        {
            // Returns the week number for the recurrence rule
            // Implement your logic here
            // This function is a placeholder and needs to be implemented based on your requirements
            return instance.ToString(); // Example placeholder return
        }

        private static string MonthNum(int monthOfYear)
        {
            // Converts month number to the correct format for iCalendar
            // Implement your logic here
            // This function is a placeholder and needs to be implemented based on your requirements
            return monthOfYear.ToString("D2"); // Example placeholder return
        }

    }
}