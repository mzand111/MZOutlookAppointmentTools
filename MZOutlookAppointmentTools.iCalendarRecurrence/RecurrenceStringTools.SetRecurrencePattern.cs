using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;

namespace MZOutlookAppointmentTools.iCalendarTools
{
    public partial class RecurrenceStringTools
    {
        /// <summary>
        /// Parses a recurrence pattern string and applies it to an AppointmentItem.
        /// </summary>
        /// <param name="recurrenceString">The recurrence pattern string.</param>
        /// <param name="appointmentItem">The appointment item to apply the recurrence pattern to.</param>
        /// <param name="start">The start date of the recurrence pattern.</param>
        /// <returns>The parsed RecurrencePattern object.</returns>
        /// <exception cref="InvalidOperationException"></exception>
        public static void SetRecurrencePattern(string recurrenceString, AppointmentItem appointmentItem, DateTime? start)
        {
            if (string.IsNullOrEmpty(recurrenceString))
            {
                appointmentItem.ClearRecurrencePattern();
            }
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

            // Split the recurrence string into key-value pairs
            string[] parts = recurrenceString.Split(';');
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
                                    bd_set = false;
                                }
                                else
                                {
                                    pattern.Instance = bsp;
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
                        else
                        //If BYSETPOS is not set and BYDAY is set, recurrence is actually a simply weekly :olRecursWeekly
                        {
                            if (bd_set)
                            {
                                pattern.RecurrenceType = OlRecurrenceType.olRecursWeekly;
                                pattern.DayOfWeekMask = bd;
                            }
                            else if (ruleBook.ContainsKey("BYMONTHDAY"))
                            {
                                pattern.RecurrenceType = OlRecurrenceType.olRecursMonthly;
                                pattern.DayOfMonth = Int16.Parse(ruleBook["BYMONTHDAY"]);
                            }
                            else
                            {
                                throw new System.Exception($"Not supported recurrence string: '{recurrenceString}' Expecting BYDAY");
                            }

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
                        if (bd_set)
                            pattern.DayOfWeekMask = bd;
                        if (bsp_set)
                        {
                            int gInstance = Convert.ToInt16(ruleBook["BYSETPOS"]);
                            pattern.Instance = (gInstance == -1) ? 5 : gInstance;
                        }
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
                else
                {
                    pattern.NoEndDate = true;
                }

            }
            finally
            {
                if (pattern != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(pattern);
            }
        }
        /// <summary>
        /// Converts recurrence days to Outlook OlDaysOfWeek mask
        /// </summary>
        /// <param name="byDay"></param>
        /// <returns></returns>
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
    }
}
