﻿using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;

namespace MZOutlookAppointmentTools.iCalendarTools
{
    public partial class RecurrenceStringTools
    {
        /// <summary>
        /// Generates an iCalendar (iCal) formatted recurrence string for a given AppointmentItem.
        /// </summary>
        /// <param name="myItem">The appointment item to generate the recurrence string for.</param>
        /// <returns>A formatted recurrence string if the item is recurring; otherwise, an empty string.</returns>
        public static string GetRecurrenceString(AppointmentItem myItem)
        {
            // Returns a properly formatted recurrence string for the item.
            if (!myItem.IsRecurring)
                return string.Empty;
            RecurrencePattern pattern = myItem.GetRecurrencePattern();
            string str = "";
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
                        //if (pattern.Instance == 5)
                        //{
                        //    str += ";BYSETPOS=" + pattern.Instance;
                        //}
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
                            str += ";BYSETPOS=-1";
                            str += ";BYDAY=" + DaysOfWeek("", pattern);
                        }
                        else
                        {
                            if (pattern.Instance > 0)
                            {
                                str += ";BYDAY=" + DaysOfWeek(WeekNum(pattern.Instance), pattern);
                                str += ";BYSETPOS=" + pattern.Instance.ToString();
                            }
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
                        str += ";INTERVAL=" + YearlyIntervalNumber(pattern.Interval);
                        var daysOfWeek = DaysOfWeek("", pattern);
                        if (!string.IsNullOrWhiteSpace(daysOfWeek))
                        {
                            str += ";BYDAY=" + DaysOfWeek("", pattern);
                        }
                        if (pattern.MonthOfYear != 0)
                        {
                            str += ";BYMONTH=" + MonthNum(pattern.MonthOfYear);
                        }
                        if (pattern.DayOfMonth > 0)
                        {
                            str += ";BYMONTHDAY=" + DayOfMonth(pattern.DayOfMonth);
                        }

                        break;

                    case OlRecurrenceType.olRecursYearNth:
                        str += "FREQ=YEARLY";
                        if (!pattern.NoEndDate)
                        {
                            str += ";UNTIL=" + FormatICalDateTime(pattern.PatternEndDate);
                        }
                        str += ";BYMONTH=" + MonthNum(pattern.MonthOfYear);
                        str += ";BYDAY=" + DaysOfWeek(WeekNum(pattern.Instance), pattern);
                        if (pattern.Instance == 5)
                        {
                            str += ";BYSETPOS=-1";
                        }
                        else
                        {
                            if (pattern.Instance > 0)
                            {
                                str += ";BYSETPOS=" + pattern.Instance.ToString();
                            }

                        }
                        str += ";INTERVAL=" + YearlyIntervalNumber(pattern.Interval);
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
        private static string DayOfMonth(int dayOfMonth)
        {
            return dayOfMonth.ToString("D2");
        }
        private static string YearlyIntervalNumber(int yearlyIntervalNumber)
        {
            return (yearlyIntervalNumber / 12).ToString("D2");
        }
    }
}
