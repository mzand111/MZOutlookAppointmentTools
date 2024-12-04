using Microsoft.Office.Interop.Outlook;

namespace MZOutlookAppointmentTools.iCalendarTools.Test;

public partial class RecurrenceStringToolsTest
{
    #region Daily
    [Fact]
    public void DAILY1()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=DAILY";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        var recPattern = aItem.GetRecurrencePattern();
        //Assert
        Assert.Equal(OlRecurrenceType.olRecursDaily, recPattern.RecurrenceType);
        Assert.Equal((OlDaysOfWeek)0, recPattern.DayOfWeekMask);
        Assert.Equal(0, recPattern.Instance);
    }
    [Fact]
    public void DAILY2()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=DAILY;INTERVAL=2";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        var recPattern = aItem.GetRecurrencePattern();
        //Assert
        Assert.Equal(OlRecurrenceType.olRecursDaily, recPattern.RecurrenceType);
        Assert.Equal((OlDaysOfWeek)0, recPattern.DayOfWeekMask);
        Assert.Equal(0, recPattern.Instance);
        Assert.Equal(2, recPattern.Interval);
    }
    [Fact]
    public void DAILY_ByWeekDay()
    {
        //Arrange

        AppointmentItem aItem = (AppointmentItem)ApplicationInstance.CreateItem(OlItemType.olAppointmentItem);
        string pattern = "FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        var recPattern = aItem.GetRecurrencePattern();
        //Assert
        Assert.Equal(OlRecurrenceType.olRecursWeekly, recPattern.RecurrenceType);
        var expectedMask = OlDaysOfWeek.olMonday | OlDaysOfWeek.olTuesday | OlDaysOfWeek.olWednesday | OlDaysOfWeek.olThursday | OlDaysOfWeek.olFriday;
        Assert.Equal(expectedMask, recPattern.DayOfWeekMask);
        Assert.Equal(0, recPattern.Instance);
        Assert.Equal(1, recPattern.Interval);
    }
    #endregion
    #region Weekly
    [Fact]
    public void Weekly1()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=WEEKLY;BYDAY=SA";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        var recPattern = aItem.GetRecurrencePattern();
        //Assert
        Assert.Equal(OlRecurrenceType.olRecursWeekly, recPattern.RecurrenceType);
        Assert.Equal(OlDaysOfWeek.olSaturday, recPattern.DayOfWeekMask);
        Assert.Equal(0, recPattern.Instance);
    }
    [Fact]
    public void Weekly2()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=WEEKLY;BYDAY=MO,TU,WE";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        var recPattern = aItem.GetRecurrencePattern();
        //Assert
        Assert.Equal(OlRecurrenceType.olRecursWeekly, recPattern.RecurrenceType);
        var expectedMask = OlDaysOfWeek.olMonday | OlDaysOfWeek.olTuesday | OlDaysOfWeek.olWednesday;
        Assert.Equal(expectedMask, recPattern.DayOfWeekMask);
        Assert.Equal(0, recPattern.Instance);
        Assert.Equal(1, recPattern.Interval);
    }
    [Fact]
    public void Weekly3()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=WEEKLY;BYDAY=MO,TU,WE;INTERVAL=2";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        var recPattern = aItem.GetRecurrencePattern();
        //Assert
        Assert.Equal(OlRecurrenceType.olRecursWeekly, recPattern.RecurrenceType);
        var expectedMask = OlDaysOfWeek.olMonday | OlDaysOfWeek.olTuesday | OlDaysOfWeek.olWednesday;
        Assert.Equal(expectedMask, recPattern.DayOfWeekMask);
        Assert.Equal(2, recPattern.Interval);
        Assert.Equal(0, recPattern.Instance);
    }
    /// <summary>
    /// WSK is not supported by Outlook, but the method should not break
    /// </summary>
    [Fact]
    public void WeeklyHandlingNotSupportedWKST()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=WEEKLY;BYDAY=SA;WKST=SU";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        var recPattern = aItem.GetRecurrencePattern();
        //Assert
        Assert.Equal(OlRecurrenceType.olRecursWeekly, recPattern.RecurrenceType);
        Assert.Equal(OlDaysOfWeek.olSaturday, recPattern.DayOfWeekMask);
        Assert.Equal(1, recPattern.Interval);
    }
    #endregion

    #region Monthly
    [Fact]
    public void Monthly1()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=MONTHLY;BYMONTHDAY=12";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        var recPattern = aItem.GetRecurrencePattern();
        //Assert
        Assert.Equal(OlRecurrenceType.olRecursMonthly, recPattern.RecurrenceType);
        Assert.Equal(12, recPattern.DayOfMonth);
        Assert.Equal(0, recPattern.Instance);
        Assert.Equal(1, recPattern.Interval);
    }

    #endregion

    #region MonthNth
    [Fact]
    public void MonthlyNth1()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=MONTHLY;BYDAY=MO,TU,WE,TH,FR;BYSETPOS=1";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        var recPattern = aItem.GetRecurrencePattern();
        //Assert

        Assert.Equal(OlRecurrenceType.olRecursMonthNth, recPattern.RecurrenceType);
        var expectedMask = OlDaysOfWeek.olMonday | OlDaysOfWeek.olTuesday | OlDaysOfWeek.olWednesday | OlDaysOfWeek.olThursday | OlDaysOfWeek.olFriday;
        Assert.Equal(0, recPattern.DayOfMonth);
        Assert.Equal(expectedMask, recPattern.DayOfWeekMask);
        Assert.Equal(1, recPattern.Instance);
        Assert.Equal(1, recPattern.Interval);
    }
    [Fact]
    public void MonthlyNth2()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=MONTHLY;BYDAY=MO,TU,WE,TH,FR;BYSETPOS=-1";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        var recPattern = aItem.GetRecurrencePattern();
        //Assert

        Assert.Equal(OlRecurrenceType.olRecursMonthNth, recPattern.RecurrenceType);
        var expectedMask = OlDaysOfWeek.olMonday | OlDaysOfWeek.olTuesday | OlDaysOfWeek.olWednesday | OlDaysOfWeek.olThursday | OlDaysOfWeek.olFriday;
        Assert.Equal(0, recPattern.DayOfMonth);
        Assert.Equal(expectedMask, recPattern.DayOfWeekMask);
        Assert.Equal(5, recPattern.Instance);
        Assert.Equal(1, recPattern.Interval);
    }
    [Fact]
    public void MonthlyNth3_WhenBySetPosIsNotSet()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=MONTHLY;INTERVAL=1;BYDAY=FR";
        var itemStart = new DateTime(2024, 11, 01, 11, 30, 0);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        var recPattern = aItem.GetRecurrencePattern();
        //Assert

        Assert.Equal(OlRecurrenceType.olRecursWeekly, recPattern.RecurrenceType);
        var expectedMask = OlDaysOfWeek.olFriday;
        Assert.Equal(0, recPattern.DayOfMonth);
        Assert.Equal(expectedMask, recPattern.DayOfWeekMask);
        Assert.Equal(0, recPattern.Instance);
        Assert.Equal(1, recPattern.Interval);
    }
    #endregion

    #region MonthNth
    [Fact]
    public void Yearly1()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=YEARLY;BYMONTHDAY=7;BYMONTH=3";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        var recPattern = aItem.GetRecurrencePattern();
        //Assert

        Assert.Equal(OlRecurrenceType.olRecursYearly, recPattern.RecurrenceType);
        Assert.Equal(7, recPattern.DayOfMonth);
        Assert.Equal(3, recPattern.MonthOfYear);
        Assert.Equal((OlDaysOfWeek)0, recPattern.DayOfWeekMask);
        Assert.Equal(0, recPattern.Instance);
        Assert.Equal(12, recPattern.Interval);
    }
    [Fact]
    public void Yearly2()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=YEARLY;BYMONTHDAY=14;BYMONTH=2;INTERVAL=2";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        var recPattern = aItem.GetRecurrencePattern();
        //Assert

        Assert.Equal(OlRecurrenceType.olRecursYearly, recPattern.RecurrenceType);
        Assert.Equal(14, recPattern.DayOfMonth);
        Assert.Equal(2, recPattern.MonthOfYear);
        Assert.Equal((OlDaysOfWeek)0, recPattern.DayOfWeekMask);
        Assert.Equal(0, recPattern.Instance);
        Assert.Equal(24, recPattern.Interval);
    }
    #endregion

    #region YearNth
    [Fact]
    public void YearNth1()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=YEARLY;BYDAY=FR;BYMONTH=11;BYSETPOS=2";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        var recPattern = aItem.GetRecurrencePattern();
        //Assert

        Assert.Equal(OlRecurrenceType.olRecursYearNth, recPattern.RecurrenceType);
        Assert.Equal(0, recPattern.DayOfMonth);
        Assert.Equal(11, recPattern.MonthOfYear);
        Assert.Equal(OlDaysOfWeek.olFriday, recPattern.DayOfWeekMask);
        Assert.Equal(2, recPattern.Instance);
        Assert.Equal(12, recPattern.Interval);// Interval *12
    }
    [Fact]
    public void YearNth2()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=YEARLY;BYDAY=FR;BYMONTH=11;BYSETPOS=3";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        var recPattern = aItem.GetRecurrencePattern();
        //Assert

        Assert.Equal(OlRecurrenceType.olRecursYearNth, recPattern.RecurrenceType);
        Assert.Equal(0, recPattern.DayOfMonth);
        Assert.Equal(11, recPattern.MonthOfYear);
        Assert.Equal(OlDaysOfWeek.olFriday, recPattern.DayOfWeekMask);
        Assert.Equal(3, recPattern.Instance);
        Assert.Equal(12, recPattern.Interval);// Interval *12
    }
    [Fact]
    public void YearNth3()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=YEARLY;BYDAY=FR;BYMONTH=11;BYSETPOS=3;INTERVAL=3";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        var recPattern = aItem.GetRecurrencePattern();
        //Assert

        Assert.Equal(OlRecurrenceType.olRecursYearNth, recPattern.RecurrenceType);
        Assert.Equal(0, recPattern.DayOfMonth);
        Assert.Equal(11, recPattern.MonthOfYear);
        Assert.Equal(OlDaysOfWeek.olFriday, recPattern.DayOfWeekMask);
        Assert.Equal(3, recPattern.Instance);
        Assert.Equal(36, recPattern.Interval);// Interval *12
    }
    [Fact]
    public void YearNth4()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=YEARLY;BYDAY=MO,TU,WE,TH,FR;BYMONTH=9;BYSETPOS=1";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        var recPattern = aItem.GetRecurrencePattern();
        //Assert

        Assert.Equal(OlRecurrenceType.olRecursYearNth, recPattern.RecurrenceType);
        Assert.Equal(0, recPattern.DayOfMonth);
        Assert.Equal(9, recPattern.MonthOfYear);
        Assert.Equal(OlDaysOfWeek.olFriday | OlDaysOfWeek.olMonday | OlDaysOfWeek.olTuesday | OlDaysOfWeek.olWednesday | OlDaysOfWeek.olThursday, recPattern.DayOfWeekMask);
        Assert.Equal(1, recPattern.Instance);
        Assert.Equal(12, recPattern.Interval);// Interval *12
    }
    #endregion
}
