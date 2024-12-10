using Microsoft.Office.Interop.Outlook;

namespace MZOutlookAppointmentTools.iCalendarTools.Test;

public partial class RecurrenceStringToolsTest
{
    #region Weekly
    [Fact]
    public void Combo_Weekly1()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=WEEKLY;BYDAY=SA";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        aItem.Save();
        var recPatternStr = RecurrenceStringTools.GetRecurrenceString(aItem);
        //Assert

        var gg = RecurrenceStringTools.AreEqual(pattern, recPatternStr);
        Assert.True(gg);
    }
    [Fact]
    public void Combo_Weekly1_1()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=WEEKLY;BYDAY=SA";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        aItem.Save();
        var recPatternStr = RecurrenceStringTools.GetRecurrenceString(aItem);
        //Assert

        var gg = RecurrenceStringTools.AreEqual("FREQ=WEEKLY;BYDAY=SA;INTERVAL=2", recPatternStr);
        Assert.False(gg);
    }
    [Fact]
    public void Combo_Weekly1_2()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=WEEKLY;BYDAY=SA";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        aItem.Save();
        var recPatternStr = RecurrenceStringTools.GetRecurrenceString(aItem);
        //Assert

        var gg = RecurrenceStringTools.AreEqual("FREQ=WEEKLY;BYDAY=SA;INTERVAL=1", recPatternStr);
        Assert.True(gg);
    }
    [Fact]
    public void Combo_Weekly2()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=WEEKLY;BYDAY=MO,TU,WE";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        aItem.Save();
        var recPatternStr = RecurrenceStringTools.GetRecurrenceString(aItem);
        //Assert
        var gg = RecurrenceStringTools.AreEqual(pattern, recPatternStr);
        Assert.True(gg);
    }
    [Fact]
    public void Combo_Weekly3()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=WEEKLY;BYDAY=MO,TU,WE;INTERVAL=2";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        aItem.Save();
        var recPatternStr = RecurrenceStringTools.GetRecurrenceString(aItem);
        //Assert
        var gg = RecurrenceStringTools.AreEqual(pattern, recPatternStr);
        Assert.True(gg);
    }
    /// <summary>
    /// WSK is not supported by Outlook, but the method should not break
    /// </summary>
    [Fact]
    public void Combo_WeeklyHandlingNotSupportedWKST()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=WEEKLY;BYDAY=SA;WKST=SU";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        aItem.Save();
        var recPatternStr = RecurrenceStringTools.GetRecurrenceString(aItem);
        //Assert
        var gg = RecurrenceStringTools.AreEqual(pattern, recPatternStr);
        Assert.True(gg);
    }
    #endregion

    #region Monthly
    [Fact]
    public void Combo_Monthly1()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=MONTHLY;BYMONTHDAY=12";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        aItem.Save();
        var recPatternStr = RecurrenceStringTools.GetRecurrenceString(aItem);
        var hh = aItem.GetRecurrencePattern();

        //Assert
        var gg = RecurrenceStringTools.AreEqual(pattern, recPatternStr);
        Assert.True(gg);
    }

    #endregion

    #region MonthNth
    [Fact]
    public void Combo_MonthlyNth1()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=MONTHLY;BYDAY=MO,TU,WE,TH,FR;BYSETPOS=1";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        aItem.Save();
        var recPatternStr = RecurrenceStringTools.GetRecurrenceString(aItem);
        //Assert
        var gg = RecurrenceStringTools.AreEqual(pattern, recPatternStr);
        Assert.True(gg);
    }
    [Fact]
    public void Combo_MonthlyNth2()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=MONTHLY;BYDAY=MO,TU,WE,TH,FR;BYSETPOS=-1";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        aItem.Save();
        var recPatternStr = RecurrenceStringTools.GetRecurrenceString(aItem);
        //Assert        
        var gg = RecurrenceStringTools.AreEqual(pattern, recPatternStr);
        Assert.True(gg);
    }
    [Fact]
    public void Combo_MonthlyNthSetPos_Existence1()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=MONTHLY;INTERVAL=1;BYDAY=FR";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        aItem.Save();
        var recPatternStr = RecurrenceStringTools.GetRecurrenceString(aItem);
        //Assert        
        var gg = RecurrenceStringTools.AreEqual("FREQ=WEEKLY;INTERVAL=1;BYDAY=FR", recPatternStr);
        Assert.True(gg);
    }
    [Fact]
    public void Combo_MonthlyNthSetPos_Existence2()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=MONTHLY;INTERVAL=1;BYDAY=FR;BYSETPOS=1";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        aItem.Save();
        var recPatternStr = RecurrenceStringTools.GetRecurrenceString(aItem);
        //Assert        
        var gg = RecurrenceStringTools.AreEqual(pattern, recPatternStr);
        Assert.True(gg);
    }
    #endregion

    #region MonthNth
    [Fact]
    public void Combo_Yearly1()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=YEARLY;BYMONTHDAY=7;BYMONTH=3";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        aItem.Save();
        var recPatternStr = RecurrenceStringTools.GetRecurrenceString(aItem);
        //Assert      
        var gg = RecurrenceStringTools.AreEqual(pattern, recPatternStr);
        //Assert.Equal(pattern, recPatternStr);
        Assert.True(gg);
    }
    [Fact]
    public void Combo_Yearly2()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=YEARLY;BYMONTHDAY=14;BYMONTH=2;INTERVAL=2";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        aItem.Save();
        var recPatternStr = RecurrenceStringTools.GetRecurrenceString(aItem);
        //Assert
        var gg = RecurrenceStringTools.AreEqual(pattern, recPatternStr);
        //Assert.Equal(pattern, recPatternStr);
        Assert.True(gg);
    }
    #endregion

    #region YearNth
    [Fact]
    public void Combo_YearNth1()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=YEARLY;BYDAY=FR;BYMONTH=11;BYSETPOS=2";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        aItem.Save();
        var recPatternStr = RecurrenceStringTools.GetRecurrenceString(aItem);
        //Assert
        var gg = RecurrenceStringTools.AreEqual(pattern, recPatternStr);
        Assert.True(gg);
    }
    [Fact]
    public void Combo_YearNth2()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=YEARLY;BYDAY=FR;BYMONTH=11;BYSETPOS=3";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        aItem.Save();
        var recPatternStr = RecurrenceStringTools.GetRecurrenceString(aItem);
        //Assert
        var gg = RecurrenceStringTools.AreEqual(pattern, recPatternStr);
        Assert.True(gg);
    }
    [Fact]
    public void Combo_YearNth3()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=YEARLY;BYDAY=FR;BYMONTH=11;BYSETPOS=3;INTERVAL=3";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        aItem.Save();
        var recPatternStr = RecurrenceStringTools.GetRecurrenceString(aItem);
        //Assert
        var gg = RecurrenceStringTools.AreEqual(pattern, recPatternStr);
        Assert.True(gg);
    }
    [Fact]
    public void Combo_YearNth4()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=YEARLY;BYDAY=MO,TU,WE,TH,FR;BYMONTH=9;BYSETPOS=1";
        var itemStart = new DateTime(2024, 11, 13);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        aItem.Save();
        var recPatternStr = RecurrenceStringTools.GetRecurrenceString(aItem);
        //Assert
        var gg = RecurrenceStringTools.AreEqual(pattern, recPatternStr);
        Assert.True(gg);
    }
    #endregion

    [Fact]
    public void Combo_Daily1()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=Daily";
        var itemStart = new DateTime(2024, 11, 27, 8, 15, 0);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        aItem.Save();
        var recPatternStr = RecurrenceStringTools.GetRecurrenceString(aItem);
        //Assert
        var gg = RecurrenceStringTools.AreEqual(pattern, recPatternStr);
        Assert.True(gg);
    }
    [Fact]
    public void Combo_DailyExceptions()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=Daily";
        var itemStart = new DateTime(2024, 11, 27, 20, 15, 0);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        aItem.Save();
        var recPatternStr = RecurrenceStringTools.GetRecurrenceString(aItem);
        //Assert
        var patternObj = aItem.GetRecurrencePattern();
        var occ1 = patternObj.GetOccurrence(new DateTime(2024, 11, 27, 20, 15, 0));
        var occ2 = patternObj.GetOccurrence(new DateTime(2024, 11, 28, 20, 15, 0));
        var gg = RecurrenceStringTools.AreEqual(pattern, recPatternStr);
        Assert.True(gg);
    }
    [Fact]
    public void Combo_MonthBySetPos()
    {
        //Arrange

        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        string pattern = "FREQ=MONTHLY;BYDAY=FR;BYSETPOS=1";
        var itemStart = new DateTime(2024, 11, 01, 11, 30, 0);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(30);
        //Act
        RecurrenceStringTools.SetRecurrencePattern(pattern, aItem, itemStart);
        aItem.Save();
        var recPatternStr = RecurrenceStringTools.GetRecurrenceString(aItem);
        //Assert
        var patternObj = aItem.GetRecurrencePattern();

        var gg = RecurrenceStringTools.AreEqual(pattern, recPatternStr);

        Assert.True(gg, "recPatternStr");
    }

    [Fact]
    public void Combo_EndTimeTest()
    {
        //Arrange
        AppointmentItem aItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)ApplicationInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
        var itemStart = new DateTime(2023, 11, 01, 11, 30, 0);
        aItem.Start = itemStart;
        aItem.End = itemStart.AddMinutes(15);
        var occ = aItem.GetRecurrencePattern();
        occ.RecurrenceType = OlRecurrenceType.olRecursMonthNth;

        occ.Interval = 1;
        occ.Instance = 1;
        var expectedMask = OlDaysOfWeek.olMonday | OlDaysOfWeek.olTuesday | OlDaysOfWeek.olWednesday | OlDaysOfWeek.olThursday | OlDaysOfWeek.olFriday;
        occ.DayOfWeekMask = expectedMask;
        occ.PatternEndDate = new DateTime(2024, 11, 1);
        aItem.Save();
        var recPat = RecurrenceStringTools.GetRecurrenceString(aItem);
        if (occ != null)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(occ);
        Assert.Equal("FREQ=MONTHLY;UNTIL=20241101T000000Z;INTERVAL=1;BYDAY=MO,TU,WE,TH,FR;BYSETPOS=1", recPat);
        var newPatternWithNoEnd = "FREQ=MONTHLY;INTERVAL=1;BYDAY=MO,TU,WE,TH,FR;BYSETPOS=1";
        RecurrenceStringTools.SetRecurrencePattern(newPatternWithNoEnd, aItem, itemStart);
        var occ2 = aItem.GetRecurrencePattern();
        var et = occ2.PatternEndDate;
        var eto = occ2.NoEndDate;
        Assert.True(eto);
        var newRec = RecurrenceStringTools.GetRecurrenceString(aItem);
        Assert.Equal(newPatternWithNoEnd, newRec);
        var hh = "";
    }
}
