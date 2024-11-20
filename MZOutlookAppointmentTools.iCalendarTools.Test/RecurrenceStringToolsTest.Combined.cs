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
        var recPattern = RecurrenceStringTools.ParseRecurrencePattern(pattern, aItem, itemStart);
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
        var recPattern = RecurrenceStringTools.ParseRecurrencePattern(pattern, aItem, itemStart);
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
        var recPattern = RecurrenceStringTools.ParseRecurrencePattern(pattern, aItem, itemStart);
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
        var recPattern = RecurrenceStringTools.ParseRecurrencePattern(pattern, aItem, itemStart);
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
        var recPattern = RecurrenceStringTools.ParseRecurrencePattern(pattern, aItem, itemStart);
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
        var recPattern = RecurrenceStringTools.ParseRecurrencePattern(pattern, aItem, itemStart);
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
        var recPattern = RecurrenceStringTools.ParseRecurrencePattern(pattern, aItem, itemStart);
        var recPatternStr = RecurrenceStringTools.GetRecurrenceString(aItem);
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
        var recPattern = RecurrenceStringTools.ParseRecurrencePattern(pattern, aItem, itemStart);
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
        var recPattern = RecurrenceStringTools.ParseRecurrencePattern(pattern, aItem, itemStart);
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
        var recPattern = RecurrenceStringTools.ParseRecurrencePattern(pattern, aItem, itemStart);
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
        var recPattern = RecurrenceStringTools.ParseRecurrencePattern(pattern, aItem, itemStart);
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
        var recPattern = RecurrenceStringTools.ParseRecurrencePattern(pattern, aItem, itemStart);
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
        var recPattern = RecurrenceStringTools.ParseRecurrencePattern(pattern, aItem, itemStart);
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
        var recPattern = RecurrenceStringTools.ParseRecurrencePattern(pattern, aItem, itemStart);
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
        var recPattern = RecurrenceStringTools.ParseRecurrencePattern(pattern, aItem, itemStart);
        var recPatternStr = RecurrenceStringTools.GetRecurrenceString(aItem);
        //Assert
        var gg = RecurrenceStringTools.AreEqual(pattern, recPatternStr);
        Assert.True(gg);
    }
    #endregion
}
