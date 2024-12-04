using Microsoft.Office.Interop.Outlook;

namespace MZOutlookAppointmentTools.iCalendarTools.Test;

public partial class RecurrenceStringToolsTest
{
    [Fact]
    public void GetString_Empty1()
    {
        AppointmentItem aItem = (AppointmentItem)ApplicationInstance.CreateItem(OlItemType.olAppointmentItem);
        var item = RecurrenceStringTools.GetRecurrenceString(aItem);
        Assert.Equal(string.Empty, item);
    }
    [Fact]
    public void GetString_Freq()
    {
        AppointmentItem aItem = (AppointmentItem)ApplicationInstance.CreateItem(OlItemType.olAppointmentItem);
        var occ = aItem.GetRecurrencePattern();
        occ.RecurrenceType = OlRecurrenceType.olRecursDaily;
        occ.Interval = 1;
        aItem.Save();
        var item = RecurrenceStringTools.GetRecurrenceString(aItem);
        Assert.Equal("FREQ=DAILY;INTERVAL=1", item);
    }
    [Fact]
    public void GetString_MonthlyNoBySetPos1()
    {
        AppointmentItem aItem = (AppointmentItem)ApplicationInstance.CreateItem(OlItemType.olAppointmentItem);
        var occ = aItem.GetRecurrencePattern();
        occ.RecurrenceType = OlRecurrenceType.olRecursMonthNth;
        occ.Interval = 1;
        occ.Instance = 1;
        occ.DayOfWeekMask = OlDaysOfWeek.olFriday;
        aItem.Save();
        var item = RecurrenceStringTools.GetRecurrenceString(aItem);
        Assert.Equal("FREQ=MONTHLY;INTERVAL=1;BYDAY=FR;BYSETPOS=1", item);
    }
    [Fact]
    public void GetString_MonthlyNoBySetPos2()
    {
        AppointmentItem aItem = (AppointmentItem)ApplicationInstance.CreateItem(OlItemType.olAppointmentItem);
        var occ = aItem.GetRecurrencePattern();
        occ.RecurrenceType = OlRecurrenceType.olRecursMonthNth;
        occ.Interval = 1;
        occ.DayOfWeekMask = OlDaysOfWeek.olFriday;
        aItem.Save();
        var item = RecurrenceStringTools.GetRecurrenceString(aItem);
        Assert.Equal("FREQ=MONTHLY;INTERVAL=1;BYDAY=FR;BYSETPOS=1", item);
    }
}
