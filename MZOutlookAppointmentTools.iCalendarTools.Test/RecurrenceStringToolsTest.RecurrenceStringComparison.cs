namespace MZOutlookAppointmentTools.iCalendarTools.Test;

public partial class RecurrenceStringToolsTest
{

    [Fact]
    public void Compare_Simple1()
    {
        //Arrange
        string pattern1 = "FREQ=WEEKLY;BYDAY=SA";
        string pattern2 = "FREQ=WEEKLY;BYDAY=SA";

        //Act & Assert
        Assert.True(RecurrenceStringTools.AreEqual(pattern1, pattern2));
    }
    [Fact]
    public void Compare_Simple2()
    {
        //Arrange
        string pattern1 = "BYDAY=SA;FREQ=WEEKLY";
        string pattern2 = "FREQ=WEEKLY;BYDAY=SA";

        //Act & Assert
        Assert.True(RecurrenceStringTools.AreEqual(pattern1, pattern2));
    }
    [Fact]
    public void Compare_EmptyPartErrorTolerance()
    {
        //Arrange
        string pattern1 = ";BYDAY=SA;FREQ=WEEKLY";
        string pattern2 = "FREQ=WEEKLY;BYDAY=SA";

        //Act & Assert
        Assert.True(RecurrenceStringTools.AreEqual(pattern1, pattern2));
    }
    [Fact]
    public void Compare_Empty1()
    {
        //Arrange
        string pattern1 = "FREQ=WEEKLY;BYDAY=SA";
        string pattern2 = "";

        //Act & Assert
        Assert.False(RecurrenceStringTools.AreEqual(pattern1, pattern2));
    }
    [Fact]
    public void Compare_Empty2()
    {
        //Arrange
        string pattern1 = "";
        string pattern2 = "FREQ=WEEKLY;BYDAY=SA";

        //Act & Assert
        Assert.False(RecurrenceStringTools.AreEqual(pattern1, pattern2));
    }
    [Fact]
    public void Compare_WeekDayPosition1()
    {
        //Arrange
        string pattern1 = "FREQ=WEEKLY;BYDAY=MO,TU,WE";
        string pattern2 = "FREQ=WEEKLY;BYDAY=TU,WE,MO";

        //Act & Assert
        Assert.True(RecurrenceStringTools.AreEqual(pattern1, pattern2));
    }
    [Fact]
    public void Compare_WeekDayDifference()
    {
        //Arrange
        string pattern1 = "FREQ=WEEKLY;BYDAY=MO,TU,WE";
        string pattern2 = "FREQ=WEEKLY;BYDAY=TU,MO";

        //Act & Assert
        Assert.False(RecurrenceStringTools.AreEqual(pattern1, pattern2));
    }
    [Fact]
    public void Compare_FreqDifference()
    {
        //Arrange
        string pattern1 = "FREQ=DAILY;BYDAY=MO,TU";
        string pattern2 = "FREQ=WEEKLY;BYDAY=TU,MO";

        //Act & Assert
        Assert.False(RecurrenceStringTools.AreEqual(pattern1, pattern2));
    }
    [Fact]
    public void Compare_FreqSameWithError()
    {
        //Arrange
        string pattern1 = "FREQ=WEEKLY;;BYDAY=MO,TU";
        string pattern2 = "FREQ=WEEKLY;BYDAY=TU,MO;";

        //Act & Assert
        Assert.True(RecurrenceStringTools.AreEqual(pattern1, pattern2));
    }
    [Fact]
    public void Compare_IntervalDifference()
    {
        //Arrange
        string pattern1 = "FREQ=WEEKLY;BYDAY=MO,TU,WE;INTERVAL=2";
        string pattern2 = "FREQ=WEEKLY;BYDAY=MO,TU,WE;INTERVAL=1";

        //Act & Assert
        Assert.False(RecurrenceStringTools.AreEqual(pattern1, pattern2));
    }
    [Fact]
    public void Compare_IntervalIgnore()
    {
        //Arrange
        string pattern1 = "FREQ=YEARLY;BYMONTHDAY=7;BYMONTH=3";
        string pattern2 = "FREQ=YEARLY;BYMONTHDAY=7;BYMONTH=3;INTERVAL=1";

        //Act & Assert
        Assert.True(RecurrenceStringTools.AreEqual(pattern1, pattern2));
    }
    [Fact]
    public void Compare_ZeroBeforeNumbersIgnore1()
    {
        //Arrange
        string pattern1 = "FREQ=YEARLY;BYMONTHDAY=07;BYMONTH=03";
        string pattern2 = "FREQ=YEARLY;BYMONTHDAY=7;BYMONTH=3;";

        //Act & Assert
        Assert.True(RecurrenceStringTools.AreEqual(pattern1, pattern2));
    }
}
