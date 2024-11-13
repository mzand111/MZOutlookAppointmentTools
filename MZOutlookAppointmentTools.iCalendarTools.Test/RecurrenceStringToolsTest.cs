namespace MZOutlookAppointmentTools.iCalendarTools.Test;

public partial class RecurrenceStringToolsTest : IDisposable
{
    private Microsoft.Office.Interop.Outlook.Application ApplicationInstance = null;
    public RecurrenceStringToolsTest()
    {
        ApplicationInstance = new Microsoft.Office.Interop.Outlook.Application();
    }

    public void Dispose()
    {
        if (ApplicationInstance != null)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ApplicationInstance);
    }
}