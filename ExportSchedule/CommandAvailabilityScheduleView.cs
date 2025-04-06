using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ExportSchedule;
internal class CommandAvailabilityScheduleView : IExternalCommandAvailability
{
    public bool IsCommandAvailable(UIApplication applicationData, CategorySet selectedCategories)
    {
        if (applicationData.ActiveUIDocument != null &&
            applicationData.ActiveUIDocument.Document.ActiveView.ViewType == ViewType.Schedule)
        {
            return true;
        }

        return false;
    }
}
