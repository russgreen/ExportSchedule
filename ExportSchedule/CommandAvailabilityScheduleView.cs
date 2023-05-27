using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
