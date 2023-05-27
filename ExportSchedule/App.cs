using System;
using System.Reflection;
using System.Windows.Media.Imaging;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ExportSchedule;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
class App : IExternalApplication
{
    public static UIControlledApplication cachedUiCtrApp;
    
    public Result OnShutdown(UIControlledApplication application)
    {
        return Result.Succeeded;
    }

    public Result OnStartup(UIControlledApplication application)
    {
        cachedUiCtrApp = application;
        var ribbonPanel = CreateRibbonPanel();

        return Result.Succeeded;
    }

    private RibbonPanel CreateRibbonPanel()
    {
        RibbonPanel panel = cachedUiCtrApp.CreateRibbonPanel(nameof(ExportSchedule));
        panel.Title = "Export Schedule";

        PushButtonData pushButtonData = new PushButtonData(nameof(ExportSchedule), 
                       "Export Schedule", 
                                  Assembly.GetExecutingAssembly().Location, 
                                             $"{nameof(ExportSchedule)}.{nameof(ExportSchedule.CommandExportSchedule)}");
        pushButtonData.AvailabilityClassName = $"{nameof(ExportSchedule)}.{nameof(ExportSchedule.CommandAvailabilityScheduleView)}";

        PushButton pushButton = (PushButton)panel.AddItem(pushButtonData);
        pushButton.ToolTip = "Export schedule to Microsoft Excel format";
        pushButton.LargeImage = PngImageSource("ExportSchedule.Images.ExpSched32.png");


        ContextualHelp contextHelp = new ContextualHelp(ContextualHelpType.Url, @"https://github.com/russgreen/ExportSchedule/wiki");

        pushButton.SetContextualHelp(contextHelp);

        return panel;
    }

    private System.Windows.Media.ImageSource PngImageSource(string embeddedPath)
    {
        var stream = GetType().Assembly.GetManifestResourceStream(embeddedPath);
        System.Windows.Media.ImageSource imageSource;
        try
        {
            imageSource = BitmapFrame.Create(stream);
        }
        catch
        {
            imageSource = null;
        }

        return imageSource;
    }
}
