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
    private readonly string _tabName = "RG Tools";
    
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
        RibbonPanel panel;

        // Check if "RG Tools" already exists and use if its there
        try
        {
            panel = cachedUiCtrApp.CreateRibbonPanel(_tabName, Guid.NewGuid().ToString());
            panel.Name = "ARBG_ExpSched_ExtApp";
            panel.Title = "Export Schedule";
        }
        catch
        {
            cachedUiCtrApp.CreateRibbonTab(_tabName);
            panel = cachedUiCtrApp.CreateRibbonPanel(_tabName, Guid.NewGuid().ToString());
            panel.Name = "ARBG_ExpSched_ExtApp";
            panel.Title = "Export Schedule";
        }

        PushButtonData pbDataExpSched = new PushButtonData("Export Schedule", "Export Schedule", Assembly.GetExecutingAssembly().Location, "ExportSchedule.cmdExportSchedule");
        PushButton pbExpSched = (PushButton)panel.AddItem(pbDataExpSched);
        pbExpSched.ToolTip = "Export schedule to Microsoft Excel format";
        pbExpSched.LargeImage = PngImageSource("ExportSchedule.Images.ExpSched32.png");

        ContextualHelp contextHelp = new ContextualHelp(ContextualHelpType.Url, @"https://github.com/russgreen/ExportSchedule/wiki");

        pbExpSched.SetContextualHelp(contextHelp);

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
