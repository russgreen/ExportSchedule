﻿using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Diagnostics;
using System.IO;

namespace ExportSchedule;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class CommandExportSchedule : IExternalCommand
{
    private UIApplication _uiApp;
    private UIDocument _uiDoc;
    private Autodesk.Revit.ApplicationServices.Application _app;
    private Document _doc;

    private Autodesk.Revit.UI.Result _result = default(Autodesk.Revit.UI.Result);

    private bool _flag = false;

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        _uiApp = commandData.Application;
        _uiDoc = _uiApp.ActiveUIDocument;
        _app = _uiApp.Application;
        _doc = _uiDoc.Document;

        ViewSchedule activeView = commandData.Application.ActiveUIDocument.ActiveView as ViewSchedule;
        if (activeView is null)
        {
            TaskDialog.Show("Error", "Open and activate a schedule view to export and run the command again");
            return Autodesk.Revit.UI.Result.Failed;
        }

        var saveFileDialog = new Microsoft.Win32.SaveFileDialog()
        {
            Filter = "Microsoft Excel|*.xlsx",
            Title = "Export schedule to file",
            FileName = activeView.Name,
            InitialDirectory = Path.GetDirectoryName(_doc.PathName)
        };
        if (saveFileDialog.ShowDialog() == false || string.IsNullOrEmpty(saveFileDialog.FileName))
        {
            return Autodesk.Revit.UI.Result.Cancelled;
        }

        var directoryName = Path.GetDirectoryName(saveFileDialog.FileName);
        var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(saveFileDialog.FileName);
        var extension = Path.GetExtension(saveFileDialog.FileName);
        try
        {
            File.Delete(saveFileDialog.FileName);
            _flag = true;
        }
        catch
        {
            TaskDialog.Show("Error", "Can't access the export file.");
            _result = Autodesk.Revit.UI.Result.Failed;
        }

        if (!_flag)
        {
            return _result;
        }

        var xlfw = new ExportSchedule.XLSXFileWriter(saveFileDialog.FileName, fileNameWithoutExtension, activeView);
        xlfw.Export();
        _result = Autodesk.Revit.UI.Result.Succeeded;

        if (_result == Autodesk.Revit.UI.Result.Succeeded)
        {
            var td = new TaskDialog("Schedule Exported");
            td.MainContent = "Your schedule has been exported successfully";
            td.CommonButtons = TaskDialogCommonButtons.Cancel;
            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Open the exported schedule file with the default application.");
            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Open the folder containing the exported schedule file");
            TaskDialogResult taskDialogResult = td.Show();
            if (taskDialogResult ==  TaskDialogResult.CommandLink1)
            {
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = saveFileDialog.FileName,
                        UseShellExecute = true
                    });
                }
                catch
                {
                }
            }
            else if (taskDialogResult == TaskDialogResult.CommandLink2)
            {
                Process.Start("explorer.exe", "/root," + directoryName);

            }
            else if (taskDialogResult == Autodesk.Revit.UI.TaskDialogResult.Cancel)
            {
                // cancel clicked
            }

            return _result;
        }

        return Autodesk.Revit.UI.Result.Failed;
    }
}
