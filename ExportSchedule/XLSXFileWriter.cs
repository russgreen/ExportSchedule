using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Autodesk.Revit.DB;
using OfficeOpenXml;

namespace ExportSchedule;

public class XLSXFileWriter 
{

    private string _fileName;
    private string _filenameWithoutExtension;
    private ViewSchedule _viewSchedule;
    private int _currentRow = 1;

    private ExcelWorksheet _worksheet;

    private List<Tuple<int, int>> _writtenCells = new List<Tuple<int, int>>();

    private Color _black => new Color(0, 0, 0);
    private Color _white => new Color(255, 255, 255);

    public XLSXFileWriter(string fileName, string filenameWithoutExtension, ViewSchedule viewSchedule)
    {
        _viewSchedule = viewSchedule;
        _fileName = fileName;
        _filenameWithoutExtension = filenameWithoutExtension;
    }

    public void Export()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var oXLFile = new FileInfo(_fileName);
        if(oXLFile.Exists == true)
        {
            oXLFile.Delete();
        }

        using(var xlPackage = new ExcelPackage(oXLFile))
        {
            //add a worksheet
            _worksheet = xlPackage.Workbook.Worksheets.Add(_filenameWithoutExtension);

            //write data to the worksheet
            //WriteSection(SectionType.Header, _worksheet);
            WriteSection(SectionType.Body, _worksheet);

            //set the file properties
            xlPackage.Workbook.Properties.Title = _filenameWithoutExtension;
            xlPackage.Workbook.Properties.Author = System.Environment.UserName;
            xlPackage.Workbook.Properties.Subject = "Revit Schedule Export";

            // save the new spreadsheet
            xlPackage.Save();
        }
    }

    private void WriteSection(SectionType sectionType, ExcelWorksheet worksheet)
    {
        _writtenCells.Clear();
        TableSectionData sectionData = _viewSchedule.GetTableData().GetSectionData(sectionType);
        
        var numberOfRows = sectionData.NumberOfRows;
        var numberOfColumns = sectionData.NumberOfColumns;
        var firstRowNumber = sectionData.FirstRowNumber;
          
            do
            {
                WriteSectionRow(sectionType, _viewSchedule, sectionData, firstRowNumber, numberOfColumns, _worksheet);
                firstRowNumber += 1;
            }
            while (firstRowNumber < numberOfRows);

    }

    private void WriteSectionRow(SectionType sectionType, ViewSchedule schedule, TableSectionData sectionData, int iRow, int numberOfColumns, ExcelWorksheet worksheet)
    {
        var currentColumn = sectionData.FirstColumnNumber;
        decimal textSize = 10m;

        do
        {
            if (!_writtenCells.Contains(new Tuple<int, int>(iRow, currentColumn)))
            {
                Autodesk.Revit.DB.TableCellStyle tableCellStyle = sectionData.GetTableCellStyle(iRow, currentColumn);
                TableMergedCell mergedCell = sectionData.GetMergedCell(iRow, currentColumn);
                int top = mergedCell.Top;
                int fromRow = _currentRow;
                int fromCol = currentColumn + 1;
                int toRow = 0;
                int toCol = 0;

                // merge cells across columns
                if (sectionType == SectionType.Header)
                {
                    toRow = _currentRow;
                    //toCol = numberOfColumns - 1;
                    toCol = numberOfColumns;

                    Debug.WriteLine($"MergeRight = {mergedCell.Right.ToString()} - ColumnCount - 1 = {toCol.ToString()}");

                    // TODO: handle the bug when header row is not merged
                    if (mergedCell.Right > 0)
                    {
                        //worksheet.Cells(_currentRow, currentColumn + 1, toRow, toCol).Merge = true;
                        worksheet.Cells[_currentRow, currentColumn + 1, toRow, toCol].Merge = true;
                    }
                        
                }
                else if (sectionType == SectionType.Body)
                {
                    if (mergedCell.Left != mergedCell.Right || mergedCell.Top != mergedCell.Bottom)
                    {
                        // so we have some cells to merge
                        // oWorkSheet.Cells.CreateRange(m_CurrentRow, firstColumnNumber + 1, mergedCell.Bottom - mergedCell.Top, mergedCell.Right - mergedCell.Left).MergeCells()
                        // Dim oCell As ExcelCell = oWorkSheet.Cell(m_CurrentRow, m_CurrentColumn + 1)
                        // oWorkSheet.MergeCells(oCell, mergedCell.Bottom - mergedCell.Top, mergedCell.Right - mergedCell.Left)
                        toRow = fromRow + (mergedCell.Bottom - mergedCell.Top);
                        toCol = fromCol + (mergedCell.Right - mergedCell.Left);
                        worksheet.Cells[fromRow, fromCol, toRow, toCol].Merge = true;
                    }
                    else
                    {
                        // no need to merge the cell
                    }

                    // Remember all written cells related to the merge 
                    for (int iMergedRow = mergedCell.Top, loopTo = mergedCell.Bottom; iMergedRow <= loopTo; iMergedRow++)
                    {
                        for (int iMergedCol = mergedCell.Left, loopTo1 = mergedCell.Right; iMergedCol <= loopTo1; iMergedCol++)
                            _writtenCells.Add(new Tuple<int, int>(iMergedRow, iMergedCol));
                    }
                }

                do
                {
                    int left = mergedCell.Left;
                    do
                    {
                        _writtenCells.Add(new Tuple<int, int>(top, left));
                        left = left + 1;
                    }
                    while (left <= mergedCell.Right);
                    top = top + 1;
                }
                while (top <= mergedCell.Bottom);
                var cellText = schedule.GetCellText(sectionType, iRow, currentColumn);
                if (cellText.Length == 0)
                {
                    cellText = " ";
                }

                Debug.WriteLine(fromRow + ", " + fromCol + ", " + toRow + ", " + toCol + " - " + cellText);
                OfficeOpenXml.Style.ExcelHorizontalAlignment horAlignment;
                switch (tableCellStyle.FontHorizontalAlignment)
                {
                    case HorizontalAlignmentStyle.Left:
                    {
                        horAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        break;
                    }

                    case HorizontalAlignmentStyle.Right:
                    {
                        horAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        break;
                    }

                    case HorizontalAlignmentStyle.Center:
                    {
                        horAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        break;
                    }

                    default:
                    {
                        horAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.General;
                        break;
                    }
                }

                OfficeOpenXml.Style.ExcelVerticalAlignment vertAlignment;
                switch (tableCellStyle.FontVerticalAlignment)
                {
                    case VerticalAlignmentStyle.Top:
                    {
                        vertAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                        break;
                    }

                    case VerticalAlignmentStyle.Middle:
                    {
                        vertAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        break;
                    }

                    case VerticalAlignmentStyle.Bottom:
                    {
                        vertAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Bottom;
                        break;
                    }

                    default:
                    {
                        vertAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                        break;
                    }
                }

                try
                {
                    textSize = (decimal)tableCellStyle.TextSize;
                }
                catch 
                {
                    textSize = 10m;
                }

                {
                    var range = worksheet.Cells[_currentRow, currentColumn + 1];

                    if (!cellText.StartsWith("0") && double.TryParse(cellText, out var numberValue))
                    {
                        // format for numeric values
                        range.Value = numberValue;
                    }
                    else
                    {
                        // format for text values
                        range.Value = cellText;
                    }

                    range.Style.WrapText = true;
                    range.Style.Font.Size = (float)textSize;
                    range.Style.Font.Name = tableCellStyle.FontName;
                    range.Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(tableCellStyle.TextColor.Red, tableCellStyle.TextColor.Green, tableCellStyle.TextColor.Blue));
                    range.Style.Font.Bold = tableCellStyle.IsFontBold;
                    range.Style.Font.Italic = tableCellStyle.IsFontItalic;
                    range.Style.Font.UnderLine = tableCellStyle.IsFontUnderline;
                    range.Style.HorizontalAlignment = horAlignment;

                    range.Style.VerticalAlignment = vertAlignment;
                    if (!ColorsEqual(tableCellStyle.BackgroundColor, _white))
                    {
                        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(tableCellStyle.BackgroundColor.Red, tableCellStyle.BackgroundColor.Green, tableCellStyle.BackgroundColor.Blue));
                    }
                }

                worksheet.Column(currentColumn + 1).Width = sectionData.GetColumnWidth(currentColumn) * 1150 / 7.5d;
            }

            currentColumn += 1;
        }
        while (currentColumn < numberOfColumns);
        _currentRow += 1;
    }

    /// <summary>
    /// Compares two colors.
    /// </summary>
    /// <param name="color1">The first color.</param>
    /// <param name="color2">The second color.</param>
    /// <returns>True if the colors are equal, false otherwise.</returns>
    private bool ColorsEqual(Color color1, Color color2)
    {
        return color1.Red == color2.Red && color1.Green == color2.Green && color1.Blue == color2.Blue;
    }
}
