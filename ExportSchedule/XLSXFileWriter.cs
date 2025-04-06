using Autodesk.Revit.DB;
using ClosedXML.Excel;
using System.Diagnostics;
using System.IO;

namespace ExportSchedule;

public class XLSXFileWriter 
{

    private string _fileName;
    private string _filenameWithoutExtension;
    private ViewSchedule _viewSchedule;
    private int _currentRow = 1;

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
        var oXLFile = new FileInfo(_fileName);
        if (oXLFile.Exists == true)
        {
            oXLFile.Delete();
        }

        using (var workbook = new XLWorkbook())
        {
            //add a worksheet
            var worksheet = workbook.AddWorksheet(_filenameWithoutExtension);

            //write data to the worksheet
            WriteSection(SectionType.Body, worksheet);

            //set the file properties
            workbook.Properties.Title = _filenameWithoutExtension;
            workbook.Properties.Author = System.Environment.UserName;
            workbook.Properties.Subject = "Revit Schedule Export";

            // save the new spreadsheet
            workbook.SaveAs(_fileName);
        }
    }

    private void WriteSection(SectionType sectionType, IXLWorksheet worksheet)
    {
        _writtenCells.Clear();
        TableSectionData sectionData = _viewSchedule.GetTableData().GetSectionData(sectionType);
        
        var numberOfRows = sectionData.NumberOfRows;
        var numberOfColumns = sectionData.NumberOfColumns;
        var firstRowNumber = sectionData.FirstRowNumber;
          
            do
            {
                WriteSectionRow(sectionType, _viewSchedule, sectionData, firstRowNumber, numberOfColumns, worksheet);
                firstRowNumber += 1;
            }
            while (firstRowNumber < numberOfRows);
    }

    private void WriteSectionRow(SectionType sectionType, ViewSchedule schedule, TableSectionData sectionData, int iRow, int numberOfColumns, IXLWorksheet worksheet)
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
                        worksheet.Range(_currentRow, currentColumn + 1, toRow, toCol).Merge();
                    }
                        
                }
                else if (sectionType == SectionType.Body)
                {
                    if (mergedCell.Left != mergedCell.Right || mergedCell.Top != mergedCell.Bottom)
                    {
                        // so we have some cells to merge
                        toRow = fromRow + (mergedCell.Bottom - mergedCell.Top);
                        toCol = fromCol + (mergedCell.Right - mergedCell.Left);
                        worksheet.Range(fromRow, fromCol, toRow, toCol).Merge();
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


                XLAlignmentHorizontalValues horAlignment;
                switch (tableCellStyle.FontHorizontalAlignment)
                {
                    case HorizontalAlignmentStyle.Left:
                    {
                        horAlignment = XLAlignmentHorizontalValues.Left;
                        break;
                    }

                    case HorizontalAlignmentStyle.Right:
                    {
                        horAlignment = XLAlignmentHorizontalValues.Right;
                        break;
                    }

                    case HorizontalAlignmentStyle.Center:
                    {
                        horAlignment = XLAlignmentHorizontalValues.Center;
                        break;
                    }

                    default:
                    {
                        horAlignment = XLAlignmentHorizontalValues.General;
                        break;
                    }
                }

                XLAlignmentVerticalValues vertAlignment;
                switch (tableCellStyle.FontVerticalAlignment)
                {
                    case VerticalAlignmentStyle.Top:
                    {
                        vertAlignment = XLAlignmentVerticalValues.Top;
                        break;
                    }

                    case VerticalAlignmentStyle.Middle:
                    {
                        vertAlignment = XLAlignmentVerticalValues.Center;
                        break;
                    }

                    case VerticalAlignmentStyle.Bottom:
                    {
                        vertAlignment = XLAlignmentVerticalValues.Bottom;
                        break;
                    }

                    default:
                    {
                        vertAlignment = XLAlignmentVerticalValues.Top;
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
                    var range = worksheet.Cell(_currentRow, currentColumn + 1);

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

                    range.Style.Alignment.WrapText = true;
                    range.Style.Font.FontSize = (float)textSize;
                    range.Style.Font.FontName = tableCellStyle.FontName;
                    range.Style.Font.SetFontColor(XLColor.FromArgb(tableCellStyle.TextColor.Red, tableCellStyle.TextColor.Green, tableCellStyle.TextColor.Blue));
                    range.Style.Font.Bold = tableCellStyle.IsFontBold;
                    range.Style.Font.Italic = tableCellStyle.IsFontItalic;

                    if (tableCellStyle.IsFontUnderline)
                    {
                        range.Style.Font.Underline = XLFontUnderlineValues.Single;
                    }
  
                    range.Style.Alignment.Horizontal = horAlignment;

                    range.Style.Alignment.Vertical = vertAlignment;
                    if (!ColorsEqual(tableCellStyle.BackgroundColor, _white))
                    {
                        range.Style.Fill.PatternType = XLFillPatternValues.Solid;
                        range.Style.Fill.SetBackgroundColor(XLColor.FromArgb(tableCellStyle.BackgroundColor.Red, tableCellStyle.BackgroundColor.Green, tableCellStyle.BackgroundColor.Blue));
                    }

                    //range.Style.Border.TopBorder = GetLineStyle(tableCellStyle.BorderTopLineStyle);
                    //range.Style.Border.BottomBorder = GetLineStyle(tableCellStyle.BorderBottomLineStyle);
                    //range.Style.Border.LeftBorder = GetLineStyle(tableCellStyle.BorderLeftLineStyle);
                    //range.Style.Border.RightBorder = GetLineStyle(tableCellStyle.BorderRightLineStyle);
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

    private static XLBorderStyleValues GetLineStyle(ElementId border)
    {
        
        //TODO get the line weight and line pattern from the border 
        //Autodesk:Revit:DB:GraphicsStyle

        //OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        //OfficeOpenXml.Style.ExcelBorderStyle.Medium;
        //OfficeOpenXml.Style.ExcelBorderStyle.Thick;
        //OfficeOpenXml.Style.ExcelBorderStyle.Dotted;
        //OfficeOpenXml.Style.ExcelBorderStyle.Dashed;
        //OfficeOpenXml.Style.ExcelBorderStyle.DashDotDot;
        //OfficeOpenXml.Style.ExcelBorderStyle.Double;

        return XLBorderStyleValues.None;
    }
}
