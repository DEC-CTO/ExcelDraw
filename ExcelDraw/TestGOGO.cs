using System;
using System.IO;
using System.Collections.Generic;
using OfficeOpenXml;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Windows.Forms;

namespace ExcelDraw
{
    [Transaction(TransactionMode.Manual)]
    public class TestGOGO : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;

            string excelFilePath = "";

            OpenFileDialog dialog = new OpenFileDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                excelFilePath = dialog.FileName;
            }

            double pointToFeet = 1.0 / 72.0 / 12.0;

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int totalRows = worksheet.Dimension.End.Row;
                int totalCols = worksheet.Dimension.End.Column;

                HashSet<string> processedCells = new HashSet<string>();

                using (Transaction trans = new Transaction(doc, "Draw Excel Table"))
                {
                    trans.Start();

                    foreach (var mergeCell in worksheet.MergedCells)
                    {
                        ExcelRange mergedRange = worksheet.Cells[mergeCell];
                        int startRow = mergedRange.Start.Row;
                        int endRow = mergedRange.End.Row;
                        int startCol = mergedRange.Start.Column;
                        int endCol = mergedRange.End.Column;

                        // 병합된 셀의 텍스트를 모두 결합
                        string cellValue = "";
                        for (int row = startRow; row <= endRow; row++)
                        {
                            for (int col = startCol; col <= endCol; col++)
                            {
                                cellValue += worksheet.Cells[row, col].Text + " ";
                            }
                        }
                        cellValue = cellValue.Trim(); // 끝의 공백 제거

                        double xPos = 0;
                        for (int col = 1; col < startCol; col++)
                        {
                            xPos += ColumnWidthToFeet(worksheet, col, pointToFeet);
                        }
                        double yPos = 0;
                        for (int row = 1; row < startRow; row++)
                        {
                            yPos -= worksheet.Row(row).Height * pointToFeet;
                        }
                        XYZ bottomLeft = new XYZ(xPos, yPos, 0);

                        double mergedCellWidth = 0;
                        for (int col = startCol; col <= endCol; col++)
                        {
                            mergedCellWidth += ColumnWidthToFeet(worksheet, col, pointToFeet);
                        }
                        double mergedCellHeight = 0;
                        for (int row = startRow; row <= endRow; row++)
                        {
                            mergedCellHeight += worksheet.Row(row).Height * pointToFeet;
                        }

                        XYZ textPosition = bottomLeft + new XYZ(mergedCellWidth / 2, -mergedCellHeight / 2, 0);
                        ElementId defaultTextTypeId = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType);
                        TextNoteOptions to = new TextNoteOptions();
                        to.HorizontalAlignment = HorizontalTextAlignment.Center;
                        to.VerticalAlignment = VerticalTextAlignment.Middle;
                        to.TypeId = defaultTextTypeId;

                        TextNote.Create(doc, doc.ActiveView.Id, textPosition, cellValue, to);

                        List<Line> cellLines = new List<Line>
                        {
                            Line.CreateBound(bottomLeft, bottomLeft + new XYZ(mergedCellWidth, 0, 0)),
                            Line.CreateBound(bottomLeft, bottomLeft + new XYZ(0, -mergedCellHeight, 0)),
                            Line.CreateBound(bottomLeft + new XYZ(mergedCellWidth, 0, 0), bottomLeft + new XYZ(mergedCellWidth, -mergedCellHeight, 0)),
                            Line.CreateBound(bottomLeft + new XYZ(0, -mergedCellHeight, 0), bottomLeft + new XYZ(mergedCellWidth, -mergedCellHeight, 0))
                        };

                        foreach (Line line in cellLines)
                        {
                            doc.Create.NewDetailCurve(doc.ActiveView, line);
                        }

                        for (int row = startRow; row <= endRow; row++)
                        {
                            for (int col = startCol; col <= endCol; col++)
                            {
                                processedCells.Add($"{row},{col}");
                            }
                        }
                    }

                    for (int row = 1; row <= totalRows; row++)
                    {
                        for (int col = 1; col <= totalCols; col++)
                        {
                            if (processedCells.Contains($"{row},{col}"))
                                continue;

                            string cellValue = worksheet.Cells[row, col].Text;

                            double xPos = 0;
                            for (int c = 1; c < col; c++)
                            {
                                xPos += ColumnWidthToFeet(worksheet, c, pointToFeet);
                            }
                            double yPos = 0;
                            for (int r = 1; r < row; r++)
                            {
                                yPos -= worksheet.Row(r).Height * pointToFeet;
                            }
                            XYZ bottomLeft = new XYZ(xPos, yPos, 0);

                            double cellWidth = ColumnWidthToFeet(worksheet, col, pointToFeet);
                            double cellHeight = worksheet.Row(row).Height * pointToFeet;

                            XYZ textPosition = bottomLeft + new XYZ(cellWidth / 2, -cellHeight / 2, 0);
                            ElementId defaultTextTypeId = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType);
                            TextNoteOptions to = new TextNoteOptions();
                            to.HorizontalAlignment = HorizontalTextAlignment.Center;
                            to.VerticalAlignment = VerticalTextAlignment.Middle;
                            to.TypeId = defaultTextTypeId;
                            TextNote.Create(doc, doc.ActiveView.Id, textPosition, cellValue, to);

                            List<Line> cellLines = new List<Line>
                            {
                                Line.CreateBound(bottomLeft, bottomLeft + new XYZ(cellWidth, 0, 0)),
                                Line.CreateBound(bottomLeft, bottomLeft + new XYZ(0, -cellHeight, 0)),
                                Line.CreateBound(bottomLeft + new XYZ(cellWidth, 0, 0), bottomLeft + new XYZ(cellWidth, -cellHeight, 0)),
                                Line.CreateBound(bottomLeft + new XYZ(0, -cellHeight, 0), bottomLeft + new XYZ(cellWidth, -cellHeight, 0))
                            };

                            foreach (Line line in cellLines)
                            {
                                doc.Create.NewDetailCurve(doc.ActiveView, line);
                            }
                        }
                    }

                    trans.Commit();
                }
            }

            return Result.Succeeded;
        }

        private double ColumnWidthToFeet(ExcelWorksheet worksheet, int col, double pointToFeet)
        {
            double columnWidthInPoints = worksheet.Column(col).Width * 7.5;
            return columnWidthInPoints * pointToFeet;
        }
    }
}
