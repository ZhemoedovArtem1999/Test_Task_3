using Application3.Exceptions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static Application3.HelpersExcel.ExcelRow;

namespace Application3.HelpersExcel
{
    public class HelperExcel
    {

        private string ReadExcelCell(Cell cell, WorkbookPart workbookPart)
        {
            string text = string.Empty;
            if (cell.CellValue is not null)
            {
                CellValue cellValue = cell.CellValue;
                text = (cellValue == null) ? cell.InnerText : cellValue.Text;
                if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
                {
                    text = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(
                            Convert.ToInt32(cell.CellValue.Text)).InnerText;
                }
                if ((cell.DataType == null))
                {
                    if (cell.StyleIndex != null)
                    {
                        CellFormat cellFormat = (CellFormat)workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ChildElements[int.Parse(cell.StyleIndex.InnerText)];
                        if (cellFormat.NumberFormatId == 14)
                        {
                            DateTime date = DateTime.FromOADate(double.Parse(cell.CellValue.Text));
                            text = date.ToString();
                        }
                    }

                }
            }

            return (text ?? string.Empty).Trim();
        }



        public ExcelData ReadExcel(string filePath, string sheetName)
        {
            ExcelData data = new ExcelData();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Where(sheet => sheet.Name == sheetName).First();

                WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id);

                Worksheet workSheet = worksheetPart.Worksheet;

                SheetData sheetData = workSheet.Elements<SheetData>().First();

                List<Row> rows = sheetData.Elements<Row>().ToList();

                if (rows.Count > 0)
                {
                    for (var i = 0; i < rows.Count; i++)
                    {
                        var cellRow = new ExcelRow();
                        var dataRow = new List<string>();
                        var row = rows[i];
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            var text = ReadExcelCell(cell, workbookPart).Trim();
                            dataRow.Add(text);
                        }
                        data.DataRowCells.Add(cellRow);
                        if (i == 0)
                        {
                            data.Headers = dataRow;
                        }
                        else
                        {
                            data.DataRows.Add(dataRow);
                        }
                    }
                }

                return data;
            }

        }

        public List<string> GetHeader(string filePath, string sheetName)
        {
            List<string> header = new List<string>();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Where(sheet => sheet.Name == sheetName).First();

                WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id);

                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                Row row = sheetData.Elements<Row>().First();
                foreach (Cell cell in row.Descendants<Cell>())
                //while (cellEnumerator.MoveNext())
                {
                    header.Add(ReadExcelCell(cell, workbookPart).Trim());
                }
            }
            return header;
        }

        public bool IsSheetInExcelFile(string filePath, string sheetName) {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Where(sheet => sheet.Name == sheetName).FirstOrDefault();

                if (sheet is null)
                    return false;
                else
                    return true;
            }
            
        }


        public void UpdateExcel(string filePath, string sheetName, int indexColumnFind, int indexColumnUpdate,  string find, string newValue)
        {
            bool isUpdate = false;
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, true))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Where(sheet => sheet.Name == sheetName).First();

                WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id);

                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                List<Row> rows = sheetData.Elements<Row>().ToList();
                rows.RemoveAt(0);

                if (rows.Count > 0)
                {
                    for (var i = 0; i < rows.Count; i++)
                    {
                        var cellRow = new ExcelRow();
                        var dataRow = new List<string>();
                        //var patternFillRow = new List<PatternFill>();
                        var row = rows[i];

                        Cell cellFind = row.Descendants<Cell>().ToList()[indexColumnFind];

                        if (ReadExcelCell(cellFind, workbookPart) == find)
                        {
                            Cell cellUpdate = row.Descendants<Cell>().ToList()[indexColumnUpdate];
                            cellUpdate.CellValue = new CellValue(newValue);
                            cellUpdate.DataType = new EnumValue<CellValues>(CellValues.String);

                            worksheetPart.Worksheet.Save();
                            isUpdate = true;
                        }
                    }
                }
                else
                {
                    throw new NotFoundDataException("В таблице нет данных!");
                }
            }
            if (!isUpdate)
            {
                throw new UpdateFailedException("Ни одна запись не обновлена");
            }
        }
    }
}
