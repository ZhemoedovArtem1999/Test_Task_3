using Application3.Exceptions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static Application3.HelpersExcel.ExcelRow;

namespace Application3.HelpersExcel
{
    public class HelperExcel
    {

        private string GetColumnName(string cellReference)
        {
            var regex = new Regex("[A-Za-z]+");
            var match = regex.Match(cellReference);

            return match.Value;
        }


        private int ConvertColumnNameToNumber(string columnName)
        {
            var alpha = new Regex("^[A-Z]+$");
            if (!alpha.IsMatch(columnName)) throw new ArgumentException();

            char[] colLetters = columnName.ToCharArray();
            Array.Reverse(colLetters);

            var convertedValue = 0;
            for (int i = 0; i < colLetters.Length; i++)
            {
                char letter = colLetters[i];
                int current = i == 0 ? letter - 65 : letter - 64; // ASCII 'A' = 65
                convertedValue += current * (int)Math.Pow(26, i);
            }

            return convertedValue;
        }

        //private IEnumerator<Cell> GetExcelCellEnumerator(Row row)
        //{
        //    int currentCount = 0;
        //    foreach (Cell cell in row.Descendants<Cell>())
        //    {
        //        string columnName = GetColumnName(cell.CellReference);

        //        int currentColumnIndex = ConvertColumnNameToNumber(columnName);

        //        for (; currentCount < currentColumnIndex; currentCount++)
        //        {
        //            var emptycell = new Cell() { DataType = null, CellValue = new CellValue(string.Empty) };
        //            yield return emptycell;
        //        }

        //        yield return cell;
        //        currentCount++;
        //    }
        //}

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

                    //text = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(
                    //        Convert.ToInt32(cell.CellValue.Text)).InnerText;
                }
            }


            return (text ?? string.Empty).Trim();
        }

        //private PatternFill GetCellPatternFill(Cell cell, WorkbookPart workbookPart)
        //{
        //    WorkbookStylesPart styles = workbookPart.WorkbookStylesPart;
        //    if (styles == null)
        //    {
        //        return null;
        //    }

        //    int cellStyleIndex;
        //    if (cell.StyleIndex == null) // I think (from testing) if the StyleIndex is null
        //    {                               // then this means use cell style index 0.
        //        cellStyleIndex = 0;           // However I did not found it in the open xml 
        //    }                               // specification.
        //    else
        //    {
        //        cellStyleIndex = (int)cell.StyleIndex.Value;
        //    }

        //    CellFormat cellFormat = (CellFormat)styles.Stylesheet.CellFormats.ChildElements[cellStyleIndex];

        //    Fill fill = (Fill)styles.Stylesheet.Fills.ChildElements[(int)cellFormat.FillId.Value];
        //    return fill.PatternFill;
        //}


        public ExcelData ReadExcel(string filePath, string sheetName)
        {
            ExcelData data = new ExcelData();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                Sheet sheetProduct = workbookPart.Workbook.Descendants<Sheet>().Where(sheet => sheet.Name == sheetName).First();

                WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheetProduct.Id);

                Worksheet workSheet = worksheetPart.Worksheet;

                Columns columns = workSheet.Descendants<Columns>().FirstOrDefault();



                SheetData sheetData = workSheet.Elements<SheetData>().First();


                List<Row> rows = sheetData.Elements<Row>().ToList();


                if (rows.Count > 0)
                {
                    for (var i = 0; i < rows.Count; i++)
                    {
                        var cellRow = new ExcelRow();
                        var dataRow = new List<string>();
                        //var patternFillRow = new List<PatternFill>();
                        var row = rows[i];
                        //var cellEnumerator = GetExcelCellEnumerator(row);
                        int z = 0;
                        foreach (Cell cell in row.Descendants<Cell>())
                        //while (cellEnumerator.MoveNext())
                        {
                            //var cell = cellEnumerator.Current;
                            var text = ReadExcelCell(cell, workbookPart).Trim();
                            //var comment = comments.FirstOrDefault(x => x.Reference == cell.CellReference);

                            dataRow.Add(text);
                            //cellRow.Cells.Add(new ExcelCell
                            //{
                            //    IsHeader = i == 0,
                            //   // Comment = comment,
                            //    Header = (i == 0 ? text : z < data.Headers.Count() ? data.Headers[z] : string.Empty),
                            //    PatternFill = patternFill,
                            //    Value = text
                            //});
                            //z++;
                        }
                        data.DataRowCells.Add(cellRow);
                        if (i == 0)
                        {
                            data.Headers = dataRow;
                        }
                        else
                        {
                            data.DataRows.Add(dataRow);
                            //data.PatternFillRows.Add(patternFillRow);
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
                Sheet sheetProduct = workbookPart.Workbook.Descendants<Sheet>().Where(sheet => sheet.Name == sheetName).First();

                WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheetProduct.Id);


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


        public void UpdateExcel(string filePath, string sheetName, int indexColumnFind, int indexColumnUpdate,  string find, string newValue)
        {
            bool isUpdate = false;
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, true))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                Sheet sheetProduct = workbookPart.Workbook.Descendants<Sheet>().Where(sheet => sheet.Name == sheetName).First();

                WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheetProduct.Id);


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
