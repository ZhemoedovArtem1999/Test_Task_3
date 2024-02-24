using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Application3.HelpersExcel
{
    public class ExcelRow
    {
        public class ExcelCell
        {
            public bool IsHeader { get; set; }
            public string Header { get; set; }
            public string Value { get; set; }
            public PatternFill PatternFill { get; set; }
            //public ExcelComment Comment { get; set; }
            public override string ToString()
            {
                return this.Value;
            }
        }
        public ExcelRow()
        {
            Cells = new List<ExcelCell>();
        }
        public List<ExcelCell> Cells { get; set; }

        public ExcelCell this[int index]
        {
            get
            {
                return Cells[index];
            }
            set
            {
                Cells[index] = value;
            }
        }


        public ExcelCell this[string header]
        {
            get
            {
                var cell = GetByHeader(header);
                return cell;
            }
            set
            {
                var cell = GetByHeader(header);
                cell = value;
            }
        }

        public ExcelCell GetByHeader(string header)
        {
            return Cells.FirstOrDefault(x => x.Header == header);
        }
    }


    public class ExcelData
    {
        public List<string> Headers { get; set; }
        public List<List<string>> DataRows { get; set; }
        public List<ExcelRow> DataRowCells { get; set; }

        //public List<ExcelComment> Comments { get; set; }
        public string SheetName { get; set; }

        public ExcelData()
        {
            Headers = new List<string>();
            DataRows = new List<List<string>>();
            DataRowCells = new List<ExcelRow>();
        }
    }
}
