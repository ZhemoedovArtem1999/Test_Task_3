using Application3.Exceptions;
using Application3.HelpersExcel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Application3.Service
{
    public class ValidateExcel : IValidate
    {
        private string _filePath;
        private Dictionary<string, List<string>> _sheets;

        public ValidateExcel(string filePath, Dictionary<string, List<string>> sheets) { 
            _filePath = filePath;
            _sheets = sheets;
        }

        private bool ValidateHeaderSheet(List<string> header, string key)
        {
            bool isValidHeader = true;
            foreach(var value in _sheets.GetValueOrDefault(key))
            {
                isValidHeader = header.Contains(value);
                if (!isValidHeader)
                    break;
            }
            return isValidHeader;
        }

        public ValidateModel Validate()
        {
            ValidateModel model = new ValidateModel();
            HelperExcel helperExcel = new();
            bool isValid = true;

            model.IsValid = isValid;
            foreach(var sheet in _sheets)
            {
                try
                {
                    isValid = helperExcel.IsSheetInExcelFile(_filePath, sheet.Key);
                }
                catch (IOException ex)
                {
                    model.IsValid = false;
                    model.Errors.Add("Файл по пути не найден!");
                    return model;
                }
                if (!isValid)
                {
                    model.IsValid = false;
                    model.Errors.Add($"В файле нет листа {sheet.Key}!"); ;
                }
                else
                {
                    var header = helperExcel.GetHeader( _filePath, sheet.Key );
                    if (!this.ValidateHeaderSheet( header, sheet.Key ))
                    {
                        model.IsValid = false;
                        model.Errors.Add($"На листе {sheet.Key} заголовок таблицы не соответствует!");
                    }
                }
            }

            return model;
        }
    }
}
