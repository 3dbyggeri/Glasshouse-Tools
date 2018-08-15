using System;
using System.Collections.Generic;
using System.Linq;
using NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;

namespace NetOffice.Excel.Extensions.Extensions
{
    //https://colinlegg.wordpress.com/2015/04/12/naughty-data-validation-lists/
    //https://andysprague.com/2017/11/30/netoffice-excel-add-validation-to-a-cell/
    public static class CellValidationExtensions
    {
        public static void AddCellListValidation(this Range cell, IList<string> allowedValues, string initialValue = null)
        {
            var flatList = String.Join(System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator, allowedValues.Select(x => x.ToString()).ToArray());
            
            // var flatList = allowedValues.Aggregate((x, y) => "{x},{y}");
            //if (flatList.Length > 255)
            //{
            //    throw new ArgumentException("Combined number of chars in the list of allowedValues can't exceed 255 characters");
            //}
            cell.AddCellListValidation(flatList, initialValue);
        }

        private static void AddCellListValidation(this Range cell, string formula, string initialValue = null)
        {
            cell.Validation.Delete();
            cell.Validation.Add(
                XlDVType.xlValidateList,
                XlDVAlertStyle.xlValidAlertInformation,
                XlFormatConditionOperator.xlEqual,
                formula,
                Type.Missing);
            cell.Validation.IgnoreBlank = true;
            cell.Validation.InCellDropdown = true;
            if (initialValue != null)
            {
                cell.Value = initialValue;
            }
        }
    }
}
