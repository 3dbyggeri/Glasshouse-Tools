﻿#region copyright notice
/*
Original work Copyright(c) 2018-2021 COWI
    
Copyright © COWI and individual contributors. All rights reserved.

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

    1) Redistributions of source code must retain the above copyright notice,
    this list of conditions and the following disclaimer.

    2) Redistributions in binary form must reproduce the above copyright notice,
    this list of conditions and the following disclaimer in the documentation
    and/or other materials provided with the distribution.

    3) Neither the name of COWI nor the names of its contributors may be used
    to endorse or promote products derived from this software without specific
    prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS “AS IS”
AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE
LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF
THE POSSIBILITY OF SUCH DAMAGE.

GlasshouseExcel may utilize certain third party software. Such third party software is copyrighted by their respective owners as indicated below.
Netoffice - MIT License - https://github.com/NetOfficeFw/NetOffice/blob/develop/LICENSE.txt
Excel DNA - zlib License - https://github.com/Excel-DNA/ExcelDna/blob/master/LICENSE.txt
RestSharp - Apache License - https://github.com/restsharp/RestSharp/blob/develop/LICENSE.txt
Newtonsoft - The MIT License (MIT) - https://github.com/JamesNK/Newtonsoft.Json/blob/master/LICENSE.md
*/
#endregion

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
