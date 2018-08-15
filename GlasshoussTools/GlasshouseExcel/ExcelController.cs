#region copyright notice
/*
Original work Copyright(c) 2018 COWI
    
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
using System.Text;
using System.Threading.Tasks;
using Application = NetOffice.ExcelApi.Application;
using ExcelDna.Integration.CustomUI;
using NetOffice.ExcelApi;
using RestSharp;
using System.IO;
using NetOffice.Excel.Extensions.Extensions;
using System.Globalization;
using GlasshouseShared;

namespace GlasshouseExcel
{
    class ExcelController : IDisposable
    {
        private readonly IRibbonUI _modelingRibbon;
        protected readonly Application _excel;
        protected readonly string _curproj;
        protected readonly string _curview;
        protected readonly string _curprojname;
        protected readonly string _curviewname;

        public ExcelController(Application excel, IRibbonUI modelingRibbon, string projectid, string viewid, string projectname, string viewname)
        {
            _modelingRibbon = modelingRibbon;
            _excel = excel;
            _curproj = projectid;
            _curview = viewid;
            _curprojname = projectname;
            _curviewname = viewname;
        }

        public void Dispose()
        {
        }


        public void GetProjects()
        {

            Dictionary<string, object> dict = Projects.GetProjects(Utils.apiKey);

            var activeSheet = _excel.ActiveSheet as Worksheet;
            //activeSheet.Range("A1").Value = "Hello, World!";

            var activeCell = _excel.ActiveCell as Range;

            int c = activeCell.Column;

            foreach (KeyValuePair<string, object> kvp in dict)
            {
                int r = activeCell.Row;
                activeSheet.Cells[r, c].Value = kvp.Key;
                r++;
                List<string> lst = kvp.Value as List<string>;

                foreach (string s in lst)
                {
                    activeSheet.Cells[r, c].Value = s;
                    r++;
                }
                c++;
            }

            /*
            MessageBox.Show("Hello!");
            var activeCell = new ExcelReference(5, 5);
            ExcelAsyncUtil.QueueAsMacro(() => XlCall.Excel(XlCall.xlcSelect, activeCell));
            //https://stackoverflow.com/questions/14896215/how-do-you-set-the-value-of-a-cell-using-excel-dna
            var write2Cell = new ExcelReference(15, 5);
            ExcelAsyncUtil.QueueAsMacro(() => { write2Cell.SetValue("Hello"); });
            */
        }

        public void GetProjectInfo()
        {


            var activeSheet = _excel.ActiveSheet as Worksheet;
            //activeSheet.Range("A1").Value = "Hello, World!";

            var activeCell = _excel.ActiveCell as Range;

            Dictionary<string, object> dict = Projects.GetProjectInfo(Utils.apiKey, _curproj);

            int c = activeCell.Column;

            int r = activeCell.Row;
            foreach (KeyValuePair<string, object> kvp in dict)
            {

                activeSheet.Cells[r, c].Value = kvp.Key;

                activeSheet.Cells[r, c + 1].Value = kvp.Value as string;
                r++;


            }
        }

        public void GetViews()
        {
            var activeSheet = _excel.ActiveSheet as Worksheet;
            //activeSheet.Range("A1").Value = "Hello, World!";

            var activeCell = _excel.ActiveCell as Range;

            Dictionary<string, object> dict = Views.GetJournalViews(Utils.apiKey, _curproj);

            int c = activeCell.Column;

            foreach (KeyValuePair<string, object> kvp in dict)
            {
                int r = activeCell.Row;
                activeSheet.Cells[r, c].Value = kvp.Key;
                r++;
                List<string> lst = kvp.Value as List<string>;

                foreach (string s in lst)
                {
                    activeSheet.Cells[r, c].Value = s;
                    r++;
                }
                c++;
            }

            // cboDepartamentos.Items.Add( New Microsoft.Office.Tools.Ribbon.RibbonDropDownItem With {.Label = line})
        }

        public void GetViewColumns()
        {
            var activeSheet = _excel.ActiveSheet as Worksheet;
            //activeSheet.Range("A1").Value = "Hello, World!";

            var activeCell = _excel.ActiveCell as Range;

            List<string> headers = JournalEntries.GetViewColumns(Utils.apiKey, _curproj, _curview);

            int c = activeCell.Column;
            int r = activeCell.Row;

            var removecols = new[] { "BIM Objects count", "BIM Objects quantity" };

            foreach (string s in headers)
            {
                if (removecols.Any(s.Contains)) continue;

                activeSheet.Cells[r, c].Value = s;

                c++;
            }
            // cboDepartamentos.Items.Add( New Microsoft.Office.Tools.Ribbon.RibbonDropDownItem With {.Label = line})
        }


        public void GetViewEntries()
        {
            var activeSheet = _excel.ActiveSheet as Worksheet;
            //activeSheet.Range("A1").Value = "Hello, World!";

            var activeCell = _excel.ActiveCell as Range;

            System.Data.DataTable table = JournalEntries.GetViewEntries(Utils.apiKey, _curproj, _curview);

            int c = activeCell.Column;

            var removecols = new[] { "BIM Objects count", "BIM Objects quantity" };
            var removehcols = new[] { "glasshousejournalguid", "short description" };
            var allowedValues = new List<string> { "---", "Update" };

            int n = table.Columns.Count;
            string s = "{0} of " + n.ToString() + " columns processed...";
            string caption = "Getting View Entries";

            using (ProgressForm pf = new ProgressForm(caption, s, n))
            {
                foreach (System.Data.DataColumn col in table.Columns)
                {
                    if (removecols.Any(col.ColumnName.Contains)) continue;

                    int r = activeCell.Row;
                    activeSheet.Cells[r, c].Value = col.ColumnName;
                    r++;
                    // add update keyword etc
                    if (!removehcols.Any(col.ColumnName.ToLower().Contains))
                    {
                        activeSheet.Cells[r, c].AddCellListValidation(allowedValues);
                    }
                    //
                    r++;
                    foreach (System.Data.DataRow row in table.Rows)
                    {
                        activeSheet.Cells[r, c].Value = ((string)row[col]);
                        r++;
                    }

                    c++;
                    pf.Increment();
                }

                table = null;
            }
        }

        public void ReadGH()
        {
            System.Windows.Forms.DialogResult dlg = System.Windows.Forms.MessageBox.Show("Are you sure you want to read from view " + _curviewname + " in project " + _curprojname,
                "Read from Glasshouse", System.Windows.Forms.MessageBoxButtons.YesNo);
            if (dlg == System.Windows.Forms.DialogResult.No) return;

            // allway read  - glasshousejournalguid, short description
            Range rngid = FindGUIDCell();
            if (rngid == null)
            {
                System.Windows.Forms.MessageBox.Show("glasshousejournalguid not found in the first 10 by 10 cells");
                return;
            }
            int idrow = rngid.Row;
            int idcol = rngid.Column;

            var removehcols = new[] { "glasshousejournalguid", "short description" };

            var activeSheet = _excel.ActiveSheet as Worksheet;
            Range usedRange = activeSheet.UsedRange;
            int maxr = usedRange.Rows[1].Row + usedRange.Rows.Count - 1;
            int maxc = usedRange.Columns[1].Column + usedRange.Columns.Count - 1;
            // Make dictinary of columns
            List<gColumns> headers = new List<gColumns>();
            for (int c = idcol; c <= maxc; c++)
            {
                if (activeSheet.Cells[idrow, c].Value2 == null) continue;
                string sc = activeSheet.Cells[idrow, c].Value2 as string;
                if (sc.Length > 0)
                {
                    gColumns gc = new gColumns();
                    gc.headerName = sc.Trim();
                    gc.headerNameLC = gc.headerName.ToLower();
                    gc.colNo = c;
                    gc.sync2gh = false;

                    //
                    if (activeSheet.Cells[idrow + 1, c].Value2 == null)
                    {
                        headers.Add(gc);
                        continue;
                    }
                    string syncway = (activeSheet.Cells[idrow + 1, c].Value2 as string).ToLower().Trim();
                    if (removehcols.Any(syncway.Contains)) continue;
                    if (syncway.Equals("update")) gc.sync2gh = true;

                    headers.Add(gc);

                }
            }


            System.Data.DataTable table = JournalEntries.GetViewEntries(Utils.apiKey, _curproj, _curview);


            var removecols = new[] { "BIM Objects count", "BIM Objects quantity" };

            int updateno = 0;
            int newno = 0;
            maxr = Math.Max(maxr, idrow + 2);

            int n = table.Rows.Count;
            string s = "{0} of " + n.ToString() + " rows processed...";
            string caption = "Getting Data From Glasshouse";

            using (ProgressForm pf = new ProgressForm(caption, s, n))
            {
                foreach (System.Data.DataRow row in table.Rows)
                {
                    string rguid = (string)row[0];
                    int foundrow = -1;
                    for (int r = idrow + 2; r <= maxr; r++)
                    {
                        var guid = activeSheet.Cells[r, idcol].Value2;
                        if (guid == null) continue;
                        string sguid = guid as string;
                        if (sguid.Length == 0) continue;


                        if (rguid.Equals(sguid) == true)
                        {
                            foundrow = r;
                            break;
                        }
                    }

                    int colno = 0;
                    int activerow = foundrow;
                    if (foundrow == -1)
                    {
                        activerow = maxr;
                        maxr++; // new line
                        newno++;
                    }
                    else
                    {
                        updateno++;
                    }
                    foreach (object col in row.ItemArray)
                    {
                        string colname = table.Columns[colno].ColumnName.ToLower().Trim();
                        colno++;
                        if (removecols.Any(colname.Contains)) continue;

                        gColumns match = headers.Find(v => v.headerNameLC.Equals(colname));

                        if (match == null) continue;
                        if (match.sync2gh == true) continue;
                        activeSheet.Cells[activerow, match.colNo].Value = col;
                    }
                    pf.Increment();

                }
            }
            table = null;
            
            System.Windows.Forms.MessageBox.Show("Updates " + updateno + " entries, and added " + newno + " new entries ", "Read From Glasshouse");
        }

        public void WriteGH(bool csv = false)
        {
            bool viewexport = true;
            string curview = _curview;
            if (csv == false)
            {
                System.Windows.Forms.DialogResult dlg = System.Windows.Forms.MessageBox.Show("Are you sure you want to write data to Glasshouse project " + _curprojname,
                   "Write to Glasshouse", System.Windows.Forms.MessageBoxButtons.YesNo);
                if (dlg == System.Windows.Forms.DialogResult.No) return;

                dlg = System.Windows.Forms.MessageBox.Show("Do you want to only update parameters in view " + _curviewname + "?",
                   "Write to Glasshouse - by specific view", System.Windows.Forms.MessageBoxButtons.YesNo);

                if (dlg == System.Windows.Forms.DialogResult.No)
                {
                    curview = "all_entries";
                    viewexport = false;
                }
            }
            else
            {
                curview = "all_entries";
                viewexport = false;
            }

            // allway read  - glasshousejournalguid, short description
            Range rngid = FindGUIDCell();
            if (rngid == null)
            {
                System.Windows.Forms.MessageBox.Show("glasshousejournalguid not found in the first 10 by 10 cells");
                return;
            }
            int idrow = rngid.Row;
            int idcol = rngid.Column;

            var removehcols = new[] { "glasshousejournalguid", "short description" };

            var activeSheet = _excel.ActiveSheet as Worksheet;
            Range usedRange = activeSheet.UsedRange;
            int maxr = usedRange.Rows[1].Row + usedRange.Rows.Count - 1;
            int maxc = usedRange.Columns[1].Column + usedRange.Columns.Count - 1;

            System.Data.DataTable table = null;
            System.Data.DataRow tablerow = null;
            if (curview != null)
            {
                 table = JournalEntries.GetViewEntries(Utils.apiKey, _curproj, curview);
                if (table.Rows.Count > 0)
                {
                    tablerow = table.Rows[0];
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Did not find any good data in view "+_curviewname);
                    return;
                }
            }

            // Make dictinary of columns
            List<gColumns> headers = new List<gColumns>();
            List<string> updates = new List<string>();
            List<string> updatesheader = new List<string>();
            updatesheader.Add("GlassHouseJournalGUID");
            for (int c = idcol; c <= maxc; c++)
            {
                if (activeSheet.Cells[idrow, c].Value2 == null) continue;
                string sc = activeSheet.Cells[idrow, c].Value2 as string;
                if (sc.Length > 0)
                {
                    gColumns gc = new gColumns();
                    gc.headerName = sc.Trim();
                    gc.headerNameLC = gc.headerName.ToLower();
                    gc.colNo = c;
                    gc.sync2gh = false;

                    if (tablerow != null)
                    {
                        int colno = 0;
                        foreach (object col in tablerow.ItemArray)
                        {
                            string colname = table.Columns[colno].ColumnName.ToLower().Trim();
                            colno++;
                            if (colname.Equals(gc.headerNameLC))
                            {
                                gc.GHcolNo = colno;
                                break;
                            }
                        }
                    }
                        //
                    if (activeSheet.Cells[idrow + 1, c].Value2 == null)
                    {
                        headers.Add(gc);
                        continue;
                    }
                    string syncway = (activeSheet.Cells[idrow + 1, c].Value2 as string).ToLower().Trim();
                    if (removehcols.Any(gc.headerNameLC.Contains))
                    {
                        headers.Add(gc);
                        continue;
                    }
                    if (syncway.Equals("update"))
                    {
                        gc.sync2gh = true;
                        updatesheader.Add(gc.headerName);
                    }

                    headers.Add(gc);

                }
            }
            if (updatesheader.Count < 2)
            {
                System.Windows.Forms.MessageBox.Show("No parameters selected for updating");
                return;
            }


            List<string> newupdatesheader = new List<string>();
            if (viewexport==true)
            {
                headers = headers.OrderBy(o => o.GHcolNo).ToList();
                foreach (gColumns gc in headers)
                {
                    if(updatesheader.Any(gc.headerName.Contains))
                    {
                        newupdatesheader.Add(gc.headerName);
                    }
                }
                updates.Add(String.Join(",", newupdatesheader.Select(x => x.ToString()).ToArray()));
            }
            else updates.Add(String.Join(",", updatesheader.Select(x => x.ToString()).ToArray()));
           


            var removecols = new[] { "BIM Objects count", "BIM Objects quantity" };

            maxr = Math.Max(maxr, idrow + 2);

            int n = table.Rows.Count;
            string s = "{0} of " + n.ToString() + " rows processed...";
            string caption = "Preparing Data For Glasshouse";

            using (ProgressForm pf = new ProgressForm(caption, s, n))
            {
                foreach (System.Data.DataRow row in table.Rows)
                {
                    string rguid = (string)row[0];
                    int foundrow = -1;
                    for (int r = idrow + 2; r <= maxr; r++)
                    {
                        var guid = activeSheet.Cells[r, idcol].Value2;
                        if (guid == null) continue;
                        string sguid = guid as string;
                        if (sguid.Length == 0) continue;


                        if (rguid.Equals(sguid) == true)
                        {
                            foundrow = r;
                            break;
                        }
                    }

                    int colno = 0;
                    int activerow = foundrow;
                    if (foundrow == -1) continue;

                    List<string> updatescol = new List<string>();

                    if (viewexport == true)
                    {
                        foreach (object col in row.ItemArray)
                        {
                            string colname = table.Columns[colno].ColumnName.ToLower().Trim();
                            colno++;
                            if (removecols.Any(colname.Contains)) continue;

                            gColumns match = headers.Find(v => v.headerNameLC.Equals(colname));

                            if (match == null) continue;
                            if (match.sync2gh == false && match.headerNameLC.Equals("glasshousejournalguid") == false) continue;

                            //add to update
                            var val = activeSheet.Cells[foundrow, match.colNo].Value2;
                            string sval = "-";
                            if (val != null) sval = Utils.FormatWithQuotes(val.ToString());
                            updatescol.Add(sval as string);
                        }
                    }
                    else
                    {
                        foreach (gColumns gc in headers)
                        {
                            if (removecols.Any(gc.headerName.Contains)) continue;
                            if (gc.sync2gh == false && gc.headerNameLC.Equals("glasshousejournalguid") == false) continue;
                            //add to update
                            var val = activeSheet.Cells[foundrow, gc.colNo].Value2;
                            string sval = "-";
                            if (val != null) sval = Utils.FormatWithQuotes(val.ToString());
                            updatescol.Add(sval as string);
                        }
                    }
                    updates.Add(String.Join(",", updatescol.Select(x => x.ToString()).ToArray()));
                    //updates.Add(updatescol.Aggregate("", (string agg, string it) => agg += string.Format("{0} \"{1}\"", agg == "" ? "" : ",", it)));
                    pf.Increment();
                }
            }
            if (updates.Count < 2)
            {
                System.Windows.Forms.MessageBox.Show("Nothing to update");
                return;
            }

            string path = @"C:\temp\" + System.IO.Path.GetFileNameWithoutExtension(_excel.ActiveWorkbook.Name) + "_updateglasshouse.csv";

            n = updates.Count;
            s = "{0} of " + n.ToString() + " rows processed...";
            caption = "Writing Data For Glasshouse";

            using (ProgressForm pf = new ProgressForm(caption, s, n))
            {
                using (var w = new StreamWriter(path, false, Encoding.UTF8))
                {
                    foreach (string us in updates)
                    {
                        w.WriteLine(us);
                        w.Flush();
                    }
                }
            }

            if (csv == false)
            {
                if (UpdateJournalCVS(_curproj, path) == true)
                {
                    System.Windows.Forms.MessageBox.Show("Glashouse updated", "Write To Glasshouse");
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Hmmm...something went wrong!", "Write To Glasshouse");
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("CSV dumped at " + path, "Write To CSV File");
            }

            updates = null;
            updatesheader = null;
            headers = null;
            table = null;
        }


        public void Validator()
        {
            var activeSheet = _excel.ActiveSheet as Worksheet;
            //activeSheet.Range("A1").Value = "Hello, World!";

            var activeCell = _excel.ActiveCell as Range;

            int c = activeCell.Column;
            int r = activeCell.Row;

            var removecols = new[] { "glasshousejournalguid", "short description", "bim objects count", "bim objects quantity" };

            var allowedValues = new List<string> { "---", "Update" };

            if (r > 1)
            {
                var val = activeSheet.Cells[r - 1, c].Value;
                string sval = "---";
                if (val != null) sval = val.ToString();
                if (removecols.Any(sval.ToLower().Contains)) return;
            }

            activeSheet.Cells[r, c].AddCellListValidation(allowedValues);


        }

        public static bool UpdateJournalCVS(string projectId, string fullpath)
        {
            var client = new RestClient(Utils.urlApi);

            string text = System.IO.File.ReadAllText(fullpath);

            var request = new RestRequest(string.Format("projects/{0}/new_journal/entries/csv_import", projectId), Method.POST);

            request.AddHeader("access-token", Utils.apiKey);

            request.AddParameter("text/plain", text, "text/plain", ParameterType.RequestBody);

            request.RequestFormat = DataFormat.Xml;

            // execute the request
            IRestResponse response = client.Execute(request);
            var content = response.Content; // raw content as string

            if (content.Contains("project")) return true;

            return false;
        }

        public Range FindGUIDCell()
        {
            var activeSheet = _excel.ActiveSheet as Worksheet;
            for (int c = 1; c < 11; c++)
            {
                for (int r = 1; r < 11; r++)
                {
                    var val = activeSheet.Cells[r, c].Value2;
                    if (val == null) continue;
                    if (activeSheet.Cells[r, c].Value2.ToString().ToLower().Trim().Equals("glasshousejournalguid") == true)
                    {
                        return activeSheet.Cells[r, c];
                    }
                }
            }
            return null;
        }

        public void Login()
        {
            SettingsForm dlg = new SettingsForm();

            // Show testDialog as a modal dialog and determine if DialogResult = OK.
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // Read the contents of testDialog's TextBox.
                //this.txtResult.Text = dlg.TextBox1.Text;
            }
            else
            {
                //this.txtResult.Text = "Cancelled";
            }
            dlg.Dispose();
        }

        public void Logout()
        {
            Utils.apiKey = "Login";
        }


        public void About()
        {
            AboutForm dlg = new AboutForm();
            dlg.ShowDialog();
            dlg.Dispose();
        }
    }

    public class gColumns
    {
        public string headerName { get; set; }
        public string headerNameLC { get; set; }
        public int colNo { get; set; }
        public int GHcolNo { get; set; }
        public bool sync2gh { get; set; }
    }

}


