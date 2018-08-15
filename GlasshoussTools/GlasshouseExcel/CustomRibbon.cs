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
using ExcelDna.Integration;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;
using System.IO;
using System.Resources;
using Application = NetOffice.ExcelApi.Application;
using GlasshouseShared;


namespace GlasshouseExcel
{
    //https://github.com/Excel-DNA/Samples/tree/master/Ribbon
    //http://www.addinx.org/addinx/control_combobox.html
    //https://exceloffthegrid.com/inserting-a-dynamic-drop-down-in-ribbon/
    //http://www.andypope.info/vba/ribboneditor.htm

    //    https://ichihedge.wordpress.com/2016/12/01/excel-dna-enable-configuration/

    [ComVisible(true)]
    public class CustomRibbon : ExcelRibbon
    {
        private Application _excel;
        private IRibbonUI _thisRibbon;

        private Dictionary<string, object> _projects = new Dictionary<string, object>();
        private Dictionary<string, object> _views = new Dictionary<string, object>();
        private string _curproj = null;
        private string _curview = null;
        private string _curprojname = null;
        private string _curviewname = null;
        /*
        public override string GetCustomUI(string ribbonId)
        {
            _excel = new Application(null, ExcelDna.Integration.ExcelDnaUtil.Application);
            string ribbonXml = GetCustomRibbonXML();
            return ribbonXml;
        }

        private string GetCustomRibbonXML()
        {
            string ribbonXml;
            var thisAssembly = typeof(CustomRibbon).Assembly;
            var resourceName = typeof(CustomRibbon).Namespace + ".CustomRibbon.xml";

            using (Stream stream = thisAssembly.GetManifestResourceStream(resourceName))
            using (StreamReader reader = new StreamReader(stream))
            {
                ribbonXml = reader.ReadToEnd();
            }

            if (ribbonXml == null)
            {
                throw new MissingManifestResourceException(resourceName);
            }
            return ribbonXml;
        }
        */
        public void OnLoad(IRibbonUI ribbon)
        {
            if (ribbon == null)
            {
                throw new ArgumentNullException(nameof(ribbon));
            }

            _thisRibbon = ribbon;

            _excel = new Application(null, ExcelDna.Integration.ExcelDnaUtil.Application); //added this
            //_excel.WorkbookActivateEvent += OnInvalidateRibbon;
            //_excel.WorkbookDeactivateEvent += OnInvalidateRibbon;
            //_excel.SheetActivateEvent += OnInvalidateRibbon;
            //_excel.SheetDeactivateEvent += OnInvalidateRibbon;
            _thisRibbon.Invalidate();
            //if (_excel.ActiveWorkbook == null)
            //{
            //    _excel.Workbooks.Add();
            //}

            //get apikey
            Utils.apiKey = System.Configuration.ConfigurationManager.AppSettings["apiKey"];

            //get projects;
            _projects = Projects.GetProjects(Utils.apiKey);
        }

        private void OnInvalidateRibbon(object obj)
        {
            _thisRibbon.Invalidate();
        }


        public void btnGetProjects(IRibbonControl control)
        {
            // _excel = new Application(null, ExcelDna.Integration.ExcelDnaUtil.Application);
            if (_excel.ActiveWorkbook == null) return;
            using (var controller = new ExcelController(_excel, _thisRibbon, _curproj,_curview, _curprojname, _curviewname)) {controller.GetProjects();}
        }

        public void btnGetProjectInfo(IRibbonControl control)
        {
            // _excel = new Application(null, ExcelDna.Integration.ExcelDnaUtil.Application);
            if (_excel.ActiveWorkbook == null) return;
            using (var controller = new ExcelController(_excel, _thisRibbon, _curproj, _curview, _curprojname, _curviewname)) { controller.GetProjectInfo(); }

        }

        public void btnGetViews(IRibbonControl control)
        {
            //_excel = new Application(null, ExcelDna.Integration.ExcelDnaUtil.Application);
            if (_excel.ActiveWorkbook == null) return;
            using (var controller = new ExcelController(_excel, _thisRibbon, _curproj, _curview, _curprojname, _curviewname)) { controller.GetViews(); }

        }

        public void btnGetViewColumns(IRibbonControl control)
        {
            //_excel = new Application(null, ExcelDna.Integration.ExcelDnaUtil.Application);
            if (_excel.ActiveWorkbook == null) return;
            using (var controller = new ExcelController(_excel, _thisRibbon, _curproj, _curview, _curprojname, _curviewname)) { controller.GetViewColumns(); }

        }

        public void btnGetViewEntries(IRibbonControl control)
        {
            //_excel = new Application(null, ExcelDna.Integration.ExcelDnaUtil.Application);
            if (_excel.ActiveWorkbook == null) return;
            using (var controller = new ExcelController(_excel, _thisRibbon, _curproj, _curview, _curprojname, _curviewname)) { controller.GetViewEntries(); }

        }
        

        public void btnRefreshProjectList(IRibbonControl control)
        {
            _projects = Projects.GetProjects(Utils.apiKey);
            _thisRibbon.Invalidate();

        }

        public void btnRefreshViewList(IRibbonControl control)
        {
            _views = Views.GetJournalViews(Utils.apiKey, _curproj);
            _thisRibbon.Invalidate();
        }



        public void btnRead(IRibbonControl control)
        {
            if (_excel.ActiveWorkbook == null) return;
            if (_curproj == null || _curview==null) return;
            List<string> plst = _projects["id"] as List<string>;
            List<string> pname = _projects["name"] as List<string>;

            int i = plst.FindIndex(v => v.Equals(_curproj));
            _curprojname = pname.ElementAt(i);

            List<string> vlst = _views["system_name"] as List<string>;
            List<string> vname = _views["human_name"] as List<string>;

            i = vlst.FindIndex(v => v.Equals(_curview));
            _curviewname = vname.ElementAt(i);

            //_excel = new Application(null, ExcelDna.Integration.ExcelDnaUtil.Application);
            using (var controller = new ExcelController(_excel, _thisRibbon, _curproj, _curview, _curprojname, _curviewname)) { controller.ReadGH(); }

            plst = null;
            pname = null;
            vlst = null;
            vname = null;
        }

        public void btnWrite(IRibbonControl control)
        {
            if (_excel.ActiveWorkbook == null) return;
            if (_curproj == null) return;

            if (_curproj == null || _curview == null) return;
            List<string> plst = _projects["id"] as List<string>;
            List<string> pname = _projects["name"] as List<string>;

            int i = plst.FindIndex(v => v.Equals(_curproj));
            _curprojname = pname.ElementAt(i);

            List<string> vlst = _views["system_name"] as List<string>;
            List<string> vname = _views["human_name"] as List<string>;

            i = vlst.FindIndex(v => v.Equals(_curview));
            _curviewname = vname.ElementAt(i);

            //_excel = new Application(null, ExcelDna.Integration.ExcelDnaUtil.Application);
            using (var controller = new ExcelController(_excel, _thisRibbon, _curproj, _curview, _curprojname, _curviewname)) { controller.WriteGH(false); }

            plst = null;
            pname = null;
            vlst = null;
            vname = null;
        }

        public void btnWriteCSV(IRibbonControl control)
        {
            if (_excel.ActiveWorkbook == null) return;
            if (_curproj == null) return;

            if (_curproj == null || _curview == null) return;
            List<string> plst = _projects["id"] as List<string>;
            List<string> pname = _projects["name"] as List<string>;

            int i = plst.FindIndex(v => v.Equals(_curproj));
            _curprojname = pname.ElementAt(i);

            // _excel = new Application(null, ExcelDna.Integration.ExcelDnaUtil.Application);
            using (var controller = new ExcelController(_excel, _thisRibbon, _curproj, _curview, _curprojname, _curviewname)) { controller.WriteGH(true); }

            plst = null;
            pname = null;
        }

        public void btnValidator(IRibbonControl control)
        {
            //_excel = new Application(null, ExcelDna.Integration.ExcelDnaUtil.Application);
            if (_excel.ActiveWorkbook == null) return;
            using (var controller = new ExcelController(_excel, _thisRibbon, _curproj, _curview, _curprojname, _curviewname)) { controller.Validator(); }
        }

        public void btnLogin(IRibbonControl control)
        {
            //_excel = new Application(null, ExcelDna.Integration.ExcelDnaUtil.Application);
            using (var controller = new ExcelController(_excel, _thisRibbon, _curproj, _curview, _curprojname, _curviewname)) { controller.Login(); }
            _projects = Projects.GetProjects(Utils.apiKey);
            _thisRibbon.Invalidate();
        }

        public void btnLogout(IRibbonControl control)
        {
            // _excel = new Application(null, ExcelDna.Integration.ExcelDnaUtil.Application);
            using (var controller = new ExcelController(_excel, _thisRibbon, _curproj, _curview, _curprojname, _curviewname)) { controller.Logout(); }
            _projects = Projects.GetProjects(Utils.apiKey);
            // clear ribbon
            _thisRibbon.Invalidate();
        }

        public void btnAbout(IRibbonControl control)
        {
            //_excel = new Application(null, ExcelDna.Integration.ExcelDnaUtil.Application);
            using (var controller = new ExcelController(_excel, _thisRibbon, _curproj, _curview, _curprojname, _curviewname)) { controller.About(); }
        }


        //
        //
        public int cbProjects_GetItemCount(IRibbonControl control)
        {
            // This Callback will create the number of drop-down items as determined by the returned value
            if (_projects == null || _projects.Count==0) return 0;
            List<string> lst = _projects["id"] as List<string>;
            return lst.Count;
        }

        public string cbProjects_GetItemID(IRibbonControl control, int index)
        {
            //This Callback will set the id for each item created. It provides the index value within the Callback.
            //The index is the position within the drop-down list. The index can be used to create the id.
            List<string> lst = _projects["id"] as List<string>;
            return lst.ElementAt(index);
        }

        public string cbProjects_GetItemLabel(IRibbonControl control, int index)
        {
            //This Callback will set the displayed label for each item created. It provides the index value within the Callback.
            //The index is the position within the drop-down list. The index can be used to create the id.
            //string label = "Unknow Workbook";

            //if (index == 1) label = "1 Workbook";
            //if (index == 2) label = "S. Workbook";
            //if (index == 3) label = "A. Workbook";
            //_myRibbon.InvalidateControl("cbScenarioWorkbook");
            //return label;
             //_thisRibbon.Invalidate();
            List<string> name = _projects["name"] as List<string>;
            return name.ElementAt(index);
            
        }

        public string cbProjects_GetSelectedItemID(IRibbonControl control)
        {
            //This Callback will change the drop-down to be set to a specific the id. This could be used to set a default value or reset the first item in the list
            //This example will set the selected item to the id with "wB2"
            if (_curproj == null)
            {
                List<string> lst = _projects["id"] as List<string>;
                _curproj = lst[0];
                _views = Views.GetJournalViews(Utils.apiKey, _curproj);
                _curview = null;
                _thisRibbon.Invalidate();
            }
            return _curproj;
        }

        public void cbProjects_onAction(IRibbonControl control, string selectedId, int selectedIndex)
        {
            //MessageBox.Show("My Dropdown Selected on control " + control.Id + " with selection " + selectedId + " at index " + selectedIndex);
            _curproj = selectedId;
            _views = Views.GetJournalViews(Utils.apiKey, _curproj);
            _curview = null;
            _thisRibbon.Invalidate();
        }
        //
        //
        public int cbViews_GetItemCount(IRibbonControl control)
        {
            // This Callback will create the number of drop-down items as determined by the returned value
            if (_views == null || _views.Count==0) return 0;
            List<string> lst = _views["system_name"] as List<string>;
            return lst.Count;
        }

        public string cbViews_GetItemID(IRibbonControl control, int index)
        {
            //This Callback will set the id for each item created. It provides the index value within the Callback.
            //The index is the position within the drop-down list. The index can be used to create the id.
            List<string> lst = _views["system_name"] as List<string>;
            return lst.ElementAt(index);
        }

        public string cbViews_GetItemLabel(IRibbonControl control, int index)
        {
            //This Callback will set the displayed label for each item created. It provides the index value within the Callback.
            //The index is the position within the drop-down list. The index can be used to create the id.
            List<string> name = _views["human_name"] as List<string>;
            return name.ElementAt(index);
        }

        public string cbViews_GetSelectedItemID(IRibbonControl control)
        {
            //This Callback will change the drop-down to be set to a specific the id. This could be used to set a default value or reset the first item in the list
            //This example will set the selected item to the id with "wB2"
            if (_curview == null)
            {
                List<string> lst = _views["system_name"] as List<string>;
                _curview = lst[0];
                _thisRibbon.Invalidate();
            }
            return _curview;
        }

        public void cbViews_onAction(IRibbonControl control, string selectedId, int selectedIndex)
        {
            //MessageBox.Show("My Dropdown Selected on control " + control.Id + " with selection " + selectedId + " at index " + selectedIndex);
            _curview = selectedId;
        }

    }
}

