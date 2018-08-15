using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using RestSharp;
using Autodesk.DesignScript.Runtime;
using Microsoft.VisualBasic.FileIO;
using System.Data;

namespace GlasshouseDynamo
{
    
    public class JournalEntries
    {
        [MultiReturn(new[] { "apikey", "headers", "table" })]
        public static Dictionary<string, object> GetViewEntries(string apiKey,string projectId, string viewname)
        {
            List<string> headers = new List<string>();
            DataTable table = GlasshouseShared.JournalEntries.GetViewEntries(apiKey, projectId, viewname);
            //

            List<List<string>> tableList = new List<List<string>>();

            foreach (DataColumn col in table.Columns)
            {
                headers.Add(col.ColumnName);

                List<string> myList = new List<string>();
                foreach (DataRow row in table.Rows)
                {
                    myList.Add((string)row[col]);
                }

                tableList.Add(myList);
            }

            return new Dictionary<string, object>()
                {
                 {"apikey", apiKey  },
                    { "headers",headers},
                    { "table",tableList}
                };

            
        }
    }
}
