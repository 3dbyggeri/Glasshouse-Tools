using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using RestSharp;
using Autodesk.DesignScript.Runtime;

namespace GlasshouseDynamo
{
    
    public class Views
    {
        [MultiReturn(new[] { "apikey", "system_name", "human_name" })]
        public static Dictionary<string, object> GetJournalViews(string apiKey,string projectId)
        {
            Dictionary<string, object> dict = new Dictionary<string, object>();
            dict.Add("apikey", apiKey);

            foreach (var kvp in GlasshouseShared.Views.GetJournalViews(apiKey, projectId))
                dict.Add(kvp.Key, kvp.Value);

            return dict;
        }
    }
}
