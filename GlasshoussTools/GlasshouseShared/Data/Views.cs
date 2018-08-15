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
using Newtonsoft.Json;
using RestSharp;


namespace GlasshouseShared
{
    
    public class Views
    {
        
        public static Dictionary<string, object> GetJournalViews(string apiKey,string projectId)
        {
            var client = new RestClient(GlasshouseShared.Utils.urlApi);

            var request = new RestRequest(string.Format("projects/{0}/new_journal/entries/views",projectId), Method.GET);
            //request.AddParameter("name", "value"); // adds to POST or URL querystring based on Method
            //request.AddUrlSegment("id", "123"); // replaces matching token in request.Resource

            // easily add HTTP Headers
            request.AddHeader("access-token", apiKey);

            request.RequestFormat = DataFormat.Json;

            // execute the request
            IRestResponse response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK) return null;
            var content = response.Content; // raw content as string

            //RestSharp.Deserializers.JsonDeserializer deserialCount = new JsonDeserializer();
            //dbCount = deserialCount.Deserialize<DBCount>(response);

            cViews views = JsonConvert.DeserializeObject<cViews>(content);

            Dictionary<string, string> dict = new Dictionary<string, string>();

            foreach (cView c in views.views)
            {
                dict.Add(c.system_name,c.human_name);
            }

            var items = from pair in dict
                        orderby pair.Value ascending
                        select pair;

            return new Dictionary<string, object>()
                {
                    { "system_name", items.Select(x => x.Key).ToList()},
                    { "human_name",items.Select(x => x.Value).ToList()}
                };

            
        }



        //http://json2csharp.com/
        
        public class cView
        {
            public string system_name { get; set; }
            public string human_name { get; set; }
        }

        
        public class cViews
        {
            public List<cView> views { get; set; }
        }

    }
}
