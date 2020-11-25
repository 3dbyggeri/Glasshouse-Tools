#region copyright notice
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
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using RestSharp;
using Microsoft.VisualBasic.FileIO;
using System.Data;
using Newtonsoft.Json.Linq;

namespace GlasshouseShared
{

    public class JournalEntries
    {

        public static List<string> GetViewColumns(string apiKey, string projectId, string viewname)
        {
            var client = new RestClient(GlasshouseShared.Utils.urlApi);

            var request = new RestRequest(string.Format("projects/{0}/new_journal/entries.csv?view={1}", projectId, viewname), Method.GET);

            //request.AddParameter("name", "value"); // adds to POST or URL querystring based on Method
            //request.AddUrlSegment("id", "123"); // replaces matching token in request.Resource

            // easily add HTTP Headers
            request.AddHeader("access-token", apiKey);

            // request.RequestFormat = DataFormat.Json;

            // execute the request
            IRestResponse response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK) return null;
            var content = response.Content; // raw content as string


            List<string> headers = new List<string>();


            //
            using (System.IO.MemoryStream ms = new System.IO.MemoryStream(Encoding.UTF8.GetBytes(content)))
            {
                using (TextFieldParser csvReader = new TextFieldParser(ms))
                {
                    csvReader.SetDelimiters(new string[] { "," });
                    csvReader.HasFieldsEnclosedInQuotes = true;
                    string[] colFields = csvReader.ReadFields();
                    foreach (string s in colFields)
                    {
                        //headers
                        headers.Add(s);


                    }


                }
            }


            return headers;
        }

        public static DataTable GetViewEntries(string apiKey, string projectId, string viewname)
        {
            var client = new RestClient(GlasshouseShared.Utils.urlApi);

            var request = new RestRequest(string.Format("projects/{0}/new_journal/entries.csv?view={1}", projectId, viewname), Method.GET);

            //request.AddParameter("name", "value"); // adds to POST or URL querystring based on Method
            //request.AddUrlSegment("id", "123"); // replaces matching token in request.Resource

            // easily add HTTP Headers
            request.AddHeader("access-token", apiKey);

            // request.RequestFormat = DataFormat.Json;

            // execute the request
            IRestResponse response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK) return null;
            var content = response.Content; // raw content as string


            List<string> headers = new List<string>();

            DataTable table = new DataTable();
            //
            using (System.IO.MemoryStream ms = new System.IO.MemoryStream(Encoding.UTF8.GetBytes(content)))
            {
                using (TextFieldParser csvReader = new TextFieldParser(ms))
                {
                    csvReader.SetDelimiters(new string[] { "," });
                    csvReader.HasFieldsEnclosedInQuotes = true;
                    string[] colFields = csvReader.ReadFields();
                    foreach (string s in colFields)
                    {
                        //headers
                        headers.Add(s);

                        table.Columns.Add(s, typeof(string));
                    }



                    while (!csvReader.EndOfData)
                    {
                        string[] split = csvReader.ReadFields();

                        for (int j = 0; j < split.Length; j++)
                        {
                            bool flag = split[j] == null;

                            if (flag)
                            {
                                split[j] = string.Empty;
                            }

                        }

                        table.Rows.Add(split);

                    }

                }
            }


            return table;
        }


        public static DataTable GetAllViewEntries(string apiKey, string projectId)
        {
            var client = new RestClient(GlasshouseShared.Utils.urlApi);

            var request = new RestRequest(string.Format("projects/{0}/new_journal/entries.json", projectId), Method.GET);

            //request.AddParameter("name", "value"); // adds to POST or URL querystring based on Method
            //request.AddUrlSegment("id", "123"); // replaces matching token in request.Resource

            // easily add HTTP Headers
            request.AddHeader("access-token", apiKey);

            //request.RequestFormat = DataFormat.Json;

            // execute the request
            IRestResponse response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK) return null;
            var content = response.Content; // raw content as string

            var journal = JsonConvert.DeserializeObject<dynamic>(content);

            var converter = new Newtonsoft.Json.Converters.ExpandoObjectConverter();
            dynamic message = JsonConvert.DeserializeObject<System.Dynamic.ExpandoObject>(content, converter);

            dynamic new_journal = journal["new_journal"];

            dynamic entries = new_journal["entries"];


            DataTable table = new DataTable();

            foreach (dynamic entry in entries)
            {
                foreach (dynamic prop in entry)
                {
                    JObject jObject = prop.First as JObject;

                    string name = "";

                    if (jObject == null)
                    {
                        name = prop.Name;
                    }
                    else if (jObject.First ==null)
                    {
                        name = prop.Name;
                    }
                    else
                    {
                        name = jObject.Property("human_name").Value.ToString();
                    }

                    if (table.Columns.Contains(name) == false && name!="")
                    {
                        table.Columns.Add(name, typeof(string));
                    }
                }

                string[] split = new string[table.Columns.Count];
                for (int j = 0; j < split.Length; j++)
                {
                    split[j] = string.Empty;
                }

                foreach (dynamic prop in entry)
                {
                    JObject jObject = prop.First as JObject;

                    string name = "";
                    string value = "";

                    if (jObject == null)
                    {
                        name = prop.Name;

                        if (prop.Value == null)
                        {
                            value = string.Empty;
                        }
                        else
                        {
                            value = prop.Value.ToString();
                        }
                    }
                    else if (jObject.First == null)
                    {
                        name = prop.Name;

                        if (prop.Value == null)
                        {
                            value = string.Empty;
                        }
                        else
                        {
                            value = prop.Value.ToString();
                        }
                    }
                    else
                    {
                        name = jObject.Property("human_name").Value.ToString();

                        if (jObject.Property("value").Value == null)
                        {
                            value = string.Empty;
                        }
                        else
                        {
                            value = jObject.Property("value").Value.ToString();
                        }
                    }

                    //bug in API? Should not return empty names
                    if (name!="") split[table.Columns.IndexOf(name)] = value;

                }

                table.Rows.Add(split);

            }

            return table;
        }

        public static DataTable GetEntry(string apiKey, string projectId, string jeguid)
        {
            var client = new RestClient(GlasshouseShared.Utils.urlApi);

            var request = new RestRequest(string.Format("projects/{0}/new_journal/entries/{1}", projectId,jeguid), Method.GET);

            //request.AddParameter("name", "value"); // adds to POST or URL querystring based on Method
            //request.AddUrlSegment("id", "123"); // replaces matching token in request.Resource

            // easily add HTTP Headers
            request.AddHeader("access-token", apiKey);

            //request.RequestFormat = DataFormat.Json;

            // execute the request
            IRestResponse response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK) return null;
            var content = response.Content; // raw content as string

            var journal = JsonConvert.DeserializeObject<dynamic>(content);

            // not done below this line !!!
            var converter = new Newtonsoft.Json.Converters.ExpandoObjectConverter();
            dynamic message = JsonConvert.DeserializeObject<System.Dynamic.ExpandoObject>(content, converter);

            dynamic new_journal = journal["new_journal"];

            dynamic entries = new_journal["entries"];


            DataTable table = new DataTable();

            foreach (dynamic entry in entries)
            {
                foreach (dynamic prop in entry)
                {
                    JObject jObject = prop.First as JObject;

                    string name = "";

                    if (jObject == null)
                    {
                        name = prop.Name;

                    }
                    else
                    {
                        name = jObject.Property("human_name").Value.ToString();
                    }

                    if (table.Columns.Contains(name) == false)
                    {
                        table.Columns.Add(name, typeof(string));
                    }
                }

                string[] split = new string[table.Columns.Count];
                for (int j = 0; j < split.Length; j++)
                {
                    split[j] = string.Empty;
                }

                foreach (dynamic prop in entry)
                {
                    JObject jObject = prop.First as JObject;

                    string name = "";
                    string value = "";

                    if (jObject == null)
                    {
                        name = prop.Name;

                        if (prop.Value == null)
                        {
                            value = string.Empty;
                        }
                        else
                        {
                            value = prop.Value.ToString();
                        }
                    }
                    else
                    {
                        name = jObject.Property("human_name").Value.ToString();

                        if (jObject.Property("value").Value == null)
                        {
                            value = string.Empty;
                        }
                        else
                        {
                            value = jObject.Property("value").Value.ToString();
                        }
                    }

                    split[table.Columns.IndexOf(name)] = value;

                }

                table.Rows.Add(split);

            }

            return table;
        }
    }
}
