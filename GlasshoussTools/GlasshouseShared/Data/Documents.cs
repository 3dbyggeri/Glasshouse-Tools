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

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GlasshouseShared
{
    public class Documents
    {
        public static cDocuments GetDocuments(string apiKey, string projectId)
        {
            var client = new RestClient(GlasshouseShared.Utils.urlApi);

            var request = new RestRequest(string.Format("projects/{0}/descriptions.json?without_descriptions=true", projectId), Method.GET);

            // easily add HTTP Headers
            request.AddHeader("access-token", apiKey);

            request.RequestFormat = DataFormat.Json;

            // execute the request
            IRestResponse response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK) return null;

            var content = response.Content; // raw content as string

            cDocuments documents = JsonConvert.DeserializeObject<cDocuments>(content);

            return documents;
        }

        public static cChapter GetChapter(string apiKey, string projectId, string chapterId)
        {
            var client = new RestClient(GlasshouseShared.Utils.urlApi);

            var request = new RestRequest(string.Format("projects/{0}/descriptions.json?without_descriptions=false", projectId), Method.GET);

            // easily add HTTP Headers
            request.AddHeader("access-token", apiKey);

            request.RequestFormat = DataFormat.Json;

            // execute the request
            IRestResponse response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK) return null;

            var content = response.Content; // raw content as string
            // not done below this line !!!
            //var chapter = desc.First(s => s["id"].Equals(chapterId));
            //cDocuments documents = JsonConvert.DeserializeObject<cDocuments>(content);

            return null;
        }

        public static cChapter GetChapterJE(string apiKey, string projectId, string chapterId)
        {
            var client = new RestClient(GlasshouseShared.Utils.urlApi);

            var request = new RestRequest(string.Format("projects/{0}/descriptions.json?without_descriptions=false", projectId), Method.GET);

            // easily add HTTP Headers
            request.AddHeader("access-token", apiKey);

            request.RequestFormat = DataFormat.Json;

            // execute the request
            IRestResponse response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK) return null;

            var content = response.Content; // raw content as string

            JObject rss = JObject.Parse(content);
            List<JToken> jTokens = new List<JToken>();

            var desc = rss["documents"].SelectMany(d => d["descriptions"]);
            foreach (var d in desc)
            {
                JToken token = d as JToken;
                string name = (string)token.SelectToken("id");
                if (name.Equals(chapterId))
                {
                    jTokens =  token.SelectToken("new_journal_entry_ids").ToList();
                    break;
                }
            }

            List<string> ret = new List<string>();
            foreach(JToken token in jTokens)
            {
                ret.Add(token.ToString());
            }

            return null;
        }

        public static cDocumentNChapters GetDocumentNChapters(string apiKey, string projectId, string documentId)
        {
            var client = new RestClient(GlasshouseShared.Utils.urlApi);

            var request = new RestRequest(string.Format("projects/{0}/documents/{1}?IncludeMetaData_LinkedFileLocation=true", projectId,documentId), Method.GET);

            // easily add HTTP Headers
            request.AddHeader("access-token", apiKey);
            //request.AddParameter("IncludeMetaData_LinkedFileLocation", true);
            
            request.RequestFormat = DataFormat.Json;

            // execute the request
            IRestResponse response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK) return null;

            var content = response.Content; // raw content as string

            cDocumentNChapters document = JsonConvert.DeserializeObject<cDocumentNChapters>(content);

            return document;
        }


        public static cDocumentNChapters CreateDocument(string apiKey, string projectId, string documentName)
        {
            var client = new RestClient(GlasshouseShared.Utils.urlApi);

            var request = new RestRequest(string.Format("projects/{0}/documents.json", projectId), Method.POST);

            // easily add HTTP Headers
            request.AddHeader("access-token", apiKey);

            
            
            request.AddParameter("document[name]", documentName);

            request.RequestFormat = DataFormat.Json;



            // execute the request
            IRestResponse response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK) return null;

            var content = response.Content; // raw content as string

            return JsonConvert.DeserializeObject<cDocumentNChapters>(content);
        }

        public static void UpdateDocumentPostion(string apiKey, string projectId, string documentId,int postion)
        {
            var client = new RestClient(GlasshouseShared.Utils.urlApi);

            var request = new RestRequest(string.Format("projects/{0}/documents/{1}", projectId, documentId), Method.PUT);

            // easily add HTTP Headers
            request.AddHeader("access-token", apiKey);

            request.AddParameter("document[position]", postion);

            request.RequestFormat = DataFormat.Json;

            // execute the request
            IRestResponse response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK) return;

            var content = response.Content; // raw content as string
        }

        public static void UpdateChapterPostion(string apiKey, string projectId, string chapterId, int postion)
        {
            var client = new RestClient(GlasshouseShared.Utils.urlApi);

            var request = new RestRequest(string.Format("projects/{0}/descriptions/{1}", projectId, chapterId), Method.PUT);

            // easily add HTTP Headers
            request.AddHeader("access-token", apiKey);

            request.AddParameter("description[position]", postion);
            
            request.RequestFormat = DataFormat.Json;

            // execute the request
            IRestResponse response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK) return;

            var content = response.Content; // raw content as string
        }

        public static void UpdateChapterPath(string apiKey, string projectId, string chapterId, string fullpath)
        {
            var client = new RestClient(GlasshouseShared.Utils.urlApi);

            var request = new RestRequest(string.Format("projects/{0}/descriptions/{1}", projectId, chapterId), Method.PUT);

            // easily add HTTP Headers
            request.AddHeader("access-token", apiKey);

            request.AddParameter("description[linked_file_location]", fullpath);

            request.RequestFormat = DataFormat.Json;

            // execute the request
            IRestResponse response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK) return;

            var content = response.Content; // raw content as string
        }


        public static string CreateChapter(string apiKey, string projectId, string documentId, string chapterName, int postion, string path, string filename)
        {
            var client = new RestClient(GlasshouseShared.Utils.urlApi);

            var request = new RestRequest(string.Format("projects/{0}/descriptions.json", projectId), Method.POST);

            // easily add HTTP Headers
            request.AddHeader("access-token", apiKey);
                       
            request.AddParameter("description[name]", chapterName);
            

            request.RequestFormat = DataFormat.Json;
                       
            // execute the request
            IRestResponse response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK) return null;

            var content = response.Content; // raw content as string

            cChapter chapter= JsonConvert.DeserializeObject<cChapter>(content);

            if (chapter != null && filename.Length>4)
            {
                ConnectChapter(apiKey, projectId, documentId, chapter.description.id, postion,Path.Combine(path,filename));
            }

            return chapter.description.id;

        }

        public static void ConnectChapter(string apiKey, string projectId, string documentId, string chapterId, int postion, string path)
        {
            var client = new RestClient(GlasshouseShared.Utils.urlApi);

            var request = new RestRequest(string.Format("projects/{0}/descriptions/{1}", projectId,chapterId), Method.PUT);

            // easily add HTTP Headers
            request.AddHeader("access-token", apiKey);

            request.AddParameter("description[document_id]", documentId);
            request.AddParameter("description[linked_file_location]", path);
            if (postion > 0)
            {
                request.AddParameter("description[position]", postion);
            }
            request.RequestFormat = DataFormat.Json;

            // execute the request
            IRestResponse response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK) return;

            var content = response.Content; // raw content as string
        }

        public static void UpdateRefsChapter(string apiKey, string projectId, string chapterId, List<string> refsChaptersId)
        {
            var client = new RestClient(GlasshouseShared.Utils.urlApi);

            var request = new RestRequest(string.Format("projects/{0}/descriptions/{1}", projectId, chapterId), Method.PUT);

            // easily add HTTP Headers
            request.AddHeader("access-token", apiKey);

            //PUT http://bskriver-qa.herokuapp.com/api/v1/projects/<project-id>/descriptions/<description-id>
            //description[update_reference_specs]=1
            //description[spec_ids]=1
            request.AddParameter("description[update_reference_specs]", 1);
            // entriesId: [0]="" - clears entry list
            foreach (string id in refsChaptersId)
            {
                request.AddParameter("description[spec_ids][]", id);
            }

            request.RequestFormat = DataFormat.Json;

            // execute the request
            IRestResponse response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK) return;

            var content = response.Content; // raw content as string
        }

        public static void ConnectJournalEntriesChapter(string apiKey, string projectId, string chapterId, List<string> entriesId)
        {
            var client = new RestClient(GlasshouseShared.Utils.urlApi);

            var request = new RestRequest(string.Format("projects/{0}/descriptions/{1}", projectId, chapterId), Method.PUT);

            // easily add HTTP Headers
            request.AddHeader("access-token", apiKey);
            // entriesId: [0]="" - clears entry list
            foreach (string id in entriesId)
            {
                request.AddParameter("description[new_journal_entry_ids][]", id);
            }
            request.RequestFormat = DataFormat.Json;

            // execute the request
            IRestResponse response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK) return;

            var content = response.Content; // raw content as string
        }

        public static void UpdateDocumentBimProperties(string apiKey, string projectId, string documentId, Dictionary<string, string> properties)
        {
            var client = new RestClient(GlasshouseShared.Utils.urlApi);

            var request = new RestRequest(string.Format("projects/{0}/documents/{1}", projectId, documentId), Method.PUT);

            // easily add HTTP Headers
            request.AddHeader("access-token", apiKey);
            request.AlwaysMultipartFormData = true;

            foreach (KeyValuePair<string, string> kvp in properties)
            {
                request.AddParameter("document[bim_property_values_attributes][][human_name]", kvp.Key);
                request.AddParameter("document[bim_property_values_attributes][][value]", kvp.Value);
            }
            request.RequestFormat = DataFormat.Json;

            // execute the request
            IRestResponse response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK)
            {
                return;
            }
            var content = response.Content; // raw content as string
        }

        public static void UpdateChapterBimProperties(string apiKey, string projectId, string chapterId, Dictionary<string, string> properties)
        {
            var client = new RestClient(GlasshouseShared.Utils.urlApi);

            var request = new RestRequest(string.Format("projects/{0}/descriptions/{1}", projectId, chapterId), Method.PUT);

            // easily add HTTP Headers
            request.AddHeader("access-token", apiKey);
            request.AlwaysMultipartFormData = true;

            foreach (KeyValuePair<string, string> kvp in properties)
            {
                request.AddParameter("description[bim_property_values_attributes][][human_name]", kvp.Key);
                request.AddParameter("description[bim_property_values_attributes][][value]", kvp.Value);
            }
            request.RequestFormat = DataFormat.Json;

            // execute the request
            IRestResponse response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK)
            {
                return;
            }
            var content = response.Content; // raw content as string
        }


        public class cDocument
        {
            public string id { get; set; }
            public string name { get; set; }
            public int position { get; set; }
            public string descriptions_url { get; set; }
            public bool? @default { get; set; }
            public object bim_property_values { get; set; }
        }

        public class cDocuments
        {
            public List<cDocument> documents { get; set; }
        }

        //
        public class cDocChapter
        {
            public string id { get; set; }
            public string document_id { get; set; } // from create
            public string document_name { get; set; }
            public string number { get; set; }
            public string parent_id { get; set; }
            public int position { get; set; }
            public string name { get; set; }
            public object grouping_property { get; set; }
            public List<object> new_journal_entry_ids { get; set; }
            public string journal_entries { get; set; }
            public List<object> reference_specs { get; set; }
            public List<object> descriptions { get; set; }
            public string linked_file_location { get; set; } // from create
            public string guid { get; set; } // from create
            public object bim_property_values { get; set; }
        }

        public class cBimPropertyValues
        {
            public object cBimPropertyValue;
        }

        public class cBimPropertyValue
        {
            [JsonProperty("id")]
            public long Id { get; set; }

            [JsonProperty("project_id")]
            public string ProjectId { get; set; }

            [JsonProperty("propertyable_id")]
            public string PropertyableId { get; set; }

            [JsonProperty("propertyable_type")]
            public string PropertyableType { get; set; }

            [JsonProperty("system_name")]
            public string SystemName { get; set; }

            [JsonProperty("human_name")]
            public string HumanName { get; set; }

            [JsonProperty("value")]
            public string Value { get; set; }

            [JsonProperty("extra_data")]
            public object ExtraData { get; set; }

            [JsonProperty("created_at")]
            public DateTimeOffset CreatedAt { get; set; }

            [JsonProperty("updated_at")]
            public DateTimeOffset UpdatedAt { get; set; }
        }

        public class cDocInFo
        {
            public string id { get; set; }
            public string name { get; set; }
            public object spec_standard_xml_file { get; set; }
            public bool numbering { get; set; }
            public int first_page_number { get; set; }
            public bool first_description_as_cover { get; set; }
            public int position { get; set; }
            public List<cDocChapter> descriptions { get; set; }
        }

        public class cDocumentNChapters
        {
            public cDocInFo document { get; set; }
        }

        //
        public class cChapter
        {
            public cDocChapter description { get; set; }
        }
    }
}
