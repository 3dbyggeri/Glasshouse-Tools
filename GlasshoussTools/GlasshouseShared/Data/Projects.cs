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


namespace GlasshouseShared
{
    public class Projects
    {
        public static cTheProjects _project = null;

        /// <summary>
        /// Gets the projects.
        /// </summary>
        /// <param name="apiKey">The API key.</param>
        /// <returns></returns>
        public static Dictionary<string, object> GetProjectsAllInfo(string apiKey)
        {

            var client = new RestClient(GlasshouseShared.Utils.urlApi);

            var request = new RestRequest("projects", Method.GET);


            // easily add HTTP Headers
            request.AddHeader("access-token", apiKey);

            request.RequestFormat = DataFormat.Json;

            // execute the request
            IRestResponse response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK) return null;

            var content = response.Content; // raw content as string

            cProjects projects = JsonConvert.DeserializeObject<cProjects>(content);

            //var defaultStrings = (new int[10]).Select(x => "my value").ToList();

            return new Dictionary<string, object>()
                {
                    { "id", projects.projects.Select(x => x.id).Concat(projects.invited_projects.Select(x => x.id)).ToList()},
                    { "name",projects.projects.Select(x => x.name).Concat(projects.invited_projects.Select(x => x.name)).ToList() },
                    { "created_at", projects.projects.Select(x => x.created_at.ToString()).Concat(projects.invited_projects.Select(x => x.created_at.ToString())).ToList()},
                    { "is_processing", projects.projects.Select(x => x.is_processing.ToString()).Concat(projects.invited_projects.Select(x => x.is_processing.ToString())).ToList()},
                    { "url", projects.projects.Select(x => x.url).Concat(projects.invited_projects.Select(x => x.url)).ToList()},
                    { "invited",  (new string[projects.projects.Count()]).Select(x => "False").Concat(projects.invited_projects.Select(x => "True")).ToList() }
                };
        }

        /// <summary>
        /// Gets the projects.
        /// </summary>
        /// <param name="apiKey">The API key.</param>
        /// <returns></returns>
        public static Dictionary<string, object> GetProjects(string apiKey)
        {

            var client = new RestClient(GlasshouseShared.Utils.urlApi);

            var request = new RestRequest("projects", Method.GET);


            // easily add HTTP Headers
            request.AddHeader("access-token", apiKey);

            request.RequestFormat = DataFormat.Json;

            // execute the request
            IRestResponse response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK) return null;

            var content = response.Content; // raw content as string

            cProjects projects = JsonConvert.DeserializeObject<cProjects>(content);

            //var defaultStrings = (new int[10]).Select(x => "my value").ToList();

            Dictionary<string, string> dict = new Dictionary<string, string>();

            foreach (cProject p in projects.projects)
            {
                dict.Add(p.id, p.name);
            }

            foreach (cInvitedProject p in projects.invited_projects)
            {
                if (dict.ContainsKey(p.id)==false) dict.Add(p.id, p.name);
            }


            var items = from pair in dict
                        orderby pair.Value ascending
                        select pair;


            return new Dictionary<string, object>()
                {
                //{ "names",names },
                //{ "ids", ids}
               
                    { "id",items.Select(x => x.Key).ToList()},
                    { "name",items.Select(x => x.Value).ToList() },
                };
        }

        /* test
        public class GetProjects
        {
            public static Projects AsProjects(string apikey)
            {
                var client = new RestClient("https://app.glasshousebim.com/api/v1");

                var request = new RestRequest("projects", Method.GET);

                // easily add HTTP Headers
                request.AddHeader("access-token", apikey);

                request.RequestFormat = DataFormat.Json;

                // execute the request
                IRestResponse response = client.Execute(request);
                var content = response.Content; // raw content as string

                return JsonConvert.DeserializeObject<Projects>(content);
            }

            /// <summary>
            /// Gets the projects.
            /// </summary>
            /// <param name="apiKey">The API key.</param>
            /// <returns></returns>
            public static Dictionary<string, object> AsDictionary(string apiKey)
            {

                var client = new RestClient(GlasshouseShared.Utils.urlApi);

                var request = new RestRequest("projects", Method.GET);


                // easily add HTTP Headers
                request.AddHeader("access-token", apiKey);

                request.RequestFormat = DataFormat.Json;

                // execute the request
                IRestResponse response = client.Execute(request);
                if (response.StatusCode != System.Net.HttpStatusCode.OK) return null;

                var content = response.Content; // raw content as string

                cProjects projects = JsonConvert.DeserializeObject<cProjects>(content);

                //var defaultStrings = (new int[10]).Select(x => "my value").ToList();

                Dictionary<string, string> dict = new Dictionary<string, string>();

                foreach (cProject p in projects.projects)
                {
                    dict.Add(p.id, p.name);
                }

                foreach (cInvitedProject p in projects.invited_projects)
                {
                    dict.Add(p.id, p.name);
                }


                var items = from pair in dict
                            orderby pair.Value ascending
                            select pair;


                return new Dictionary<string, object>()
                {
                //{ "names",names },
                //{ "ids", ids}
               
                    { "id",items.Select(x => x.Key).ToList()},
                    { "name",items.Select(x => x.Value).ToList() },
                };
            }
        }

        */

        /// <summary>
        /// Gets the project information.
        /// </summary>
        /// <param name="apiKey">The API key.</param>
        /// <param name="projectId">The project identifier.</param>
        /// <returns></returns>
        public static Dictionary<string, object> GetProjectInfo(string apiKey, string projectId)
        {
            FetchProjectInfo(apiKey, projectId);

            return new Dictionary<string, object>()
            {

                { "id",_project.project.id },
                { "name",_project.project.name },
                { "created_at", _project.project.created_at.ToString()},
                { "is_processing", _project.project.is_processing.ToString()},
                { "url", _project.project.url},
                { "model_containers_url", _project.project.model_containers_url},
                { "groupings_url", _project.project.groupings_url},
                { "selected", _project.project.selected.ToString()},
                { "collaborator_role", _project.project.collaborator_role}
            };
        }

        public static void FetchProjectInfo(string apiKey, string projectId)
        {
            var client = new RestClient(GlasshouseShared.Utils.urlApi);

            var request = new RestRequest(string.Format("projects/{0}", projectId), Method.GET);

            // easily add HTTP Headers
            request.AddHeader("access-token", apiKey);

            request.RequestFormat = DataFormat.Json;

            // execute the request
            IRestResponse response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK)
            {
                _project = null;
                return;
            }
            var content = response.Content; // raw content as string

            //content = content.Replace("selected?", "selected");
            _project = JsonConvert.DeserializeObject<cTheProjects>(content);
        }

        public static void UpdateBimProperties(string apiKey, string projectId, Dictionary<string,string> properties)
        {
            var client = new RestClient(GlasshouseShared.Utils.urlApi);

            var request = new RestRequest(string.Format("projects/{0}.json", projectId), Method.PUT);

            // easily add HTTP Headers
            request.AddHeader("access-token", apiKey);
            request.AlwaysMultipartFormData = true;

            foreach (KeyValuePair<string, string> kvp in properties)
            {
                request.AddParameter("project[bim_property_values_attributes][][human_name]", kvp.Key);
                request.AddParameter("project[bim_property_values_attributes][][value]", kvp.Value);
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



        //http://json2csharp.com/

        public class cProject
        {
            public string id { get; set; }
            public string name { get; set; }
            public DateTime created_at { get; set; }
            public bool is_processing { get; set; }
            public string url { get; set; }
        }


        public class cInvitedProject
        {
            public string id { get; set; }
            public string name { get; set; }
            public DateTime created_at { get; set; }
            public bool is_processing { get; set; }
            public string url { get; set; }
        }

        public class cProjects
        {
            public List<cProject> projects { get; set; }
            public List<cInvitedProject> invited_projects { get; set; }
        }

        //

        public class cTheProject
        {
            public string id { get; set; }
            public string name { get; set; }
            public DateTime created_at { get; set; }
            public bool is_processing { get; set; }
            public string url { get; set; }
            public string model_containers_url { get; set; }
            public string groupings_url { get; set; }
            [JsonProperty(PropertyName = "selected?")]
            public bool selected { get; set; }
            public string collaborator_role { get; set; }
            public string client { get; set; }
            public string location { get; set; }
            public string size { get; set; }
            public string architect { get; set; }
            public string engineer { get; set; }
            public string contractor { get; set; }
            //public cBimPropertyValues bim_property_values { get; set; }
            public object bim_property_values { get; set; }
        }

        public class cBimPropertyValues
        {
            public List<cBimPropertyValue> cBimPropertyValue;
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

        public class cTheProjects
        {
            public cTheProject project { get; set; }
        }

    }
}


