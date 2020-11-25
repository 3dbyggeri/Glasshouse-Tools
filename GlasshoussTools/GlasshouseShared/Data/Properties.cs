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

namespace GlasshouseShared
{

    public class HumanNameTranslations
    {
        public object en { get; set; }
    }

    public class Property
    {
        public string name { get; set; }
        public string id { get; set; }
        public string original_name { get; set; }
        public object human_name { get; set; }
        public HumanNameTranslations human_name_translations { get; set; }
        public bool visible { get; set; }
        public string type { get; set; }
        public string unit { get; set; }
        public string description { get; set; }
        public string guid { get; set; }
        public string example_value { get; set; }
        public string reference_systems { get; set; }
        public string reference_url { get; set; }
        public string alternative_guid { get; set; }
        public string revit_guid { get; set; }
        public string revit_data_type { get; set; }
        public string revit_parameter_group { get; set; }
        public string revit_name { get; set; }
        public bool revit_is_instance { get; set; }
        public string archicad_property_name { get; set; }
        public string archicad_property_type { get; set; }
        public string archicad_value_type { get; set; }
        public string archicad_property_set { get; set; }
        public string archicad_bsdd_guid { get; set; }
    }

    public class PropertySet
    {
        public string id { get; set; }
        public string name { get; set; }
        public List<Property> properties { get; set; }
    }

    

    public class Metadata
    {
        public RevitSharedParameters revit_shared_parameters { get; set; }
        public RevitMappingFile revit_mapping_file { get; set; }
        public ArchicadPropertiesFile archicad_properties_file { get; set; }

        public class RevitSharedParameters
        {
            public string url { get; set; }
        }

        public class RevitMappingFile
        {
            public string url { get; set; }
        }

        public class ArchicadPropertiesFile
        {
            public string url { get; set; }
        }
    }

    public class Properties
    {
        public List<PropertySet> property_sets { get; set; }
        public Metadata metadata { get; set; }

        public static Properties GetProperties(string apiKey, string projectId)
        {
            var client = new RestClient(GlasshouseShared.Utils.urlApi);

            var request = new RestRequest(string.Format("projects/{0}/property_sets", projectId), Method.GET);
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

            return JsonConvert.DeserializeObject<Properties>(content);

            

        }

        public static Dictionary<string,string> GetPropertyMapping(string apiKey, string projectId)
        {
            Properties props= GetProperties(apiKey, projectId);

            Dictionary<string, string> ghrvt = new Dictionary<string, string>();

            foreach(PropertySet pset in props.property_sets)
            {
                foreach(Property prop in pset.properties)
                {
                    if (prop.revit_name == null) continue;
                    if (prop.revit_name.Equals("")) continue;

                    ghrvt.Add(prop.name, prop.revit_name); //mayby orignalname?
                }
            }
                                          
            return ghrvt;

        }

        public static List<string> GetPropertyNames(string apiKey, string projectId)
        {
            Properties props = GetProperties(apiKey, projectId);

            List<string> names = new List<string>();
            names.Add("GlassHouseJournalGUID");
            foreach (PropertySet pset in props.property_sets)
            {
                if (pset.id == null) continue; // we do not use unsorted!
                foreach (Property prop in pset.properties)
                {
                    names.Add(prop.original_name);
                }
            }

            return names;

        }

        public static bool AddPropertySet(string apiKey, string projectId, string propertySet)
        {
            var client = new RestClient(GlasshouseShared.Utils.urlApi);

            var request = new RestRequest(string.Format("projects/{0}/property_sets", projectId), Method.POST);
            request.AddHeader("access-token", apiKey);

            request.AddParameter("property_set[name]", propertySet);

            request.RequestFormat = DataFormat.Json;

            // execute the request
            IRestResponse response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK) return false;

            var content = response.Content; // raw content as string

            //return JsonConvert.DeserializeObject<cDocumentNChapters>(content);
            return true;
        }

        public static bool AddProperty(string apiKey, string projectId, string propertySet, string property)
        {
            var client = new RestClient(GlasshouseShared.Utils.urlApi);

            var request = new RestRequest(string.Format("projects/{0}/bim_property_refs", projectId), Method.POST);
            request.AddHeader("access-token", apiKey);

            request.AddParameter("bim_property_ref[property_set_name]", propertySet);
            request.AddParameter("bim_property_ref[original_name]", property);
            request.RequestFormat = DataFormat.Json;

            // execute the request
            IRestResponse response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK) return false;

            var content = response.Content; // raw content as string

            //return JsonConvert.DeserializeObject<cDocumentNChapters>(content);
            return true;
        }

    }
}
