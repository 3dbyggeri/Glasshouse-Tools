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
    public class Projects
    {

        /// <summary>
        /// Gets the projects.
        /// </summary>
        /// <param name="apiKey">The API key.</param>
        /// <returns></returns>
        [MultiReturn(new[] {"apikey", "id", "name", "created_at", "is_processing", "url", "invited"})]
        public static Dictionary<string, object> GetProjectsAllInfo(string apiKey)
        {

            Dictionary<string, object> dict = new Dictionary<string, object>();
            dict.Add("apikey", apiKey);

            foreach (var kvp in GlasshouseShared.Projects.GetProjectsAllInfo(apiKey))
                dict.Add(kvp.Key, kvp.Value);

            return dict;

        }

        /// <summary>
        /// Gets the projects.
        /// </summary>
        /// <param name="apiKey">The API key.</param>
        /// <returns></returns>
        [MultiReturn(new[] { "apikey", "id", "name" })]
        public static Dictionary<string, object> GetProjects(string apiKey)
        {

            Dictionary<string, object> dict = new Dictionary<string, object>();
            dict.Add("apikey", apiKey);

            foreach (var kvp in GlasshouseShared.Projects.GetProjects(apiKey))
                dict.Add(kvp.Key, kvp.Value);

            return dict;

        }


        /// <summary>
        /// Gets the project information.
        /// </summary>
        /// <param name="apiKey">The API key.</param>
        /// <param name="projectId">The project identifier.</param>
        /// <returns></returns>
        [MultiReturn(new[] { "apikey", "id", "name", "created_at", "is_processing", "url", "model_containers_url", "groupings_url", "selected", "collaborator_role" })]
        public static Dictionary<string, object> GetProjectInfo(string apiKey, string projectId)
        {
            Dictionary<string, object> dict = new Dictionary<string, object>();
            dict.Add("apikey", apiKey);

            foreach (var kvp in GlasshouseShared.Projects.GetProjectInfo(apiKey, projectId))
                dict.Add(kvp.Key, kvp.Value);

            return dict;
        }
    }
}


