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
using System.Text;
using System.Globalization;
using RestSharp;
using Newtonsoft.Json;
using System.Windows.Forms;


namespace GlasshouseShared
{
    public class Utils
    {
        public const string urlApi = "https://app.glasshousebim.com/api/v1";
        public static string apiKey = "Login"; //

        private readonly static string reservedCharacters = "!*'();:@&=+$,/?%#[]";

        public static string UrlEncode(string value)
        {
            if (String.IsNullOrEmpty(value))
                return String.Empty;

            var sb = new StringBuilder();

            foreach (char @char in value)
            {
                if (reservedCharacters.IndexOf(@char) == -1)
                    sb.Append(@char);
                else
                    sb.AppendFormat("%{0:X2}", (int)@char);
            }
            return sb.ToString();
        }

        public static bool LogInGlassHouse(string email, string password)
        {
            bool loggedIn = false;
            try
            {
                var client = new RestClient(GlasshouseShared.Utils.urlApi);

                if (null == client)
                {
                    MessageBox.Show("'RestClient' failed!");
                    return loggedIn;
                }

                email = UrlEncode(email);
                password = UrlEncode(password);

                string strRelativePath = string.Format("users/sign_in.json?email={0}&password={1}", email, password);

                var request = new RestRequest(strRelativePath, Method.POST);

                request.RequestFormat = DataFormat.Json;
                //request.Timeout = 

                // execute the request
                IRestResponse response = client.Execute(request);
                var content = response.Content; // raw content as string

                cLoginUserProfile profileJson = JsonConvert.DeserializeObject<cLoginUserProfile>(content);


                if (null != profileJson && null != profileJson.user && !string.IsNullOrEmpty(profileJson.user.api_key))
                {
                    Utils.apiKey = profileJson.user.api_key;
                    loggedIn = true;
                }
                else
                {
                    MessageBox.Show("'Response' failed to get " + response.ResponseUri + "\nError is " + response.StatusCode + " : " + response.ErrorMessage + "\nUsername:'" + email + "'\nPassword:'" + password + "'\nContent:\n" + content);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(email + " failed to log in. Please check user name and password and try again.\nError:\n" + ex.Message);
                return loggedIn;
            }
            return loggedIn;
        }


        /// <summary>
        /// Formats a string to an invariant culture
        /// </summary>
        /// <param name="formatString">The format string.</param>
        /// <param name="objects">The objects.</param>
        /// <returns></returns>
        public static string FormatInvariant(string formatString, params object[] objects)
        {
            return string.Format(CultureInfo.InvariantCulture, formatString, objects);
        }

        /// <summary>
        /// Formats a string to the current culture.
        /// </summary>
        /// <param name="formatString">The format string.</param>
        /// <param name="objects">The objects.</param>
        /// <returns></returns>
        public static string FormatCurrentCulture(string formatString, params object[] objects)
        {
            return string.Format(CultureInfo.CurrentCulture, formatString, objects);
        }

        public static string FormatEN_USCulture(string formatString, params object[] objects)
        {
            return string.Format(System.Globalization.CultureInfo.GetCultureInfo("en-US"), formatString, objects);
        }

        public static string FormatWithQuotes(string val)
        {
            //string decimalSeparator = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
            //string listSeparator = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator;

            if (val.Contains(","))
            {
                return string.Format("\"{0}\"", val);
            }
            return val;
        }

        
    }

    public class LoginUser
    {
        public string api_key { get; set; }
    }

    public class cLoginUserProfile
    {
        public LoginUser user { get; set; }
        public string access_token { get; set; }
        public string token_type { get; set; }
    }
}
