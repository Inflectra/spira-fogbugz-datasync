using System;
using System.Collections;
using System.Collections.Generic;
using System.Net;
using System.Net.Security;
using System.IO;
using System.Xml;
using System.Web;
using System.Security;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace Inflectra.SpiraTest.PlugIns.FogBugzDataSync
{
    /// <summary>
    /// Provides an OO abstraction layer to access the FogBugz XML API
    /// </summary>
    public class FogBugz
    {
		protected bool verifyCertificate = true;
        protected bool enableKeepAlives = false;
        protected bool traceLogging = false;
		protected string url = "";
        protected string apiUrlSuffix = "";
        protected string token = "";
        protected DataSync dataSync = null;

        protected const int API_VERSION = 5; //The API Version this class was developed against

		/// <summary>
		/// Constructor
		/// </summary>
        public FogBugz(DataSync dataSync, bool traceLogging)
		{
            this.traceLogging = traceLogging;
            this.dataSync = dataSync;
		}

        /// <summary>
        /// Gets the current FogBugz API suffix
        /// </summary>
        public string ApiUrlSuffix
        {
            get
            {
                return this.apiUrlSuffix;
            }
        }

        /// <summary>
        /// Constructor with initial url
        /// </summary>
        /// <param name="url">The url of the FogBugz API</param>
        public FogBugz(string url)
        {
            //cookieContainer = new CookieContainer();
            this.url = url;
        }


        /// <summary>
        /// Whether to verify the SSL certificate hosted by the FogBugz instance
        /// </summary>
        /// <remarks>Set to False if you're using a self-signed certificate</remarks>
        public bool VerifyCertificate
        {
            get
            {
                return this.verifyCertificate;
            }
            set
            {
                this.verifyCertificate = value;
            }
        }

        /// <summary>
        /// Should we use HTTP 1.1 keep alive connections
        /// </summary>
        public bool EnableKeepAlives
        {
            get
            {
                return this.enableKeepAlives;
            }
            set
            {
                this.enableKeepAlives = value;
            }
        }

        /// <summary>
        /// Gets and Sets the URL to the FogBugz URL for accessing the XML API
        /// </summary>
        public string Url
        {
            get
            {
                return this.url;
            }
            set
            {
                this.url = value;
            }
        }

        /// <summary>
        /// Verifies that the API exists and is a compatible Version
        /// </summary>
        /// <returns>True if the API Version is compatible</returns>
        /// <remarks>Should be called before other methods</remarks>
        public bool VerifyApi()
        {
            //First we need need to hit the API with the api.xml document request
            XmlElement response = CallMethod("verifyapi", null);

            //Get the API Version and suffix
            XmlElement xmlVersion = (XmlElement)response.SelectSingleNode("version");
            int version = Int32.Parse(xmlVersion.InnerText);
            XmlElement xmlMinVersion = (XmlElement)response.SelectSingleNode("minversion");
            int minVersion = Int32.Parse(xmlMinVersion.InnerText);

            //Make sure that we support this API
            //We need to have an API Version >= 5 and a minimum Version <= 5
            if (version >= API_VERSION && minVersion <= API_VERSION)
            {
                XmlElement xmlApiUrlSuffix = (XmlElement)response.SelectSingleNode("url");
                this.apiUrlSuffix = xmlApiUrlSuffix.InnerText;
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Gets an existing case in the system
        /// </summary>
        /// <param name="id">The id of the case to retrieve</param>
        /// <returns>The case object</returns>
        public FogBugzCase GetCase(int id)
        {
            string searchString = id.ToString();
            List<FogBugzCase> fogBugzCases = this.Search(searchString);
            if (fogBugzCases.Count > 0)
            {
                return fogBugzCases[0];
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Gets a list of all the cases that have been changed since the specified start date
        /// </summary>
        /// <param name="startDate">The date after which any changed item gets returned</param>
        /// <param name="fogBugzProjectId">The id of the fogbugz project</param>
        /// <returns>The case list</returns>
        public List<FogBugzCase> GetNewCases(int fogBugzProjectId, DateTime startDate)
        {
            string searchString = "project:=" + fogBugzProjectId + " edited:\"" + startDate.ToShortDateString() + ".." + DateTime.Now.ToShortDateString() + "\"";
            return this.Search(searchString);
        }

        /// <summary>
        /// Gets an existing case in the system
        /// </summary>
        /// <param name="id">The id of the case to retrieve</param>
        /// <returns>The case object</returns>
        protected List<FogBugzCase> Search(string searchString)
        {
            //First create the list of items
            List<FogBugzCase> fogBugzCases = new List<FogBugzCase>();

            //Make sure we have a token
            if (this.token == "")
            {
                throw new FogBugzApiException("You need to logon to the API before calling this method");
            }

            //Now construct the various parameters
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("token", this.token);
            parameters.Add("q", searchString);
            parameters.Add("cols", "ixBug,fOpen,sTitle,sLatestTextSummary,ixProject,ixArea,ixPersonOpenedBy,ixPersonAssignedTo,ixStatus,ixPriority,ixFixFor,sVersion,sComputer,hrsCurrEst,ixCategory,dtClosed,dtDue,dtLastUpdated");

            //Finally call the method and get the response xml
            XmlElement response = CallMethod("search", parameters);

            //Extract the bug/case id and populate
            XmlNodeList xmlCases = (XmlNodeList)response.ChildNodes[0].ChildNodes;
            foreach (XmlElement xmlCase in xmlCases)
            {
                FogBugzCase fogBugzCase = new FogBugzCase();
                fogBugzCase.Id = Int32.Parse(xmlCase.Attributes["ixBug"].Value);
                fogBugzCase.Title = xmlCase.SelectSingleNode("sTitle").InnerText;
                fogBugzCase.Project = GetIntValue(xmlCase, "ixProject");
                fogBugzCase.Area = GetIntValue(xmlCase, "ixArea");
                fogBugzCase.FixFor = GetIntValue(xmlCase, "ixFixFor");
                fogBugzCase.Status = GetIntValue(xmlCase, "ixStatus");
                fogBugzCase.Category = GetIntValue(xmlCase, "ixCategory");
                fogBugzCase.PersonOpenedBy = GetIntValue(xmlCase, "ixPersonOpenedBy");
                fogBugzCase.PersonAssignedTo = GetIntValue(xmlCase, "ixPersonAssignedTo");
                fogBugzCase.Priority = GetIntValue(xmlCase, "ixPriority");
                fogBugzCase.Due = GetDateTimeValue(xmlCase, "dtDue");
                fogBugzCase.HrsCurrEst = GetIntValue(xmlCase, "hrsCurrEst");
                fogBugzCase.Version = xmlCase.SelectSingleNode("sVersion").InnerText;
                fogBugzCase.Computer = xmlCase.SelectSingleNode("sComputer").InnerText;
                fogBugzCase.Description = xmlCase.SelectSingleNode("sLatestTextSummary").InnerText;
                fogBugzCase.Closed = GetDateTimeValue(xmlCase, "dtClosed");
                fogBugzCase.LastUpdated = GetDateTimeValue(xmlCase, "dtLastUpdated");
                fogBugzCases.Add(fogBugzCase);
            }
            
            //Return the case
            return fogBugzCases;
        }

        /// <summary>
        /// Gets an existing fixFor in the system
        /// </summary>
        /// <param name="id">The id of the fixFor to retrieve</param>
        /// <returns>The fixFor object</returns>
        public FogBugzFixFor GetFixFor(int id)
        {
            //First create an empty instance to populate
            FogBugzFixFor fogBugzFixFor = new FogBugzFixFor();

            //Make sure we have a token
            if (this.token == "")
            {
                throw new FogBugzApiException("You need to logon to the API before calling this method");
            }

            //Now construct the various parameters
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("token", this.token);
            parameters.Add("ixFixFor", id);

            //Finally call the method and get the response xml
            XmlElement response = CallMethod("viewFixFor", parameters);

            //Extract the fixFor and populate
            if (response.ChildNodes.Count > 0)
            {
                XmlElement xmlFixFor = (XmlElement)response["fixfor"];
                if (xmlFixFor == null)
                {
                    return null;
                }
                else
                {
                    fogBugzFixFor.Id = GetIntValue(xmlFixFor, "ixFixFor");
                    fogBugzFixFor.Project = GetIntValue(xmlFixFor, "ixProject");
                    fogBugzFixFor.Name = xmlFixFor.SelectSingleNode("sFixFor").InnerText;
                    if (GetDateTimeValue(xmlFixFor, "dt").HasValue)
                    {
                        fogBugzFixFor.ReleaseDate = GetDateTimeValue(xmlFixFor, "dt").Value;
                    }

                    //Return the fixFor
                    return fogBugzFixFor;
                }
            }
            else
            {
                //No matching FixFors
                return null;
            }
        }

        /// <summary>
        /// Gets a date-time value safely
        /// </summary>
        /// <param name="childNodeName">The node containing the value we're interested in</param>
        /// <param name="xmlElement">The parent XML element</param>
        /// <returns>A datetime or null</returns>
        DateTime? GetDateTimeValue(XmlElement xmlElement, string childNodeName)
        {
            //Make sure we have a child node with this 
            if (xmlElement.SelectSingleNode(childNodeName) != null)
            {
                string textValue = xmlElement.SelectSingleNode(childNodeName).InnerText;

                DateTime returnValue;
                if (DateTime.TryParse(textValue, out returnValue))
                {
                    return returnValue;
                }
                else
                {
                    return null;
                }
            }
            return null;
        }

        /// <summary>
        /// Gets an integer value safely
        /// </summary>
        /// <param name="childNodeName">The node containing the value we're interested in</param>
        /// <param name="xmlElement">The parent XML element</param>
        /// <returns>An integer or -1 if not found</returns>
        int GetIntValue(XmlElement xmlElement, string childNodeName)
        {
            //Make sure we have a child node with this 
            if (xmlElement.SelectSingleNode(childNodeName) != null)
            {
                string textValue = xmlElement.SelectSingleNode(childNodeName).InnerText;

                int returnValue;
                if (Int32.TryParse(textValue, out returnValue))
                {
                    return returnValue;
                }
                else
                {
                    return -1;
                }
            }
            return -1;
        }

        /// <summary>
        /// Adds a new fixfor and populates the ID field
        /// </summary>
        /// <param name="fixFor"></param>
        /// <returns></returns>
        public FogBugzFixFor AddFixFor(FogBugzFixFor fixFor)
        {
            //Make sure we have a token
            if (this.token == "")
            {
                throw new FogBugzApiException("You need to logon to the API before calling this method");
            }

            //Now construct the various parameters
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("token", this.token);
            parameters.Add("ixProject", fixFor.Project);
            parameters.Add("sFixFor", fixFor.Name);
            parameters.Add("dtRelease", fixFor.ReleaseDate);
            parameters.Add("fAssignable", fixFor.Assignable);

            //Finally call the method and get the id back
            XmlElement response = CallMethod("newFixFor", parameters);

            //Extract the fix-for id and populate
            XmlElement xmlNewCase = (XmlElement)response.SelectSingleNode("fixFor");
            fixFor.Id = Int32.Parse(xmlNewCase.Attributes["ixFixFor"].Value);

            return fixFor;
        }

        /// <summary>
        /// Adds the newly created case to the system and adds the ID to the class
        /// </summary>
        /// <returns>The case with the id populated</returns>
        public FogBugzCase Add(FogBugzCase fogBugzCase)
        {
            //Make sure we have a token
            if (this.token == "")
            {
                throw new FogBugzApiException("You need to logon to the API before calling this method");
            }

            //Now construct the various parameters
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("token", this.token);
            parameters.Add("sTitle", fogBugzCase.Title);
            parameters.Add("ixProject", fogBugzCase.Project);
            parameters.Add("ixArea", fogBugzCase.Area);
            parameters.Add("ixFixFor", fogBugzCase.FixFor);
            parameters.Add("ixCategory", fogBugzCase.Category);
            parameters.Add("ixPersonAssignedTo", fogBugzCase.PersonAssignedTo);
            parameters.Add("ixPriority", fogBugzCase.Priority);
            parameters.Add("dtDue", fogBugzCase.Due);
            parameters.Add("hrsCurrEst", fogBugzCase.HrsCurrEst);
            parameters.Add("sVersion", fogBugzCase.Version);
            parameters.Add("sComputer", fogBugzCase.Computer);
            parameters.Add("sEvent", fogBugzCase.Description);

            //Finally call the method and get the id back
            XmlElement response = CallMethod("new", parameters);

            //Extract the bug/case id and populate
            XmlElement xmlNewCase = (XmlElement)response.SelectSingleNode("case");
            fogBugzCase.Id = Int32.Parse(xmlNewCase.Attributes["ixBug"].Value);
            return fogBugzCase;
        }

        /// <summary>
        /// Logs-off the user. Does nothing if there is no user logged in.
        /// </summary>
        public void Logoff()
        {
            //Create the parameter list
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("token", this.token);

            //Call the logoff command
            XmlElement response = CallMethod("logoff", parameters);
        }

        /// <summary>
        /// Logging on, with a username and password is required for searching for cases,
        /// posting new cases, etc. This method logs in a user.
        /// </summary>
        /// <param name="email">The user's email address</param>
        /// <param name="password">The user's password</param>
        /// <remarks>Throws an exception if the logon fails</remarks>
        public void Logon(string email, string password)
        {
            //Create the parameter list
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("email", email);
            parameters.Add("password", password);

            //Call the logon command and get the session token
            XmlElement response = CallMethod("logon", parameters);
            XmlElement xmlToken = (XmlElement)response.SelectSingleNode("token");
            this.token = xmlToken.InnerText;
        }

        /// <summary>
        /// Calls a specific API method
        /// </summary>
        /// <param name="cmdName">The command name</param>
        /// <param name="parameters">Any provided parameters</param>
        /// <returns>The XML response element</returns>
        protected XmlElement CallMethod(string cmdName, Dictionary<string, object> parameters)
        {
            //Handle the case where there are certificate errors (e.g. self-signed SSL certs)
            ServicePointManager.ServerCertificateValidationCallback = new System.Net.Security.RemoteCertificateValidationCallback(certificateValidation_Handler);
            
            //Create the full URL to the API
            string url = this.url;
            if (cmdName == "verifyapi")
            {
                //In this case we always add on the api.xml suffix since we don't know the real API suffix yet
                url += "/api.xml";
            }
            else
            {
                url += "/" + this.apiUrlSuffix;

                //Add the command name as a parameter
                parameters.Add("cmd", cmdName);
            }

            //Create a new web request and add the parameters as form GET elements if provided
            string parameterClause = "";
            if (parameters != null)
            {
                foreach (KeyValuePair<string, object> parameter in parameters)
                {
                    //Handle the different possible data-types
                    //Don't add parameters with empty values (-1, "") or null values
                    if (parameter.Value != null)
                    {
                        if (parameter.Value.GetType() == typeof(string))
                        {
                            string parameterValue = HttpUtility.UrlEncode((string)parameter.Value);
                            if (parameterValue != "")
                            {
                                if (parameterClause != "")
                                {
                                    parameterClause += "&";
                                }
                                parameterClause += parameter.Key + "=" + parameter.Value;
                            }
                        }
                        if (parameter.Value.GetType() == typeof(int))
                        {
                            if ((int)parameter.Value != -1)
                            {
                                string parameterValue = ((int)parameter.Value).ToString();
                                if (parameterClause != "")
                                {
                                    parameterClause += "&";
                                }
                                parameterClause += parameter.Key + "=" + parameter.Value;
                            }
                        }
                        if (parameter.Value.GetType() == typeof(DateTime))
                        {
                            string parameterValue = ((DateTime)parameter.Value).ToString("yyyyMMddTHH:mm:ss");
                            if (parameterClause != "")
                            {
                                parameterClause += "&";
                            }
                            parameterClause += parameter.Key + "=" + parameter.Value;
                        }
                    }
                }
                url += parameterClause;
            }

            HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(url);

            //Now send the web post and get the response
			webRequest.Method = "GET";
			webRequest.ContentType = "text/html";
			//Set timeouts for 20 minutes
			webRequest.Timeout = 1200000;
			webRequest.ReadWriteTimeout = 1200000;
            webRequest.KeepAlive = this.EnableKeepAlives;

            //Log out the request
            this.dataSync.LogTraceEvent("Request - URL: " + url, System.Diagnostics.EventLogEntryType.Information);

			//Get the response
			string result;
			HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();
			using (StreamReader sr = new StreamReader(webResponse.GetResponseStream()) )
			{
				result = sr.ReadToEnd();

				// Close and clean up the StreamReader
				sr.Close();
			}

			//Now load the result into an XML Document
			XmlDocument xmlDoc = new XmlDocument();
			try
			{
				xmlDoc.LoadXml (result);
			}
			catch (Exception)
			{
				//Log a more meaningful exception
				throw new Exception ("Unable to load in the method response XML: " + result);
			}
			//Get the command response node
			XmlNodeList xmlNodeList = xmlDoc.GetElementsByTagName ("response");
			if (xmlNodeList.Count < 1)
			{
                throw new Exception("Unable to get response element");
			}
			XmlElement xmlMethodResponse = (XmlElement) xmlNodeList[0];

			//Now see if we have any errors
			XmlNodeList xmlErrorList = xmlDoc.GetElementsByTagName ("error");
            if (xmlErrorList.Count > 0)
            {
                string errorCode = "(unknown)";
                if (xmlErrorList[0].Attributes["code"] != null)
                {
                    errorCode = xmlErrorList[0].Attributes["code"].Value;
                }
                throw new FogBugzApiException("Error code " + errorCode + " returned from FogBugz API: " + xmlErrorList[0].InnerText);
            }

            return xmlMethodResponse;
        }

        /// <summary>
        /// Determine whether we are allowed to ignore SSL certificate errors
        /// </summary>
        /// <param name="sender">An object that contains state information for this validation</param>
        /// <param name="cert">The certificate used to authenticate the remote party</param>
        /// <param name="chain">The chain of certificate authorities associated with the remote certificate</param>
        /// <param name="Errors">The chain of certificate authorities associated with the remote certificate</param>
        /// <returns>True if we are to ignore the error</returns>
        protected bool certificateValidation_Handler(object sender, X509Certificate cert, X509Chain chain, SslPolicyErrors Errors)
        {
            //See if we supposed to ignore invalid SSL certificates (used when we have self-signed SSL certs)
            if (this.verifyCertificate)
            {
                return false;
            }
            else
            {
                //Ignore any errors
                return true;
            }
        }
    }

    /// <summary>
    /// This exception is thrown when you try and insert a test case
    /// parameter with a name that is already in use for that test case
    /// </summary>
    public class FogBugzApiException : ApplicationException
    {
        public FogBugzApiException()
        {
        }
        public FogBugzApiException(string message)
            : base(message)
        {
        }
        public FogBugzApiException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}
