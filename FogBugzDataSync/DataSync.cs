using System;
using System.Configuration;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using System.Xml;
using System.Web;
using System.Net;
using System.Web.Services;
using System.Web.Services.Protocols;

namespace Inflectra.SpiraTest.PlugIns.FogBugzDataSync
{
    /// <summary>
    /// Contains all the logic necessary to sync SpiraTest incidents with FogBugz
    /// </summary>
    public class DataSync : IServicePlugIn
    {
        //Constant containing data-sync name
        protected const string DATA_SYNC_NAME = "FogBugzDataSync";

        //Special FogBugz fields that we map to Spira custom properties
        private const string FOGBUGZ_SPECIAL_FIELD_AREA = "Area";
        private const string FOGBUGZ_SPECIAL_FIELD_VERSION = "Version";
        private const string FOGBUGZ_SPECIAL_FIELD_COMPUTER = "Computer";
        private const string FOGBUGZ_SPECIAL_INCIDENT_STATUS_CLOSED = "Closed";
        private const int FOGBUGZ_SPECIAL_USER_CLOSED = 1;

        // Track whether Dispose has been called.
        private bool disposed = false;

        //Configuration data passed through from calling service
        private EventLog eventLog;
        private bool traceLogging;
        private int dataSyncSystemId;
        private string webServiceBaseUrl;
        private string internalLogin;
        private string internalPassword;
        private string connectionString;
        private string externalLogin;
        private string externalPassword;
        private int timeOffsetHours;
        private bool autoMapUsers;
        private bool enableKeepAlives = true;
        private bool verifyCertificate = true;
        private bool supportsRichText = false;
        private bool getNewItemsFromFogBugz = true;

        #region Class Lifecycle Methods

        /// <summary>
        /// Constructor, does nothing - all setup in the Setup() method instead
        /// </summary>
        public DataSync()
        {
            //Does Nothing - all setup in the Setup() method instead
        }

        // Implement IDisposable.
        // Do not make this method virtual.
        // A derived class should not be able to override this method.
        public void Dispose()
        {
            Dispose(true);
            // Take yourself off the Finalization queue 
            // to prevent finalization code for this object
            // from executing a second time.
            GC.SuppressFinalize(this);
        }

        // Use C# destructor syntax for finalization code.
        // This destructor will run only if the Dispose method 
        // does not get called.
        // It gives your base class the opportunity to finalize.
        // Do not provide destructors in types derived from this class.
        ~DataSync()
        {
            // Do not re-create Dispose clean-up code here.
            // Calling Dispose(false) is optimal in terms of
            // readability and maintainability.
            Dispose(false);
        }

        // Dispose(bool disposing) executes in two distinct scenarios.
        // If disposing equals true, the method has been called directly
        // or indirectly by a user's code. Managed and unmanaged resources
        // can be disposed.
        // If disposing equals false, the method has been called by the 
        // runtime from inside the finalizer and you should not reference 
        // other objects. Only unmanaged resources can be disposed.
        protected virtual void Dispose(bool disposing)
        {
            // Check to see if Dispose has already been called.
            if (!this.disposed)
            {
                // If disposing equals true, dispose all managed 
                // and unmanaged resources.
                if (disposing)
                {
                    //Remove the event log reference
                    this.eventLog = null;
                }
                // Release unmanaged resources. If disposing is false, 
                // only the following code is executed.

                //This class doesn't have any unmanaged resources to worry about
            }
            disposed = true;
        }

        #endregion

        /// <summary>
        /// Loads in all the configuration information passed from the calling service
        /// </summary>
        /// <param name="eventLog">Handle to the event log to use</param>
        /// <param name="dataSyncSystemId">The id of the plug-in used when accessing the mapping repository</param>
        /// <param name="webServiceBaseUrl">The base URL of the Spira web service</param>
        /// <param name="internalLogin">The login to Spira</param>
        /// <param name="internalPassword">The password used for the Spira login</param>
        /// <param name="connectionString">The web service URL for the FogBugz API</param>
        /// <param name="externalLogin">The login used for accessing FogBugz</param>
        /// <param name="externalPassword">The password for the FogBugz login</param>
        /// <param name="timeOffsetHours">Any time offset to apply between Spira and FogBugz</param>
        /// <param name="autoMapUsers">Should we auto-map users</param>
        /// <param name="custom01">Whether we should use HTTP keep-alives</param>
        /// <param name="custom02">Whether we need to verify the authenticity of SSL certificates</param>
        /// <param name="custom03">Whether we support rich text or not</param>
        /// <param name="custom04">Whether we should get new items from FogBugz or not</param>
        /// <param name="custom05">Not used by this plug-in</param>
        public void Setup(
            EventLog eventLog,
            bool traceLogging,
            int dataSyncSystemId,
            string webServiceBaseUrl,
            string internalLogin,
            string internalPassword,
            string connectionString,
            string externalLogin,
            string externalPassword,
            int timeOffsetHours,
            bool autoMapUsers,
            string custom01,
            string custom02,
            string custom03,
            string custom04,
            string custom05
            )
        {
            //Make sure the object has not been already disposed
            if (this.disposed)
            {
                throw new ObjectDisposedException(DATA_SYNC_NAME + " has been disposed already.");
            }

            try
            {
                //Set the member variables from the passed-in values
                this.eventLog = eventLog;
                this.traceLogging = traceLogging;
                this.dataSyncSystemId = dataSyncSystemId;
                this.webServiceBaseUrl = webServiceBaseUrl;
                this.internalLogin = internalLogin;
                this.internalPassword = internalPassword;
                this.connectionString = connectionString;
                this.externalLogin = externalLogin;
                this.externalPassword = externalPassword;
                this.timeOffsetHours = timeOffsetHours;
                this.autoMapUsers = autoMapUsers;
                if (!String.IsNullOrEmpty(custom01) && custom01.ToLowerInvariant() == "false")
                {
                    this.enableKeepAlives = false;
                }
                if (!String.IsNullOrEmpty(custom02) && custom02.ToLowerInvariant() == "false")
                {
                    this.verifyCertificate = false;
                }
                if (!String.IsNullOrEmpty(custom03) && custom03.ToLowerInvariant() == "true")
                {
                    this.supportsRichText = true;
                }
                if (!String.IsNullOrEmpty(custom04) && custom04.ToLowerInvariant() == "false")
                {
                    this.getNewItemsFromFogBugz = false;
                }
            }
            catch (Exception exception)
            {
                //Log and rethrow the exception
                LogErrorEvent("Unable to setup the " + DATA_SYNC_NAME + " plug-in ('" + exception.Message + "')\n" + exception.StackTrace, EventLogEntryType.Error);
                throw exception;
            }
        }

        /// <summary>
        /// Executes the data-sync functionality between the two systems
        /// </summary>
        /// <param name="LastSyncDate">The last date/time the plug-in was successfully executed</param>
        /// <param name="serverDateTime">The current date/time on the server</param>
        /// <returns>Code denoting success, failure or warning</returns>
        public ServiceReturnType Execute(Nullable<DateTime> lastSyncDate, DateTime serverDateTime)
        {
            //Make sure the object has not been already disposed
            if (this.disposed)
            {
                throw new ObjectDisposedException(DATA_SYNC_NAME + " has been disposed already.");
            }

            try
            {
				LogTraceEvent ("Starting " + DATA_SYNC_NAME + " data synchronization", EventLogEntryType.Information);

				//Instantiate the SpiraTest web-service proxy class
                SpiraImportExport.ImportExport spiraImportExport = new SpiraImportExport.ImportExport();
                spiraImportExport.Url = this.webServiceBaseUrl + Constants.WEB_SERVICE_URL_SUFFIX;

                //Instantiate the FogBugz API proxy class
                FogBugz fogBugz = new FogBugz(this, this.traceLogging);
                fogBugz.Url = this.connectionString;
                //Determine if we should verify the SSL certificate being used by fogbugz, and whether
                //HTTP Keep-alives should be used
                fogBugz.VerifyCertificate = this.verifyCertificate;
                fogBugz.EnableKeepAlives = this.enableKeepAlives;

				//Create new cookie container to hold the session handles
				CookieContainer cookieContainer = new CookieContainer();
                spiraImportExport.CookieContainer = cookieContainer;
				
                //First lets get the product name we should be referring to
                string productName = spiraImportExport.System_GetProductName();

				//**** Next lets load in the project and user mappings ****
                bool success = spiraImportExport.Connection_Authenticate(internalLogin, internalPassword);
                if (!success)
                {
                    //We can't authenticate so end
                    LogErrorEvent("Unable to authenticate with " + productName + " API, stopping data-synchronization", EventLogEntryType.Error);
                    return ServiceReturnType.Error;
                }
                SpiraImportExport.RemoteDataMapping[] projectMappings = spiraImportExport.DataMapping_RetrieveProjectMappings(dataSyncSystemId);
                SpiraImportExport.RemoteDataMapping[] userMappings = spiraImportExport.DataMapping_RetrieveUserMappings(dataSyncSystemId);

                //Next verify the api version and login to FogBugz
                if (!fogBugz.VerifyApi())
                {
                    LogErrorEvent("FogBugz API not compatible with this plug-in", EventLogEntryType.Error);
                    return ServiceReturnType.Error;
                }
                try
                {
                    fogBugz.Logon(externalLogin, externalPassword);
                }
                catch (Exception exception)
                {
                    //We can't authenticate so end
                    LogErrorEvent("Unable to logon to FogBugz (" + exception.Message + ")", EventLogEntryType.Error);
                    return ServiceReturnType.Error;
                }

				//Loop for each of the projects in the project mapping
                foreach (SpiraImportExport.RemoteDataMapping projectMapping in projectMappings)
				{
					//Get the SpiraTest project id equivalent Jira project identifier
                    int projectId = projectMapping.InternalId;
                    int fogBugzProject = -1;
                    if (!Int32.TryParse(projectMapping.ExternalKey, out fogBugzProject))
                    {
                        LogErrorEvent("The FogBugz project external key needs to be numeric", EventLogEntryType.Error);
                        return ServiceReturnType.Error;
                    }

					//Connect to the SpiraTest project
                    success = spiraImportExport.Connection_ConnectToProject(projectId);
                    if (!success)
                    {
                        //We can't connect so go to next project
                        LogErrorEvent("Unable to connect to " + productName + " project, please check that the " + productName + " login has the appropriate permissions", EventLogEntryType.Error);
                        continue;
                    }

                    //Get the list of project-specific mappings from the data-mapping repository
                    SpiraImportExport.RemoteDataMapping[] severityMappings = spiraImportExport.DataMapping_RetrieveFieldValueMappings(dataSyncSystemId, (int)Constants.ArtifactField.Severity);
                    SpiraImportExport.RemoteDataMapping[] priorityMappings = spiraImportExport.DataMapping_RetrieveFieldValueMappings(dataSyncSystemId, (int)Constants.ArtifactField.Priority);
                    SpiraImportExport.RemoteDataMapping[] statusMappings = spiraImportExport.DataMapping_RetrieveFieldValueMappings(dataSyncSystemId, (int)Constants.ArtifactField.Status);
                    SpiraImportExport.RemoteDataMapping[] typeMappings = spiraImportExport.DataMapping_RetrieveFieldValueMappings(dataSyncSystemId, (int)Constants.ArtifactField.Type);

                    //Get the list of custom properties configured for this project and the corresponding data mappings
                    SpiraImportExport.RemoteCustomProperty[] projectCustomProperties = spiraImportExport.CustomProperty_RetrieveProjectProperties((int)Constants.ArtifactType.Incident);
                    Dictionary<int, SpiraImportExport.RemoteDataMapping> customPropertyMappingList = new Dictionary<int, SpiraImportExport.RemoteDataMapping>();
                    Dictionary<int, SpiraImportExport.RemoteDataMapping[]> customPropertyValueMappingList = new Dictionary<int, SpiraImportExport.RemoteDataMapping[]>();
                    foreach (SpiraImportExport.RemoteCustomProperty customProperty in projectCustomProperties)
                    {
                        //Get the mapping for this custom property
                        SpiraImportExport.RemoteDataMapping customPropertyMapping = spiraImportExport.DataMapping_RetrieveCustomPropertyMapping(dataSyncSystemId, (int)Constants.ArtifactType.Incident, customProperty.CustomPropertyId);
                        customPropertyMappingList.Add(customProperty.CustomPropertyId, customPropertyMapping);

                        //For list types need to also get the property value mappings
                        if (customProperty.CustomPropertyTypeId == (int)Constants.CustomPropertyType.List)
                        {
                            SpiraImportExport.RemoteDataMapping[] customPropertyValueMappings = spiraImportExport.DataMapping_RetrieveCustomPropertyValueMappings(dataSyncSystemId, (int)Constants.ArtifactType.Incident, customProperty.CustomPropertyId);
                            customPropertyValueMappingList.Add(customProperty.CustomPropertyId, customPropertyValueMappings);
                        }
                    }

                    //Now get the list of releases and incidents that have already been mapped
                    SpiraImportExport.RemoteDataMapping[] incidentMappings = spiraImportExport.DataMapping_RetrieveArtifactMappings(dataSyncSystemId, (int)Constants.ArtifactType.Incident);
                    SpiraImportExport.RemoteDataMapping[] releaseMappings = spiraImportExport.DataMapping_RetrieveArtifactMappings(dataSyncSystemId, (int)Constants.ArtifactType.Release);

					//**** First we need to get the list of recently created incidents in SpiraTest ****
                    if (!lastSyncDate.HasValue)
                    {
                        lastSyncDate = DateTime.Parse("1/1/1900");
                    }
                    SpiraImportExport.RemoteIncident[] incidentList = spiraImportExport.Incident_RetrieveNew(lastSyncDate.Value);

                    //Create the mapping collections to hold any new items that need to get added to the mappings
                    //or any old items that need to get removed from the mappings
                    List<SpiraImportExport.RemoteDataMapping> newIncidentMappings = new List<SpiraImportExport.RemoteDataMapping>();
                    List<SpiraImportExport.RemoteDataMapping> newReleaseMappings = new List<SpiraImportExport.RemoteDataMapping>();
                    List<SpiraImportExport.RemoteDataMapping> oldReleaseMappings = new List<SpiraImportExport.RemoteDataMapping>();

					//Iterate through each record
                    foreach (SpiraImportExport.RemoteIncident remoteIncident in incidentList)
					{
						try
						{
							//Get certain incident fields into local variables (if used more than once)
                            int incidentId = remoteIncident.IncidentId.Value;
                            int incidentStatusId = remoteIncident.IncidentStatusId;

							//Make sure we've not already loaded this case
                            if (FindMappingByInternalId (projectId, incidentId, incidentMappings) == null)
							{
								//We need to add the incident number and owner in the description to make it easier
								//to track cases between the two systems. Also include a link
                                string incidentUrl = spiraImportExport.System_GetWebServerUrl() + "/IncidentDetails.aspx?incidentId=" + incidentId.ToString();
                                string description;
                                if (this.supportsRichText)
                                {
                                    description = "Incident <a href=\"" + incidentUrl + "\">[IN" + incidentId.ToString() + ":" + incidentUrl + "]</a> detected by " + remoteIncident.OpenerName + " in " + productName + ".<br/>\n" + remoteIncident.Description;
                                }
                                else
                                {
                                    description = "Incident [IN" + incidentId.ToString() + "|" + incidentUrl + "] detected by " + remoteIncident.OpenerName + " in " + productName + ".\n" + HtmlRenderAsPlainText(remoteIncident.Description);
                                }

                                //Create the FogBugz case and populate the standard fields (that don't need mapping)
                                FogBugzCase fogBugzCase = new FogBugzCase();
								fogBugzCase.Project = fogBugzProject;
								fogBugzCase.LastUpdated = remoteIncident.LastUpdateDate;
                                fogBugzCase.Title = remoteIncident.Name;
                                fogBugzCase.Description = description;
                                if (remoteIncident.StartDate.HasValue)
                                {
                                    fogBugzCase.Due = remoteIncident.StartDate.Value;
                                }
                                if (remoteIncident.EstimatedEffort.HasValue)
                                {
                                    //Convert from mins to hours
                                    fogBugzCase.HrsCurrEst = remoteIncident.EstimatedEffort.Value / 60;
                                }

                                //Now get the case category from the mapping
                                SpiraImportExport.RemoteDataMapping dataMapping = FindMappingByInternalId(projectId, remoteIncident.IncidentTypeId, typeMappings);
                                if (dataMapping == null)
                                {
                                    //We can't find the matching item so log and move to the next incident
                                    LogErrorEvent("Unable to locate mapping entry for incident type " + remoteIncident.IncidentTypeId + " in project PR" + projectId, EventLogEntryType.Error);
                                    continue;
                                }
                                int categoryId = -1;
                                if (!Int32.TryParse(dataMapping.ExternalKey, out categoryId))
                                {
                                    //The external key needs to be numeric
                                    LogErrorEvent("The external key for incident type " + remoteIncident.IncidentTypeId + " in project PR" + projectId + " needs to be numeric", EventLogEntryType.Error);
                                    continue;
                                }
                                fogBugzCase.Category = categoryId;

                                //Now get the case status from the mapping
                                dataMapping = FindMappingByInternalId(projectId, remoteIncident.IncidentStatusId, statusMappings);
                                if (dataMapping == null)
                                {
                                    //We can't find the matching item so log and move to the next incident
                                    LogErrorEvent("Unable to locate mapping entry for incident status " + remoteIncident.IncidentStatusId + " in project PR" + projectId, EventLogEntryType.Error);
                                    continue;
                                }
                                
                                //Handle the special case of an incident that needs to get mapped to the closed user
                                if (dataMapping.ExternalKey == FOGBUGZ_SPECIAL_INCIDENT_STATUS_CLOSED)
                                {
                                    fogBugzCase.Status = 0;
                                    fogBugzCase.PersonAssignedTo = FOGBUGZ_SPECIAL_USER_CLOSED;
                                }
                                else
                                {
                                    int statusId = -1;
                                    if (!Int32.TryParse(dataMapping.ExternalKey, out statusId))
                                    {

                                        //The external key needs to be numeric
                                        LogErrorEvent("The external key for incident status " + remoteIncident.IncidentStatusId + " in project " + projectId + " needs to be numeric", EventLogEntryType.Error);
                                        continue;
                                    }
                                    fogBugzCase.Status = statusId;
                                }

                                //Now get the case priority from the mapping (if priority is set)
								if (remoteIncident.PriorityId.HasValue)
								{
                                    dataMapping = FindMappingByInternalId(projectId, remoteIncident.PriorityId.Value, priorityMappings);
                                    if (dataMapping == null)
                                    {
                                        //We can't find the matching item so log and just don't set the priority
                                        LogErrorEvent("Unable to locate mapping entry for incident priority " + remoteIncident.PriorityId.Value + " in project " + projectId, EventLogEntryType.Warning);
                                    }
                                    else
                                    {
                                        int priorityId = -1;
                                        if (Int32.TryParse(dataMapping.ExternalKey, out priorityId))
                                        {
                                            fogBugzCase.Priority = priorityId;
                                        }
                                        else
                                        {
                                            //The external key needs to be numeric
                                            LogErrorEvent("The external key for incident priority " + remoteIncident.PriorityId + " in project " + projectId + " needs to be numeric", EventLogEntryType.Warning);
                                        }
                                    }
								}

                                //The reporter of the case can't be externally set

                                //Now set the assignee
                                if (remoteIncident.OwnerId.HasValue && fogBugzCase.PersonAssignedTo != FOGBUGZ_SPECIAL_USER_CLOSED)
                                {
                                    dataMapping = FindMappingByInternalId(remoteIncident.OwnerId.Value, userMappings);
                                    if (dataMapping == null)
                                    {
                                        //We can't find the matching user so ignore
                                        LogErrorEvent("Unable to locate mapping entry for user id " + remoteIncident.OwnerId.Value + " so leaving unset", EventLogEntryType.Warning);
                                    }
                                    else
                                    {
                                        int personAssignedTo = -1;
                                        if (Int32.TryParse(dataMapping.ExternalKey, out personAssignedTo))
                                        {
                                            fogBugzCase.PersonAssignedTo = personAssignedTo;
                                        }
                                        else
                                        {
                                            //The external key needs to be numeric
                                            LogErrorEvent("The external key for user " + remoteIncident.OwnerId + " in project " + projectId + " needs to be numeric", EventLogEntryType.Error);
                                            continue;
                                        }
                                    }
                                }

								//Specify the detected-in version/release if applicable
                                if (remoteIncident.ResolvedReleaseId.HasValue)
                                {
                                    int resolvedReleaseId = remoteIncident.ResolvedReleaseId.Value;
                                    dataMapping = FindMappingByInternalId(projectId, resolvedReleaseId, releaseMappings);
                                    if (dataMapping == null)
                                    {
                                        //See if it's in the list of new mappings
                                        dataMapping = FindMappingByInternalId(projectId, resolvedReleaseId, newReleaseMappings.ToArray());
                                    }
                                    int fixForId = -1;
                                    if (dataMapping == null)
                                    {
                                        //We can't find the matching item so need to create a new fixfor in FogBugz and add to mappings
                                        LogTraceEvent("Adding new FixFor in FogBugz for release " + resolvedReleaseId + "\n", EventLogEntryType.Information);
                                        FogBugzFixFor newFixFor = new FogBugzFixFor();
                                        newFixFor.Name = remoteIncident.ResolvedReleaseVersionNumber + "(" + remoteIncident.ResolvedReleaseId + ")";
                                        newFixFor.ReleaseDate = DateTime.Now.AddMonths(1);
                                        newFixFor.Project = fogBugzProject;
                                        newFixFor.Assignable = true;
                                        newFixFor = fogBugz.AddFixFor(newFixFor);

                                        //Add a new mapping entry
                                        SpiraImportExport.RemoteDataMapping newReleaseMapping = new SpiraImportExport.RemoteDataMapping();
                                        newReleaseMapping.ProjectId = projectId;
                                        newReleaseMapping.InternalId = resolvedReleaseId;
                                        newReleaseMapping.ExternalKey = newFixFor.Id.ToString();
                                        newReleaseMappings.Add(newReleaseMapping);
                                        fogBugzCase.FixFor = newFixFor.Id;
                                    }
                                    else
                                    {
                                        if (Int32.TryParse(dataMapping.ExternalKey, out fixForId))
                                        {
                                            fogBugzCase.FixFor = fixForId;
                                        }
                                        else
                                        {
                                            //The external key needs to be numeric
                                            LogErrorEvent("The external key for resolved release " + remoteIncident.ResolvedReleaseId + " in project " + projectId + " needs to be numeric", EventLogEntryType.Warning);
                                        }
                                    }
                                }
                                else if (remoteIncident.DetectedReleaseId.HasValue)
                                {
                                    //If we have no resolved release, use the detected release instead
                                    int detectedReleaseId = remoteIncident.DetectedReleaseId.Value;
                                    dataMapping = FindMappingByInternalId(projectId, detectedReleaseId, releaseMappings);
                                    if (dataMapping == null)
                                    {
                                        //See if it's in the list of new mappings
                                        dataMapping = FindMappingByInternalId(projectId, detectedReleaseId, newReleaseMappings.ToArray());
                                    }
                                    int fixForId = -1;
                                    if (dataMapping == null)
                                    {
                                        //We can't find the matching item so need to create a new fixfor in FogBugz and add to mappings
                                        LogTraceEvent("Adding new FixFor in FogBugz for release " + detectedReleaseId + "\n", EventLogEntryType.Information);
                                        FogBugzFixFor newFixFor = new FogBugzFixFor();
                                        newFixFor.Name = remoteIncident.DetectedReleaseVersionNumber + "(" + remoteIncident.DetectedReleaseId + ")";
                                        newFixFor.ReleaseDate = DateTime.Now.AddMonths(1);
                                        newFixFor.Project = fogBugzProject;
                                        newFixFor.Assignable = true;
                                        newFixFor = fogBugz.AddFixFor(newFixFor);

                                        //Add a new mapping entry
                                        SpiraImportExport.RemoteDataMapping newReleaseMapping = new SpiraImportExport.RemoteDataMapping();
                                        newReleaseMapping.ProjectId = projectId;
                                        newReleaseMapping.InternalId = detectedReleaseId;
                                        newReleaseMapping.ExternalKey = newFixFor.Id.ToString();
                                        newReleaseMappings.Add(newReleaseMapping);
                                        fogBugzCase.FixFor = newFixFor.Id;
                                    }
                                    else
                                    {
                                        if (Int32.TryParse(dataMapping.ExternalKey, out fixForId))
                                        {
                                            fogBugzCase.FixFor = fixForId;
                                        }
                                        else
                                        {
                                            //The external key needs to be numeric
                                            LogErrorEvent("The external key for detected release " + remoteIncident.DetectedReleaseId + " in project " + projectId + " needs to be numeric", EventLogEntryType.Warning);
                                        }
                                    }
                                }

                                //Setup the dictionary to hold the various custom properties to set on the FogBugz case
                                Dictionary<string, string> customPropertyValues = new Dictionary<string, string>();

                                //Now iterate through the project custom properties
                                foreach (SpiraImportExport.RemoteCustomProperty customProperty in projectCustomProperties)
                                {
                                    //Handle list and text ones separately
                                    if (customProperty.CustomPropertyTypeId == (int)Constants.CustomPropertyType.Text)
                                    {
                                        //See if we have a custom property value set
                                        String customPropertyValue = GetCustomPropertyTextValue(remoteIncident, customProperty.CustomPropertyName);
                                        if (!String.IsNullOrEmpty(customPropertyValue))
                                        {
                                            //Get the corresponding external custom field (if there is one)
                                            if (customPropertyMappingList.ContainsKey(customProperty.CustomPropertyId))
                                            {
                                                string externalCustomField = customPropertyMappingList[customProperty.CustomPropertyId].ExternalKey;

                                                if (!String.IsNullOrEmpty(externalCustomField))
                                                {
                                                    //See if we have one of the special standard FogBugz fields that it maps to
                                                    if (externalCustomField == FOGBUGZ_SPECIAL_FIELD_COMPUTER)
                                                    {
                                                        fogBugzCase.Computer = customPropertyValue;
                                                    }
                                                    if (externalCustomField == FOGBUGZ_SPECIAL_FIELD_VERSION)
                                                    {
                                                        fogBugzCase.Version = customPropertyValue;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    if (customProperty.CustomPropertyTypeId == (int)Constants.CustomPropertyType.List)
                                    {
                                        //See if we have a custom property value set
                                        Nullable<int> customPropertyValue = GetCustomPropertyListValue(remoteIncident, customProperty.CustomPropertyName);

                                        //Get the corresponding external custom field (if there is one)
                                        if (customPropertyValue.HasValue && customPropertyMappingList.ContainsKey(customProperty.CustomPropertyId))
                                        {
                                            string externalCustomField = customPropertyMappingList[customProperty.CustomPropertyId].ExternalKey;

                                            //Get the corresponding external custom field value (if there is one)
                                            if (!String.IsNullOrEmpty(externalCustomField) && customPropertyValueMappingList.ContainsKey(customProperty.CustomPropertyId))
                                            {
                                                SpiraImportExport.RemoteDataMapping[] customPropertyValueMappings = customPropertyValueMappingList[customProperty.CustomPropertyId];
                                                SpiraImportExport.RemoteDataMapping customPropertyValueMapping = FindMappingByInternalId(projectId, customPropertyValue.Value, customPropertyValueMappings);
                                                if (customPropertyValueMapping != null)
                                                {
                                                    string externalCustomFieldValue = customPropertyValueMapping.ExternalKey;

                                                    //See if we have one of the special standard FogBugz fields that it maps to
                                                    if (externalCustomField == FOGBUGZ_SPECIAL_FIELD_AREA)
                                                    {
                                                        //Now set the value of the case's area
                                                        int areaId = -1;
                                                        if (Int32.TryParse(externalCustomFieldValue, out areaId))
                                                        {
                                                            fogBugzCase.Area = areaId;
                                                        }
                                                        else
                                                        {
                                                            //The external key needs to be numeric
                                                            LogErrorEvent("The external key for area custom property value " + customPropertyValue.Value + " in project " + projectId + " needs to be numeric", EventLogEntryType.Warning);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                                //Now actually add the new FogBugz case and capture the id generated
                                fogBugzCase = fogBugz.Add(fogBugzCase);

								//Extract the FogBugz Case id and add to mappings table
                                SpiraImportExport.RemoteDataMapping newIncidentMapping = new SpiraImportExport.RemoteDataMapping();
                                newIncidentMapping.ProjectId = projectId;
                                newIncidentMapping.InternalId = incidentId;
                                newIncidentMapping.ExternalKey = fogBugzCase.Id.ToString();
                                newIncidentMappings.Add(newIncidentMapping);
							}
						}
						catch (Exception exception)
						{
							//Log and continue execution
							LogErrorEvent ("Error Adding " + productName + " Incident to FogBugz: " + exception.Message + "\n" + exception.StackTrace, EventLogEntryType.Error);
						}
					}

                    //Finally we need to update the mapping data on the server before starting the second phase
                    //of the data-synchronization
                    //At this point we have potentially added incidents, added releases and removed releases
                    spiraImportExport.DataMapping_AddArtifactMappings(dataSyncSystemId, (int)Constants.ArtifactType.Incident, newIncidentMappings.ToArray());
                    spiraImportExport.DataMapping_AddArtifactMappings(dataSyncSystemId, (int)Constants.ArtifactType.Release, newReleaseMappings.ToArray());
                    spiraImportExport.DataMapping_RemoveArtifactMappings(dataSyncSystemId, (int)Constants.ArtifactType.Release, oldReleaseMappings.ToArray());

					//**** Next we need to see if any of the previously mapped incidents has changed or any new items added to FogBugz ****
                    incidentMappings = spiraImportExport.DataMapping_RetrieveArtifactMappings(dataSyncSystemId, (int)Constants.ArtifactType.Incident);

                    //Need to create a list to hold any new releases and new incidents
                    newIncidentMappings = new List<SpiraImportExport.RemoteDataMapping>();
                    newReleaseMappings = new List<SpiraImportExport.RemoteDataMapping>();
					
					//Call the FogBugz API to get the list of recently added/changed cases
                    List<FogBugzCase> fogBugzCases = fogBugz.GetNewCases(fogBugzProject, lastSyncDate.Value.AddHours(-timeOffsetHours));

                    foreach (FogBugzCase fogBugzCase in fogBugzCases)
					{
                        //Make sure the project matches
                        if (fogBugzCase.Project == fogBugzProject)
                        {
                            //See if we have an existing mapping or not
                            SpiraImportExport.RemoteDataMapping incidentMapping = FindMappingByExternalKey(projectId, fogBugzCase.Id.ToString(), incidentMappings, false);

                            int incidentId = -1;
                            SpiraImportExport.RemoteIncident remoteIncident = null;
                            if (incidentMapping == null)
                            {
                                //If we have the option to get new items from FB disabled, don't create a new one
                                if (this.getNewItemsFromFogBugz)
                                {
                                    //This case needs to be inserted into SpiraTest
                                    remoteIncident = new SpiraImportExport.RemoteIncident();
                                    remoteIncident.ProjectId = projectId;
                                }
                            }
                            else
                            {
                                //We need to load the matching SpiraTest incident and update
                                incidentId = incidentMapping.InternalId;

                                //Now retrieve the SpiraTest incident using the Import APIs
                                try
                                {
                                    remoteIncident = spiraImportExport.Incident_RetrieveById(incidentId);
                                }
                                catch (Exception)
                                {
                                    //Ignore as it will leave the remoteIncident as null
                                }
                            }

                            try
                            {
                                //Make sure we have retrieved or created the incident
                                if (remoteIncident != null)
                                {
                                    //Debug logging - comment out for production code
                                    LogTraceEvent("Retrieved incident in " + productName + "\n", EventLogEntryType.Information);

                                    //Update the incident with the text fields
                                    if (String.IsNullOrEmpty(fogBugzCase.Title))
                                    {
                                        remoteIncident.Name = "Untitled Incident";
                                    }
                                    else
                                    {
                                        remoteIncident.Name = fogBugzCase.Title;
                                    }
                                    //Only update the spiratest description if this is a new item
                                    //as it will get overwritten by FogBugz comments otherwise since
                                    //Fogbugz doesn't keep the description separate from the comments
                                    if (incidentId == -1)
                                    {
                                        if (String.IsNullOrEmpty(fogBugzCase.Description))
                                        {
                                            remoteIncident.Description = "Empty Description in FogBugz";
                                        }
                                        else
                                        {
                                            remoteIncident.Description = fogBugzCase.Description;
                                        }
                                    }

                                    //Update any dates and effort values
                                    if (fogBugzCase.HrsCurrEst != -1)
                                    {
                                        //Need to convert into minutes
                                        remoteIncident.EstimatedEffort = fogBugzCase.HrsCurrEst * 60;
                                    }
                                    remoteIncident.StartDate = fogBugzCase.Due;
                                    remoteIncident.ClosedDate = fogBugzCase.Closed;

                                    //Debug logging - comment out for production code
                                    LogTraceEvent("Got the name and description\n", EventLogEntryType.Information);

                                    //Now get the case priority from the mapping (if priority is set)
                                    SpiraImportExport.RemoteDataMapping dataMapping;
                                    if (fogBugzCase.Priority == -1)
                                    {
                                        remoteIncident.PriorityId = null;
                                    }
                                    else
                                    {
                                        dataMapping = FindMappingByExternalKey(projectId, fogBugzCase.Priority.ToString(), priorityMappings, true);
                                        if (dataMapping == null)
                                        {
                                            //We can't find the matching item so log and just don't set the priority
                                            LogErrorEvent("Unable to locate mapping entry for case priority " + fogBugzCase.Priority + " in project " + projectId, EventLogEntryType.Warning);
                                        }
                                        else
                                        {
                                            remoteIncident.PriorityId = dataMapping.InternalId;
                                        }
                                    }

                                    //Debug logging - comment out for production code
                                    LogTraceEvent("Got the priority\n", EventLogEntryType.Information);

                                    //Now get the case status from the mapping
                                    //First handle the special case of a case assigned to the closed user
                                    //This maps to the special status that has an external key of Closed
                                    if (fogBugzCase.PersonAssignedTo == FOGBUGZ_SPECIAL_USER_CLOSED)
                                    {
                                        dataMapping = FindMappingByExternalKey(projectId, FOGBUGZ_SPECIAL_INCIDENT_STATUS_CLOSED, statusMappings, true);
                                        if (dataMapping == null)
                                        {
                                            dataMapping = FindMappingByExternalKey(projectId, fogBugzCase.Status.ToString(), statusMappings, true);
                                        }
                                    }
                                    else
                                    {
                                        dataMapping = FindMappingByExternalKey(projectId, fogBugzCase.Status.ToString(), statusMappings, true);
                                    }
                                    if (dataMapping == null)
                                    {
                                        //We can't find the matching item so log and ignore
                                        LogErrorEvent("Unable to locate mapping entry for case status " + fogBugzCase.Status + " in project PR" + projectId, EventLogEntryType.Error);
                                    }
                                    else
                                    {
                                        remoteIncident.IncidentStatusId = dataMapping.InternalId;
                                    }

                                    //Debug logging - comment out for production code
                                    LogTraceEvent("Got the status\n", EventLogEntryType.Information);

                                    //Now get the case type from the mapping
                                    dataMapping = FindMappingByExternalKey(projectId, fogBugzCase.Category.ToString(), typeMappings, true);
                                    if (dataMapping == null)
                                    {
                                        //If this is a new case in FogBugz, and there is no type
                                        //mapping then it just means that we are to ignore this type
                                        if (incidentId == -1)
                                        {
                                            continue;
                                        }
                                        //We can't find the matching item so log and ignore
                                        LogErrorEvent("Unable to locate mapping entry for case category " + fogBugzCase.Category + " in project " + projectId, EventLogEntryType.Error);
                                    }
                                    else
                                    {
                                        remoteIncident.IncidentTypeId = dataMapping.InternalId;
                                    }
                                    //Debug logging - comment out for production code
                                    LogTraceEvent("Got the type\n", EventLogEntryType.Information);

                                    //Iterate through all the comments and see if we need to add any to SpiraTest
                                    List<SpiraImportExport.RemoteIncidentResolution> newIncidentResolutions = new List<SpiraImportExport.RemoteIncidentResolution>();

                                    //FogBugz only makes the most recent comment available
                                    if (!String.IsNullOrEmpty(fogBugzCase.Description))
                                    {
                                        //Now get the list of comments attached to the SpiraTest incident
                                        SpiraImportExport.RemoteIncidentResolution[] incidentResolutions = null;
                                        if (incidentId != -1)
                                        {
                                            incidentResolutions = spiraImportExport.Incident_RetrieveResolutions(incidentId);
                                        }

                                        //See if we already have this resolution inside SpiraTest
                                        bool alreadyAdded = false;
                                        if (incidentResolutions != null)
                                        {
                                            foreach (SpiraImportExport.RemoteIncidentResolution incidentResolution in incidentResolutions)
                                            {
                                                if (incidentResolution.Resolution == fogBugzCase.Description)
                                                {
                                                    alreadyAdded = true;
                                                }
                                            }
                                        }
                                        if (!alreadyAdded)
                                        {
                                            //Get the resolution author mapping
                                            int creatorId = -1;
                                            dataMapping = FindMappingByExternalKey(fogBugzCase.PersonAssignedTo.ToString(), userMappings);
                                            if (dataMapping == null)
                                            {
                                                //We can't find the matching item so log and ignore
                                                LogTraceEvent("Unable to locate mapping entry for assignee " + fogBugzCase.PersonAssignedTo, EventLogEntryType.Warning);
                                            }
                                            else
                                            {
                                                creatorId = dataMapping.InternalId;
                                            }
                                            if (creatorId != -1)
                                            {
                                                LogTraceEvent("Got the resolution creator: " + creatorId.ToString() + "\n", EventLogEntryType.Information);

                                                //Handle nullable dates safely
                                                DateTime createdDate = DateTime.Now;
                                                if (fogBugzCase.LastUpdated.HasValue)
                                                {
                                                    createdDate = fogBugzCase.LastUpdated.Value;
                                                }

                                                //Add the comment to SpiraTest
                                                SpiraImportExport.RemoteIncidentResolution newIncidentResolution = new SpiraImportExport.RemoteIncidentResolution();
                                                newIncidentResolution.IncidentId = incidentId;
                                                newIncidentResolution.CreatorId = creatorId;
                                                newIncidentResolution.CreationDate = createdDate;
                                                newIncidentResolution.Resolution = fogBugzCase.Description;
                                                newIncidentResolutions.Add(newIncidentResolution);
                                            }
                                        }
                                    }
                                    //Debug logging - comment out for production code
                                    LogTraceEvent("Got the comments/resolution\n", EventLogEntryType.Information);

                                    //Only set the 'detector' if this case originates in FB and it's a fresh insert into Spira
                                    if (incidentId == -1)
                                    {
                                        dataMapping = FindMappingByExternalKey(fogBugzCase.PersonOpenedBy.ToString(), userMappings);
                                        if (dataMapping == null)
                                        {
                                            //We can't find the matching user so stop the insert since this is needed
                                            LogErrorEvent("Unable to locate mapping entry for FogBugz user " + fogBugzCase.PersonOpenedBy + " so not able to load new case into " + productName, EventLogEntryType.Error);
                                            continue;
                                        }
                                        else
                                        {
                                            remoteIncident.OpenerId = dataMapping.InternalId;
                                            LogTraceEvent("Got the detector " + remoteIncident.OpenerId.ToString() + "\n", EventLogEntryType.Information);
                                        }
                                    }

                                    if (fogBugzCase.PersonAssignedTo == -1)
                                    {
                                        remoteIncident.OwnerId = null;
                                    }
                                    else
                                    {
                                        dataMapping = FindMappingByExternalKey(fogBugzCase.PersonAssignedTo.ToString(), userMappings);
                                        if (dataMapping == null)
                                        {
                                            //We can't find the matching user so log and ignore - unless we're adding a new incident
                                            if (incidentId == -1)
                                            {
                                                LogErrorEvent("Unable to locate mapping entry for FogBugz user " + fogBugzCase.PersonAssignedTo + " so not able to load new case into " + productName, EventLogEntryType.Error);
                                                continue;
                                            }
                                            else
                                            {
                                                LogTraceEvent("Unable to locate mapping entry for FogBugz user " + fogBugzCase.PersonAssignedTo + " so ignoring the assignee change", EventLogEntryType.Warning);
                                            }
                                        }
                                        else
                                        {
                                            remoteIncident.OwnerId = dataMapping.InternalId;
                                            LogTraceEvent("Got the assignee " + remoteIncident.OwnerId.ToString() + "\n", EventLogEntryType.Information);
                                        }
                                    }

                                    //Specify the resolved-in release if applicable
                                    if (fogBugzCase.FixFor != -1)
                                    {
                                        //See if we have a mapped SpiraTest release
                                        dataMapping = FindMappingByExternalKey(projectId, fogBugzCase.FixFor.ToString(), releaseMappings, false);
                                        if (dataMapping == null)
                                        {
                                            //See if it's in the list of new mappings
                                            dataMapping = FindMappingByExternalKey(projectId, fogBugzCase.FixFor.ToString(), newReleaseMappings.ToArray(), false);
                                        }
                                        if (dataMapping == null)
                                        {
                                            //We can't find the matching item so need to create a new release in SpiraTest and add to mappings
                                            FogBugzFixFor fogBugzFixFor = fogBugz.GetFixFor(fogBugzCase.FixFor);
                                            if (fogBugzFixFor != null)
                                            {
                                                LogTraceEvent("Adding new release in " + productName + " for version " + fogBugzFixFor.Name + "\n", EventLogEntryType.Information);
                                                SpiraImportExport.RemoteRelease remoteRelease = new SpiraImportExport.RemoteRelease();
                                                remoteRelease.Name = fogBugzFixFor.Name;
                                                remoteRelease.VersionNumber = "FB-" + fogBugzFixFor.Id.ToString();
                                                remoteRelease.Active = true;
                                                remoteRelease.StartDate = DateTime.Now.Date;
                                                remoteRelease.EndDate = DateTime.Now.Date.AddDays(5);
                                                remoteRelease.CreatorId = remoteIncident.OpenerId;
                                                remoteRelease.CreationDate = DateTime.Now;
                                                remoteRelease.ResourceCount = 1;
                                                remoteRelease.DaysNonWorking = 0;
                                                remoteRelease = spiraImportExport.Release_Create(remoteRelease, null);

                                                //Add a new mapping entry
                                                SpiraImportExport.RemoteDataMapping newReleaseMapping = new SpiraImportExport.RemoteDataMapping();
                                                newReleaseMapping.ProjectId = projectId;
                                                newReleaseMapping.InternalId = remoteRelease.ReleaseId.Value;
                                                newReleaseMapping.ExternalKey = fogBugzFixFor.Id.ToString();
                                                newReleaseMappings.Add(newReleaseMapping);
                                                remoteIncident.ResolvedReleaseId = newReleaseMapping.InternalId;

                                                //If this is an item originating in FB, also set the detected release
                                                if (incidentId == -1)
                                                {
                                                    remoteIncident.DetectedReleaseId = newReleaseMapping.InternalId;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            remoteIncident.ResolvedReleaseId = dataMapping.InternalId;
                                            //If this is an item originating in FB, also set the detected release
                                            if (incidentId == -1)
                                            {
                                                remoteIncident.DetectedReleaseId = dataMapping.InternalId;
                                            }
                                        }
                                    }

                                    //Now we need to see if any of the SpiraTest custom properties that map to FogBugz-specific Fields have changed
                                    foreach (SpiraImportExport.RemoteCustomProperty customProperty in projectCustomProperties)
                                    {
                                        //First the text fields
                                        if (customProperty.CustomPropertyTypeId == (int)Constants.CustomPropertyType.Text)
                                        {
                                            if (customProperty.Alias == FOGBUGZ_SPECIAL_FIELD_COMPUTER)
                                            {
                                                //Now we need to set the value on the SpiraTest incident
                                                SetCustomPropertyTextValue(remoteIncident, customProperty.CustomPropertyName, fogBugzCase.Computer);
                                            }
                                            if (customProperty.Alias == FOGBUGZ_SPECIAL_FIELD_VERSION)
                                            {
                                                //Now we need to set the value on the SpiraTest incident
                                                SetCustomPropertyTextValue(remoteIncident, customProperty.CustomPropertyName, fogBugzCase.Version);
                                            }
                                        }

                                        //Next the list fields
                                        if (customProperty.CustomPropertyTypeId == (int)Constants.CustomPropertyType.List)
                                        {
                                            if (customProperty.Alias == FOGBUGZ_SPECIAL_FIELD_AREA)
                                            {
                                                //Now we need to set the value on the SpiraTest incident
                                                SpiraImportExport.RemoteDataMapping[] customPropertyValueMappings = customPropertyValueMappingList[customProperty.CustomPropertyId];
                                                SpiraImportExport.RemoteDataMapping customPropertyValueMapping = FindMappingByExternalKey(projectId, fogBugzCase.Area.ToString(), customPropertyValueMappings, false);
                                                if (customPropertyValueMapping != null)
                                                {
                                                    SetCustomPropertyListValue(remoteIncident, customProperty.CustomPropertyName, customPropertyValueMapping.InternalId);
                                                }
                                            }
                                        }
                                    }

                                    //Finally add or update the incident in SpiraTest
                                    if (incidentId == -1)
                                    {
                                        //Debug logging - comment out for production code
                                        try
                                        {
                                            remoteIncident = spiraImportExport.Incident_Create(remoteIncident);
                                        }
                                        catch (Exception exception)
                                        {
                                            LogErrorEvent("Error Adding FogBugz case " + fogBugzCase.Id + " to " + productName + " (" + exception.Message + ")\n" + exception.StackTrace, EventLogEntryType.Error);
                                            continue;
                                        }
                                        LogTraceEvent("Successfully added FogBugz case " + fogBugzCase.Id + " to " + productName + "\n", EventLogEntryType.Information);

                                        //Extract the SpiraTest incident and add to mappings table
                                        SpiraImportExport.RemoteDataMapping newIncidentMapping = new SpiraImportExport.RemoteDataMapping();
                                        newIncidentMapping.ProjectId = projectId;
                                        newIncidentMapping.InternalId = remoteIncident.IncidentId.Value;
                                        newIncidentMapping.ExternalKey = fogBugzCase.Id.ToString();
                                        newIncidentMappings.Add(newIncidentMapping);

                                        //Now add any resolutions (need to set the ID)
                                        foreach (SpiraImportExport.RemoteIncidentResolution newResolution in newIncidentResolutions)
                                        {
                                            newResolution.IncidentId = remoteIncident.IncidentId.Value;
                                        }
                                        spiraImportExport.Incident_AddResolutions(newIncidentResolutions.ToArray());
                                    }
                                    else
                                    {
                                        try
                                        {
                                            spiraImportExport.Incident_Update(remoteIncident);
                                        }
                                        catch (Exception exception)
                                        {
                                            LogErrorEvent("Error Upating FogBugz case " + fogBugzCase.Id + " in " + productName + " (" + exception.Message + ")\n" + exception.StackTrace, EventLogEntryType.Error);
                                            continue;
                                        }

                                        //Now add any resolutions
                                        spiraImportExport.Incident_AddResolutions(newIncidentResolutions.ToArray());

                                        //Debug logging - comment out for production code
                                        LogTraceEvent("Successfully updated FogBugz case " + fogBugzCase.Id + " in " + productName + "\n", EventLogEntryType.Information);
                                    }
                                }
                            }
                            catch (Exception exception)
                            {
                                //Log and continue execution
                                LogErrorEvent("Error Adding/Updating FogBugz Case in " + productName + ": " + exception.Message + "\n" + exception.StackTrace, EventLogEntryType.Error);
                            }
                        }
					}

                    //Finally we need to update the mapping data on the server
                    //At this point we have potentially added incidents and releases
                    spiraImportExport.DataMapping_AddArtifactMappings(dataSyncSystemId, (int)Constants.ArtifactType.Release, newReleaseMappings.ToArray());
                    spiraImportExport.DataMapping_AddArtifactMappings(dataSyncSystemId, (int)Constants.ArtifactType.Incident, newIncidentMappings.ToArray());
                }

				//The following code is only needed during debugging
                LogTraceEvent("Import Completed", EventLogEntryType.Warning);

				//Mark objects ready for garbage collection
				spiraImportExport.Dispose();
                spiraImportExport = null;
				fogBugz = null;

                //Let the service know that we ran correctly
                return ServiceReturnType.Success;
			}
			catch (Exception exception)
			{
				//Log the exception and return as a failure
				LogErrorEvent ("General Error: " + exception.Message + "\n" + exception.StackTrace, EventLogEntryType.Error);
                return ServiceReturnType.Error;
			}
        }

        /// <summary>
        /// Logs a trace event message if the configuration option is set
        /// </summary>
        /// <param name="eventLog">The event log handle</param>
        /// <param name="message">The message to log</param>
        /// <param name="type">The type of event</param>
        public void LogTraceEvent(string message, EventLogEntryType type = EventLogEntryType.Information)
        {
            if (traceLogging && eventLog != null)
            {
                if (message.Length > 31000)
                {
                    //Split into smaller lengths
                    int index = 0;
                    while (index < message.Length)
                    {
                        try
                        {
                            string messageElement = message.Substring(index, 31000);
                            this.LogErrorEvent(messageElement, type);
                        }
                        catch (ArgumentOutOfRangeException)
                        {
                            string messageElement = message.Substring(index);
                            this.LogErrorEvent(messageElement, type);
                        }
                        index += 31000;
                    }
                }
                else
                {
                    this.LogErrorEvent(message, type);
                }
            }
        }

        /// <summary>
        /// Logs a trace event message if the configuration option is set
        /// </summary>
        /// <param name="message">The message to log</param>
        /// <param name="type">The type of event</param>
        public void LogErrorEvent(string message, EventLogEntryType type = EventLogEntryType.Error)
        {
            if (this.eventLog != null)
            {
                if (message.Length > 31000)
                {
                    //Split into smaller lengths
                    int index = 0;
                    while (index < message.Length)
                    {
                        try
                        {
                            string messageElement = message.Substring(index, 31000);
                            eventLog.WriteEntry(messageElement, type);
                        }
                        catch (ArgumentOutOfRangeException)
                        {
                            string messageElement = message.Substring(index);
                            eventLog.WriteEntry(messageElement, type);
                        }
                        index += 31000;
                    }
                }
                else
                {
                    eventLog.WriteEntry(message, type);
                }
            }
        }

        /// <summary>
        /// Finds a mapping entry from the internal id and project id
        /// </summary>
        /// <param name="projectId">The project id</param>
        /// <param name="internalId">The internal id</param>
        /// <param name="dataMappings">The list of mappings</param>
        /// <returns>The matching entry or Null if none found</returns>
        private SpiraImportExport.RemoteDataMapping FindMappingByInternalId(int projectId, int internalId, SpiraImportExport.RemoteDataMapping[] dataMappings)
        {
            foreach (SpiraImportExport.RemoteDataMapping dataMapping in dataMappings)
            {
                if (dataMapping.InternalId == internalId && dataMapping.ProjectId == projectId)
                {
                    return dataMapping;
                }
            }
            return null;
        }

        /// <summary>
        /// Finds a mapping entry from the external key and project id
        /// </summary>
        /// <param name="projectId">The project id</param>
        /// <param name="externalKey">The external key</param>
        /// <param name="dataMappings">The list of mappings</param>
        /// <param name="onlyPrimaryEntries">Do we only want to locate primary entries</param>
        /// <returns>The matching entry or Null if none found</returns>
        private SpiraImportExport.RemoteDataMapping FindMappingByExternalKey(int projectId, string externalKey, SpiraImportExport.RemoteDataMapping[] dataMappings, bool onlyPrimaryEntries)
        {
            foreach (SpiraImportExport.RemoteDataMapping dataMapping in dataMappings)
            {
                if (dataMapping.ExternalKey == externalKey && dataMapping.ProjectId == projectId)
                {
                    //See if we're only meant to return primary entries
                    if (!onlyPrimaryEntries || dataMapping.Primary)
                    {
                        return dataMapping;
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Finds a mapping entry from the internal id
        /// </summary>
        /// <param name="internalId">The internal id</param>
        /// <param name="dataMappings">The list of mappings</param>
        /// <returns>The matching entry or Null if none found</returns>
        /// <remarks>Used when no project id stored in the mapping collection</remarks>
        private SpiraImportExport.RemoteDataMapping FindMappingByInternalId(int internalId, SpiraImportExport.RemoteDataMapping[] dataMappings)
        {
            foreach (SpiraImportExport.RemoteDataMapping dataMapping in dataMappings)
            {
                if (dataMapping.InternalId == internalId)
                {
                    return dataMapping;
                }
            }
            return null;
        }

        /// <summary>
        /// Finds a mapping entry from the external key
        /// </summary>
        /// <param name="externalKey">The external key</param>
        /// <param name="dataMappings">The list of mappings</param>
        /// <returns>The matching entry or Null if none found</returns>
        /// <remarks>Used when no project id stored in the mapping collection</remarks>
        private SpiraImportExport.RemoteDataMapping FindMappingByExternalKey(string externalKey, SpiraImportExport.RemoteDataMapping[] dataMappings)
        {
            foreach (SpiraImportExport.RemoteDataMapping dataMapping in dataMappings)
            {
                if (dataMapping.ExternalKey == externalKey)
                {
                    return dataMapping;
                }
            }
            return null;
        }

        /// <summary>
        /// Extracts the matching custom property text value from an artifact
        /// </summary>
        /// <param name="remoteArtifact">The artifact</param>
        /// <param name="customPropertyName">The name of the custom property</param>
        /// <returns></returns>
        private String GetCustomPropertyTextValue(SpiraImportExport.RemoteArtifact remoteArtifact, string customPropertyName)
        {
            if (customPropertyName == "TEXT_01")
            {
                return remoteArtifact.Text01;
            }
            if (customPropertyName == "TEXT_02")
            {
                return remoteArtifact.Text02;
            }
            if (customPropertyName == "TEXT_03")
            {
                return remoteArtifact.Text03;
            }
            if (customPropertyName == "TEXT_04")
            {
                return remoteArtifact.Text04;
            }
            if (customPropertyName == "TEXT_05")
            {
                return remoteArtifact.Text05;
            }
            if (customPropertyName == "TEXT_06")
            {
                return remoteArtifact.Text06;
            }
            if (customPropertyName == "TEXT_07")
            {
                return remoteArtifact.Text07;
            }
            if (customPropertyName == "TEXT_08")
            {
                return remoteArtifact.Text08;
            }
            if (customPropertyName == "TEXT_09")
            {
                return remoteArtifact.Text09;
            }
            if (customPropertyName == "TEXT_10")
            {
                return remoteArtifact.Text10;
            }
            return null;
        }

        /// <summary>
        /// Sets the matching custom property text value on an artifact
        /// </summary>
        /// <param name="remoteArtifact">The artifact</param>
        /// <param name="customPropertyName">The name of the custom property</param>
        /// <param name="value">The value to set</param>
        /// <returns></returns>
        private void SetCustomPropertyTextValue(SpiraImportExport.RemoteArtifact remoteArtifact, string customPropertyName, string value)
        {
            if (customPropertyName == "TEXT_01")
            {
                remoteArtifact.Text01 = value;
            }
            if (customPropertyName == "TEXT_02")
            {
                remoteArtifact.Text02 = value;
            }
            if (customPropertyName == "TEXT_03")
            {
                remoteArtifact.Text03 = value;
            }
            if (customPropertyName == "TEXT_04")
            {
                remoteArtifact.Text04 = value;
            }
            if (customPropertyName == "TEXT_05")
            {
                remoteArtifact.Text05 = value;
            }
            if (customPropertyName == "TEXT_06")
            {
                remoteArtifact.Text06 = value;
            }
            if (customPropertyName == "TEXT_07")
            {
                remoteArtifact.Text07 = value;
            }
            if (customPropertyName == "TEXT_08")
            {
                remoteArtifact.Text08 = value;
            }
            if (customPropertyName == "TEXT_09")
            {
                remoteArtifact.Text09 = value;
            }
            if (customPropertyName == "TEXT_10")
            {
                remoteArtifact.Text10 = value;
            }
        }

        /// <summary>
        /// Sets the matching custom property list value on an artifact
        /// </summary>
        /// <param name="remoteArtifact">The artifact</param>
        /// <param name="customPropertyName">The name of the custom property</param>
        /// <param name="value">The value to set</param>
        /// <returns></returns>
        private void SetCustomPropertyListValue(SpiraImportExport.RemoteArtifact remoteArtifact, string customPropertyName, Nullable<int> value)
        {
            if (customPropertyName == "LIST_01")
            {
                remoteArtifact.List01 = value;
            }
            if (customPropertyName == "LIST_02")
            {
                remoteArtifact.List02 = value;
            }
            if (customPropertyName == "LIST_03")
            {
                remoteArtifact.List03 = value;
            }
            if (customPropertyName == "LIST_04")
            {
                remoteArtifact.List04 = value;
            }
            if (customPropertyName == "LIST_05")
            {
                remoteArtifact.List05 = value;
            }
            if (customPropertyName == "LIST_06")
            {
                remoteArtifact.List06 = value;
            }
            if (customPropertyName == "LIST_07")
            {
                remoteArtifact.List07 = value;
            }
            if (customPropertyName == "LIST_08")
            {
                remoteArtifact.List08 = value;
            }
            if (customPropertyName == "LIST_09")
            {
                remoteArtifact.List09 = value;
            }
            if (customPropertyName == "LIST_10")
            {
                remoteArtifact.List10 = value;
            }
        }

        /// <summary>
        /// Extracts the matching custom property list value from an artifact
        /// </summary>
        /// <param name="remoteArtifact">The artifact</param>
        /// <param name="customPropertyName">The name of the custom property</param>
        /// <returns></returns>
        private Nullable<int> GetCustomPropertyListValue(SpiraImportExport.RemoteArtifact remoteArtifact, string customPropertyName)
        {
            if (customPropertyName == "LIST_01")
            {
                return remoteArtifact.List01;
            }
            if (customPropertyName == "LIST_02")
            {
                return remoteArtifact.List02;
            }
            if (customPropertyName == "LIST_03")
            {
                return remoteArtifact.List03;
            }
            if (customPropertyName == "LIST_04")
            {
                return remoteArtifact.List04;
            }
            if (customPropertyName == "LIST_05")
            {
                return remoteArtifact.List05;
            }
            if (customPropertyName == "LIST_06")
            {
                return remoteArtifact.List06;
            }
            if (customPropertyName == "LIST_07")
            {
                return remoteArtifact.List07;
            }
            if (customPropertyName == "LIST_08")
            {
                return remoteArtifact.List08;
            }
            if (customPropertyName == "LIST_09")
            {
                return remoteArtifact.List09;
            }
            if (customPropertyName == "LIST_10")
            {
                return remoteArtifact.List10;
            }
            return null;
        }

        /// <summary>
        /// Renders HTML content as plain text, since Bugzilla cannot handle tags
        /// </summary>
        /// <param name="source">The HTML markup</param>
        /// <returns>Plain text representation</returns>
        /// <remarks>Handles line-breaks, etc.</remarks>
        protected string HtmlRenderAsPlainText(string source)
        {
            try
            {
                string result;

                // Remove HTML Development formatting
                // Replace line breaks with space
                // because browsers inserts space
                result = source.Replace("\r", " ");
                // Replace line breaks with space
                // because browsers inserts space
                result = result.Replace("\n", " ");
                // Remove step-formatting
                result = result.Replace("\t", string.Empty);
                // Remove repeating speces becuase browsers ignore them
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"( )+", " ");

                // Remove the header (prepare first by clearing attributes)
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"<( )*head([^>])*>", "<head>",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"(<( )*(/)( )*head( )*>)", "</head>",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    "(<head>).*(</head>)", string.Empty,
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // remove all scripts (prepare first by clearing attributes)
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"<( )*script([^>])*>", "<script>",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"(<( )*(/)( )*script( )*>)", "</script>",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                //result = System.Text.RegularExpressions.Regex.Replace(result, 
                //         @"(<script>)([^(<script>\.</script>)])*(</script>)",
                //         string.Empty, 
                //         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"(<script>).*(</script>)", string.Empty,
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // remove all styles (prepare first by clearing attributes)
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"<( )*style([^>])*>", "<style>",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"(<( )*(/)( )*style( )*>)", "</style>",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    "(<style>).*(</style>)", string.Empty,
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // insert tabs in spaces of <td> tags
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"<( )*td([^>])*>", "\t",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // insert line breaks in places of <BR> and <LI> tags
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"<( )*br( )*>", "\r",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"<( )*li( )*>", "\r",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // insert line paragraphs (double line breaks) in place
                // if <P>, <DIV> and <TR> tags
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"<( )*div([^>])*>", "\r\r",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"<( )*tr([^>])*>", "\r\r",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"<( )*p([^>])*>", "\r\r",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // Remove remaining tags like <a>, links, images,
                // comments etc - anything thats enclosed inside < >
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"<[^>]*>", string.Empty,
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // replace special characters:
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"&nbsp;", " ",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"&bull;", " * ",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"&lsaquo;", "<",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"&rsaquo;", ">",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"&trade;", "(tm)",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"&frasl;", "/",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"<", "<",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @">", ">",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"&copy;", "(c)",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"&reg;", "(r)",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                // Remove all others. More can be added, see
                // http://hotwired.lycos.com/webmonkey/reference/special_characters/
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"&(.{2,6});", string.Empty,
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // for testng
                //System.Text.RegularExpressions.Regex.Replace(result, 
                //       this.txtRegex.Text,string.Empty, 
                //       System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // make line breaking consistent
                result = result.Replace("\n", "\r");

                // Remove extra line breaks and tabs:
                // replace over 2 breaks with 2 and over 4 tabs with 4. 
                // Prepare first to remove any whitespaces inbetween
                // the escaped characters and remove redundant tabs inbetween linebreaks
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    "(\r)( )+(\r)", "\r\r",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    "(\t)( )+(\t)", "\t\t",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    "(\t)( )+(\r)", "\t\r",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    "(\r)( )+(\t)", "\r\t",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                // Remove redundant tabs
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    "(\r)(\t)+(\r)", "\r\r",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                // Remove multible tabs followind a linebreak with just one tab
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    "(\r)(\t)+", "\r\t",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                // Initial replacement target string for linebreaks
                string breaks = "\r\r\r";
                // Initial replacement target string for tabs
                string tabs = "\t\t\t\t\t";
                for (int index = 0; index < result.Length; index++)
                {
                    result = result.Replace(breaks, "\r\r");
                    result = result.Replace(tabs, "\t\t\t\t");
                    breaks = breaks + "\r";
                    tabs = tabs + "\t";
                }

                // Thats it.
                return result;

            }
            catch
            {
                return source;
            }
        }
    }
}
