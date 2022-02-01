
using Microsoft.MetadirectoryServices;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using Vigo.Bas.Agresso.WebServices;
using Vigo.Bas.Agresso.WebServices.UserAdministrationV200702;
using Vigo.Bas.ManagementAgent.Log;

namespace FimSync_Ezma
{
    public class EzmaExtension :
    IMAExtensible2CallExport,
    IMAExtensible2CallImport,
    IMAExtensible2GetSchema,
    IMAExtensible2GetCapabilities,
    IMAExtensible2GetParameters
    //IMAExtensible2GetPartitions
    //IMAExtensible2FileImport,
    //IMAExtensible2FileExport,
    //IMAExtensible2GetHierarchy,
    {
        UserAdministration _userAdministrationWS = new UserAdministration();
        WSCredentials _wsCredentials = new WSCredentials();
        List<string> _companies = new List<string>();
        string _delfaultCompany;
        List<string> _filters = new List<string>();
        Dictionary<string, string> _companyFilters = new Dictionary<string, string>();
        List<String> _excludedResourceTypes = new List<string>();
        List<String> _excludedEmploymentTypes = new List<string>();
        List<String> _excludedPositionCodes = new List<string>();
        List<string> _alreadyimportedPersons = new List<string>();
        List<string> _alreadyimportedbusinessUnits = new List<string>();
        Dictionary<string, BusinessUnit> _businessUnits = new Dictionary<string, BusinessUnit>();
        Dictionary<string, OrganizationUnit> _organizationUnits = new Dictionary<string, OrganizationUnit>();


        private string _serviceUrl;
        KeyedCollection<string, ConfigParameter> _exportConfigParameters;


        public Dictionary<string, GroupInfo> groups = new Dictionary<string, GroupInfo>();
        public Dictionary<string, EmploymentInfo> employments = new Dictionary<string, EmploymentInfo>();

        public Dictionary<string, List<string>> groupmembers = new Dictionary<string, List<string>>();

        #region config parameter definition keys

        //private string _checkBoxKey01 = "Ressursdata fra web service";
        private string _checkBoxKey02 = "Bruk arbeidsstedsinfo til å generere virksomheter";

        private string _checkBoxKey03 = "Bruk organisasjonsinfo til å generere nærmeste leder";
        private string _checkBoxKey04 = "Hent ansatte fram i tid";
        private string _checkBoxKey05 = "Opprett stillingsinformasjon";
        //private string _checkBoxKey03 = "Bruk Organizations-info til å finne nærmeste leder";
        //private string _labelKey01 = "Uten avhuking hentes data fra filen angitt i Inndatafilfeltet.";
        private string _stringKey01 = "Webservice url";
        private string _stringKey02 = "Selskap(er)";
        private string _stringKey03 = "Ansattnummerfilter(e) per selskap";
        private string _stringKey04 = "Username";
        private string _stringKey05 = "Password";
        private string _stringKey06 = "Client";
        //private string _stringKey07 = "Inndatafil";
        private string _stringKey08 = "Lag grupper";
        private string _stringKey09 = "Gruppeprefix";
        private string _stringKey10 = "Gruppesuffix";
        private string _stringKey11 = "Antall dager fram i tid";
        private string _stringKey12 = "Ignorer ressurstyper";
        private string _stringKey13 = "Ignorer ansettelsestyper";
        private string _stringKey14 = "Ignorer stillingskoder";
        
        #endregion

        #region Page Size
        // Variables for page size
        private int m_importDefaultPageSize = 12;
        private int m_importMaxPageSize = 50;
        private int m_exportDefaultPageSize = 10;
        private int m_exportMaxPageSize = 20;

        public int ImportMaxPageSize
        {
            get
            {
                return m_importMaxPageSize;
            }
        }
        public int ImportDefaultPageSize
        {
            get
            {
                return m_importDefaultPageSize;
            }
        }
        public int ExportDefaultPageSize
        {
            get
            {
                return m_exportDefaultPageSize;
            }
            set
            {
                m_exportDefaultPageSize = value;
            }
        }
        public int ExportMaxPageSize
        {
            get
            {
                return m_exportMaxPageSize;
            }
            set
            {
                m_exportMaxPageSize = value;
            }
        }

        #endregion
        //
        // Constructor
        //
        public EzmaExtension()
        {
            //
            // TODO: Add constructor logic here
            //
        }
        public MACapabilities Capabilities
        {
            // Returns the capabilities that this management agent has
            get
            {
                MACapabilities myCapabilities = new MACapabilities();

                myCapabilities.ConcurrentOperation = true;
                myCapabilities.ObjectRename = false;
                myCapabilities.DeleteAddAsReplace = false;
                myCapabilities.DeltaImport = false;
                myCapabilities.DistinguishedNameStyle = MADistinguishedNameStyle.None;
                //myCapabilities.ExportType = MAExportType.ObjectReplace;
                //myCapabilities.ExportType = MAExportType.AttributeUpdate;
                //myCapabilities.ExportType = MAExportType.AttributeReplace;
                myCapabilities.NoReferenceValuesInFirstExport = true;
                myCapabilities.Normalizations = MANormalizations.None;
                myCapabilities.ObjectConfirmation = MAObjectConfirmation.NoDeleteConfirmation;

                return myCapabilities;
            }
        }

        public IList<ConfigParameterDefinition> GetConfigParameters(KeyedCollection<string, ConfigParameter> configParameters, ConfigParameterPage page)
        {
            List<ConfigParameterDefinition> configParametersDefinitions = new List<ConfigParameterDefinition>();

            switch (page)
            {
                case ConfigParameterPage.Connectivity:
                    //configParametersDefinitions.Add(ConfigParameterDefinition.CreateCheckBoxParameter(_checkBoxKey01, true));
                    //configParametersDefinitions.Add(ConfigParameterDefinition.CreateLabelParameter(_labelKey01));

                    //configParametersDefinitions.Add(ConfigParameterDefinition.CreateDividerParameter());
                    // Parametere for webservice
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter(_stringKey01, "", ""));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter(_stringKey02, "", "Skriv inn selskaper adskilt med,"));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter(_stringKey03, "", "Skriv inn filter adskilt med ,"));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter(_stringKey12, "", "Adskilt med ,"));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter(_stringKey13, "", "Adskilt med ,"));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter(_stringKey14, "", "Adskilt med ,"));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateCheckBoxParameter(_checkBoxKey04, false));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter(_stringKey11, ""));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateDividerParameter());

                    // Credentials for the webservice
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter(_stringKey04, ""));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateEncryptedStringParameter(_stringKey05, ""));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter(_stringKey06, ""));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateDividerParameter());
                    // Parametere for filimport Agresso
                    //configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter(_stringKey07, "", "Legg inn fullstendig sti"));
                    //hardcoded in the person constructor for file input instead
                    //configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter("Begrepsnøkler for gruppedannelse", "", "Skriv inn typer adskilt med ,"));
                    //configParametersDefinitions.Add(ConfigParameterDefinition.CreateDividerParameter());
                    //configParametersDefinitions.Add(ConfigParameterDefinition.CreateLabelParameter("Felles parametre for begge inndatametodene:"));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateCheckBoxParameter(_stringKey08, false));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter(_stringKey09, "", ""));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter(_stringKey10, "", "_HRM"));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateDividerParameter());

                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateCheckBoxParameter(_checkBoxKey02, false));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateCheckBoxParameter(_checkBoxKey03, false));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateCheckBoxParameter(_checkBoxKey05, false));


                    break;
                case ConfigParameterPage.Global:
                case ConfigParameterPage.Partition:
                case ConfigParameterPage.RunStep:
                    break;
            }

            return configParametersDefinitions;
        }

        public ParameterValidationResult ValidateConfigParameters(KeyedCollection<string, ConfigParameter> configParameters, ConfigParameterPage page)
        {

            // Configuration validation

            ParameterValidationResult myResults = new ParameterValidationResult();
            try
            {
                //Logger.Log.Info(configParameters);
                //bool _useWebService = configParameters[_checkBoxKey01].Value == "1";

                //if (_useWebService == true)
                //{
                //    //TODO: add validation when using web service
                //}
                //else
                //{
                //    String xmlcontent = File.ReadAllText(configParameters[_stringKey07].Value);
                //    XDocument.Parse(xmlcontent);
                //}
            }
            catch (Exception e)
            {
                myResults.ErrorMessage = "Parameter Validation failed: " + e.Message;
                myResults.Code = ParameterValidationResultCode.Failure;
                //Logger.Log.ErrorFormat("Parameter Validation failed: {0}", myResults.ErrorMessage);
            }

            return myResults;
        }

        public Schema GetSchema(KeyedCollection<string, ConfigParameter> configParameters)
        {
            // Create CS Schema type person
            SchemaType person = SchemaType.Create("person", true);

            // Anchor
            person.Attributes.Add(SchemaAttribute.CreateAnchorAttribute("ResourceId", AttributeType.String));

            // Attributes
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("FirstName", AttributeType.String, AttributeOperation.ImportExport));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Surname", AttributeType.String, AttributeOperation.ImportExport));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("SocialSecurityNumber", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Birthdate", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Sex", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("DateFrom", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("DateTo", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Street", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("ZipCode", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Place", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Email", AttributeType.String, AttributeOperation.ImportExport));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Mobile", AttributeType.String, AttributeOperation.ImportExport));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Telephone", AttributeType.String, AttributeOperation.ImportExport));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Home", AttributeType.String, AttributeOperation.ImportExport));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("LastUpdate", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("TotalEmploymentPercentage", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("CompanyCode", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateMultiValuedAttribute("Employment", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("HasMainPositionSet", AttributeType.Boolean, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("NumberOfActiveEmployments", AttributeType.Integer, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("MainPositionEmploymentType", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("MainPositionEmploymentTypeDescription", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("MainPositionPercentage", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("MainPositionPositionStartDate", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("MainPositionPositionEndDate", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("MainPositionPostCode", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("MainPositionPostCodeDescription", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("MainPositionOrganizationUnitId", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("MainPositionWorkPlaceDescription", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("MainPositionWorkPlaceId", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("MainPositionBusinessUnitId", AttributeType.Reference, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("MainPositionBusinessUnitNumber", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("MainPositionBusinessUnitName", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("MainPositionOrganizationUnitDescription", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("MainPositionManagerRef", AttributeType.Reference, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("MainPositionManagerId", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("OrganizationID", AttributeType.Reference, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("ResourceTypeValue", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("ResourceTypeDescription", AttributeType.String, AttributeOperation.ImportOnly));

            // Create CS type group
            SchemaType gruppe = SchemaType.Create("group", true);

            // Anchor
            gruppe.Attributes.Add(SchemaAttribute.CreateAnchorAttribute("GroupId", AttributeType.String, AttributeOperation.ImportOnly));

            // Attributes
            gruppe.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("GroupName", AttributeType.String, AttributeOperation.ImportOnly));
            gruppe.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("GroupType", AttributeType.String, AttributeOperation.ImportOnly));
            gruppe.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("GroupManager", AttributeType.Reference, AttributeOperation.ImportOnly));
            gruppe.Attributes.Add(SchemaAttribute.CreateMultiValuedAttribute("member", AttributeType.Reference, AttributeOperation.ImportOnly));

            // Create CS Schema type organizationUnit
            SchemaType organizationUnit = SchemaType.Create("organizationUnit", true);

            // Anchor
            organizationUnit.Attributes.Add(SchemaAttribute.CreateAnchorAttribute("organizationUnitId", AttributeType.String));

            // Attributes
            organizationUnit.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("organizationUnitLevel", AttributeType.Integer, AttributeOperation.ImportOnly));
            organizationUnit.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("organizationUnitName", AttributeType.String, AttributeOperation.ImportOnly));
            organizationUnit.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("headOfOrganizationUnit", AttributeType.String, AttributeOperation.ImportOnly));
            organizationUnit.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("parentOrganizationUnitValue", AttributeType.String, AttributeOperation.ImportOnly));
            organizationUnit.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("parentOrganizationUnitName", AttributeType.String, AttributeOperation.ImportOnly));
            organizationUnit.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("companyCode", AttributeType.String, AttributeOperation.ImportExport));

            // Create CS Schema type workplace
            SchemaType workplace = SchemaType.Create("workplace", true);

            // Anchor
            workplace.Attributes.Add(SchemaAttribute.CreateAnchorAttribute("workplaceId", AttributeType.String));

            // Attributes
            workplace.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("companyCode", AttributeType.String, AttributeOperation.ImportExport));
            workplace.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("workplaceDescription", AttributeType.String, AttributeOperation.ImportExport));
            workplace.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("organizationNumberValue", AttributeType.String, AttributeOperation.ImportOnly));
            workplace.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("organizationNumberDescription", AttributeType.String, AttributeOperation.ImportOnly));
            workplace.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("legalOrganizationNumberValue", AttributeType.String, AttributeOperation.ImportOnly));



            // Create CS type businessUnit
            SchemaType businessUnit = SchemaType.Create("businessUnit", true);

            // Anchor
            businessUnit.Attributes.Add(SchemaAttribute.CreateAnchorAttribute("BusinessUnitid", AttributeType.String, AttributeOperation.ImportOnly));

            // Attributes
            businessUnit.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("BusinessUnitName", AttributeType.String, AttributeOperation.ImportOnly));
            businessUnit.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("BusinessNumber", AttributeType.String, AttributeOperation.ImportOnly));
            businessUnit.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("CompanyCode", AttributeType.String, AttributeOperation.ImportOnly));
            businessUnit.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("LegalNumber", AttributeType.String, AttributeOperation.ImportOnly));
            businessUnit.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("OrganizationId", AttributeType.Reference, AttributeOperation.ImportOnly));

            // Create CS type organization
            SchemaType organization = SchemaType.Create("organization", true);

            // Anchor
            organization.Attributes.Add(SchemaAttribute.CreateAnchorAttribute("Organizationid", AttributeType.String, AttributeOperation.ImportOnly));

            // Attributes
            organization.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("OrganizationName", AttributeType.String, AttributeOperation.ImportOnly));
            organization.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("OrganizationNumber", AttributeType.String, AttributeOperation.ImportOnly));

            // Create CS type employment
            SchemaType employment = SchemaType.Create("employment", false);

            // Anchor
            employment.Attributes.Add(SchemaAttribute.CreateAnchorAttribute("EmploymentId", AttributeType.String, AttributeOperation.ImportOnly));

            // Attributes
            employment.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("EmployeeId", AttributeType.String, AttributeOperation.ImportOnly));
            employment.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("EmploymentTypeId", AttributeType.String, AttributeOperation.ImportOnly));
            employment.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("EmploymentTypeDescription", AttributeType.String, AttributeOperation.ImportOnly));
            employment.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("DateFrom", AttributeType.String, AttributeOperation.ImportOnly));
            employment.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("DateTo", AttributeType.String, AttributeOperation.ImportOnly));
            employment.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("WorkPlaceId", AttributeType.String, AttributeOperation.ImportOnly)); 
            employment.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("BusinessUnitId", AttributeType.String, AttributeOperation.ImportOnly));
            employment.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("OrgUnitId", AttributeType.String, AttributeOperation.ImportOnly));
            employment.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("MainPosition", AttributeType.Boolean, AttributeOperation.ImportOnly));            
            employment.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("PositionPercentage", AttributeType.String, AttributeOperation.ImportOnly));
            employment.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("ManagerId", AttributeType.String, AttributeOperation.ImportOnly));
            employment.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("PostCode", AttributeType.String, AttributeOperation.ImportOnly));
            employment.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("PostCodeDescription", AttributeType.String, AttributeOperation.ImportOnly));


            // Return schema
            Schema schema = Schema.Create();
            schema.Types.Add(person);
            schema.Types.Add(gruppe);
            schema.Types.Add(organizationUnit);
            schema.Types.Add(businessUnit);
            schema.Types.Add(organization);
            schema.Types.Add(employment);
            schema.Types.Add(workplace);
            
            return schema;
        }

        #region Import methods
        /*
         * Attributes used during import 
         */
        List<ImportListItem> ImportedObjectsList;
        int GetImportEntriesIndex, GetImportEntriesPageSize;

        public OpenImportConnectionResults OpenImportConnection(KeyedCollection<string, ConfigParameter> configParameters,
                                                Schema types,
                                                OpenImportConnectionRunStep importRunStep)
        {
            Logger.Log.InfoFormat("OpenImportConnection started with import type {0} and page size {1}", importRunStep.ImportType.ToString(), importRunStep.PageSize.ToString());
            if (configParameters[_stringKey02].Value.Contains(","))
            {
                _companies = GetCompanies(configParameters);
                Logger.Log.Debug("More than one company in input");
                foreach (string filter in configParameters[_stringKey03].Value.Split(','))
                {
                    _filters.Add(filter.Trim());
                }
                //TODO: Validation should check that number of companies and number of filters are equal. 
                //      If not the code below could break
                int i = 0;
                foreach (string company in _companies)
                {
                    _companyFilters.Add(company, _filters[i]);
                    i++;
                }
            }
            else
            {
                string company = configParameters[_stringKey02].Value.Trim();
                Logger.Log.Debug("Only one company in input");
                string filter = configParameters[_stringKey03].Value.Trim();
                _companies.Add(company);
                _companyFilters.Add(company, filter);
            }

            string _excludedResourceTypesInput = configParameters[_stringKey12].Value;
            if (_excludedResourceTypesInput.Contains(","))
            {
                foreach (string _resType in _excludedResourceTypesInput.Split(','))
                {
                    _excludedResourceTypes.Add(_resType);
                }
            }
            else if (!(string.IsNullOrEmpty(_excludedResourceTypesInput)))
            {
                _excludedResourceTypes.Add(_excludedResourceTypesInput);
            }

            string _excludedEmpTypesInput = configParameters[_stringKey13].Value;
            if (_excludedEmpTypesInput.Contains(","))
            {
                foreach (string _empType in _excludedEmpTypesInput.Split(','))
                {
                    _excludedEmploymentTypes.Add(_empType);
                }
            }
            else if (!(string.IsNullOrEmpty(_excludedEmpTypesInput)))
            {
                _excludedEmploymentTypes.Add(_excludedEmpTypesInput);
            }

            string _excludedPositionCodesInput = configParameters[_stringKey14].Value;
            if (_excludedPositionCodesInput.Contains(","))
            {
                foreach (string _posCode in _excludedPositionCodesInput.Split(','))
                {
                    _excludedPositionCodes.Add(_posCode);
                }
            }
            else if (!(string.IsNullOrEmpty(_excludedPositionCodesInput)))
            {
                _excludedPositionCodes.Add(_excludedPositionCodesInput);
            }


            //bool _useWebService = configParameters[_checkBoxKey01].Value=="1";
            bool _useWebService = true;

            // Instantiate ImportedObjectsList
            ImportedObjectsList = new List<ImportListItem>();

            if (_useWebService == true)
            {
                Logger.Log.Debug("_useWebService == true hit");

                bool _useWorkplaces = configParameters[_checkBoxKey02].Value == "1";
                bool _useOrgInfo = configParameters[_checkBoxKey03].Value == "1";
                bool _getFutureEmployees = configParameters[_checkBoxKey04].Value == "1";

                string _serviceUrl = configParameters[_stringKey01].Value.Trim();
                Logger.Log.DebugFormat("Webservice url: {0}", _serviceUrl);
                var userAdministrationWS = new UserAdministration();
                var _wsCredentials = userAdministrationWS.GetCredentials(configParameters);
                var _alreadyimportedOrganizations = new List<string>();
                var _workplaceValueTobusinessUnitId = new Dictionary<string, string>();
                var _managerToManagerGroupId = new Dictionary<string, string>();
                var _employedDateFrom = new DateTime();
                var _employedDateTo = new DateTime();
                try
                {
                    DateTime _today = DateTime.Today;
                    _employedDateFrom = _today;
                    if (_getFutureEmployees)
                    {
                        int _noOfDaysAhead = Int32.Parse(configParameters[_stringKey11].Value.Trim());
                        _employedDateTo = _today.AddDays(_noOfDaysAhead);
                    }
                    else
                    {
                        _employedDateTo = _today;
                    }
                    foreach (var company in _companies)
                    {
                        if (_useWorkplaces == true)
                        {
                            var workplaces = userAdministrationWS.GetWorkplaces(company, _wsCredentials, _serviceUrl);
                            // generate organization and businessUnit objects from workplace
                            foreach (var workplace in workplaces)
                            {
                                var workplaceInfo = new WorkplaceInfo(workplace);
                                Logger.Log.InfoFormat("Added workplace {0} with workplaceid {1} to CS", workplace.WorkplaceDescription, workplace.WorkplaceValue);
                                ImportedObjectsList.Add(new ImportListItem() { workplaceInfo = workplaceInfo });
                                // Generate organization object(s), there should be only one?
                                var _organizationNumber = workplace?.LegalOrganizationNumberValue;
                                if (!string.IsNullOrEmpty(_organizationNumber) && !_alreadyimportedOrganizations.Contains(_organizationNumber))
                                {
                                    LegalOrganization organization = new LegalOrganization(workplace);
                                    _alreadyimportedOrganizations.Add(_organizationNumber);
                                    Logger.Log.InfoFormat("Added organization {0} with organization number {1} to CS", organization._organizationName, organization._organizationNumber);
                                    ImportedObjectsList.Add(new ImportListItem() { organization = organization });
                                }
                                // generate businessUnits objects and workplace to businessUnit mapping
                                var _workplaceValue = workplace.WorkplaceValue;
                                var _businessUnitId = workplace.OrganizationNumberValue;

                                if (!_workplaceValueTobusinessUnitId.ContainsKey(_workplaceValue))
                                {
                                    _workplaceValueTobusinessUnitId.Add(_workplaceValue, _businessUnitId);
                                }
                                //create business units
                                if (!_alreadyimportedbusinessUnits.Contains(_businessUnitId))
                                {
                                    BusinessUnit businessUnit = new BusinessUnit(workplace);
                                    Logger.Log.InfoFormat("Added businessUnit {0} with orgnumber {1} to CS", workplace.OrganizationNumberDescription, _businessUnitId);
                                    _businessUnits.Add(_businessUnitId, businessUnit);
                                    _alreadyimportedbusinessUnits.Add(_businessUnitId);
                                    ImportedObjectsList.Add(new ImportListItem() { businessUnit = businessUnit });
                                }
                            }

                        }
                        if (_useOrgInfo)
                        {
                            var _orgUnits = userAdministrationWS.GetOrganization(company, _wsCredentials, _serviceUrl);
                            if (_orgUnits != null)
                            {
                                foreach (var _orgUnit in _orgUnits)
                                {
                                    var _organizationUnitID = _orgUnit?.OrganizationUnitValue;

                                    if (!string.IsNullOrEmpty(_organizationUnitID) && !_organizationUnits.ContainsKey(_organizationUnitID))
                                    {
                                        var _organizationUnit = new OrganizationUnit(_orgUnit);
                                        _organizationUnits.Add(_organizationUnitID, _organizationUnit);
                                        ImportedObjectsList.Add(new ImportListItem() { organizationUnit = _organizationUnit });
                                    }
                                }
                            }
                        }

                        var filter = _companyFilters[company];
                        Logger.Log.DebugFormat("GetResources called with company {0} and filter {1}", company, filter);

                        var resources = userAdministrationWS.GetResources(company, filter, _employedDateFrom, _employedDateTo, _wsCredentials, _serviceUrl);

                        foreach (var resource in resources)
                        {
                            var resourceId = resource.ResourceId;
                            string _resourceType;
                            var missingResourceType = (resource.ResourceTypes.Length == 0);
                            if (missingResourceType)
                            {
                                _resourceType="";
                                Logger.Log.InfoFormat("Resource with id {0} is missing resourcetype", resourceId);
                            }
                            else
                            {
                                _resourceType = resource.ResourceTypes[0].Value;
                            }                          
                            
                            if (!missingResourceType && _excludedResourceTypes.Contains(_resourceType))
                            {
                                Logger.Log.InfoFormat("Resource with id {0} has resourcetype {1} and is not added to CS", resourceId, _resourceType);
                            }
                            else
                            {
                                if (!_alreadyimportedPersons.Contains(resourceId))
                                {
                                    Person person = new Person(
                                        resource,
                                        _employedDateFrom,
                                        _employedDateTo,
                                        _excludedEmploymentTypes,
                                        _excludedPositionCodes,
                                        configParameters,
                                        _workplaceValueTobusinessUnitId,
                                        _managerToManagerGroupId,
                                        _businessUnits,
                                        _organizationUnits,
                                        groupmembers,
                                        groups,
                                        employments
                                        );

                                    if (person.NoOfActiveEmployments > 0)
                                    {
                                        Logger.Log.InfoFormat("Added person {0} to CS", resourceId);

                                        ImportedObjectsList.Add(new ImportListItem() { person = person });
                                        _alreadyimportedPersons.Add(resourceId);
                                    }
                                    else
                                    {
                                        Logger.Log.InfoFormat("Person {0} has no active employments and is not added to CS", resourceId);
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log.ErrorFormat("Error importing resources from web service: {0}", ex.Message);
                    throw;
                }

            }
            //else
            //{
            //    // Hent alle personer fra fil
            //    string _fileName = configParameters[_stringKey07].Value.Trim();
            //    String _AgressoGetAllPersons = GetAllAgressoResourcesFromFile(_fileName);

            //    // Parser alle elementer i /ExportInfo/Resources/Resource
            //    foreach (XElement xperson in XDocument.Parse(_AgressoGetAllPersons).Element("ExportInfo").Element("Resources").Elements("Resource"))
            //    {
            //        if (!_alreadyimportedPersons.Contains(xperson.Element("ResourceId").Value))
            //        {
            //            // Parser som person
            //            Person person = new Person(xperson, configParameters, groupmembers, groups);
            //            Logger.Log.DebugFormat("La til person: {0}",person.ToString());
            //            ImportedObjectsList.Add(new ImportListItem() { person = person });
            //            _alreadyimportedPersons.Add(person.ToString());
            //        }
            //    }
            //}
            if (configParameters[_stringKey08].Value == "1")
            {
                Logger.Log.Info("Lager grupper");
                foreach (var _group in groups)
                {
                    var _groupInfo = _group.Value;
                    var _groupId = configParameters[_stringKey09].Value + _group.Key + configParameters[_stringKey10].Value;
                    var _noOfMembers = _groupInfo.groupMembers.Count;

                    // We only want nonempty groups to CS
                    if (_noOfMembers > 0)
                    {
                        var _groupToCS = true;

                        if (_groupInfo.groupType == "OrgUnit")
                        {
                            if (_noOfMembers == 1)
                            {
                                var _groupManager = _groupInfo?.groupManager;
                                var _groupMember = _groupInfo.groupMembers[0];
                                // Dont add "artificial" orgunit-groups, eg groups without a manager or where the same resource is both member and manager
                                if (string.IsNullOrEmpty(_groupManager) || _groupManager == _groupMember)
                                {
                                    _groupToCS = false;
                                    Logger.Log.Info("OrgUnit-group: " + _groupId + " is not added to CS. The group is lacking a manager or has same the person as manager and member");
                                }
                            }
                        }
                        if (_groupToCS)
                        {
                            Group group = new Group(_groupInfo, configParameters);
                            Logger.Log.Info("Added group: " + _groupId);
                            ImportedObjectsList.Add(new ImportListItem() { group = group });
                        }
                    }
                    else
                    {
                        Logger.Log.Info("OrgUnit-group: " + _groupId + " is not added to CS. The group has no members");
                    }
                }
            }
            if (configParameters[_checkBoxKey05].Value=="1")
            {
                Logger.Log.Info("Lager stillingsinformasjon");
                foreach(var _employment in employments)
                {
                    var _employmentInfo = _employment.Value;
                    Logger.Log.Info("La til stilling: " + _employment.Key);
                    ImportedObjectsList.Add(new ImportListItem() { employmentInfo = _employmentInfo });
                }
            }
            // Set index values and page size
            GetImportEntriesIndex = 0;
            GetImportEntriesPageSize = importRunStep.PageSize;

            return new OpenImportConnectionResults();
        }

        public CloseImportConnectionResults CloseImportConnection(CloseImportConnectionRunStep importRunStepInfo)
        {
            return new CloseImportConnectionResults();
        }

        public GetImportEntriesResults GetImportEntries(GetImportEntriesRunStep importRunStep)
        {
            /* This method will be invoked multiple times, once for each "page" */

            List<CSEntryChange> csentries = new List<CSEntryChange>();
            GetImportEntriesResults importReturnInfo = new GetImportEntriesResults();

            // If no result, return empty success
            if (ImportedObjectsList == null || ImportedObjectsList.Count == 0)
            {
                importReturnInfo.CSEntries = csentries;
                return importReturnInfo;
            }

            // Parse a full page or to the end
            for (int currentPage = 0;
                GetImportEntriesIndex < ImportedObjectsList.Count && currentPage < GetImportEntriesPageSize;
                GetImportEntriesIndex++, currentPage++)
            {
                if (ImportedObjectsList[GetImportEntriesIndex].person != null)
                {
                    csentries.Add(ImportedObjectsList[GetImportEntriesIndex].person.GetCSEntryChange());
                }

                if (ImportedObjectsList[GetImportEntriesIndex].organizationUnit != null)
                {
                    csentries.Add(ImportedObjectsList[GetImportEntriesIndex].organizationUnit.GetCSEntryChange());
                }

                if (ImportedObjectsList[GetImportEntriesIndex].group != null)
                {
                    csentries.Add(ImportedObjectsList[GetImportEntriesIndex].group.GetCSEntryChange());
                }

                if (ImportedObjectsList[GetImportEntriesIndex].workplaceInfo != null)
                {
                    csentries.Add(ImportedObjectsList[GetImportEntriesIndex].workplaceInfo.GetCSEntryChange());
                }

                if (ImportedObjectsList[GetImportEntriesIndex].businessUnit != null)
                {
                    csentries.Add(ImportedObjectsList[GetImportEntriesIndex].businessUnit.GetCSEntryChange());
                }

                if (ImportedObjectsList[GetImportEntriesIndex].organization != null)
                {
                    csentries.Add(ImportedObjectsList[GetImportEntriesIndex].organization.GetCSEntryChange());
                }

                if (ImportedObjectsList[GetImportEntriesIndex].employmentInfo != null)
                {
                    csentries.Add(ImportedObjectsList[GetImportEntriesIndex].employmentInfo.GetCSEntryChange());
                }
            }

            // Store and return
            importReturnInfo.CSEntries = csentries;
            importReturnInfo.MoreToImport = GetImportEntriesIndex < ImportedObjectsList.Count;
            return importReturnInfo;
        }

        #endregion

        #region Export methods        
        public void OpenExportConnection(KeyedCollection<string, ConfigParameter> configParameters, Schema types, OpenExportConnectionRunStep exportRunStep)
        {
            Logger.Log.Info("Starting export");
            _exportConfigParameters = configParameters;
            _serviceUrl = _exportConfigParameters[_stringKey01].Value;
            _companies = GetCompanies(_exportConfigParameters);
            // TODO: Add support for multiple clients, using _defaultCompany in the meantime
            _delfaultCompany = _companies[0];
            _wsCredentials = _userAdministrationWS.GetCredentials(_exportConfigParameters);
        }

        public PutExportEntriesResults PutExportEntries(IList<CSEntryChange> csentries)
        {
            Logger.Log.Debug("Opening PutExportEntries");

            Dictionary<string, Dictionary<string, string>> personsToModify = new Dictionary<string, Dictionary<string, string>>();

            foreach (CSEntryChange csentry in csentries)
            {
                Logger.Log.DebugFormat("Exporting csentry {0} with modificationType {1}", csentry.DN, csentry.ObjectModificationType);

                switch (csentry.ObjectModificationType)
                {
                    case ObjectModificationType.Add:
                        Logger.Log.Debug("UpdateAdd hit");
                        break;
                    case ObjectModificationType.Delete:
                        Logger.Log.Debug("UpdateDelete hit");
                        break;
                    case ObjectModificationType.Replace:
                        Logger.Log.Debug("UpdateReplace hit");
                        AgressoPersonUpdate(csentry, ref personsToModify);
                        break;
                    case ObjectModificationType.Update:
                        Logger.Log.Debug("UpdateUpdate hit");
                        AgressoPersonUpdate(csentry, ref personsToModify);
                        break;
                    case ObjectModificationType.Unconfigured:
                        Logger.Log.Debug("UpdateUnconfigured hit");
                        break;
                    case ObjectModificationType.None:
                        Logger.Log.Debug("UpdateNone hit");
                        break;
                }
            }
            Logger.Log.Debug("Passed Foreach loop PutExportEntries");
            if (personsToModify != null)
            {
                ModifyAgressoPersons(personsToModify, _wsCredentials, _serviceUrl);
                PutExportEntriesResults exportEntriesResults = new PutExportEntriesResults();
                return exportEntriesResults;

            }
            return null;
        }

        public void CloseExportConnection(CloseExportConnectionRunStep exportRunStep)
        {
            Logger.Log.Info("Ending export");
        }
        #endregion

        #region private
        private void ModifyAgressoPersons(Dictionary<string, Dictionary<string, string>> personsAndAttributesToModify,
                                        WSCredentials credentials,
                                        string serviceUrl)
        {
            var resourcesToModify = new List<Resource>();
            //get all persons to be modified from RR
            foreach (string resourceId in personsAndAttributesToModify.Keys)
            {
                Resource resource = null;
                var resources = _userAdministrationWS.GetResources(_delfaultCompany, resourceId, DateTime.Today,
                                                                    DateTime.Today, _wsCredentials, _serviceUrl);
                if (resources == null || resources.Length == 0)
                {
                    Logger.Log.ErrorFormat("Resource with id {0} not found in Resource Registry", resourceId);
                    return;
                }
                resource = resources[0];
                resourcesToModify.Add(resource);
            }
            var personsToModify = resourcesToModify.ToArray();
            UpdatePersons(ref personsToModify, personsAndAttributesToModify);
            var response = _userAdministrationWS.ModifyResources(personsToModify, _wsCredentials, _serviceUrl);
            var responseMessage = WSObjectToXML(response);
            Logger.Log.DebugFormat("ModifyResources responded with message: {0}", responseMessage);

        }

        private void UpdatePersons(ref Resource[] persons,
                                    Dictionary<string, Dictionary<string, string>> userAttributes)
        {
            int noPersons = persons.Length;
            Logger.Log.DebugFormat("Updating {0} persons", noPersons.ToString());
            for (int i = 0; i < noPersons; i++)
            {
                string resourceId = persons[i].ResourceId;
                // find the corresponding attribute, value pairs
                Dictionary<string, string> elems = userAttributes[resourceId];
                if (elems != null)
                {
                    // The work is done here:
                    UpdatePerson(ref persons[i], elems);
                }
            }
        }

        private void UpdatePerson(ref Resource person, Dictionary<string, string> elems)
        {
            var personToUpdate = WSObjectToXML(person);
            Logger.Log.DebugFormat("UpdatePerson try to update Person: {0}", personToUpdate);
            string attributesAndValues = "";
            foreach (var attribute in elems.Keys)
            {
                attributesAndValues += attribute + ":" + elems[attribute] + ", ";
            }
            char[] trimChars = { ',', ' ' };
            attributesAndValues = attributesAndValues.TrimEnd(trimChars);
            Logger.Log.DebugFormat("Attribute(s) to be updated: {0}", attributesAndValues);

            try
            {
                foreach (string attribute in elems.Keys)
                {
                    switch (attribute.ToLower())
                    {
                        case "firstName":
                            person.FirstName = elems[attribute];
                            break;
                        case "surname":
                            person.Surname = elems[attribute];
                            break;
                        case "name":
                            person.Name = elems[attribute];
                            break;
                        case "shortname":
                            person.ShortName = elems[attribute];
                            break;
                        case "email":
                            if (person.Addresses.Length > 0)
                            {
                                var mail = elems[attribute];
                                Address homeAddress = person.Addresses[0];
                                ArrayOfString eMails = new ArrayOfString();
                                eMails.Add(mail);
                                homeAddress.EMailList = eMails;
                            }
                            break;
                        //TODO: Add more attributes that can be changed eg mobile, Telephone
                        default:
                            break;
                    }
                }
                var updatedPerson = WSObjectToXML(person);
                Logger.Log.DebugFormat("Updated person: {0}", updatedPerson);
            }
            catch (Exception ex)
            {
                Logger.Log.ErrorFormat("Error in UpdatePerson: {0} ", ex.Message);
            }
        }

        /// <summary>
        /// Creates a nested dictinary object containing userid as outer key and a inner dictionary with all the changed attributes and corresponding values
        /// </summary>
        /// <param name="csentry"></param>
        /// <param name="usersToModify"></param>
        private void AgressoPersonUpdate(CSEntryChange csentry, ref Dictionary<string, Dictionary<string, string>> usersToModify)
        {
            string userId = csentry.AnchorAttributes[0].Value.ToString();
            Logger.Log.DebugFormat("AgressoPersonUpdate started for person: {0}", userId);
            Dictionary<string, string> changedAttributes = new Dictionary<string, string>();
            foreach (var attrib in csentry.ChangedAttributeNames)
            {
                string changedValue = string.Empty;
                switch (attrib.ToLower())
                {
                    case "firstname":
                        changedValue = csentry.AttributeChanges["FirstName"].ValueChanges[0].Value.ToString();
                        Logger.Log.DebugFormat("Export firstname: " + changedValue);
                        changedAttributes.Add("FirstName", changedValue);
                        break;
                    case "surname":
                        changedValue = csentry.AttributeChanges["Surname"].ValueChanges[0].Value.ToString();
                        Logger.Log.DebugFormat("Export surname: " + changedValue);
                        changedAttributes.Add("Surname", changedValue);
                        break;
                    case "name":
                        changedValue = csentry.AttributeChanges["name"].ValueChanges[0].Value.ToString();
                        Logger.Log.DebugFormat("Export name: " + changedValue);
                        changedAttributes.Add("name", changedValue);
                        break;
                    case "shortname":
                        changedValue = csentry.AttributeChanges["shortName"].ValueChanges[0].Value.ToString();
                        Logger.Log.DebugFormat("Export shortName: " + changedValue);
                        changedAttributes.Add("shortName", changedValue);
                        break;
                    case "email":
                        changedValue = csentry.AttributeChanges["email"].ValueChanges[0].Value.ToString();
                        Logger.Log.DebugFormat("Export Email: " + changedValue);
                        changedAttributes.Add("email", changedValue);
                        break;
                }
            }
            usersToModify.Add(userId, changedAttributes);
        }

        private static string WSObjectToXML(object obj)
        {
            string filter = "/";
            String returnStr = "";
            XmlSerializer serializer = new XmlSerializer(obj.GetType());
            StringBuilder result = new StringBuilder();
            XmlDocument xDoc = new XmlDocument();
            using (var writer = XmlWriter.Create(result))
            {
                serializer.Serialize(writer, obj);
            }
            xDoc.LoadXml(result.ToString());

            XmlNodeList nodelist = xDoc.SelectNodes(filter);

            if (nodelist != null)
            {
                foreach (XmlNode node in nodelist)
                {
                    returnStr += node.OuterXml;
                }
            }
            return returnStr;
        }

        private List<string> GetCompanies(KeyedCollection<string, ConfigParameter> configParameters)
        {
            var companies = new List<string>();
            if (configParameters[_stringKey02].Value.Contains(","))
            {
                foreach (string company in configParameters[_stringKey02].Value.Split(','))
                {
                    companies.Add(company.Trim());
                }
            }
            else
            {
                string company = configParameters[_stringKey02].Value.Trim();
                companies.Add(company);
            }
            return companies;
        }

        private String GetAllAgressoResourcesFromFile(string fileName)
        {
            string xmlcontent = File.ReadAllText(fileName);

            return xmlcontent;
        }

        #endregion

    };
}
