using Microsoft.MetadirectoryServices;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Xml.Linq;
using Vigo.Bas.Agresso.WebServices.UserAdministrationV200702;
using Vigo.Bas.ManagementAgent.Log;
using Newtonsoft.Json;

namespace FimSync_Ezma
{
    class Person
    {
        // Definisjoner for bruk i koden
        string _resourceId, _firstName, _surName, _name, _socialSecurityNumber, _street, _zipcode, _place, _email, _mobile,
            _home, _telephone, _lastUpdate, _resourceTypeValue, _resourceTypeDescription, _sex, _birthDate, _companyCode, _dateFrom, _dateTo,
            _employmentPercentage, _grpUnitName, _grpUnitType;

        string _mainPositionPositionStartDate, _mainPositionPositionEndDate, _mainPositionEmploymentType, _mainPositionEmploymentTypeDescription, _mainPositionPercentage, _mainPositionPostCode,
            _mainPositionPostCodeDescription, _mainPositionOrganizationUnitId, _mainPositionOrganizationUnitDescription, _mainPositionManager,
            _mainPositionWorkPlaceId, _mainPositionWorkPlaceDescription, _mainPositionbusinessUnitId, _mainPositionbusinessUnitName,
            _mainPositionLegalOrganizationNumber;

        List<string> _employments = new List<string>();

        DateTime today = DateTime.Today;
        DateTime _employedDateFrom = new DateTime();
        DateTime _employedDateTo = new DateTime();
        List<string> _excludedEmploymentTypes;
        List<string> _excludedPositionCodes;

        double _calculatedEmploymentPercentage = 0;
        int _noOfActiveEmployments = 0;
        bool _hasMainPositionSet;

        // if this number is null the person is not imported to CS
        public int NoOfActiveEmployments
        {
            get { return _noOfActiveEmployments; }
        }

        List<string> alias = new List<string>();
        List<string> initials = new List<string>();
        List<string> logonId = new List<string>();



        string _dateFormat = "yyyy-MM-dd";

        /// <summary>
        /// Constructor to be used when the person is collected from the web service GetResources
        /// </summary>
        /// <param name=""></param>
        /// <param name="configParameter"></param>
        /// <param name="groupmembers"></param>
        /// <param name="groups"></param>
        public Person(
            Resource person,
            DateTime employedDateFrom,
            DateTime employedDateTo,
            List<string> excludedEmploymentTypes,
            List<string> excludedPositionCodes,
            KeyedCollection<string, ConfigParameter> configParameter,
            Dictionary<string, string> workplaceValuesTobusinessUnitId,
            Dictionary<string, string> _managerToManagerGroupId,
            Dictionary<string, BusinessUnit> businessUnits,
            Dictionary<string, OrganizationUnit> organizationUnits,
            Dictionary<string, List<string>> groupmembers,
            Dictionary<string, GroupInfo> groups,
            Dictionary<string, EmploymentInfo> employments
                )
        {
            _employedDateFrom = employedDateFrom;
            _employedDateTo = employedDateTo;
            _excludedEmploymentTypes = excludedEmploymentTypes;
            _excludedPositionCodes = excludedPositionCodes;

            // resourceId is used as anchor in th CS
            _resourceId = person.ResourceId;

            // Finner Fornavn
            try { _firstName = person.FirstName; }
            catch { }

            // Finner Etternavn
            try { _surName = person.Surname; }
            catch { }

            try { _name = person.Name; }
            catch { }

            // Finner kjønnkode
            try { _sex = person.Sex; }
            catch { }

            // Finner fødselsnummer
            try { _socialSecurityNumber = person.SocialSecurityNumber; }
            catch { }

            try { _birthDate = person.Birthdate.ToString(_dateFormat); }
            catch { }

            // get resourcetype info
            try
            {
                if (person?.ResourceTypes.Length > 0)
                    foreach (var resourceType in person.ResourceTypes)
                    {
                        DateTime dateFrom = resourceType.DateFrom;
                        DateTime dateTo = resourceType.DateTo;
                        if (DateTime.Compare(_employedDateTo, dateFrom) >= 0 && DateTime.Compare(_employedDateFrom, dateTo) <= 0)
                        {
                            // Active resourcetype found, there should be one and only one?
                            _resourceTypeValue = resourceType.Value;
                            _resourceTypeDescription = resourceType.Description;
                            break;
                        }
                    }
            }
            catch
            { }
            // get info about employments 
            try
            {
                var _noOfEmployments = person?.Employments.Length;
                if (_noOfEmployments > 0)
                {
                    _hasMainPositionSet = false;
                    int _employmentNo = 0;
                    double _tmpPercentage = 0.0;
                    string _tmpEmployment ="";
                    string _mainEmployment ="";

                    foreach (var employment in person.Employments)
                    {
                        var _employmentType = employment.EmploymentType;
                        if (!(_excludedEmploymentTypes.Contains(_employmentType)))
                        {
                            var _positionCode = employment.PostCode;
                            if (!(_excludedPositionCodes.Contains(_positionCode)))
                            {
                                _employmentNo++;
                                DateTime dateFrom = employment.DateFrom;
                                DateTime dateTo = employment.DateTo;

                                if (DateTime.Compare(_employedDateTo, dateFrom) >= 0 && DateTime.Compare(_employedDateFrom, dateTo) <= 0)
                                {
                                    _noOfActiveEmployments++;

                                    var _employmentAsString = JsonConvert.SerializeObject(employment);
                                    _employments.Add(_employmentAsString);

                                    var _employmentInfo = new EmploymentInfo(employment, _resourceId, true, organizationUnits, workplaceValuesTobusinessUnitId);

                                    var _initEmploymentId = _employmentInfo.EmploymentId;
                                    var _employmentId = _initEmploymentId;

                                    // generate unique employment key 
                                    int _i = 0;
                                    while (employments.ContainsKey(_employmentId))
                                    {
                                        _i++;
                                        _employmentId = _initEmploymentId + '_' + _i.ToString();
                                    }

                                    // update employmentId on the employmentInfo object
                                    _employmentInfo.EmploymentId = _employmentId;
                                    employments.Add(_employmentId, _employmentInfo);

                                    var _thisPercentage = _employmentInfo.Percentage;
                                    _calculatedEmploymentPercentage += _thisPercentage;

                                    // save employment with highest percentage so far as a candidate for main position
                                    if (_thisPercentage >= _tmpPercentage)
                                    {
                                        _tmpPercentage = _thisPercentage;
                                        _tmpEmployment = _employmentId;
                                    }

                                    //locate main position
                                    if (!(_hasMainPositionSet) && (_employmentInfo.MainPosition == true || _noOfActiveEmployments == 1 || _noOfActiveEmployments == _employmentNo))
                                    {
                                        if (_employmentInfo.MainPosition == true)
                                        {
                                            _hasMainPositionSet = true;
                                        }

                                        // have to test _tmpPercentage > 0 in case all employments has percentage 0. If so _tmpEmployment is uninitialized                                
                                        if (!(_hasMainPositionSet) && _noOfActiveEmployments == _employmentNo && _tmpPercentage > 0)
                                        {
                                            _mainEmployment = _tmpEmployment;
                                        }
                                        else
                                        {
                                            _mainEmployment = _employmentId;
                                        }
                                        // update MainPosition in employmentInfo
                                        employments[_mainEmployment].MainPosition = true;

                                        _mainPositionEmploymentType = employments[_mainEmployment].EmploymentTypeId;
                                        _mainPositionEmploymentTypeDescription = employments[_mainEmployment].EmploymentTypeDescription;
                                        _mainPositionPercentage = employments[_mainEmployment].Percentage.ToString();
                                        _mainPositionPostCode = employments[_mainEmployment].PostCode;
                                        _mainPositionPostCodeDescription = employments[_mainEmployment].PostCodeDescription;
                                        _mainPositionPositionStartDate = employments[_mainEmployment].DateFrom;
                                        _mainPositionPositionEndDate = employments[_mainEmployment].DateTo;

                                        _mainPositionManager = employments[_mainEmployment]?.ManagerId;

                                        _mainPositionWorkPlaceId = employments[_mainEmployment].WorkPlaceId;
                                        _mainPositionWorkPlaceDescription = employments[_mainEmployment].WorkPlaceDescription;

                                        _mainPositionbusinessUnitId = employments[_mainEmployment]?.BusinessUnitId;

                                        if (!string.IsNullOrEmpty(_mainPositionbusinessUnitId))
                                        {
                                            BusinessUnit _businessUnit = businessUnits[_mainPositionbusinessUnitId];
                                            _mainPositionbusinessUnitName = _businessUnit?._businessUnitName;
                                            _mainPositionLegalOrganizationNumber = _businessUnit?._legalOrganizationNumber;
                                        }
                                        else
                                        {
                                            Logger.Log.DebugFormat("Main position workplace {0} for employee {1} has no assosiated business unit", _mainPositionWorkPlaceId, _resourceId);
                                        }

                                        _mainPositionOrganizationUnitId = employments[_mainEmployment].OrgUnitId;
                                        _mainPositionOrganizationUnitDescription = employments[_mainEmployment].OrgUnitDescription;
                                    }

                                    var _thisPositionWorkPlaceId = _employmentInfo.WorkPlaceId;
                                    var _thisPositionWorkPlaceDescription = _employmentInfo.WorkPlaceDescription;

                                    //Create or update Workplace groups
                                    if (!string.IsNullOrEmpty(_thisPositionWorkPlaceId) && !string.IsNullOrEmpty(_thisPositionWorkPlaceDescription))
                                    {
                                        AddOrUpdateGroup(_thisPositionWorkPlaceId, _thisPositionWorkPlaceDescription, "Workplace", null, ref groups);
                                    }
                                    //TODO: add businessunit if not already present
                                    //var _thisPositionbusinessUnitId = _employmentInfo?.BusinessUnitId;


                                    var _thisPositionOrganizationUnitId = _employmentInfo.OrgUnitId;
                                    var _thisPositionOrganizationUnitDescription = _employmentInfo.OrgUnitDescription;
                                    var _thisPositionManager = _employmentInfo?.ManagerId;

                                    // Update group info orgunits  
                                    if (!string.IsNullOrEmpty(_thisPositionOrganizationUnitId) && !string.IsNullOrEmpty(_thisPositionOrganizationUnitDescription))
                                    {
                                        AddOrUpdateGroup(_mainPositionOrganizationUnitId, _mainPositionOrganizationUnitDescription, "OrgUnit", _thisPositionManager, ref groups);

                                        // Create or update manager group 
                                        if (!string.IsNullOrEmpty(_thisPositionManager))
                                        {
                                            string _mgrGroupId;

                                            if (_managerToManagerGroupId.ContainsKey(_thisPositionManager))
                                            {
                                                _mgrGroupId = _managerToManagerGroupId[_thisPositionManager];
                                                var group = groups[_mgrGroupId];
                                                if (!group.groupMembers.Contains(_resourceId))
                                                {
                                                    group.groupMembers.Add(_resourceId);
                                                }
                                            }
                                            else
                                            {
                                                var _managerOrgUnit = GetOrgUnitForManager(_resourceId, _thisPositionOrganizationUnitId, organizationUnits);
                                                string _grpUnitID = _managerOrgUnit._organizationUnitId;
                                                string _grpUnitName = _managerOrgUnit._organizationUnitName;

                                                if (!string.IsNullOrEmpty(_grpUnitID) && !string.IsNullOrEmpty(_grpUnitName))
                                                {
                                                    _mgrGroupId = "MGR_" + _grpUnitID;
                                                    _grpUnitType = "ManagerGroup";
                                                    var _groupInfo = new GroupInfo(_mgrGroupId, _grpUnitName, _grpUnitType, _thisPositionManager, new List<string>() { _resourceId });
                                                    groups.Add(_mgrGroupId, _groupInfo);

                                                    _managerToManagerGroupId.Add(_thisPositionManager, _mgrGroupId);
                                                }
                                            }
                                        }
                                    }
                                    
                                }
                                else
                                {
                                    Logger.Log.DebugFormat("Resource {0} has multiple employments, only primary position is added to CS", _resourceId);

                                    // TODO: Add support for multiple employements
                                }
                                }
                            }
                        }                    
                        _employmentPercentage = _calculatedEmploymentPercentage.ToString();
                }
                else
                {
                    Logger.Log.DebugFormat("Resource {0} has no registered employments", _resourceId);
                }
            }
            catch { }

            try
            {
                _companyCode = person.CompanyCode;
            }
            catch
            {
            }

            // Adresses-elementet på ressurs kan inneholde flere adressetyper. AFK ser kun ut til å benytte Adressetype=1:
            try
            {
                foreach (Address address in person.Addresses)
                {
                    if (address.Type == "1")
                    {
                        //_street, _zipcode, _place,
                        if (!string.IsNullOrEmpty(address.Street))
                        {
                            _street = address.Street;
                        }
                        if (!string.IsNullOrEmpty(address.ZipCode))
                        {
                            _zipcode = address.ZipCode;
                        }
                        if (!string.IsNullOrEmpty(address.Place))
                        {
                            _place = address.Place;
                        }
                        if (!string.IsNullOrEmpty(address.EMailList[0]))
                        {
                            _email = address.EMailList[0];
                        }
                        if (!string.IsNullOrEmpty(address.Mobile))
                        {
                            _mobile = address.Mobile;
                        }
                        if (!string.IsNullOrEmpty(address.Telephone))
                        {
                            _telephone = address.Telephone;
                        }
                        if (!string.IsNullOrEmpty(address.Home))
                        {
                            _home = address.Home;
                        }

                    }
                }
            }
            catch { }

            try { _dateFrom = person.DateFrom.ToString(_dateFormat); }
            catch { }

            try { _dateTo = person.DateTo.ToString(_dateFormat); }
            catch { }

            // Finner sist endret tidspunkt
            try { _lastUpdate = person.LastUpdate.ToString(_dateFormat); }
            catch { }
        }



        public string Anchor()
        {
            return _resourceId;
        }

        public override string ToString()
        {
            return _resourceId;
        }

        internal Microsoft.MetadirectoryServices.CSEntryChange GetCSEntryChange()
        {
            CSEntryChange csentry = CSEntryChange.Create();
            csentry.ObjectModificationType = ObjectModificationType.Add;
            csentry.ObjectType = "person";

            //get _resourceId (unique ID for resources in Agresso)
            csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("ResourceId", _resourceId));

            // get resourcetype id and description
            if (!string.IsNullOrEmpty(_resourceTypeValue))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("ResourceTypeValue", _resourceTypeValue));

            if (!string.IsNullOrEmpty(_resourceTypeDescription))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("ResourceTypeDescription", _resourceTypeDescription));

            //Henter Fornavn
            if (!string.IsNullOrEmpty(_firstName))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("FirstName", _firstName));

            //Henter Etternavn
            if (!string.IsNullOrEmpty(_surName))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Surname", _surName));

            //Henter fullt fødselsnummer
            if (!string.IsNullOrEmpty(_socialSecurityNumber))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("SocialSecurityNumber", _socialSecurityNumber));

            //Henter fødelsdato
            if (!string.IsNullOrEmpty(_birthDate))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Birthdate", _birthDate));

            if (!string.IsNullOrEmpty(_street))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Street", _street));

            if (!string.IsNullOrEmpty(_zipcode))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("ZipCode", _zipcode));

            if (!string.IsNullOrEmpty(_place))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Place", _place));

            //Henter epost
            if (!string.IsNullOrEmpty(_email))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Email", _email));

            if (!string.IsNullOrEmpty(_mobile))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Mobile", _mobile));

            if (!string.IsNullOrEmpty(_telephone))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Telephone", _telephone));

            if (!string.IsNullOrEmpty(_home))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Home", _home));

            if (!string.IsNullOrEmpty(_sex))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Sex", _sex));

            if (!string.IsNullOrEmpty(_dateFrom))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("DateFrom", _dateFrom));

            if (!string.IsNullOrEmpty(_dateTo))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("DateTo", _dateTo));

            if (!string.IsNullOrEmpty(_employmentPercentage))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("TotalEmploymentPercentage", _employmentPercentage));

            if (!string.IsNullOrEmpty(_lastUpdate))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("LastUpdate", _lastUpdate));

            csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("HasMainPositionSet", _hasMainPositionSet));

            csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("NumberOfActiveEmployments", _noOfActiveEmployments));

            if (!string.IsNullOrEmpty(_mainPositionEmploymentType))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("MainPositionEmploymentType", _mainPositionEmploymentType));

            if (!string.IsNullOrEmpty(_mainPositionEmploymentType))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("MainPositionEmploymentTypeDescription", _mainPositionEmploymentTypeDescription));
            if (!string.IsNullOrEmpty(_mainPositionPostCode))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("MainPositionPostCode", _mainPositionPostCode));

            if (!string.IsNullOrEmpty(_mainPositionPostCodeDescription))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("MainPositionPostCodeDescription", _mainPositionPostCodeDescription));

            if (!string.IsNullOrEmpty(_mainPositionOrganizationUnitId))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("MainPositionOrganizationUnitId", _mainPositionOrganizationUnitId));

            if (!string.IsNullOrEmpty(_mainPositionOrganizationUnitDescription))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("MainPositionOrganizationUnitDescription", _mainPositionOrganizationUnitDescription));

            //_mainPositionManager
            if (!string.IsNullOrEmpty(_mainPositionManager))
            {
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("MainPositionManagerRef", _mainPositionManager));
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("MainPositionManagerId", _mainPositionManager));
            }

            if (!string.IsNullOrEmpty(_mainPositionWorkPlaceId))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("MainPositionWorkPlaceId", _mainPositionWorkPlaceId));

            if (!string.IsNullOrEmpty(_mainPositionWorkPlaceDescription))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("MainPositionWorkPlaceDescription", _mainPositionWorkPlaceDescription));

            if (!string.IsNullOrEmpty(_mainPositionbusinessUnitId))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("MainPositionBusinessUnitId", _mainPositionbusinessUnitId));

            if (!string.IsNullOrEmpty(_mainPositionbusinessUnitId))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("MainPositionBusinessUnitNumber", _mainPositionbusinessUnitId));

            if (!string.IsNullOrEmpty(_mainPositionbusinessUnitName))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("MainPositionBusinessUnitName", _mainPositionbusinessUnitName));

            if (!string.IsNullOrEmpty(_mainPositionLegalOrganizationNumber))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("OrganizationID", _mainPositionLegalOrganizationNumber));
            //Henter stillingstørrelse på primærstilling
            if (!string.IsNullOrEmpty(_mainPositionPercentage))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("MainPositionPercentage", _mainPositionPercentage));

            //Henter startdato på primærstilling
            if (!string.IsNullOrEmpty(_mainPositionPositionStartDate))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("MainPositionPositionStartDate", _mainPositionPositionStartDate));

            //Henter sluttdato på primærstilling
            if (!string.IsNullOrEmpty(_mainPositionPositionEndDate))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("MainPositionPositionEndDate", _mainPositionPositionEndDate));

            if (_employments != null)
            {
                IList<object> _employmentList = new List<object>();
                foreach (var _employment in _employments)
                {
                    _employmentList.Add(_employment);
                }
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Employment", _employmentList));

            }

            if (!string.IsNullOrEmpty(_companyCode))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("CompanyCode", _companyCode));

            return csentry;

        }

    #region private

        private OrganizationUnit GetOrgUnitForManager(string _resourceId, string _orgUnitId, Dictionary<string, OrganizationUnit> organizationUnits)
        {
            var _orgUnit = new OrganizationUnit();
            if (organizationUnits.ContainsKey(_orgUnitId))
            {
                _orgUnit = organizationUnits[_orgUnitId];
                var _headOfOrgUnit = _orgUnit._headOfOrganizationUnit;
                var _orgUnitLevel = _orgUnit._organizationUnitLevel;
                var _parentOrgUnitId = _orgUnit._parentOrganizationUnitValue;

                while (_headOfOrgUnit.Equals(_resourceId) && _orgUnitLevel != 1)
                {
                    _orgUnit = organizationUnits[_parentOrgUnitId];
                    _headOfOrgUnit = _orgUnit._headOfOrganizationUnit;
                    _orgUnitLevel = _orgUnit._organizationUnitLevel;
                    _parentOrgUnitId = _orgUnit._parentOrganizationUnitValue;
                }
            }
            return _orgUnit;
        }

        private void UpdateOrgUnitGroups(EmploymentInfo employmentInfo, ref Dictionary<string, GroupInfo> groups)
        {
            var _orgUnitId = employmentInfo.OrgUnitId;
            _grpUnitName = employmentInfo.OrgUnitDescription;
            _grpUnitType = "OrgUnit";
            if (!groups.ContainsKey(_orgUnitId))
            {
                var _groupInfo = new GroupInfo(_orgUnitId, _grpUnitName, _grpUnitType, null, new List<string>() { _resourceId });
                groups.Add(_orgUnitId, _groupInfo);
            }
            else
            {
                var group = groups[_orgUnitId];
                if (!group.groupMembers.Contains(_resourceId))
                {
                    group.groupMembers.Add(_resourceId);
                }
            }

        }
        private void AddOrUpdateGroup(string groupId, string groupName, string groupType, string manager, ref Dictionary<string, GroupInfo> groups)
        {
            var _groupId = groupId;
            var _groupName = groupName;
            var _groupType = groupType;
            if (!groups.ContainsKey(_groupId))
            {
                var _groupInfo = new GroupInfo(_groupId, _groupName, _groupType, manager, new List<string>() { _resourceId });
                groups.Add(_groupId, _groupInfo);

            }
            else
            {
                var group = groups[_groupId];
                if (!group.groupMembers.Contains(_resourceId))
                {
                    group.groupMembers.Add(_resourceId);
                }
            }
        }

        #endregion
    }

}




