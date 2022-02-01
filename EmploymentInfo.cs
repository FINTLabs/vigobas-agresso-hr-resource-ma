using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.MetadirectoryServices;
using Vigo.Bas.Agresso.WebServices.UserAdministrationV200702;

namespace FimSync_Ezma
{
    public class EmploymentInfo
    {
        string _dateFormat = "yyyy-MM-dd";

        public string EmploymentId; // To be used as anchor
        public string EmployeeId;
        public string EmploymentTypeId;
        public string EmploymentTypeDescription;
        public string DateFrom;
        public string DateTo;
        public string WorkPlaceId;
        public string WorkPlaceDescription;
        public string BusinessUnitId;
        public string OrgUnitId;
        public string OrgUnitDescription;
        public bool MainPosition;
        public double Percentage;
        public string PostCode;
        public string PostCodeDescription;
        public string ManagerId;

        public EmploymentInfo(
            Employment employment, 
            string employeeId, 
            bool workPlaceAsDepartment, 
            Dictionary<string, OrganizationUnit> organizationUnits, 
            Dictionary<string, string> workplaceValuesTobusinessUnitId
            )
        {
            EmployeeId = employeeId;

            EmploymentTypeId = employment?.EmploymentType;
            EmploymentTypeDescription = employment?.EmploymentTypeDescription;
            DateFrom = employment.DateFrom.ToString(_dateFormat);
            DateTo = employment.DateTo.ToString(_dateFormat);

            if (employment.Workplaces.Length > 0)
            {
                WorkPlaceId = employment.Workplaces[0]?.Value.ToString();
                WorkPlaceDescription = employment.Workplaces[0]?.Description.ToString();
                if (workplaceValuesTobusinessUnitId.ContainsKey(WorkPlaceId))
                {
                    BusinessUnitId = workplaceValuesTobusinessUnitId[WorkPlaceId];
                    // Need for BusinessUnit name?                    
                }
            }
            if (employment.OrganizationUnits.Length > 0)
            {
                OrgUnitId = employment.OrganizationUnits[0]?.Value.ToString();
                OrgUnitDescription = employment.OrganizationUnits[0]?.Description.ToString();
                ManagerId = GetOrgUnitManager(EmployeeId, OrgUnitId, organizationUnits);
            }

            MainPosition = employment.MainPosition;
            Percentage = employment.Percentage;
            PostCode = employment?.PostCode;
            PostCodeDescription = employment?.PostCodeDescription;

            string _depID = (workPlaceAsDepartment) ? WorkPlaceId : OrgUnitId;
            EmploymentId = EmployeeId + '_' + _depID;

        }

        public string Anchor()
        {
            return EmploymentId;
        }


        public override string ToString()
        {
            return EmploymentId;
        }

        internal Microsoft.MetadirectoryServices.CSEntryChange GetCSEntryChange()
        {
            CSEntryChange csentry = CSEntryChange.Create();
            csentry.ObjectModificationType = ObjectModificationType.Add;
            csentry.ObjectType = "employment";

            csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("EmploymentId", EmploymentId));

            csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("EmployeeId", EmployeeId));

            if (!string.IsNullOrEmpty(EmploymentTypeId))
            {
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("EmploymentTypeId", EmploymentTypeId));
            }

            if (!string.IsNullOrEmpty(EmploymentTypeDescription))
            {
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("EmploymentTypeDescription", EmploymentTypeDescription));
            }
            if (!string.IsNullOrEmpty(DateFrom))
            {
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("DateFrom", DateFrom));
            }
            if (!string.IsNullOrEmpty(DateTo))
            {
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("DateTo", DateTo));
            }
            if (!string.IsNullOrEmpty(OrgUnitId))
            {
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("OrgUnitId", OrgUnitId));
            }
            if (!string.IsNullOrEmpty(WorkPlaceId))
            {
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("WorkPlaceId", WorkPlaceId));
            }            
            if (!string.IsNullOrEmpty(BusinessUnitId))
            {
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("BusinessUnitId", BusinessUnitId));
            }

            csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("MainPosition", MainPosition));

            csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("PositionPercentage", Percentage.ToString()));
            
            if (!string.IsNullOrEmpty(ManagerId))
            {
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("ManagerId", ManagerId));
            }
            if (!string.IsNullOrEmpty(PostCode))
            {
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("PostCode", PostCode));
            }
            if (!string.IsNullOrEmpty(PostCodeDescription))
            {
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("PostCodeDescription", PostCodeDescription));
            }
            return csentry;
        }

        private string GetOrgUnitManager(string _resourceId, string _orgUnitId, Dictionary<string, OrganizationUnit> organizationUnits)
        {
            string _manager = "";
            if (organizationUnits.ContainsKey(_orgUnitId))
            {
                var _currentOrgUnit = organizationUnits[_orgUnitId];
                var _headOfOrgUnit = _currentOrgUnit._headOfOrganizationUnit;
                var _orgUnitLevel = _currentOrgUnit._organizationUnitLevel;
                var _parentOrgUnitId = _currentOrgUnit._parentOrganizationUnitValue;

                while (organizationUnits.ContainsKey(_parentOrgUnitId) && _headOfOrgUnit.Equals(_resourceId) && _orgUnitLevel != 1)
                {
                    _currentOrgUnit = organizationUnits[_parentOrgUnitId];
                    _headOfOrgUnit = _currentOrgUnit._headOfOrganizationUnit;
                    _orgUnitLevel = _currentOrgUnit._organizationUnitLevel;
                    _parentOrgUnitId = _currentOrgUnit._parentOrganizationUnitValue;
                }

                if (!(_headOfOrgUnit.Equals(_resourceId)))
                {
                    _manager = _headOfOrgUnit;
                }
                
            }
            return _manager;
        }
    }
}
