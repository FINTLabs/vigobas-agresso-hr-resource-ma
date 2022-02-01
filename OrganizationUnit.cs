using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.MetadirectoryServices;
using Vigo.Bas.Agresso.WebServices.UserAdministrationV200702;

namespace FimSync_Ezma
{
    public class OrganizationUnit
    {
        public string _organizationUnitId;
        public string _companyCode;
        public int _organizationUnitLevel;
        public string _organizationUnitValue;
        public string _organizationUnitName;
        public string _headOfOrganizationUnit;
        public string _parentOrganizationUnitValue;
        public string _parentOrganizationUnitName;

        public OrganizationUnit(Organization orgUnit)
        {
            _organizationUnitId = orgUnit.OrganizationUnitValue;

            if (orgUnit.OrganizationUnitLevel != 0)
            {
                _organizationUnitLevel = orgUnit.OrganizationUnitLevel;
            }

            _companyCode = orgUnit?.CompanyCode;
            _organizationUnitValue = orgUnit?.OrganizationUnitValue;
            _organizationUnitName = orgUnit?.OrganizationUnitName;
            _headOfOrganizationUnit = orgUnit?.HeadOfOrganizationUnit;
            _parentOrganizationUnitValue = orgUnit?.ParentOrganizationUnitValue;
            _parentOrganizationUnitName = orgUnit?.ParentOrganizationUnitName;
        }

        public OrganizationUnit()
        {
        }

        public string Anchor()
        {
            return _organizationUnitId;
        }

        public override string ToString()
        {
            return _organizationUnitId;
        }

        // Returns CSEntryChange for use by the FIM Sync Engine directly
        internal Microsoft.MetadirectoryServices.CSEntryChange GetCSEntryChange()
        {
            CSEntryChange csentry = CSEntryChange.Create();
            csentry.ObjectModificationType = ObjectModificationType.Add;
            csentry.ObjectType = "organizationUnit";

            csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("organizationUnitId", _organizationUnitId));

            if (_organizationUnitLevel != 0)
            {
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("organizationUnitLevel", _organizationUnitLevel));
            }
            if (!string.IsNullOrEmpty(_companyCode))
            {
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("companyCode", _companyCode));
            }

            if (!string.IsNullOrEmpty(_organizationUnitName))
            {
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("organizationUnitName", _organizationUnitName));
            }

            if (!string.IsNullOrEmpty(_headOfOrganizationUnit))
            {
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("headOfOrganizationUnit", _headOfOrganizationUnit));
            }

            if (!string.IsNullOrEmpty(_parentOrganizationUnitValue))
            {
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("parentOrganizationUnitValue", _parentOrganizationUnitValue));
            }

            if (!string.IsNullOrEmpty(_parentOrganizationUnitName))
            {
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("parentOrganizationUnitName", _parentOrganizationUnitName));
            }

            return csentry;
        }
        }

}
