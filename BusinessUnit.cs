using Microsoft.MetadirectoryServices;
using System.Xml.Linq;
using Vigo.Bas.Agresso.WebServices.UserAdministrationV200702;

namespace FimSync_Ezma
{
    class BusinessUnit
    {
        string _businessUnitId;
        public string _businessUnitName;
        public string _businessOrganizationNumber;
        public string _legalOrganizationNumber;
        public string  _companycode;

        public BusinessUnit(WorkPlace1 workplace)
        {
            _businessUnitId = workplace.OrganizationNumberValue;

            try
            {
                _businessUnitName = workplace.OrganizationNumberDescription;
            }
            catch { }

            try
            {
                _businessOrganizationNumber = workplace.OrganizationNumberValue;
            }
            catch { }

            try
            {
                _legalOrganizationNumber = workplace.LegalOrganizationNumberValue;
            }
            catch { }
            try
            {
                _companycode = workplace.CompanyCode;
            }
            catch { }
        }

        public string Anchor()
        {
            return _businessUnitId;
        }


        public override string ToString()
        {
            return _businessUnitId;
        }

        internal Microsoft.MetadirectoryServices.CSEntryChange GetCSEntryChange()
        {
            CSEntryChange csentry = CSEntryChange.Create();
            csentry.ObjectModificationType = ObjectModificationType.Add;
            csentry.ObjectType = "businessUnit";

            csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("BusinessUnitid", _businessUnitId));

            if (!string.IsNullOrEmpty (_businessUnitName))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("BusinessUnitName", _businessUnitName));

            if (!string.IsNullOrEmpty (_businessOrganizationNumber))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("BusinessNumber", _businessOrganizationNumber));

            if (!string.IsNullOrEmpty (_companycode))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("CompanyCode", _companycode));

            if (!string.IsNullOrEmpty (_legalOrganizationNumber))
            {
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("LegalNumber", _legalOrganizationNumber));
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("OrganizationId", _legalOrganizationNumber));

            }

            return csentry;

        }
    }
}
