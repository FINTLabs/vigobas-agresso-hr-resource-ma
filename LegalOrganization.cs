using Microsoft.MetadirectoryServices;
using System.Xml.Linq;
using Vigo.Bas.Agresso.WebServices.UserAdministrationV200702;

namespace FimSync_Ezma
{
    class LegalOrganization
    {
        string _organizationId;
        public string _organizationName;
        public string _organizationNumber;

        public LegalOrganization(WorkPlace1 workplace)
        {
            _organizationId = workplace.LegalOrganizationNumberValue;

            try
            {
                _organizationName = workplace.LegalOrganizationNumberDescription;
            }
            catch { }

            try
            {
                _organizationNumber = workplace.LegalOrganizationNumberValue;
            }
            catch { }

        }

        public string Anchor()
        {
            return _organizationId;
        }


        public override string ToString()
        {
            return _organizationId;
        }

        internal Microsoft.MetadirectoryServices.CSEntryChange GetCSEntryChange()
        {
            CSEntryChange csentry = CSEntryChange.Create();
            csentry.ObjectModificationType = ObjectModificationType.Add;
            csentry.ObjectType = "organization";

            csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Organizationid", _organizationId));

            if (!string.IsNullOrEmpty (_organizationName))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("OrganizationName", _organizationName));

            if (!string.IsNullOrEmpty (_organizationNumber))
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("OrganizationNumber", _organizationNumber));

            return csentry;

        }
    }
}
