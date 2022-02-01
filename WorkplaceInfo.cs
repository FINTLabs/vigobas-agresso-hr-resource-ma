using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.MetadirectoryServices;
using Vigo.Bas.Agresso.WebServices.UserAdministrationV200702;

namespace FimSync_Ezma
{
    class WorkplaceInfo : WorkPlace1
    {
        public WorkplaceInfo()
        {

        }
        public WorkplaceInfo(WorkPlace1 workplace)
        {
            this.WorkplaceValue = workplace.WorkplaceValue;
            this.WorkplaceDescription = workplace.WorkplaceDescription;
            this.OrganizationNumberValue = workplace.OrganizationNumberValue;
            this.OrganizationNumberDescription = workplace.OrganizationNumberDescription;
            this.LegalOrganizationNumberValue = workplace.LegalOrganizationNumberValue;
        }
        public string Anchor()
        {
            return base.WorkplaceValue;
        }


        public override string ToString()
        {
            return base.WorkplaceValue;
        }

        internal Microsoft.MetadirectoryServices.CSEntryChange GetCSEntryChange()
        {
            CSEntryChange csentry = CSEntryChange.Create();
            csentry.ObjectModificationType = ObjectModificationType.Add;
            csentry.ObjectType = "workplace";

            var _workplaceId = base.WorkplaceValue;

            var _companyCode = base.CompanyCode;
            var _workplaceDescription = base.WorkplaceDescription;
            var _organizationNumberValue = base.OrganizationNumberValue;
            var _organizationNumberDescription = base.OrganizationNumberDescription;
            var _legalOrganizationNumberValue = base.LegalOrganizationNumberValue;

            csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("workplaceId", _workplaceId));

            if (!string.IsNullOrEmpty(_companyCode))
            {
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("companyCode", _companyCode));
            }
            if (!string.IsNullOrEmpty(_workplaceDescription))
            {
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("workplaceDescription", _workplaceDescription));
            }
            if (!string.IsNullOrEmpty(_organizationNumberValue))
            {
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("organizationNumberValue", _organizationNumberValue));
            }
            if (!string.IsNullOrEmpty(_organizationNumberDescription))
            {
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("organizationNumberDescription", _organizationNumberDescription));
            }
            if (!string.IsNullOrEmpty(_legalOrganizationNumberValue))
            {
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("legalOrganizationNumberValue", _legalOrganizationNumberValue));
            }

            return csentry;
        }
    }
}
