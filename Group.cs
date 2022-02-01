using Microsoft.MetadirectoryServices;
using System.Collections.ObjectModel;
using System.Collections.Generic;

namespace FimSync_Ezma
{
    class Group
    {
        string _groupId, _group_prefix, _group_suffix;
        GroupInfo _group;

        public Group( GroupInfo group, KeyedCollection<string, ConfigParameter> configparameter)
        {
            _group = group;
            _groupId = _group.groupId;
            _group_prefix = configparameter["Gruppeprefix"].Value;
            _group_suffix = configparameter["Gruppesuffix"].Value;
        }

        public string Anchor()
        {
            return _groupId;
        }

        public override string ToString()
        {
            return _groupId;
        }

        // Returns CSEntryChange for use by the FIM Sync Engine directly
        internal Microsoft.MetadirectoryServices.CSEntryChange GetCSEntryChange()
        {
            CSEntryChange csentry = CSEntryChange.Create();
            csentry.ObjectModificationType = ObjectModificationType.Add;
            csentry.ObjectType = "group";

            csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("GroupId", _group_prefix + _groupId + _group_suffix));

            if (!string.IsNullOrEmpty(_group.groupName))
            {
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("GroupName", _group.groupName));
            }
            if (!string.IsNullOrEmpty(_group.groupType))
            {
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("GroupType", _group.groupType));
            }

            if (!string.IsNullOrEmpty(_group.groupManager))
            {
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("GroupManager", _group.groupManager));
            }

            IList<object> members = new List<object>();
            foreach (var member in _group.groupMembers)
            {
                members.Add(member.ToString());
            }

            if (members.Count > 0 )
            {
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("member", members));
            }

            return csentry;
        }
    }
}
