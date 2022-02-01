using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FimSync_Ezma
{
    public class GroupInfo
    {
        public string groupId;
        public string groupName;
        public string groupType;
        public string groupManager;
        public List<string> groupMembers { set; get; } 

        public GroupInfo(string groupId, string groupName, string groupType, string groupManager, List<string> groupMembers)
        {
            this.groupId = groupId;
            this.groupName = groupName;
            this.groupType = groupType;
            this.groupManager = groupManager;
            this.groupMembers = groupMembers;
        }
    }
}
