using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UserRoleAudit
{
    public class AuditRole
    {
        public Guid AuditId { get; set; }
        public string User { get; set; }
        public string Role { get; set; }
        public string AsignedBy { get; set; }
        public DateTime AssignedOn { get; set; }

    }
}
