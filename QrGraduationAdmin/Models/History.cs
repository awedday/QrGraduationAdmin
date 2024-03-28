using System;
using System.Collections.Generic;

namespace QrGraduationAdmin.Models
{
    public partial class History
    {
        public int IdHistory { get; set; }
        public string DateStartHistory { get; set; } = null!;
        public string? DateFinishHistory { get; set; }
        public int EmployeeId { get; set; }

    }
}
