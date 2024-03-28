using System;
using System.Collections.Generic;

namespace QrGraduationAdmin.Models
{
    public partial class Information
    {
        public int IdInformation { get; set; }
        public string? LocationInformation { get; set; }
        public double? DistanceInformation { get; set; }
        public string AndroidInformation { get; set; } = null!;
        public int EmployeeId { get; set; }

    }
}
