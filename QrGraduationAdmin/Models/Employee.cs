using System;
using System.Collections.Generic;

namespace QrGraduationAdmin.Models
{
    public partial class Employee
    {
       

        public int IdEmployee { get; set; }
        public string FirstNameEmployee { get; set; } = null!;
        public string SecondNameEmployee { get; set; } = null!;
        public string? MiddleNameEmployee { get; set; }
        public string MailEmployee { get; set; } = null!;
        public string PasswordEmployee { get; set; } = null!;
        public string PhoneEmployee { get; set; } = null!;

       
    }
}
