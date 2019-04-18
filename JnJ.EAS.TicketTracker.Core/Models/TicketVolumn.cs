using System;
using System.Collections.Generic;
using System.Text;

namespace JnJ.EAS.TicketTracker.Core.Models
{
    public class TicketVolumn : Application
    {
        public int ServiceRequest { get; set; }
        public int ChangeRequest { get; set; }
        public int NonTicketed { get; set; }
        public int Incidents { get; set; }
        public int Problem { get; set; }
        public int Total { get; set; }
        public bool IsActive { get; set; }
    }
}
