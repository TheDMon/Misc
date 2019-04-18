using System;
using System.Collections.Generic;
using System.Text;

namespace JnJ.EAS.TicketTracker.Core.Models
{
    public class IncidentReport : Application
    {
        public string Numnber { get; set; }
        public string Priority { get; set; }
        public string ShortDescription { get; set; }
        public string State { get; set; }
        public string AssignedTo { get; set; }
        public string Opened { get; set; }
        public string Resolved { get; set; }
        public string Closed { get; set; }
        public string Category { get; set; }
        public string Status { get; set; }
        public string AssignmentGroup { get; set; }
        public string ResolutionCategory { get; set; }
        public string ResolutionCode { get; set; }
        public string KCS { get; set; }
    }
}
