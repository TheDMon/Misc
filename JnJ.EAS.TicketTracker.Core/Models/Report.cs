using System;
using System.Collections.Generic;
using System.Text;

namespace JnJ.EAS.TicketTracker.Core.Models
{
    public class Report : Application
    {
        public string Numnber { get; set; }
        public string Priority { get; set; }
        public string State { get; set; }
        public string AssignedTo { get; set; }
        public string ShortDescription { get; set; }
        public string AssignmentGroup { get; set; }
        public string TaskType { get; set; }
        public string Opened { get; set; }
        public string CustTime { get; set; }
        public string Duration { get; set; }
        public string Status { get; set; }
        public string TaskState { get; set; }
    }
}
