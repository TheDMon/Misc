using System;
using System.Collections.Generic;
using System.Text;

namespace JnJ.EAS.TicketTracker.Core.Models
{
    public class ChangeRequest
    {
        public int ID { get; set; }
        public string Application { get; set; }
        public string Title { get; set; }
        public string SOW { get; set; }
        public string Assignee { get; set; }
        public string Status { get; set; }
        public string ResolutionMethod { get; set; }
        public DateTime? ActualResolutionDate { get; set; }
        public DateTime Created { get; set; }
        public DateTime Modified { get; set; }
        public DateTime? Resolved { get; set; }
    }
}
