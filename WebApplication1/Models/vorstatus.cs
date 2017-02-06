using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebApplication1.Models
{
    public class vorstatus
    {
        public string Region  { get; set; }
        public string Segment { get; set; }
        public DateTime? InDate { get; set; }
        public string Status { get; set; }
        public string Reason { get; set; }
        public int? NegAmount { get; set; }
        public string Customer { get; set; }
        public string Remarks { get; set; }
        public string Dealer { get; set; }
        public string TypeOfRepair { get; set; } 
                
    }
}