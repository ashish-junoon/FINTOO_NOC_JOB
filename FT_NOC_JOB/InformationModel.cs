using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LOAN_CLOSER_APPLICATION
{
   public class InformationModel
   {
        public string loan_id { get; set; } = string.Empty;
        public string full_name { get; set; } = string.Empty;
        public string lead_id { get; set; } = string.Empty;
        public bool NOC_sent_over_email { get; set; }
     
   }
}
