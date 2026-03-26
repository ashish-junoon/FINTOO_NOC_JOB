using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LOAN_CLOSER_APPLICATION
{
   public class LoanCloserModel
    {
        public string user_id { get; set; } = string.Empty;
        public string lead_id { get; set; } = string.Empty;
        public string name { get; set; } = string.Empty;
        public string current_date { get; set; } = string.Empty;
        public string created_date { get; set; } = string.Empty; 
        public string amount { get; set; } = string.Empty;
        public string tenure { get; set; } = string.Empty;
        public string disbursed_amount { get; set; } = string.Empty;
        public string _tenure { get; set; } = string.Empty;
        public string mobile_number { get; set; } = string.Empty;
        public string email_id { get; set; } = string.Empty;
        public string loan_id { get; set; } = string.Empty;
        public string product_name { get; set; } = string.Empty;
        public string loan_amount { get; set; } = string.Empty;
        public string address { get; set; } = string.Empty;
        public int is_active { get; set; }
        public string doucumentname { get; set; } = string.Empty;
    }
}
