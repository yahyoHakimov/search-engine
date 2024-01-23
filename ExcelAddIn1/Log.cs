using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using User;



namespace Log
{
    public class LogData
    {
        public int LogId { get; set; }
        public int UserId {  get; set; }
        public string ActivityType {  get; set; }
        public string SearchTerm {  get; set; }
        public DateTime TimeStap {  get; set; }

        public Employee User { get; set; }
    }
}
