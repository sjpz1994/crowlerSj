using System;

namespace crowlerSj.Models
{
    public class Log
    {
        public long Id { get; set; }
        public string Message { get; set; }
        public string Level { get; set; }
        public DateTime Timestamp { get; set; } = DateTime.Now;
        public long? CrowlId { get; set; } // nullable
        public virtual Crowl Crowl { get; set; }
    }
}
