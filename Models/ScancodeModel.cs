using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace locationapi.Models
{
    public class ScancodeModel
    {
        public string CodeId { get; set; }
        public DateTime? ImportTime { get; set; }
        public DateTime? ExportTime { get; set; }
    }
}
