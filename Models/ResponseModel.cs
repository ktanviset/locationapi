using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace locationapi.Models
{
    public class ResponseModel
    {
        public bool IsSuccess { get; set; } = true;
        public string Message { get; set; } = "";
    }
}