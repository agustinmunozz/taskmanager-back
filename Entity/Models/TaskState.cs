using System;
using System.Collections.Generic;

namespace Entity.Models
{
    public partial class TaskState
    {
        public int Id { get; set; }
        public string Status { get; set; }
        public string StatusDesc { get; set; }
    }
}
