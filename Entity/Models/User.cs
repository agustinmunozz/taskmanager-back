using System;
using System.Collections.Generic;

namespace Entity.Models
{
    public partial class User
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string SurName { get; set; }
        public string Email { get; set; }
        public string UserName { get; set; }
        public string Pass { get; set; }
    }
}
