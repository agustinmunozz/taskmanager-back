using System;
using System.Collections.Generic;

namespace Entity.Models
{
    public partial class Commentss
    {
        public int Id { get; set; }
        public DateTime CreatedDate { get; set; }
        public string Comment { get; set; }
        public int CommentTypeId { get; set; }
        public DateTime ReminderDate { get; set; }
        public int UserId { get; set; }
        public int TaskId { get; set; }
        public char State { get; set; }
    }
}
