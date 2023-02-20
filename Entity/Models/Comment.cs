using System;
using System.Collections.Generic;

namespace Entity.Models
{
    public partial class Comment
    {
        public int Id { get; set; }
        public DateTime CreatedDate { get; set; }
        public string Comment1 { get; set; }
        public int CommentTypeId { get; set; }
        public DateTime? ReminderDate { get; set; }
        public int UserId { get; set; }
        public int TaskId { get; set; }
        public string State { get; set; }
    }
}
