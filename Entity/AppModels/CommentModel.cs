using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Entity
{
	public partial class CommentModel
	{
		public int? Id { get; set; }
		public DateTime? CreatedDate { get; set; }
		public string Comment { get; set; }
		public int? CommentTypeId { get; set; }
		public DateTime? ReminderDate { get; set; }
		public int? UserId { get; set; }
		public int? TaskId { get; set; }
	}

	public partial class CommentsId
	{
		public List<int> Ids { get; set; }
	}
}
