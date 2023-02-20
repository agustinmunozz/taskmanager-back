using System;
using System.Collections.Generic;
using System.Text;

namespace Entity
{
	public partial class TaskModel
	{
		public int? Id { get; set; }
		public string Description { get; set; }
		public int? TaskTypeId { get; set; }
		public int? TaskStatusId { get; set; }
		public DateTime? CreatedDate { get; set; }
		public DateTime? RequiredDate { get; set; }
		public int? UserId { get; set; }
		public DateTime? DateClose { get; set; }
		public DateTime? NextActionDate { get; set; }
	}

	public class TaskIds
	{
		public List<int> Ids { get; set; }
	}
}
