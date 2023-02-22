using Entity;
using Entity.Context;
using Entity.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using static NuGet.Packaging.PackagingConstants;

namespace DAL
{
	public class CommentDAL
	{

		public Comment getCommentById(long commentId)
		{
			try
			{
				using (VueTaskContext _context = new VueTaskContext())
				{

					var query = from cs in _context.Comments
								where cs.Id == commentId
					select cs;

					var task = query.SingleOrDefault<Comment>();
					return task;
				}
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		public List<Comment> getComments()
		{
			try
			{
				using (VueTaskContext _context = new VueTaskContext())
				{
					var Comments = new List<Comment>();
					Comments = _context.Comments.ToList();
					return Comments;
				}
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		public List<dynamic> getCommentsByTaskId(int taskId)
		{
			try
			{
				using (VueTaskContext _context = new VueTaskContext())
				{
					var query = from cs in _context.Comments
					join ct in _context.CommentTypes on cs.CommentTypeId equals ct.Id
					where cs.TaskId == taskId
					&& cs.State == "A"
					select new { cs.ReminderDate, cs.CreatedDate, cs.UserId, cs.Comment1, cs.CommentTypeId,
								 cs.TaskId, commentTypeDesc = ct.Description, cs.Id };

					var comments = query.ToList<dynamic>();
					return comments;
				}
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		public List<CommentType> getCommentTypes()
		{
			try
			{
				using (VueTaskContext _context = new VueTaskContext())
				{
					var CommentTypes = new List<CommentType>();
					CommentTypes = _context.CommentTypes.ToList();
					return CommentTypes;
				}
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		public object searchComments(CommentModel filters)
		{
			try
			{
				using (VueTaskContext _context = new VueTaskContext())
				{
					var query = from cs in _context.Comments
								join ct in _context.CommentTypes on cs.CommentTypeId equals ct.Id
								where (filters.Id == null || cs.Id == filters.Id)
								&& cs.State == "A"
								&& (filters.TaskId == null || cs.TaskId == filters.TaskId)
								&& (String.IsNullOrEmpty(filters.Comment) || cs.Comment1.Contains(filters.Comment))
								&& (filters.CreatedDate == null || cs.CreatedDate.Date == filters.CreatedDate.Value.Date)
								&& (filters.CommentTypeId == null || filters.CommentTypeId == 0 || cs.CommentTypeId == filters.CommentTypeId)
								&& (filters.ReminderDate == null || cs.ReminderDate.Value.Date == filters.ReminderDate.Value.Date)
								select new { cs.ReminderDate, cs.CreatedDate, cs.UserId, cs.Comment1, cs.CommentTypeId, 
											 cs.TaskId, commentTypeDesc = ct.Description, cs.Id };

					var comments = query.ToList();
					return comments;
				}
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		public void saveOrUpdateComment(CommentModel comment)
		{
			try
			{
				using (VueTaskContext _context = new VueTaskContext())
				{
					var c = new Comment();
					if (comment.Id == null || comment.Id == 0)
					{
						c.Comment1 = comment.Comment;
						c.CreatedDate = DateTime.Now;
						c.ReminderDate = Convert.ToDateTime(comment.ReminderDate);
						c.CommentTypeId = Convert.ToInt32(comment.CommentTypeId);
						c.TaskId = Convert.ToInt32(comment.TaskId);
						c.State = "A";
						c.UserId = Convert.ToInt32(comment.UserId);
						_context.Comments.Add(c);
					}
					else
					{
						var query = from cs in _context.Comments
									where cs.Id == comment.Id
									select cs;

						c = query.SingleOrDefault<Comment>();
						c.Comment1 = comment.Comment;
						c.CommentTypeId = Convert.ToInt16(comment.CommentTypeId);
						c.ReminderDate = Convert.ToDateTime(comment.ReminderDate);
						c.UserId = Convert.ToInt32(comment.UserId);
						_context.Comments.Update(c);
					}

					_context.SaveChanges();
				}
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		public void deleteComment(int id)
		{
			try
			{
				using (VueTaskContext _context = new VueTaskContext())
				{
					var comment = getCommentById(id);
					comment.State = "I";
					_context.Comments.Update(comment);
					_context.SaveChanges();
				}
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}
	}
}
