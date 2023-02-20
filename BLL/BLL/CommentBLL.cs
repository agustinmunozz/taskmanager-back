using Entity;
using Entity.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL
{
	public class CommentBLL
	{
		private static DAL.CommentDAL oComment = new DAL.CommentDAL();

		public Comment getCommentById(long commentId)
		{
			try
			{
				return oComment.getCommentById(commentId);
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
				return oComment.getComments();
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		public object getCommentsByTaskId(int taskId)
		{
			try
			{
				return oComment.getCommentsByTaskId(taskId);
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
				return oComment.getCommentTypes();
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
				return oComment.searchComments(filters);
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
				oComment.saveOrUpdateComment(comment);
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		public void deleteComment(List<int> ids)
		{
			try
			{
				foreach (var id in ids)				
					oComment.deleteComment(id);
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}
	}
}
