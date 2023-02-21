using Entity;
using Entity.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System;

namespace API.Controllers
{
	[Produces("application/json")]
	[ResponseCache(NoStore = true)]
	public class CommentController : Controller
	{
		private static BLL.CommentBLL oComment = new BLL.CommentBLL();

		//[Authorize]
		[Authorize]
		[Route("api/CommentController/getCommentById/{id}")]
		[HttpGet]
		public IActionResult getCommentById(int id)
		{
			try
			{
				//var Task = oTask.getTaskById(id);
				//var Vendedor = oUsuarios.getUsuarioById(model.toID);

				//BLL.EnvioMail.Enviar(Vendedor.Email, model.mensaje, model.asunto, string.Empty);

				return Ok();
			}
			catch (Exception ex)
			{
				return BadRequest(ex.InnerException != null ? ex.InnerException.Message : ex.Message);
			}
		}

		//[Authorize]
		[Authorize]
		[Route("api/CommentController/getComments/")]
		[HttpGet]
		public IActionResult getTasks()
		{
			try
			{
				//var Task = oTask.getTasks();
				return Ok();
			}
			catch (Exception ex)
			{
				return BadRequest(ex.InnerException != null ? ex.InnerException.Message : ex.Message);
			}
		}

		[Authorize]
		[Route("api/CommentController/getCommentsByTaskId/{id}")]
		[HttpGet]
		public IActionResult getCommentsByTaskId(int id)
		{
			try
			{
				//var Task = oTask.getTaskById(id);
				//var Vendedor = oUsuarios.getUsuarioById(model.toID);

				//BLL.EnvioMail.Enviar(Vendedor.Email, model.mensaje, model.asunto, string.Empty);
				var comments = oComment.getCommentsByTaskId(id);
				return Ok(comments);
			}
			catch (Exception ex)
			{
				return BadRequest(ex.InnerException != null ? ex.InnerException.Message : ex.Message);
			}
		}

		[Authorize]
		[Route("api/CommentController/getCommentTypes/")]
		[HttpGet]
		public IActionResult getCommentTypes()
		{
			try
			{
				var CommentTypes = oComment.getCommentTypes();
				return Ok(CommentTypes);
			}
			catch (Exception ex)
			{
				return BadRequest(ex.InnerException != null ? ex.InnerException.Message : ex.Message);
			}
		}

		[Authorize]
		[Route("api/CommentController/searchComments/")]
		[HttpPost]
		public IActionResult searchComments([FromBody] CommentModel filters)
		{
			try
			{
				var Comments = oComment.searchComments(filters);
				return Ok(Comments);
			}
			catch (Exception ex)
			{
				return BadRequest(ex.InnerException != null ? ex.InnerException.Message : ex.Message);
			}
		}

		[Authorize]
		[Route("api/CommentController/saveOrUpdateComment")]
		[HttpPost]
		public IActionResult saveOrUpdateComment([FromBody] CommentModel comment)
		{
			try
			{
				oComment.saveOrUpdateComment(comment);
				return Ok();
			}
			catch (Exception ex)
			{
				return BadRequest(ex.InnerException != null ? ex.InnerException.Message : ex.Message);
			}
		}

		[Authorize]
		[Route("api/CommentController/deleteComment/")]
		[HttpPost]
		public IActionResult deleteComment([FromBody] CommentsId model)
		{
			try
			{
				oComment.deleteComment(model.Ids);
				return Ok();
			}
			catch (Exception ex)
			{
				return BadRequest(ex.InnerException != null ? ex.InnerException.Message : ex.Message);
			}
		}
	}
}
