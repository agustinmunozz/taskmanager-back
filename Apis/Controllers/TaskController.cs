using Entity;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace API.Controllers
{

    [Produces("application/json")]
    [ResponseCache(NoStore = true)]
    public class TaskController : Controller
    {
        private static BLL.TaskBLL oTask = new BLL.TaskBLL();

        //[Authorize]
        [Route("api/TaskController/getTaskById/{id}")]
        [HttpGet]
        public IActionResult getTaskById(int id)
        {
            try
            {
                var Task = oTask.getTaskById(id);
				return Ok(Task);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.InnerException != null ? ex.InnerException.Message : ex.Message);
            }
        }

		//[Authorize]
		//[Route("api/TaskController/getTasks/")]
		//[HttpGet]
		//public IActionResult getTasks()
		//{
		//    try
		//    {
		//        var Tasks = oTask.getTasks();
		//        return Ok(Tasks);
		//    }
		//    catch (Exception ex)
		//    {
		//        return BadRequest(ex.InnerException != null ? ex.InnerException.Message : ex.Message);
		//    }
		//}

		//[AllowAnonymous]
		[Authorize]
		[Route("api/TaskController/saveOrUpdateTask")]
        [HttpPost]
        public IActionResult saveOrUpdateTask([FromBody]TaskModel task)
        {
            try
            {
                oTask.saveOrUpdateTask(task);
                return Ok();
            }
            catch (Exception ex)
            {
                return BadRequest(ex.InnerException != null ? ex.InnerException.Message : ex.Message);
            }
        }

		[Authorize]
		[Route("api/TaskController/getTaskTypes/")]
		[HttpGet]
		public IActionResult getTaskTypes()
		{
			try
			{
				var Tasks = oTask.getTaskTypes();
				return Ok(Tasks);
			}
			catch (Exception ex)
			{
				return BadRequest(ex.InnerException != null ? ex.InnerException.Message : ex.Message);
			}
		}

		[Authorize]
		[Route("api/TaskController/getTaskStates/")]
		[HttpGet]
		public IActionResult getTaskStates()
		{
			try
			{
				var Tasks = oTask.getTaskStates();
				return Ok(Tasks);
			}
			catch (Exception ex)
			{
				return BadRequest(ex.InnerException != null ? ex.InnerException.Message : ex.Message);
			}
		}

		[Authorize]
		[Route("api/TaskController/searchTasks/")]
		[HttpPost]
		public IActionResult searchTasks([FromBody] TaskModel filters)
		{
			try
			{
				var Tasks = oTask.searchTasks(filters);
				return Ok(Tasks);
			}
			catch (Exception ex)
			{
				return BadRequest(ex.InnerException != null ? ex.InnerException.Message : ex.Message);
			}
		}

		[Authorize]
		[Route("api/TaskController/deleteTask/{id}")]
		[HttpGet]
		public IActionResult deleteTask(int id)
		{
			try
			{
				oTask.deleteTask(id);
				return Ok();
			}
			catch (Exception ex)
			{
				return BadRequest(ex.InnerException != null ? ex.InnerException.Message : ex.Message);
			}
		}

		[Authorize]
		[Route("api/TaskController/export")]
		[HttpPost]
		public IActionResult export([FromBody] TaskModel task)
		{
			try
			{
				oTask.export(task);
				return Ok();
			}
			catch (Exception ex)
			{
				return BadRequest(ex.InnerException != null ? ex.InnerException.Message : ex.Message);
			}
		}
	}
}