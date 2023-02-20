using Entity;
using Entity.Models;
using Newtonsoft.Json.Linq;
using NuGet.Protocol;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
//using VueTareas;
//using System.Threading.Tasks;
//using VueTareas.Models;
//using vue;


namespace BLL
{
    public class TaskBLL
    {
        private static DAL.TaskDAL oTask = new DAL.TaskDAL();
		private static DAL.CommentDAL oComment = new DAL.CommentDAL();

		public Task getTaskById(long TaskId)
        {
            try
            {
                return oTask.getTaskById(TaskId);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //public List<Task> getTasks()
        //{
        //    try
        //    {
        //        return oTask.getTasks();
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }
        //}

        public void saveOrUpdateTask(TaskModel task)
        {
            try
            {
                oTask.saveOrUpdateTask(task);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

		public List<TaskType> getTaskTypes()
		{
			try
			{
				return oTask.getTaskTypes();
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		public List<TaskState> getTaskStates()
		{
			try
			{
				return oTask.getTaskStates();
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		public object searchTasks(TaskModel filters)
		{
			try
			{
				return oTask.searchTasks(filters);
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		public void deleteTask(int id)
		{
			try
			{
				var comments = oComment.getCommentsByTaskId(id);
				if(comments.Count > 0)
					foreach (var comment in comments)
						oComment.deleteComment((int)comment.GetType().GetProperty("Id").GetValue(comment, null));
				
				oTask.deleteTask(id);
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		public string export(TaskModel filters)
		{
			try
			{
				Object[] array = oTask.searchTasks(filters);
				string fecha = " Fecha Desde:  Fecha Hasta: ";// + fechaHasta.ToString("dd/MM/yyyy");
				string titulo = "Listado de Pedidos De NC-ND-Reintegros - " + fecha;
				string[] columnas = { "CLIENTEID", "APELLIDO_Y_NOMBRE", "DOMICILIO", "ZONATECNICA", "LOCALIDAD", "ZONACOMERCIAL", "FORMADEPAGO", "TIPO", "NRO_PEDIDO", "FECHA_PEDIDO", "MOTIVO_PEDIDO", "IMPORTE", "OBSERVACION", "USUARIO_PEDIDO", "ESTADO", "MOTIVO_CIERRE", "FECHA_CIERRE", "USUARIO_CIERRE", "FECHAVENTA", "FECHAINSTALACION", "FECHAULTIMOESTADO", "COMENTARIOS_DEL_CIERRE" };
				string nombre = System.DateTime.Now.ToShortDateString().Replace("/", "_") + System.DateTime.Now.Hour.ToString() + System.DateTime.Now.Minute.ToString() + System.DateTime.Now.Second.ToString() + "ReportePedidosNC_ND_Reintegros";
				//Object[] array = {new { algo }};
				//new Dictionary<>
				var props = new Dictionary<string, string>
				{
					["DateClose"] = "DateClose",
					["CreatedDate"] = "CreatedDate",
					["NextActionDate"] = "NextActionDate",
					["Description"] = "Description",
					["StatusDesc"] = "StatusDesc",
					["Status"] = "Status",				
				};
				var dt = Exportador.ToDataTable(array.ToList(), props);
				return Exportador.GenerarExcel(dt, columnas, titulo, nombre, "", null);

				//return oTask.export(filters);
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}
	}
}
