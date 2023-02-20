using Entity;
using Entity.Context;
using Entity.Models;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using static NuGet.Packaging.PackagingConstants;
//using System.Threading.Tasks;

namespace DAL
{
    public class TaskDAL
    {
        public Task getTaskById(long taskId)
        {
            try
            {
                using (VueTaskContext _context = new VueTaskContext())
                {
                    var query = from ts in _context.Tasks
                                where ts.Id == taskId
                                select ts;

                    var task = query.SingleOrDefault<Task>();
                    return task;
                }
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
        //        using (VueTareasContext _context = new VueTareasContext())
        //        {
        //            var Tasks = new List<Task>();
        //            Tasks = _context.Tasks.ToList();
        //            return Tasks;
        //        }
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
                using (VueTaskContext _context = new VueTaskContext())
                {
                    var t = new Task();
                    if (task.Id == null || task.Id == 0)
                    {
                        t.Description = task.Description;
                        t.TaskTypeId = Convert.ToInt32(task.TaskTypeId);
                        t.TaskStatusId = Convert.ToInt32(task.TaskStatusId);
                        t.CreatedDate = DateTime.Now;
                        t.UserId = Convert.ToInt32(task.UserId);
                        t.RequiredDate = task.RequiredDate;
                        t.DateClose = Convert.ToDateTime(task.DateClose);
                        t.State = "A";
                        //t.NextActionDate = Convert.ToDateTime(task.NextActionDate);
                        _context.Tasks.Add(t);
                    }
                    else
                    {
                        var query = from ts in _context.Tasks
                                    where ts.Id == task.Id
                                    select ts;

                        t = query.SingleOrDefault<Task>();
                        t.Description = task.Description;
                        t.TaskTypeId = Convert.ToInt32(task.TaskTypeId);
                        t.TaskStatusId = Convert.ToInt32(task.TaskStatusId);
                        t.RequiredDate = Convert.ToDateTime(task.RequiredDate);
                        t.DateClose = Convert.ToDateTime(task.DateClose);
                        t.UserId = Convert.ToInt32(task.UserId);
                        _context.Tasks.Update(t);
                    }
                    _context.SaveChanges();
                }
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
                using (VueTaskContext _context = new VueTaskContext())
                {
                    var TaskTypes = new List<TaskType>();
                    TaskTypes = _context.TaskTypes.ToList();
                    return TaskTypes;
                }
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
                using (VueTaskContext _context = new VueTaskContext())
                {
                    var TaskStates = new List<TaskState>();
                    TaskStates = _context.TaskStates.ToList();
                    return TaskStates;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public Object[] searchTasks(TaskModel filters)
        {
            try
            {
                using (VueTaskContext _context = new VueTaskContext())
                {
                    var query = from ts in _context.Tasks
                                join tst in _context.TaskStates on ts.TaskStatusId equals tst.Id
                                join tt in _context.TaskTypes on ts.TaskTypeId equals tt.Id
                                where (filters.Id == null || ts.Id == filters.Id)
                                && ts.State == "A"
                                && (String.IsNullOrEmpty(filters.Description) || ts.Description == filters.Description)
                                && (filters.TaskTypeId == null || filters.TaskTypeId == 0 || ts.TaskTypeId == filters.TaskTypeId)
                                && (filters.TaskStatusId == null || filters.TaskStatusId == 0 || ts.TaskStatusId == filters.TaskStatusId)
                                && (filters.CreatedDate == null || ts.CreatedDate == filters.CreatedDate)
                                && (filters.RequiredDate == null || ts.RequiredDate == filters.RequiredDate)
                                && (filters.DateClose == null || ts.DateClose == filters.DateClose)
                                select new
                                {
                                    ts.DateClose,
                                    ts.RequiredDate,
                                    ts.CreatedDate,
                                    ts.NextActionDate,
                                    ts.Description,
                                    ts.Id,
                                    ts.TaskStatusId,
                                    ts.TaskTypeId,
                                    ts.UserId,
                                    taskTypeDesc = tt.Description,
                                    tst.StatusDesc,
                                    tst.Status
                                };

                    //var task = Getdt

                    var task = query.ToArray();
                    return task;
                }
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
                using (VueTaskContext _context = new VueTaskContext())
                {
                    var task = getTaskById(id);
                    task.State = "I";
                    _context.Tasks.Update(task);
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
