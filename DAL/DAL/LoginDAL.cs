using Entity.Context;
using Entity.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAL
{
	public class LoginDAL
	{
		public User login(string userName, string password)
		{
			try
			{
				using (VueTaskContext _context = new VueTaskContext())
				{

					//var query = from us in _context.Users
					//			where us.UserName == userName
					//			&& us.Password == password
					//			select us;
					var query = from us in _context.Users
								where us.UserName == userName
								&& us.Pass == password
								select us;

					var user = query.SingleOrDefault<User>();
					return user;
				}
			}
			catch (Exception ex)
			{
				throw;
			}
		}
	}
}
