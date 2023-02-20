using Entity.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace BLL
{
	public class LoginBLL
	{
		private static DAL.LoginDAL oLogin = new DAL.LoginDAL();

		public User login(string user, string password)
		{
			try
			{
				return oLogin.login(user, password);
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}
	}
}
