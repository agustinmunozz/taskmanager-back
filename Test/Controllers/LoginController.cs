using Entity;
using Microsoft.AspNetCore.Mvc;
using System.Xml.Linq;
using System;
using Entity.AppModels;
using System.Security.Claims;
using Microsoft.IdentityModel.Tokens;
using System.IdentityModel.Tokens.Jwt;
using System.Text;
using Entity.Models;
using Microsoft.DiaSymReader;
using Microsoft.Extensions.Configuration;

namespace VueTareas.Controllers
{

	[Produces("application/json")]
	[ResponseCache(NoStore = true)]
	public class LoginController : Controller
	{
		private IConfiguration config;
		private static BLL.LoginBLL oLogin = new BLL.LoginBLL();
		public LoginController(IConfiguration configuration)
		{
			config = configuration;
		}

		[Route("api/LoginController/authenticate/")]
		[HttpPost]
		public IActionResult authenticate([FromBody] LoginModel model)
		{
			try
			{
				
				var user = oLogin.login(model.UserLogin, model.Password);
				if (user == null)
					return BadRequest(new { message = "Invalid credentials" });
				string jwtToken = GenerateToken(user);
				return Ok(new {token = jwtToken, userId = user.Id});
			}
			catch (Exception ex)
			{
				return BadRequest(ex.InnerException != null ? ex.InnerException.Message : ex.Message);
			}
		}

		private string GenerateToken(User user)
		{
			var claims = new[]
			{
				new Claim(ClaimTypes.Name, user.Name),
				new Claim(ClaimTypes.Email, user.SurName)
			};

			var key = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(config.GetSection("JWT:Key").Value));
			var creds = new SigningCredentials(key, SecurityAlgorithms.HmacSha512Signature);
			var securityToken = new JwtSecurityToken(
									claims: claims,
									expires: DateTime.Now.AddMinutes(120),
									signingCredentials: creds);
			string token = new JwtSecurityTokenHandler().WriteToken(securityToken);
			return token;
		}
	}
}
