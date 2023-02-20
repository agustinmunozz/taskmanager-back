using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Azure;
using Entity;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.HttpsPolicy;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.IdentityModel.Tokens;

namespace Test
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }

		// This method gets called by the runtime. Use this method to add services to the container.
		public void ConfigureServices(IServiceCollection services)
		{
			//services.AddDbContext<UserContext>(ServiceLifetime.Transient);
			services.AddControllersWithViews();
			services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme).AddJwtBearer(options =>
			{
				options.TokenValidationParameters = new TokenValidationParameters
				{
					ValidateIssuerSigningKey = true,
					IssuerSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(Configuration["Jwt:Key"])),
					ValidateIssuer = false,
					ValidateAudience = false
				};
			});

			string[] arrOrigins = { "https://www.vuetask.somee.com",
									"http://www.vuetask.somee.com",
									"https://taskmanager.onlinewebshop.net",
									"http://taskmanager.onlinewebshop.net",
									"https://www.vuetask.somee.com/",
									"http://www.vuetask.somee.com/",
									"https://taskmanager.onlinewebshop.net/",
									"http://taskmanager.onlinewebshop.net/",
									"*"};
			var MyAllowSpecificOrigins = "_myAllowSpecificOrigins";
			services.AddCors(options =>
			{
				options.AddPolicy(name: MyAllowSpecificOrigins,
								  builder =>
								  {
									  builder
									 .WithOrigins(arrOrigins)
									 .AllowAnyHeader()
									 .AllowAnyMethod()
									 .SetIsOriginAllowed(origin => true)
									 .SetIsOriginAllowedToAllowWildcardSubdomains()
									 .AllowCredentials();
								  });
			});

			services.AddControllers();
		}

		// This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
		public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
		{
			if (env.IsDevelopment())
			{
				app.UseDeveloperExceptionPage();
			}
			else
			{
				app.UseExceptionHandler("/Home/Error");
				// The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
				app.UseHsts();
			}
			app.UseHttpsRedirection();
			app.UseStaticFiles();

			string[] arrOrigins = { "https://www.vuetask.somee.com",
									"http://www.vuetask.somee.com",
									"https://taskmanager.onlinewebshop.net",
									"http://taskmanager.onlinewebshop.net",
									"https://www.vuetask.somee.com/",
									"http://www.vuetask.somee.com/",
									"https://taskmanager.onlinewebshop.net/",
									"http://taskmanager.onlinewebshop.net/",
									"*"};
			app.UseCors(builder => builder
				.WithOrigins(arrOrigins)
				.AllowAnyHeader()
				.AllowAnyMethod()
				.SetIsOriginAllowed(origin => true)
				.SetIsOriginAllowedToAllowWildcardSubdomains()
				.AllowCredentials()
				.Build());

			app.UseAuthentication();

			app.UseRouting();
			app.UseAuthorization();
			app.UseEndpoints(endpoints =>
			{
				endpoints.MapControllerRoute(
					name: "default",
					pattern: "{controller=Home}/{action=Index}/{id?}");
			});
		}
	}
}
