using LessonLearntPortalWeb.Repository;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.HttpsPolicy;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace LessonLearntPortalWeb
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }
        public void ConfigureServices(IServiceCollection services)
        {
           // services.AddCors();
            services.AddCors(options =>
            {
                options.AddDefaultPolicy(
                    builder =>
                    {
                        //builder.WithOrigins("https://localhost:44348")
                         builder.WithOrigins("http://localhost:62292" , "https://localhost:44348", "https://teali4metricstest.azurewebsites.net/", "*")
                                            .AllowAnyHeader()
                                            .AllowAnyMethod()
                                           .AllowCredentials();

                        //.SetIsOriginAllowedToAllowWildcardSubdomains()
                        // .AllowAnyHeader()
                        // .AllowCredentials()
                        // .WithMethods("GET", "PUT", "POST", "DELETE", "OPTIONS");
                    });
            });
            services.AddControllers();
            services.AddControllersWithViews();
            services.AddMvc();
            services.AddTransient<ExcelReportRepo>();
            services.AddSwaggerGen();
           
        }

        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {

            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
                app.UseSwagger();
                app.UseSwaggerUI(c =>
                {
                    c.SwaggerEndpoint("/swagger/v1/swagger.json", "MyLessonLearntPortalWebv1");
                });

            }
            else
            {
                app.UseExceptionHandler("/Home/Error");
                app.UseHsts();
            }
          
            //app.UseCors(policy => policy.AllowAnyHeader().AllowAnyMethod());
            app.UseHttpsRedirection();
            app.UseStaticFiles();

            app.UseRouting();

            app.UseAuthorization();

            // Shows UseCors with CorsPolicyBuilder.
            //app.UseCors(builder =>
            //{
            //    builder.AllowAnyOrigin()
            //           .AllowAnyMethod()
            //           .AllowAnyHeader();
            //});
            app.UseCors();
            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllerRoute(
                    name: "default",
                    pattern: "{controller=Home}/{action=Index}/{id?}");
            });
        }
    }
}
