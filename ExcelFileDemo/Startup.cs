using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(ExcelFileDemo.Startup))]
namespace ExcelFileDemo
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
