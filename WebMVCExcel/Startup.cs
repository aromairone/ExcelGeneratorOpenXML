using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(WebMVCExcel.Startup))]
namespace WebMVCExcel
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
