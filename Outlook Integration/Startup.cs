using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(Outlook_Integration.Startup))]
namespace Outlook_Integration
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            
        }
    }
}
