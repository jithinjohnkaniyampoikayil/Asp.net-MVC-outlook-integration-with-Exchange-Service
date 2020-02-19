using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Exchange.WebServices.Data;
using System.Net;

namespace Outlook_Integration.Controllers
{
    public class HomeController : Controller
    {
     
        public ActionResult Index()
        {
            //###########  Move this piece of code to inner layers  ###########

            ExchangeService service = new ExchangeService();
            ServicePointManager.ServerCertificateValidationCallback =
            delegate (object s, X509Certificate certificate,
                     X509Chain chain, SslPolicyErrors sslPolicyErrors)
            { return true; };
            // Setting credentials is unnecessary when you connect from a computer that is
            // logged on to the domain.
            service.Credentials = new WebCredentials("email", "password", "domain");
            // Or use NetworkCredential directly (WebCredentials is a wrapper
            // around NetworkCredential).
            service.Url = new Uri("domain.com/EWS/Exchange.asmx");
            Folder inbox = Folder.Bind(service, WellKnownFolderName.Inbox);

            ItemView view = new ItemView(50);
            FindItemsResults<Item> findResults;
            List<string> email = new List<string>(); 
            do
            {
                findResults = service.FindItems(WellKnownFolderName.Inbox, view);

                foreach (EmailMessage item in findResults.Items)
                {
                    email.Add(item.Sender.Address);
                }

                view.Offset += 50;
            } while (findResults.MoreAvailable);

            //###########  Move this piece of code to inner layers  ###########

            return View();
        }
        
    }
}