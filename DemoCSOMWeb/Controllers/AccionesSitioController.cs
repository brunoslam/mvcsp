using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;

namespace DemoCSOMWeb.Controllers
{
    public class AccionesSitioController : Controller
    {
        // GET: AccionesSitio
        public ActionResult Index()
        {
            return View();
        }


        public ActionResult ObtenerUsuarios()
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                clientContext.Load(clientContext.Web.SiteGroups);
                clientContext.ExecuteQuery();

                GroupCollection oSiteCollectionGroups = clientContext.Web.SiteGroups;
                clientContext.Load(oSiteCollectionGroups,
                groups => groups.Include(
                group => group.Users));

                foreach (Group oGroup in oSiteCollectionGroups)
                {
                    System.Diagnostics.Debug.WriteLine(oGroup.Title);
                    try
                    {
                        clientContext.Load(oGroup.Users);
                        clientContext.ExecuteQuery();

                        foreach (User oUser in oGroup.Users)
                        {
                            System.Diagnostics.Debug.WriteLine(oUser.Title);
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(ex.Message);
                    }
                }
            }
            return View("Index");
        }

        public ActionResult ObtenerSitios()
        {
            string mainpath = "https://falabella.sharepoint.com";
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                WebCollection oWebsite = clientContext.Web.GetSubwebsForCurrentUser(new SubwebQuery());
                clientContext.Load(oWebsite, n => n.Include(o => o.Title));
                clientContext.ExecuteQuery();
                foreach (Web orWebsite in oWebsite)
                {
                    string newpath = mainpath + orWebsite.Title;
                    System.Diagnostics.Debug.WriteLine(newpath + "\n" + orWebsite.Title);
                }
            }
            return View("Index");
        }


    }
}