using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DemoCSOMWeb.Models;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;

using System.Web.Mvc;
namespace DemoCSOMWeb.Controllers
{
    public class SharePointController
    {


        public void InsertarElemento(Persona persona, HttpContextBase HttpContext)
        {
            string siteUrl = "http://MyServer/sites/MySiteCollection";

            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);


            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                //var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

                //ClientContext clientContext = new ClientContext(siteUrl);
                SP.List oList = clientContext.Web.Lists.GetByTitle("Noticias");

                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem oListItem = oList.AddItem(itemCreateInfo);
                oListItem["Title"] = "My New Item!";

                oListItem.Update();

                clientContext.ExecuteQuery();
            }
        }
        
        public void ObtenerElementos(HttpContextBase HttpContext)
        {
            //string siteUrl = "https://latinshare.sharepoint.com/sites/dev/";

            //ClientContext clientContext = new ClientContext(siteUrl);
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);


            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                //var creds = new SharePointOnlineCredentials("user@tenant.onmicrosoft.com", password); // Requires SecureString() for password
                //context.Credentials = creds;
                SP.List oList = clientContext.Web.Lists.GetByTitle("Noticias");

                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = "<View><Query><OrderBy><FieldRef Name='Title' Ascending='TRUE'></FieldRef></OrderBy><Where><IsNotNull><FieldRef Name='ID' /></IsNotNull></Where></Query></View>";
                ListItemCollection collListItem = oList.GetItems(camlQuery);

                clientContext.Load(collListItem);

                clientContext.ExecuteQuery();

                foreach (ListItem oListItem in collListItem)
                {
                    Console.WriteLine("ID: {0} \nTitle: {1} ", oListItem.Id, oListItem["Title"]);
                }
            }
        }

        public void EliminarElemento()
        {
            
        }

        public void MigrarDatos(List<Persona> ListPersona, HttpContextBase HttpContext)
        {
            LimpiarTabla(HttpContext);
            foreach (Persona persona in ListPersona)
            {
                string siteUrl = "http://MyServer/sites/MySiteCollection";

                var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);


                using (var clientContext = spContext.CreateUserClientContextForSPHost())
                {
                    //var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

                    //ClientContext clientContext = new ClientContext(siteUrl);
                    SP.List oList = clientContext.Web.Lists.GetByTitle("Noticias");

                    ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                    ListItem oListItem = oList.AddItem(itemCreateInfo);
                    oListItem["Title"] = persona.Id;

                    oListItem.Update();

                    clientContext.ExecuteQuery();
                }
            }


        }

        public void LimpiarTabla(HttpContextBase HttpContext)
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);


            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {

                SP.List oList = clientContext.Web.Lists.GetByTitle("Noticias");
                ListItemCollection listItems = oList.GetItems(CamlQuery.CreateAllItemsQuery());
                clientContext.Load(listItems,
                                    eachItem => eachItem.Include(
                                    item => item["ID"]));
                clientContext.ExecuteQuery();

                var totalListItems = listItems.Count;
                if (totalListItems > 0)
                {
                    for (var counter = totalListItems - 1; counter > -1; counter--)
                    {
                        listItems[counter].DeleteObject();
                        clientContext.ExecuteQuery();
                        Console.WriteLine("Row: " + counter + " Item Deleted");
                    }
                }

            }

        }


    }
}