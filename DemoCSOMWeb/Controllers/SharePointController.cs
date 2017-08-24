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
            //string siteUrl = "https://latinshare.sharepoint.com/sites/dev/";
            //ClientContext clientContext = new ClientContext(siteUrl);
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
            //var creds = new SharePointOnlineCredentials("user@tenant.onmicrosoft.com", password); // Requires SecureString() for password
            //context.Credentials = creds;

            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                SP.List oList = clientContext.Web.Lists.GetByTitle("Persona");

                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem oListItem = oList.AddItem(itemCreateInfo);
                oListItem["Title"] = persona.Id;
                oListItem["Nombre"] = persona.Nombre;
                oListItem["Comuna"] = persona.Comuna;
                oListItem["Direccion"] = persona.Direccion;
                oListItem["FechaNacimiento"] = persona.FechaNacimiento;
                oListItem["EsHumano"] = persona.EsHumano;

                oListItem.Update();

                clientContext.ExecuteQuery();
            }
        }
        
        public void ObtenerElementos(HttpContextBase HttpContext)
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                SP.List oList = clientContext.Web.Lists.GetByTitle("Persona");

                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = "<View><Query><OrderBy><FieldRef Name='Title' Ascending='TRUE'></FieldRef></OrderBy><Where><IsNotNull><FieldRef Name='ID' /></IsNotNull></Where></Query></View>";
                ListItemCollection collListItem = oList.GetItems(camlQuery);

                clientContext.Load(collListItem);
                clientContext.ExecuteQuery();

                foreach (ListItem oListItem in collListItem)
                {
                    Console.WriteLine("ID BBDD: {0}, Nombre: {1} ", oListItem["Title"], oListItem["Nombre"]);
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
                var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
                using (var clientContext = spContext.CreateUserClientContextForSPHost())
                {
                    //var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

                    //ClientContext clientContext = new ClientContext(siteUrl);
                    SP.List oList = clientContext.Web.Lists.GetByTitle("Persona");

                    ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                    ListItem oListItem = oList.AddItem(itemCreateInfo);
                    oListItem["Title"] = persona.Id;
                    oListItem["Nombre"] = persona.Nombre;
                    oListItem["Comuna"] = persona.Comuna;
                    oListItem["Direccion"] = persona.Direccion;
                    oListItem["FechaNacimiento"] = persona.FechaNacimiento;
                    oListItem["EsHumano"] = persona.EsHumano;

                    oListItem.Update();

                    clientContext.ExecuteQuery();

                    if (persona.Id % 2 == 0)
                    {
                        oListItem.BreakRoleInheritance(false, false);

                        User spUser = clientContext.Web.CurrentUser;
                        clientContext.Load(spUser, user => user.LoginName);
                        clientContext.ExecuteQuery();

                        //spUser.LoginName
                        User oUser = clientContext.Web.SiteUsers.GetByLoginName("i:0#.f|membership|sleiva@latinshare.com");
                        RoleDefinitionBindingCollection collRoleDefinitionBinding = new RoleDefinitionBindingCollection(clientContext);
                        collRoleDefinitionBinding.Add(clientContext.Web.RoleDefinitions.GetByType(RoleType.Reader));
                        oListItem.RoleAssignments.Add(oUser, collRoleDefinitionBinding);
                        clientContext.ExecuteQuery();
                    }
                    

                }
            }


        }

        public void LimpiarTabla(HttpContextBase HttpContext)
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {

                SP.List oList = clientContext.Web.Lists.GetByTitle("Persona");
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
                    }
                }

            }

        }

        


    }
}