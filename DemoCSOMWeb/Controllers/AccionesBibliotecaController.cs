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
    public class AccionesBibliotecaController : Controller
    {
        // GET: AccionesBiblioteca
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult SubirArchivo()
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                string rutaArchivo = @"C:\Demo\demo.pdf";

                FileCreationInformation newFile = new FileCreationInformation();
                newFile.Overwrite = true;
                newFile.Content = System.IO.File.ReadAllBytes(rutaArchivo);
                newFile.Url = System.IO.Path.GetFileName(rutaArchivo);

                Web web = clientContext.Web;
                List docs = web.Lists.GetByTitle("Documentos");
                Microsoft.SharePoint.Client.File uploadFile = docs.RootFolder.Files.Add(newFile);
                clientContext.Load(uploadFile, w => w.MajorVersion, w => w.MinorVersion);
                clientContext.ExecuteQuery();

                //Solicitar Url Absoluta archivo
                clientContext.Load(uploadFile, f => f.ListItemAllFields["EncodedAbsUrl"]);
                clientContext.ExecuteQuery();
               
                ViewBag.Url = uploadFile.ListItemAllFields["EncodedAbsUrl"];
                return View("Index");
            }
        }
        //Comentar a sukis que tiene que hablar sobre limites, storage metrics, versionamiento

        public ActionResult ObtenerTamano()
        {
            Double totalSize = 0;
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                Web web = clientContext.Web;
                List oList = clientContext.Web.Lists.GetByTitle("Documentos");
                CamlQuery oQuery = new CamlQuery();

                FileCollection allFile = oList.RootFolder.Files;
                clientContext.Load(allFile);

                clientContext.ExecuteQuery();
                foreach (File file in allFile)
                {
                    FileVersionCollection versions = file.Versions;
                    totalSize += file.Length;
                    clientContext.Load(versions);

                    clientContext.ExecuteQuery();
                    foreach (FileVersion fileVersion in versions)
                    {
                        totalSize += fileVersion.Size;
                    }
                }
            }
            ViewBag.TotalSize = Math.Round(totalSize / 1024, 2);
            return View("Index");

        }
    }
}