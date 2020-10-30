using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;
using System.Configuration;
using System.Net;

namespace ActualizaDocSharePointCC
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                ArchivoLog.EscribirLog(null, "Inicia Proceso");

                string siteUrl = ConfigurationManager.AppSettings["SiteURL"];
                string bibliotecaDocumentoSP = ConfigurationManager.AppSettings["BibliotecaDocumentosSP"];
                string listaFlujoCronograma = ConfigurationManager.AppSettings["ListaFlujoCronograma"];
                string columnaLoop = ConfigurationManager.AppSettings["ColumnaDocumentoSP"];

                ClientContext clientContext = new ClientContext(siteUrl);

                //clientContext.Credentials = new NetworkCredential(ConfigurationManager.AppSettings["UsuarioSP"],
                //                                                    ConfigurationManager.AppSettings["PasswordSP"]);

                ArchivoLog.EscribirLog(null, "Carga URL de Sharepoint " + siteUrl);

                SP.Web myWeb = clientContext.Web;
                

                List myListFlujoCronograma = myWeb.Lists.GetByTitle(listaFlujoCronograma);

                ListItemCollection listItems = myListFlujoCronograma.GetItems(CamlQuery.CreateAllItemsQuery());

                clientContext.Load(listItems);

                clientContext.ExecuteQuery();

                foreach (var item in listItems)
                {
                    item["Title"] = "Si";

                    item.Update();
                }

                ListItem it = myListFlujoCronograma.AddItem(new ListItemCreationInformation());

                it["Title"] = "Nos empinamos al chetomi";

                it.Update();

                clientContext.ExecuteQuery();

                ListItemCreationInformation listItem = new ListItemCreationInformation();
                
                

                myListFlujoCronograma.AddItem(listItem);

                clientContext.ExecuteQuery();

                List myList_2 = myWeb.Lists.GetByTitle(bibliotecaDocumentoSP);

                ArchivoLog.EscribirLog(null, "Carga Biblioteca de documentos " + bibliotecaDocumentoSP);

                ListItemCollection items = myList_2.GetItems(CamlQuery.CreateAllItemsQuery());

                clientContext.Load(items);
                clientContext.ExecuteQuery();

                ArchivoLog.EscribirLog(null, "Cargar items " + items.Count);

                foreach (var item in items)
                {
                    item[columnaLoop] = "Si";

                    item.Update();
                }

                clientContext.ExecuteQuery();

                ArchivoLog.EscribirLog(null, "Se actualizo correctamente la columna " + columnaLoop);
            }
            catch (Exception ex)
            {
                ArchivoLog.EscribirLog(null, ex.Message);
            }
        }
    }
}
