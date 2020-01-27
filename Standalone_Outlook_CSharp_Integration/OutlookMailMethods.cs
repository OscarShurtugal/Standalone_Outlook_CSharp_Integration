using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;

namespace Standalone_Outlook_CSharp_Integration
{
    public class OutlookMailMethods
    {
        public Boolean SendEmailCumpleanios(string rutaImagen, string nombreImagen, string mailSubject, string mailDirection)
        {
            try
            {
                var oApp = new Application();

                NameSpace ns = oApp.GetNamespace("MAPI");
                var f = ns.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

                System.Threading.Thread.Sleep(1000);

                var mailItem = (MailItem)oApp.CreateItem(OlItemType.olMailItem);

                // Adjuntar como  archivo
                // mailItem.Attachments.Add("C:\\RPA\\RPA.png", Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue);

                Microsoft.Office.Interop.Outlook.Attachment attachment = mailItem.Attachments.Add(
     @""+rutaImagen+"\\"+nombreImagen
    , OlAttachmentType.olEmbeddeditem
    , null
    , "Some image display name"
    );

                string imageCid = nombreImagen+"@123";

                attachment.PropertyAccessor.SetProperty(
                  "http://schemas.microsoft.com/mapi/proptag/0x3712001E"
                 , imageCid
                 );

                string mailContent;
                //aviso de privacidad
                mailContent=String.Format("<body><center><img src=\"cid:{0}\"><p style=\"font-family:calibri; font-size:11px;\"></p></center><br/><br/>", imageCid);

                /*attachment =*/

                Microsoft.Office.Interop.Outlook.Attachment attachment2 = mailItem.Attachments.Add(
     @"C:\RPA\RPA.png"
    , OlAttachmentType.olEmbeddeditem
    , null
    , "Some image display name"
    );

                //imagen por default
                string imageRPA = "RPA.png@456";

                attachment2.PropertyAccessor.SetProperty(
                  "http://schemas.microsoft.com/mapi/proptag/0x3712001E"
                 , imageRPA
                 );

                mailContent=String.Format("{0}<img src=\"cid:{1}\"></p><\body>", mailContent, imageRPA);

                mailContent+=String.Format("<p><h6 style=\"margin:0; font-family:Arial; font-weight:normal;\">Por favor no responder este correo</h6><br>");



                mailItem.Subject=mailSubject;
                mailItem.HTMLBody=mailContent;
                mailItem.To=mailDirection;

                mailItem.Send();

            }
            catch (System.Exception ex)
            {
                return false;
            }
            finally
            {
            }
            return true;
        }

        public Boolean SendEmailLeyendaRPA(string mailDirection, string mailSubject, string mailContent, string mailGreeting, string mailFarewell)
        {
            try
            {
                var oApp = new Application();

                NameSpace ns = oApp.GetNamespace("MAPI");
                var f = ns.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

                System.Threading.Thread.Sleep(1000);



                var mailItem = (MailItem)oApp.CreateItem(OlItemType.olMailItem);


                Microsoft.Office.Interop.Outlook.Attachment attachment = mailItem.Attachments.Add(
     @"C:\RPA\RPA.png"
    , OlAttachmentType.olEmbeddeditem
    , null
    , "Some image display name"
    );

                string imageCid = "RPA.png@123";

                attachment.PropertyAccessor.SetProperty(
                  "http://schemas.microsoft.com/mapi/proptag/0x3712001E"
                 , imageCid
                 );

                mailGreeting=String.Format("<p style=\"font-family:Arial;\">{0}</p>", mailGreeting);

                mailContent=String.Format("<p align=\"justify\" style=\"font-family:Arial; \">{0}</p>", mailContent);

                mailFarewell=String.Format("<p style=\"font-family:Arial; \">{0}</p><p><img src=\"cid:{1}\"><h6 style=\"margin:0; font-family:Arial; font-weight:normal;\">Por favor no responda este correo</h6><br></p>"
                              , mailFarewell, imageCid);

                mailContent=String.Format("<body>{0}{1}{2}<\body>", mailGreeting, mailContent, mailFarewell);

                mailItem.Subject=mailSubject;
                mailItem.HTMLBody=mailContent;
                mailItem.To=mailDirection;

                mailItem.Send();

            }
            catch (System.Exception ex)
            {
                return false;
            }
            finally
            {
            }
            return true;
        }

        public Boolean SendEmailWithAttachment(string rutaArchivo, string mailSubject, string mailDirection, string mailContent, string mailGreeting, string mailFarewell, string rutaImagen)
        {
            try
            {

                rutaArchivo.Replace("\"", "\\");

                if (rutaImagen!="")
                {
                    rutaImagen.Replace("\"", "\\");

                }
                else
                {
                    //Default footer image
                    rutaImagen="C:\\RPA\\RPA.png";
                }



                var oApp = new Application();

                NameSpace ns = oApp.GetNamespace("MAPI");
                var f = ns.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

                System.Threading.Thread.Sleep(1000);

                var mailItem = (MailItem)oApp.CreateItem(OlItemType.olMailItem);

                //Aqui añadimos el archivo que queremos enviar

                //Aqui añadimos el archivo que queremos enviar
                //Hago un split para que podamos enviar varios archivos separados por ;
                foreach (var archivo in rutaArchivo.Split(';'))
                {
                    mailItem.Attachments.Add(archivo);
                }


                //El saludo
                mailGreeting=String.Format("<p style=\"font-family:sans-serif; font-weight:300; font-size:14px;\">{0}</p>", mailGreeting);

                // El cuerpo del correo
                mailContent=String.Format("<p align=\"justify\" style=\"font-family:sans-serif; font-weight:300; font-size:14px;\">{0}</p>", mailContent);

                //La despedida Y LA FIRMA PREDEFINIDA
                mailFarewell=String.Format("<p style=\"font-family:sans-serif; font-weight:300; font-size:14px;\">{0}</p><p><h5 style=\"color:#0568AE; font-family:sans-serif; font-weight:normal; margin:0; \">Robot</h5>"+
                    "<h6 style=\"margin:0; font-family:sans-serif; font-weight:bold;color:RGB(5,104,174); \">Robotic Process Automation</h6>"
                              , mailFarewell);


                mailContent=String.Format("<body>{0}{1}{2}<\body>", mailGreeting, mailContent, mailFarewell);

                //Este es el aviso de "No responder al correo"
                mailContent+=String.Format("<br><p><h6 style=\"margin:0; font-family:Arial; font-weight:normal;\">Por favor no responder este correo</h6><br>");

                //Aqui ponemos la firma de RPA, en caso de no poder mostrarla se manda "Imagen No Disponible"
                mailContent+=String.Format("<img src=\""+rutaImagen+"\" alt=\"Imagen no disponible\">");

                //El aviso de privacidad de ATT
                mailContent+=String.Format("<br><body><center><p style=\"font-family:sans-serif; font-weight:300; font-size:10px;\"></p></center><br/><br/>");


                mailItem.Subject=mailSubject;
                mailItem.HTMLBody=mailContent;
                mailItem.To=mailDirection;

                mailItem.Send();

            }
            catch (System.Exception ex)
            {
                //Console.WriteLine(ex.Message);
                return false;
            }
            finally
            {
            }
            return true;
        }

        public Boolean SendSimpleMailNoAttachment(string mailSubject, string mailDirection, string mailContent, string mailGreeting, string mailFarewell, string rutaImagen)
        {
            try
            {

                //rutaArchivo.Replace("\"", "\\");

                if (rutaImagen!="")
                {
                    rutaImagen.Replace("\"", "\\");

                }
                else
                {
                    //Default footer image
                    rutaImagen="C:\\RPA\\RPA.png";
                }



                var oApp = new Application();

                NameSpace ns = oApp.GetNamespace("MAPI");
                var f = ns.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

                System.Threading.Thread.Sleep(1000);

                var mailItem = (MailItem)oApp.CreateItem(OlItemType.olMailItem);

                //Aqui añadimos el archivo que queremos enviar

                //mailItem.Attachments.Add(rutaArchivo);

                //El saludo
                mailGreeting=String.Format("<p style=\"font-family:sans-serif; font-weight:300; font-size:14px;\">{0}</p>", mailGreeting);

                // El cuerpo del correo
                mailContent=String.Format("<p align=\"justify\" style=\"font-family:sans-serif; font-weight:300; font-size:14px;\">{0}</p>", mailContent);


                //La despedida Y LA FIRMA PREDEFINIDA
                mailFarewell=String.Format("<p style=\"font-family:sans-serif; font-weight:300; font-size:14px;\">{0}</p><p><h5 style=\"color:#0568AE; font-family:sans-serif; font-weight:normal; margin:0; \">Robot</h5>"+
                    "<h6 style=\"margin:0; font-family:sans-serif; font-weight:bold;color:RGB(5,104,174); \">Robotic Process Automation</h6>"
                              , mailFarewell);


                mailContent=String.Format("<body>{0}{1}{2}<\body>", mailGreeting, mailContent, mailFarewell);

                //Este es el aviso de "No responder al correo"
                mailContent+=String.Format("<br><p><h6 style=\"margin:0; font-family:Arial; font-weight:normal;\">Por favor no responder este correo</h6><br>");

                //Aqui ponemos la firma de RPA, en caso de no poder mostrarla se manda "Imagen No Disponible"
                mailContent+=String.Format("<img src=\""+rutaImagen+"\" alt=\"Imagen no disponible\">");

                //Insertar acuerdo de confidencialidad en caso de aplicar
                mailContent+=String.Format("<br><body><center><p style=\"font-family:sans-serif; font-weight:300; font-size:10px;\"></p></center><br/><br/>");


                mailItem.Subject=mailSubject;
                mailItem.HTMLBody=mailContent;
                mailItem.To=mailDirection;


                mailItem.Send();

            }
            catch (System.Exception ex)
            {
                //Console.WriteLine(ex.Message);
                return false;
            }
            finally
            {
            }
            return true;
        }


    }
}
