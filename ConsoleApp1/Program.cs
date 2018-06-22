using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;
namespace ConsoleApp1
{
    class Program
    {


        static void Main(string[] args)
        {
            Test test = new Test();
            //test.SendEmailtoContacts();
            //SendEmailtoContacts();
        }       

        public class Test
        {
            public Test()
            {
               // SendEmailtoContacts();
            }

            Outlook.NameSpace outlookNameSpace;
            Outlook.MAPIFolder inbox;
            Outlook.Items items;

            private void ThisAddIn_Startup(object sender, System.EventArgs e)
            {
                Outlook.Application OutlookApplication1 = new Outlook.Application();

                outlookNameSpace = OutlookApplication1.GetNamespace("MAPI");

                //Messaging Application Programming Interface (MAPI) is a messaging architecture and a Component Object Model based 
                //API for Microsoft Windows. MAPI allows client programs to become (e-mail) messaging-enabled, -aware, or -based by calling MAPI subsystem 
                //routines that interface with certain messaging servers.

                inbox = outlookNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
                items = inbox.Items;
                items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(items_ItemAdd);

            }

            void items_ItemAdd(object Item)
            {
                string filter = "USED CARS";
                Outlook.MailItem mail = (Outlook.MailItem)Item;
                if (Item != null)
                {
                    if (mail.MessageClass == "IPM.Note" &&
                               mail.Subject.ToUpper().Contains(filter.ToUpper()))//==="(IPM)InterPersonal Message"====
                    {
                        mail.Move(outlookNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderJunk));
                    }
                }
            }



            public void SendEmailtoContacts()
            {
               

                string subjectEmail = "Meeting has been rescheduled.";
                string bodyEmail = "Meeting is one hour later.";
                string Emailaddress = "anbarasu.intellibot@gmail.com";

                if (Emailaddress.Contains("gmail.com"))
                {
                    this.CreateEmailItem(subjectEmail, Emailaddress, bodyEmail);
                }

                //Outlook.MAPIFolder sentContacts = (Outlook.MAPIFolder)OutlookApplication.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);

                //foreach (Outlook.ContactItem contact in sentContacts.Items)
                //{
                //    if (contact.Email1Address.Contains("gmail.com"))
                //    {
                //        this.CreateEmailItem(subjectEmail, contact.Email1Address, bodyEmail);
                //    }
                //}
            }

            private void CreateEmailItem(string subjectEmail, string toEmail, string bodyEmail)
            {
                Outlook.Application OutlookApplication = new Outlook.Application();

                Outlook.MailItem eMail = (Outlook.MailItem)OutlookApplication.CreateItem(Outlook.OlItemType.olMailItem);//CreateItem(Outlook.OlItemType.olMailItem);
                eMail.Subject = subjectEmail;
                eMail.To = toEmail;
                eMail.Body = bodyEmail;
                eMail.Importance = Outlook.OlImportance.olImportanceLow;
                ((Outlook._MailItem)eMail).Send();
            }
        }    

    }
}
