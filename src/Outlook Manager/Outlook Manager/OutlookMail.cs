using System;
using System.Collections.Generic;

using Microsoft.Office.Interop.Outlook;

namespace Outlook_Manager
{
    public class OutlookMail
    {
        public string From { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        
        public static  List<OutlookMail> ReadMails()
        {
            Application outlookApplication = null;
            NameSpace outlookNameSpace = null;
            MAPIFolder inboxFolder = null;
            Items items = null;
            List<OutlookMail> emails = new List<OutlookMail>();
            OutlookMail mail;

            try
            {
                outlookApplication = new Application();
                outlookNameSpace = outlookApplication.GetNamespace("MAPI");
                inboxFolder = outlookNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                items = inboxFolder.Items;

                foreach (MailItem item in items)
                {
                    emails.Add(new OutlookMail
                    {
                        From = item.SenderEmailAddress,
                        Subject = item.Subject,
                        Body = item.Body
                    });
                    ReleaseComObject(item);
                }

            }
            catch (System.Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
            finally
            {   
                ReleaseComObject(items);
                ReleaseComObject(inboxFolder);
                ReleaseComObject(outlookNameSpace);
                ReleaseComObject(outlookApplication);
            }

            return emails;
        }

        private static void ReleaseComObject(object obj)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            obj = null;
        }
    }
}