using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Email_Read_Test
{
    class OutlookEmails
    {
        //inst variables we want to grab from each email
        public string EmailFrom { get; set; }

        public string EmailSubject { get; set; }

        public string EmailBody { get; set; }

        //function that compiles the needed info
        public static List<OutlookEmails> ReadMailItems()
        {
            Application outlookApplication = null;
            NameSpace outlookNamespace = null;
            MAPIFolder inboxFolder = null;
            Items mailItems = null;
            List<OutlookEmails> listEmailDetails = new List<OutlookEmails>();
            OutlookEmails emailDetails;

            //try block to add from subject and body into a new item
            try
            {
                outlookApplication = new Application();
                outlookNamespace = outlookApplication.GetNamespace("MAPI");
                //Select email inbox to read
                inboxFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                mailItems = inboxFolder.Items;
                foreach (MailItem item in mailItems)
                {
                    emailDetails = new OutlookEmails();
                    emailDetails.EmailFrom = item.SenderEmailAddress;
                    emailDetails.EmailSubject = item.Subject;
                    emailDetails.EmailBody = item.Body;
                    listEmailDetails.Add(emailDetails);
                    ReleaseComObject(item);
                }
            }
            catch (System.Exception ex)

            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                //releasing objects used by program
                ReleaseComObject(mailItems);
                ReleaseComObject(inboxFolder);
                ReleaseComObject(outlookNamespace);
                ReleaseComObject(outlookApplication);
            }
            return listEmailDetails;
        }
        // used to release objects in application
        private static void ReleaseComObject(object obj)
        {
            if (obj != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
        }
    }
}
